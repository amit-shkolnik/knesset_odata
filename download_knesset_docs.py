import sys
import os
import requests
import pandas as pd
import time
from pywintypes import com_error
import tabulate

from logger_configurer import configure_logger

log=configure_logger('default')

if os.name !='nt':
    raise Exception("App needs win32com.client package, hence need to run on Windows")
# Install with 'pip install pywin32'
import win32com.client
# Knesset ODATA site
main_hypelink="http://knesset.gov.il/Odata/ParliamentInfo.svc/"
# Datasources on Knesset ODATA to be scraped.
plenum_session_ref="KNS_DocumentPlenumSession"
committees_sessions="KNS_DocumentCommitteeSession"
bills="KNS_DocumentBill"

odata_download_format="format=json"
ms_words_suffix=["doc", "DOC", "docx", "DOCX"] #, "rtf"]

word_application=None 
# Documents corrupted previously downloaded
# and can't be open via 'Word'   
corrupted_docs_log="corrupted_docs_log.csv"
corrupted_docs_list=[]

false_words=['\n', '\r']

def download_dataset(source_name, skip_token:str):
    """
    Download documents from 1 source (Plenum, committees, etc),
    Paging API (100 documents per page)
    """
    try:
        # Skip token used for paging between Knesset ODATA API pages.    
        log.info(f"Downloading source {source_name}")
        corrupted_docs_list=list(pd.read_csv(corrupted_docs_log)["doc_name"])
        rounds=1
        
        while True:
            log.info(f"*** ROUND {rounds} ***")        
            rounds+=1
            skip_token, errors_list=download_one_page_docs(source_name, skip_token, corrupted_docs_list)    
            if len(errors_list)>0:
                log_erros(errors_list)            
            if not skip_token:
                break
    except Exception as err:
        log.exception(err)

def download_one_page_docs(source_name:str, skip_token:str, corrupted_docs_list)->str:    
    """
    Main method to download and extract texts from
    Knesset ODATA API,
    Each API page contains -by default- 100 documents' links 
    """
    already_downloaded=get_already_downloaded(source_name)
    # Call ODATA API
    response, num_of_docs=get_docs_list(source_name,skip_token)    
    documents_log_list=[]
    errors_list=[]
    # Per document:
    for idx, entry in  enumerate(response.json()["value"]):
        try:
            if not handle_or_skip_docs(entry,already_downloaded, num_of_docs, idx, 
                                       corrupted_docs_list):
                continue
            log.info(f"{idx}/{num_of_docs} Downloading {entry['FilePath']}")
            file_name=download_doc(source_name, entry)
            if file_name is not None:
                extract_text_from_doc(source_name, file_name)
            documents_log_list.append(entry)
        except Exception as err:
            log.exception(err)
            errors_list.append({"doc":entry, "error":err})
            continue
        continue
    if len(documents_log_list)>0:
        log_documents(source_name, documents_log_list)
    # something like "KNS_DocumentPlenumSession?$skiptoken=128985L" to move
    # to next page on ODATA
    if "odata.nextLink" in response.json():
        return response.json()["odata.nextLink"], errors_list
    return None, errors_list


def handle_or_skip_docs(entry, already_downloaded, num_of_docs, idx, corrupted_docs_list):
    # Skip file if previously donwloaded 
    file_path=entry['FilePath']
    if file_path.split("/")[-1] in already_downloaded:
        log.info(f"{idx}/{num_of_docs} {entry['FilePath']} already downloaded")
        return False
    if file_path.split(".").pop() not in ms_words_suffix:
        log.info(f"Skipping non MS Word doc {entry['FilePath']}")
        return False
    if file_path.split("/")[-1] in corrupted_docs_list:
        log.info("Skipping corrupted documnet")
        return False
    return True

def get_docs_list(source_name:str, skip_token:str):
    url=f"{main_hypelink}{source_name}?${odata_download_format}"
    # Paging through the ODATA
    if skip_token:
        token=skip_token.split("?$")[1]
        url=f"{url}&${token}"
    log.info(f"*** Download main ODATA {url} ***")
    # Call ODATA API
    while True:
        response=requests.get(url=url)    
        if 'value' not in response.json():
            log.info(f"No 'valeue' key on response.json")
            log.info(response.json())
            time.sleep(10)
            continue
        break

    num_of_docs=len(response.json()['value'])
    log.info(f"*** {num_of_docs} documents to download ***")
    return response, num_of_docs

def get_already_downloaded(source):
    already_downloaded= os.listdir(f"{source}_extracted_texts")
    for idx, doc in enumerate(already_downloaded):
        already_downloaded[idx]=doc.replace(".txt", "")
    return already_downloaded


def log_documents(source_name:str, documents_log_list:list):
    """
    Save a log of all downloaded documents 
    """
    new_docs_df=pd.DataFrame(documents_log_list)

    _file=f"{source_name}_docs_download_log.txt"
    if os.path.exists(_file):
        old_docs_df=pd.read_csv(_file)
        concated_df=pd.concat([new_docs_df, old_docs_df])
        concated_df.to_csv(_file, index=False)
    else:
        new_docs_df.to_csv(_file, index=False)

    return

def download_doc(source:str, entry:dict):
    """
    Download and save documents in original format
    """
    url=entry["FilePath"]
    response=requests.get(url)    
    if response.status_code == 200:
        # Save the document to a local file
        file_name=url.split("/")[len(url.split("/"))-1]
        with open(os.path.join(f"{source}_docs", file_name), "wb") as file:
            file.write(response.content)
        log.info("Document downloaded successfully.")
        return file_name
    else:
        log.info(f"Failed to download document. Status code: {response.status_code}")
        return None
    
def extract_text_from_doc(source_name:str, file_name:str):
    """
    Extract text from downloaded document management method.
    Handle per document format: .doc, .docx, .pdf, etc.
    """
    
    if file_name.lower().split(".")[len(file_name.split("."))-1] in ms_words_suffix:
        extract_text_from_ms_word(source_name, file_name)
        log.info("Document's text successfuly extracted")

    else:
        log.info("This file type is not handled")
    return

def  extract_text_from_ms_word(source_name:str, file_name:str):
    # Old .doc format, non OXML files.
    if file_name.lower().split(".").pop() in ms_words_suffix:
        return read_old_msword_doc(source_name, file_name)    
   

def read_old_msword_doc(source_name:str, file_path):
    """
    Text extraction from Old word format, non OXML
    """
    output_text=""
    if word_application==None:
        init_word_app()    
    
    doc=open_word_doc(source_name, file_path)
    if doc==None:
        log.info("Failed to open doc")
        return
    
    output_text=doc.Range().Text
    ''' Old, slower extraction method '''
    # for paragraph in doc.Paragraphs:
    #     full_text.append(paragraph.Range.Text.strip())    
    # output_text="\n".join([ t for t in full_text if len(t.strip())>0])

    # Extract text from Text Box, which appears on
    # old Knesset documents, originaly extracted from TIFF / PDF images
    # using OCR.
    text_boxes_texts = []
    for shape in doc.Shapes:
        # Check if there is textboxs
        if shape.Type == 17:  # 17 represents a textbox shape
            # Extract text from the textbox
            text = shape.TextFrame.TextRange.Text
            text_boxes_texts.append(text)
    filtered_text=[w for w in text_boxes_texts if w.strip()]
    if len(filtered_text)>0:
        output_text=output_text+ " " +'\n'.join([ t for t in text_boxes_texts if len(t.strip())>0])        
    doc.Close(False)    
    log.info(f"\t{len(output_text.split())} words on doc")
    #word_application.Quit()    
    if len(output_text.strip())==0:
        log.info("No text found in documet")
        return
    output_path=os.path.join(os.getcwd(), f"{source_name}_extracted_texts", f"{file_path}.txt")
    with open(output_path, "w", encoding="utf-8") as _fout:
        _fout.write(output_text)
    
    return

def open_word_doc(source_name, file_path):
    try:
        #file_path=f"19_cs_bg_325715.doc"
        document_path = os.path.join(os.getcwd(), f"{source_name}_docs", file_path)
        doc = word_application.Documents.Open(document_path, ReadOnly=True)
        return doc
    except Exception as err:
        log.exception(err)
        add_doc_to_corrupted_docs_list(file_path)
        init_word_app()
    return None
    

def init_word_app():
    # Main object to open MS WORD docs with
    global word_application
    word_application = win32com.client.gencache.EnsureDispatch('Word.Application')   
    # Avoid actualy open the docs- all work should be done in the background
    word_application.Visible=False

    # This cal init word with late binding, the code above int early binding
    # word_application = win32com.client.Dispatch('Word.Application')

def add_doc_to_corrupted_docs_list(file_path):
    _df=pd.read_csv(corrupted_docs_log)
    new_df=pd.DataFrame([{"doc_name":file_path}])
    _df2=pd.concat([_df, new_df])
    _df2.to_csv(corrupted_docs_log, index=False)
    return

def log_erros(errors_list):
    _file="errors_list.csv"
    new_df=pd.DataFrame(errors_list)
    old_df=pd.read_csv(_file)
    merged_df=pd.concat([new_df, old_df])
    merged_df.to_csv(_file, index=False)
    return
    

def mkdir_per_source(source:str):
    # Folder to store original docs downloaded from Knesset ODATA
    if not os.path.exists(f"{source}_docs"):
        os.makedirs(f"{source}_docs")
    if not os.path.exists(f"{source}_extracted_texts"):
        os.makedirs(f"{source}_extracted_texts")

def summerize_content()():
    _rslts=[]
    for idx, source in enumerate(datasets_sources):
        source_text=[]
        source_total_size=0
        number_of_words=0
        _dir=f"{source}_extracted_texts"
        files=os.listdir(_dir)
        log.info(f"Number of files in {source}: {len(files)}")
        for idx2, _file in enumerate(files):
            file_path=os.path.join(f"{source}_extracted_texts", _file)
            with open(file_path, "r", encoding="utf-8") as file_open:
                file_text=file_open.read()
                number_of_words=number_of_words+len(file_text.split())
            source_total_size+=os.path.getsize(file_path)/1024**2
            continue

        _dict={
            "source":source,
            "number of files": len(files),
            "volumne (KB)": round(source_total_size, 0),
            "number of words": number_of_words
        }
  
        _rslts.append(_dict)
  
        continue
    rslts_df=pd.DataFrame(_rslts)
    log.info(f"\n{rslts_df.to_markdown()}")
# Main call
# Datasource to download from
datasets_sources=[committees_sessions, plenum_session_ref,   bills]
# Skip tokens per source-not to re-iterate all pages
skip_tokens=[ "?$skiptoken=321264L", "?$skiptoken=4099505L",  "?$skiptoken=497312L"]

for idx, source in enumerate(datasets_sources):
    _query=f"https://knesset.gov.il/Odata/ParliamentInfo.svc/{source}/$count"
    _response=requests.get(_query)
    log.info(f"** TOTAL {_response.text} documents on {source} **")

# Loop between Knesset sources (Plenum, committees, etc)
for idx, source in enumerate(datasets_sources):
    mkdir_per_source(source)
    download_dataset(source, skip_token=skip_tokens[idx])
    continue

summerize_content()()

word_application.Quit()

