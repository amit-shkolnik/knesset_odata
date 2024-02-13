import sys
import os
import requests
import pandas as pd


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
ms_words_suffix=["doc", "DOC", "docx", "DOCX"]
word_application = win32com.client.Dispatch('Word.Application')
    

def download_dataset(source_name, skip_token:str):
    """
    Download documents from 1 source (Plenum, committees, etc),
    Paging API (100 documents per page)
    """
    # Skip token used for paging between Knesset ODATA API pages.    
    log.info(f"Downloading source {source_name}")
    rounds=1
    
    while True:
        log.info(f"*** ROUND {rounds} ***")        
        rounds+=1
        skip_token, errors_list=download_one_page_docs(source_name, skip_token)    
        if len(errors_list)>0:
            log_erros(errors_list)            
        if not skip_token:
            break


def download_one_page_docs(source_name:str, skip_token:str)->str:    
    """
    Main method to download and extract texts from
    Knesset ODATA API,
    Each API page contains -by default- 100 documents' links 
    """
    already_downloaded=get_already_downloaded(source_name)
    url=f"{main_hypelink}{source_name}?${odata_download_format}"
    # Paging through the ODATA
    if skip_token:
        token=skip_token.split("?$")[1]
        url=f"{url}&${token}"
    log.info(f"*** Download main ODATA {url} ***")
    # Call ODATA API
    response=requests.get(url=url)    
    documents_log_list=[]
    errors_list=[]
    num_of_docs=len(response.json()['value'])
    log.info(f"*** {num_of_docs} documents to download ***")
    # Per document:
    for idx, entry in  enumerate(response.json()["value"]):
        try:
            # Skip file if previously donwloaded 
            if entry["FilePath"].split("/")[-1] in already_downloaded:
                log.info(f"{idx}/{num_of_docs} {entry['FilePath']} already downloaded")
                continue
            if entry['FilePath'].split(".").pop() not in ms_words_suffix:
                log.info(f"Skipping non MS Word doc {entry['FilePath']}")
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
    
    #word_application = win32com.client.Dispatch('Word.Application')
    doc = word_application.Documents.Open(os.path.join(os.getcwd(), f"{source_name}_docs", file_path))
    full_text = []
    for paragraph in doc.Paragraphs:
        full_text.append(paragraph.Range.Text.strip())
    # Extract text from Text Box, which appears on
    # old Knesset documents.
    for shape in doc.Shapes:
        # Check if the shape is a textbox
        if shape.Type == 17:  # 17 represents a textbox shape
            # Extract text from the textbox
            text = shape.TextFrame.TextRange.Text
            full_text.append(text)
    doc.Close()
    #word_application.Quit()
    # Saving to .txt
    output_text="\n".join([ t for t in full_text if len(t.strip())>0])
    if len(output_text)==0:
        log.info("No text found in documet")
        return
    output_path=os.path.join(os.getcwd(), f"{source_name}_extracted_texts", f"{file_path}.txt")
    with open(output_path, "w", encoding="utf-8") as _fout:
        _fout.write(output_text)
    
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

# Main call
# Datasource to download from
datasets_sources=[bills, committees_sessions, plenum_session_ref]
# Skip tokens per source-not to re-iterate all pages
skip_tokens=["?$skiptoken=464328L", None, None]

# Loop between Knesset sources (Plenum, committees, etc)
for idx, source in enumerate(datasets_sources):
    mkdir_per_source(source)
    download_dataset(source, skip_token=skip_tokens[idx])
    continue
