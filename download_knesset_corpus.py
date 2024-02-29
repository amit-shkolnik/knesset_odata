'''
Script download Knesset (https://main.knesset.gov.il/EN/Pages/default.aspx)
ODATA from 3 sources:
* Plenum
* Legislation processes (Bills)
* Committees sessions
Author: Amit Shkolnik, amit.shkolnik@gmail.com, 2024
'''

import sys
import os

import config
from config import *
from logger_configurer import configure_logger

if os.name !='nt':
    raise Exception("App needs win32com.client package, hence need to run on Windows")
# Install with 'pip install pywin32'
import win32com.client
word_application=None 

class DownloadKnesetCorpus():

    def __init__(self) -> None:
        self.log=logging.getLogger('default')


    def run(self):
        try:
            ######################################################################
            # Main call                                                          #
            ######################################################################    
            # Skip tokens per source-not to re-iterate all pages, like:
            # skip_tokens=[ "?$skiptoken=321264L", "?$skiptoken=4099505L",  "?$skiptoken=497312L"]
            skip_tokens=[ None, None, None]

            # Check number of files on each source:
            for idx, source in enumerate(config.datasets_sources):
                _query=f"https://knesset.gov.il/Odata/ParliamentInfo.svc/{source}/$count"
                _response=requests.get(_query)
                self.log.info(f"** TOTAL {_response.text} documents on {source} **")

            # Loop between Knesset sources (Plenum, committees, etc)
            for idx, source in enumerate(config.datasets_sources):
                self.mkdir_per_source(source)
                self.download_dataset(source, skip_token=skip_tokens[idx])
                continue

            word_application.Quit()
            return

        except Exception as err:
            self.log.exception(err)
            self.log.info("End run")
        return

    def download_dataset(self, source_name, skip_token:str):
        """
        Download documents from 1 source (Plenum, committees, etc),
        Paging API (100 documents per page).
        Parameters:
        * source: Knesset source to download from.
        * skip_token: string, if not None, script skip all 
            pages to the skip_token page.
        """
        try:
            corrupted_docs_list=[]
            # Skip token used for paging between Knesset ODATA API pages.    
            self.log.info(f"Downloading source {source_name}")
            corrupted_docs_list=list(pd.read_csv(config.corrupted_docs_log)["doc_name"])
            rounds=1
            
            while True:
                self.log.info(f"*** ROUND {rounds} ***")        
                rounds+=1
                skip_token, errors_list=self.download_one_page_docs(
                    source_name, skip_token, corrupted_docs_list)    
                if len(errors_list)>0:
                    self.log_erros(errors_list)            
                if not skip_token:
                    break
        except Exception as err:
            log.exception(err)

    def download_one_page_docs(self, source_name:str, skip_token:str, corrupted_docs_list)->str:    
        """
        Main method to download and extract texts from
        Knesset ODATA API,
        Each API page contains -by default- 100 documents' links 
        """
        already_downloaded=self.get_already_downloaded(source_name)
        # Call ODATA API
        response, num_of_docs, url=self.get_docs_list(source_name,skip_token)    
        self.save_response_json(response, source_name, url)
        documents_log_list=[]
        errors_list=[]
        #already_downloaded_cnt, not_msword_cnt, corrupted_cnt
        skip_cntr=[0,0,0]
        # Per document:
        for idx, entry in  enumerate(response.json()["value"]):
            try:
                if not self.handle_or_skip_docs(entry,already_downloaded, 
                    num_of_docs, idx, corrupted_docs_list, skip_cntr):
                    continue
                self.log.info(f"{idx}/{num_of_docs} Downloading {entry['FilePath']}")
                file_name=self.download_doc(source_name, entry)
                if file_name is not None:
                    self.extract_text_from_doc(source_name, file_name)
                documents_log_list.append(entry)
            except Exception as err:
                log.exception(err)
                errors_list.append({"doc":entry, "error":err})
                continue
            continue
        self.log.info("{} downloaded {} already downloaded, {} not WORD format, {} corrupted ".format(
            len(documents_log_list), skip_cntr[0], skip_cntr[1], skip_cntr[2]))
        if len(documents_log_list)>0:
            self.log_documents(source_name, documents_log_list)
        # something like "KNS_DocumentPlenumSession?$skiptoken=128985L" to move
        # to next page on ODATA
        if "odata.nextLink" in response.json():
            return response.json()["odata.nextLink"], errors_list
        return None, errors_list


    def handle_or_skip_docs(self, entry, already_downloaded, num_of_docs, idx, corrupted_docs_list,
            skip_cntr:list):
        """
        Decide to skip file if previously donwloaded, format is not handled
        or document is corrupted
        """
        file_path=entry['FilePath']
        if file_path.split("/")[-1] in already_downloaded:
            self.log.debug(f"{idx}/{num_of_docs} {entry['FilePath']} already downloaded")
            skip_cntr[0]=skip_cntr[0]+1
            return False
        if file_path.split(".").pop() not in config.ms_words_suffix:
            self.log.debug(f"Skipping non MS Word doc {entry['FilePath']}")
            skip_cntr[1]=skip_cntr[1]+1
            return False
        if file_path.split("/")[-1] in corrupted_docs_list:
            self.log.debug("Skipping corrupted documnet")
            skip_cntr[2]=skip_cntr[2]+1
            return False
        return True

    def get_docs_list(self, source_name:str, skip_token:str):
        """
        HTTP request to get 1 page from Knesset ODATA.
        """
        url=f"{config.main_hypelink}{source_name}?${config.odata_download_format}"
        # Paging through the ODATA
        if skip_token:
            token=skip_token.split("?$")[1]
            url=f"{url}&${token}"
        self.log.info(f"*** Download main ODATA {url} ***")
        # Call ODATA API
        while True:
            response=requests.get(url=url)    
            if 'value' not in response.json():
                self.log.info(f"No 'valeue' key on response.json")
                self.log.info(response.json())
                time.sleep(10)
                continue
            break

        num_of_docs=len(response.json()['value'])
        self.log.info(f"*** {num_of_docs} documents to download ***")
        return response, num_of_docs, url

    def get_already_downloaded(self, source):
        """
        Avoiding re-download documents by watching a list of 
        previously dowloaded docs.
        """
        already_downloaded= os.listdir(f"{source}_extracted_texts")
        for idx, doc in enumerate(already_downloaded):
            already_downloaded[idx]=doc.replace(".txt", "")
        return already_downloaded


    def log_documents(self, source_name:str, documents_log_list:list):
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

    def download_doc(self, source:str, entry:dict):
        """
        Download document in original format,
        save it to local folder
        """
        url=entry["FilePath"]
        response=requests.get(url)    
        if response.status_code == 200:
            # Save the document to a local file
            file_name=url.split("/")[len(url.split("/"))-1]
            with open(os.path.join(f"{source}_docs", file_name), "wb") as file:
                file.write(response.content)
            self.log.info("Document downloaded successfully.")
            return file_name
        else:
            self.log.info(f"Failed to download document. Status code: {response.status_code}")
            return None
        
    def extract_text_from_doc(self, source_name:str, file_name:str):
        """
        Extract text from downloaded document management method.
        Handle per document format: .doc, .docx, .pdf, etc.
        """    
        if file_name.lower().split(".")[len(file_name.split("."))-1] in config.ms_words_suffix:
            self.extract_text_from_ms_word(source_name, file_name)
            self.log.info("Document's text successfuly extracted")

        else:
            self.log.info("This file type is not handled")
        return

    def  extract_text_from_ms_word(self, source_name:str, file_name:str):
        # Old .doc format, non OXML files.
        if file_name.lower().split(".").pop() in config.ms_words_suffix:
            return self.read_msword_with_win32com(source_name, file_name)    
    

    def read_msword_with_win32com(self, source_name:str, file_path):
        """
        Text extraction from MS WORD
        """
        output_text=""
        if word_application==None:
            self.init_word_app()    
        
        doc=self.open_word_doc(source_name, file_path)
        if doc==None:
            self.log.info("Failed to open doc")
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
        self.log.info(f"\t{len(output_text.split())} words on doc")
        #word_application.Quit()    
        if len(output_text.strip())==0:
            self.log.info("No text found in documet")
            return
        output_path=os.path.join(os.getcwd(), f"{source_name}_extracted_texts", f"{file_path}.txt")
        with open(output_path, "w", encoding="utf-8") as _fout:
            _fout.write(output_text)
        
        return

    def open_word_doc(self, source_name, file_path):
        try:
            #file_path=f"19_cs_bg_325715.doc"
            document_path = os.path.join(os.getcwd(), f"{source_name}_docs", file_path)
            doc = word_application.Documents.Open(document_path, ReadOnly=True)
            return doc
        except Exception as err:
            log.exception(err)
            self.add_doc_to_corrupted_docs_list(file_path)
            self.init_word_app()
        return None
        

    def init_word_app(self):
        # Main object to open MS WORD docs with
        global word_application
        word_application = win32com.client.gencache.EnsureDispatch('Word.Application')   
        # Avoid actualy open the docs- all work should be done in the background
        word_application.Visible=False

        # This cal init word with late binding, the code above int early binding
        # word_application = win32com.client.Dispatch('Word.Application')

    def add_doc_to_corrupted_docs_list(self, file_path):
        _df=pd.read_csv(config.corrupted_docs_log)
        new_df=pd.DataFrame([{"doc_name":file_path}])
        _df2=pd.concat([_df, new_df])
        _df2.to_csv(config.corrupted_docs_log, index=False)
        return

    def log_erros(self, errors_list):
        _file="errors_list.csv"
        new_df=pd.DataFrame(errors_list)
        old_df=pd.read_csv(_file)
        merged_df=pd.concat([new_df, old_df])
        merged_df.to_csv(_file, index=False)
        return
        

    def mkdir_per_source(self, source:str):
        # Folder to store original docs downloaded from Knesset ODATA
        if not os.path.exists(f"{source}_docs"):
            os.makedirs(f"{source}_docs")
        if not os.path.exists(f"{source}_extracted_texts"):
            os.makedirs(f"{source}_extracted_texts")

    
    def save_response_json(self, response:requests.Response, source_name:str, url:str):

        json_obj=json.dumps( response.json())
        _name=response.json()["odata.nextLink"].replace("?$skiptoken=", "_")
        _file=os.path.join(config.jsons_dir, f"{_name}.json")
        with open(_file, "w") as output_file:
            output_file.write(json_obj)
        return

if __name__=='__main__':
    log=configure_logger('default')
    log.info("Program start")

    dkc=DownloadKnesetCorpus()
    dkc.run()

    log.info("Program ends")

