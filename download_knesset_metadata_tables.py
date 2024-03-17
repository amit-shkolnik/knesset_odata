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

class DownloadMetadataTables():
    """
    Downloading metadata per Knesset's plenum sessions, 
    committees sessions, etc.
    """

    def __init__(self) -> None:
        self.log=logging.getLogger('default')


    def run(self):
        try:
            ######################################################################
            # Main call                                                          #
            ######################################################################    
            # Skip tokens per source-not to re-iterate all pages, like:
            # skip_tokens=[ "?$skiptoken=321264L", "?$skiptoken=4099505L",  "?$skiptoken=497312L"]
            skip_tokens=[None, None, None]

            # Check number of files on each source:
            for idx, source in enumerate(config.meta_data_tables):
                _query=f"https://knesset.gov.il/Odata/ParliamentInfo.svc/{source}/$count"
                _response=requests.get(_query)
                self.log.info(f"** TOTAL {_response.text} documents on {source} **")

            # Loop between Knesset sources (Plenum, committees, etc)
            for idx, source in enumerate(config.meta_data_tables):
                self.mkdir_per_source(source)
                self.download_dataset(source, skip_token=skip_tokens[idx])
                continue

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
            
            # Skip token used for paging between Knesset ODATA API pages.    
            self.log.info(f"Downloading source {source_name}")
            rounds=1
            
            while True:
                self.log.info(f"*** ROUND {rounds} ***")        
                rounds+=1
                
                skip_token=self.download_one_json(
                    source_name, skip_token)                    
                if not skip_token:
                    break
        except Exception as err:
            log.exception(err)

    def download_one_json(self, source_name:str, skip_token:str)->str:    
        """
        Main method to download and extract texts from
        Knesset ODATA API,
        Each API page contains -by default- 100 documents' links 
        """
        # Call ODATA API
        response, num_of_docs, url=self.get_metadata_json(source_name,skip_token)    
        self.save_response_json(response, source_name, url)
        # something like "KNS_DocumentPlenumSession?$skiptoken=128985L" to move
        # to next page on ODATA
        if "odata.nextLink" in response.json():
            return response.json()["odata.nextLink"]
        return None


    def get_metadata_json(self, source_name:str, skip_token:str):
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

        num_of_obj=len(response.json()['value'])
        self.log.info(f"*** {num_of_obj} documents to download ***")
        return response, num_of_obj, url

    def mkdir_per_source(self, source:str):
        # Folder to store original docs downloaded from Knesset ODATA
        if not os.path.exists(f"{source}_metadata_jsons"):
            os.makedirs(f"{source}_metadata_jsons")

    
    def save_response_json(self, response:requests.Response, source_name:str, url:str):

        json_obj=json.dumps( response.json())
        if "odata.nextLink" in response.json():
            _name=response.json()["odata.nextLink"].replace("?$skiptoken=", "_")
        else:
            _name="last_json"
        _file=os.path.join(f"{source_name}_metadata_jsons", f"{_name}.json")
        with open(_file, "w") as output_file:
            output_file.write(json_obj)
        return

if __name__=='__main__':
    log=configure_logger('default')
    log.info("Program start")

    dmt=DownloadMetadataTables()
    dmt.run()

    log.info("Program ends")

