'''
Author: Amit Shkolnik, amit.shkolnik@gmail.com, 2024
'''
import sys
import os

import config
from config import *
from logger_configurer import configure_logger


class CountFilesNContent():

    def __init__(self) -> None:
        self.log=logging.getLogger('default')

    def run(self):
        self.count_files_per_knesset()
        self.count_words_file_by_source()

    def count_files_per_knesset(self):
        jsons_dfs, urls_list=self.json_to_dfs()
        for idx, json_df in enumerate(jsons_dfs):
            jsons_dfs[idx]=self.prepare_json_df(json_df, urls_list[idx])

        full_df=pd.concat(jsons_dfs)
        self.log.info(f"{len(full_df)} records on all sources")

        source_counts_list=[]
        for source in config.datasets_sources:
            source_counts=self.count_source_per_knesset(full_df, source)
            source_counts_list.append(source_counts)

        for source in source_counts_list:
            self.log.info(f"\n{source.to_markdown()}")

        return
    
    def json_to_dfs(self):
        files=os.listdir(config.jsons_dir)
        log.info(f"Number of files in API JSONS {len(files)}")
        json_dict=None
        jsons_dfs=[]
        urls_list=[]
        for idx, _file in enumerate(files):
            file_path=os.path.join(f"{config.jsons_dir}", _file)
            with open(file_path, "r", encoding="utf-8") as read_file:
                json_dict=json.load(read_file)
            _url=json_dict["odata.metadata"]
            _value=json_dict["value"]
            _next_link=json_dict["odata.nextLink"]
            _df=pd.DataFrame(_value)
            jsons_dfs.append(_df)
            urls_list.append(_url)
        return jsons_dfs, urls_list

    def prepare_json_df(self, _df:pd.DataFrame, url:str):
        _df["source"]=url.split("$metadata#")[1]
        _df["knesset_num"]=_df.apply(lambda x:
            x["FilePath"].split("https://fs.knesset.gov.il//")[1].split("/")[0], axis=1)
        _df["file_format"]=_df.apply(lambda x: x["FilePath"].split(".")[-1].lower(),axis=1)

        return _df

    def count_source_per_knesset(self, jsons_df:pd.DataFrame, source:str):
        source_records=jsons_df.loc[jsons_df["source"]==source]
        self.log.info(f"{len(source_records)} records on {source}")
        
        grouped_counts = source_records.groupby(["knesset_num", "file_format"]).size().reset_index(
            name='Counts')
        grouped_counts["source"]=source

        return  grouped_counts


    def count_words_file_by_source(self):
        """
        Count number of files, words and disk volume 
        downloaded from Knesset ODATA
        """
        _rslts=[]
        for idx, source in enumerate(config.datasets_sources):
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
                if idx2%5000==0:
                    log.info("End counting {} fles, {} MB, {} words".
                            format(idx2, round(source_total_size,0 ), number_of_words))
                continue

            _dict={
                "source":source,
                "number of files": len(files),
                "volume (MB)": round(source_total_size, 0),
                "number of words": number_of_words
            }
    
            _rslts.append(_dict)
    
            continue
        rslts_df=pd.DataFrame(_rslts)
        log.info(f"\n{rslts_df.to_markdown()}")



if __name__=='__main__':
    log=configure_logger('default')
    log.info("Program start")

    cfc=CountFilesNContent()
    cfc.run()

    log.info("Program ends")

