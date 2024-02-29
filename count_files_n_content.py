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
            jsons_dfs[idx]=self.add_metadata_to_json_df(json_df, urls_list[idx])
            if idx%500==0:
                self.log.info(f"{idx} metadata added to DFs")
        full_df=pd.concat(jsons_dfs)
        self.log.info(f"{len(full_df)} records on all sources")
        full_df.drop_duplicates(keep='first', inplace=True)
        self.log.info(f"{len(full_df)} records after drop duplicates")

        source_counts_dict={}
        for source in config.datasets_sources:
            source_counts=self.count_source_per_knesset(full_df, source)
            source_counts_dict[source]=source_counts

        for source, source_df in source_counts_dict.items():
            self.log.info(f"SUMMARY FOR {source}")
            self.log.info(f"\n{source_df.to_markdown()}")
            _file=f"{source}_summary_per_knesset.csv"
            source_df.to_csv(_file)

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
            if idx%500==0:
                self.log.info(f"{idx} json files processed")
        return jsons_dfs, urls_list

    def add_metadata_to_json_df(self, _df:pd.DataFrame, url:str):        
        try:
            # Not all records contains Knesset number, some records are like:
            # https://fs.knesset.gov.il///FILER/E_SHARE/WMA_POOL/14/2013_04_29/2013_04_29_15_59_50_18_56_51_19.wmv
            knesseet_num=-1
            for row_index, row in _df.iterrows():
                if "https://fs.knesset.gov.il//" in row["FilePath"]:
                    candidate= row["FilePath"].split(
                        "https://fs.knesset.gov.il//")[1].split("/")[0] 
                    try:
                        candidate=int(candidate)
                        if candidate<50:
                            knesseet_num=candidate
                            break
                    except Exception as err:
                        continue

            _df["source"]=url.split("$metadata#")[1]
            _df["knesset_num"]=knesseet_num
            
            # _df.apply(lambda x:
            #     x["FilePath"].split("https://fs.knesset.gov.il//")[1].split("/")[0] \
            #         if "https://fs.knesset.gov.il//" in x["FilePath"] else default_knesset, axis=1)
            _df["file_format"]=_df.apply(lambda x: self.get_file_format(x),axis=1)
            
            return _df

        except Exception as err:
            self.log.exception(err)
            return _df

    def get_file_format(self, row:pd.Series):
        file_extension=row["FilePath"].split(".")[-1].lower()
        if "aspx" in file_extension:
            file_extension="aspx"
        return file_extension



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
                    log.info("End counting {} files, {} MB, {} words".
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

