import ast as _ast
import os as _os
import platform
import fileinput
import datetime
import requests
import pandas as pd
import time
from pywintypes import com_error
import tabulate
import logging
import json


main_path=None


log_file = "logs/g_log.txt"
"""
Level 	Numeric value
CRITICAL 	50
ERROR 	40
WARNING 	30
INFO 	20
DEBUG 	10
NOTSET 	0
"""
log_level='INFO'
default='default'

# Documents datasources on Knesset ODATA to be scraped.
plenum_session_ref="KNS_DocumentPlenumSession"
committees_sessions="KNS_DocumentCommitteeSession"
bills="KNS_DocumentBill"

# Datasource to download from    
datasets_sources=[bills, plenum_session_ref, committees_sessions]

plenum_session="KNS_PlenumSession"
knesset_committies="KNS_Committee"
meta_data_tables=[plenum_session, knesset_committies ]

# Knesset ODATA site
main_hypelink="http://knesset.gov.il/Odata/ParliamentInfo.svc/"

odata_download_format="format=json"
ms_words_suffix=["doc", "DOC", "docx", "DOCX"] #, "rtf"]

# Documents corrupted previously downloaded
# and can't be open via 'Word'   
corrupted_docs_log="corrupted_docs_log.csv"

false_words=['\n', '\r']

jsons_dir="odata_jsons"