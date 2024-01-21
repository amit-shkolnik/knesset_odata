import ast as _ast
import os as _os
import platform
import fileinput
import datetime

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

