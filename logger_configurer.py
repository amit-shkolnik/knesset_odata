import logging
import logging.config
import sys

import config


def configure_logger(name):
    logging.config.dictConfig({
        'version': 1,
        'formatters': {
            'default': {'format': '%(asctime)s - %(levelname)s - %(message)s', 'datefmt': '%Y-%m-%d %H:%M:%S'}
        },
        'handlers': {
            'console': {
                'level': config.log_level,
                'class': 'logging.StreamHandler',
                'formatter': 'default',
                'stream': 'ext://sys.stdout'
            },
            'file': {
                'level': config.log_level,
                'class': 'logging.handlers.RotatingFileHandler',
                'formatter': 'default',
                'filename': config.log_file,
                'maxBytes': 1048576,
                'backupCount': 3
            }
        },
        'loggers': {
            'default': {
                'level': config.log_level,
                'handlers': ['console', 'file']
            }
        },
        'disable_existing_loggers': False
    })
    return logging.getLogger(name)

