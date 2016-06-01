#! /usr/bin/python

'''
USAGE:
FindReplace.py <Document.docx> find_text replace_text
<Document.docx>:  MS Word docx document

ABOUT:
Find some text and replace with other text.

Progress information and errors are logged to the same directory as the program
is run from.

'''

from __future__ import print_function, unicode_literals

from WordHelpers import *

import re, sys, os.path, os, codecs, msvcrt, shutil

# Set up a logger so any errors can go to file to facilitate debugging
import logging
from logging.config import dictConfig

# By default, log to the same directory the program is run from    
if os.path.exists(os.path.dirname(sys.argv[0])):
    logfile = os.path.join(os.path.dirname(sys.argv[0]), 'FindReplace.log')
else:
    logfile = 'SplitByHighlighting.log'

logging_config = {
    'version': 1,
    'formatters': {
        'file': {'format':
              '%(asctime)s ' + os.environ['USERNAME'] + ' %(levelname)-8s %(message)s'},
        'console': {'format':
              '%(levelname)-8s %(message)s'}
        },
        
    'handlers': {
        'file': {'class': 'logging.FileHandler',
            'filename': logfile,
            'formatter': 'file',
            'level': 'INFO'},
        'console': {'class': 'logging.StreamHandler',
            'formatter': 'console',
            'level': 'DEBUG'}
        },
    'loggers': {
        'root' : {'handlers': ['file', 'console'],
            'level': 'DEBUG'}
        }
}

dictConfig(logging_config)

logger = logging.getLogger('root')

try:
    logger.info('Running %s.' % sys.argv[0])
    logger.info('Logging to file %s.' % os.path.abspath(logfile))

    '''
    The docx package documentation can be found at
    https://python-docx.readthedocs.org/en/latest/

    Installation:
    If pip is available:
    pip install python-docx

    If easy_install is available:
    easy_install python-docx

    Otherwise download the tar.gz and run setup.py install.
    Currently working with 0.8.5. Version 0.7.2 NOT working.
    As of 5/27/2016, I needed the development branch in order to get the
    highlighting code.
    '''
    from docx import Document
    from docx.enum.text import WD_COLOR_INDEX

    debug = False
    sys.stdout = codecs.getwriter(sys.stdout.encoding)(sys.stdout, errors='replace')

    if len(sys.argv) < 2:
        logger.error("Not enough input parameters.  Please include one filename when calling this program.")
        logger.error(__doc__)
        raise SystemExit
    elif len(sys.argv) > 2:
        logger.error("Too many input parameters.  Please include one filename when calling this program.")
        logger.error(__doc__)
        raise SystemExit

    documentParam = sys.argv[1]


    doc_base = re.match('(.*)\.doc[xm]$',documentParam, flags=re.I).group(1)

    logger.info('Input file: ' + documentParam)
    logger.info('Base filename:  ' + doc_base)
    #logger.info('Output files: ' + sel_save_file + ', ' + aspen_save_file)

    
    save_file = doc_base + ' (modified).docx'
    shutil.copyfile(documentParam, save_file)
    document = Document(save_file)
    
    # Hardcoded for testing.
    find_text = 'test'
    replace_text = 'replaced'
    
    find_replace(document, find_text, replace_text)

    
    logger.info('Saving output file: %s' % save_file)
    document.save(save_file)

except (SystemExit, KeyboardInterrupt):
    raise
except Exception, e:
    logger.error('Program error', exc_info=True)
finally:    
    # Code copied from http://stackoverflow.com/questions/11876618/python-press-any-key-to-exit    
    # This uses a Windows-specific library (msvrt).
    logger.info('DONE.')
    print('Press any key to exit...')
    #junk = msvcrt.getch() # Assign to a variable just to suppress output. Blocks until key press.