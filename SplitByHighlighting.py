#! /usr/bin/python

'''
USAGE:
SplitByHighlighting.py <Document.docx>
<Document.docx>:  MS Word docx document

ABOUT:
The program will take a MS Word document and split it into multiple documents
based on highlight colors. Each resulting document will include all
unhighlighted text as well as text highlighted in ONE color.

Progress information and errors are logged to the same directory as the program
is run from.

'''

from __future__ import print_function, unicode_literals

from WordHelpers import remove_highlighted, iter_all_runs

import re
import sys
import os.path
import os
import codecs
import msvcrt
import shutil

# Set up a logger so any errors can go to file to facilitate debugging
import logging
from logging.config import dictConfig

#
# The docx package documentation can be found at
# https://python-docx.readthedocs.org/en/latest/
#
# Installation:
# If pip is available:
# pip install python-docx
#
# If easy_install is available:
# easy_install python-docx
#
# Otherwise download the tar.gz and run setup.py install.
# Currently working with 0.8.5. Version 0.7.2 NOT working.
# As of 5/27/2016, I needed the development branch in order to get the
# highlighting code.
#
from docx import Document

# By default, log to the same directory the program is run from
if os.path.exists(os.path.dirname(sys.argv[0])):
    logfile = os.path.join(os.path.dirname(sys.argv[0]),
                           'SplitByHighlighting.log')
else:
    logfile = 'SplitByHighlighting.log'

logging_config = {
    'version': 1,
    'formatters': {
        'file': {'format':
                 '%(asctime)s ' + os.environ['USERNAME'] +
                 ' %(levelname)-8s %(message)s'},
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
                    'level': 'INFO'}
        },
    'loggers': {
        'root': {'handlers': ['file', 'console'],
                 'level': 'DEBUG'}
        }
}

dictConfig(logging_config)

logger = logging.getLogger('root')

try:
    logger.info('Running %s.' % sys.argv[0])
    logger.info('Logging to file %s.' % os.path.abspath(logfile))

    debug = False
    sys.stdout = codecs.getwriter(sys.stdout.encoding)(sys.stdout,
                                                       errors='replace')

    if len(sys.argv) < 2:
        logger.error("Not enough input parameters.  Please include one "
                     "filename when calling this program.")
        logger.error(__doc__)
        raise SystemExit
    elif len(sys.argv) > 2:
        logger.error("Too many input parameters.  Please include one filename "
                     "when calling this program.")
        logger.error(__doc__)
        raise SystemExit

    documentParam = sys.argv[1]

    doc_base = re.match('(.*)\.doc[xm]$', documentParam, flags=re.I).group(1)

    logger.info('Input file: ' + documentParam)
    logger.info('Base filename:  ' + doc_base)

    document = Document(documentParam)

    run_dict = dict()

    for p, r in iter_all_runs(document):
        if r.font.highlight_color is not None:
            logger.debug(r.text + ' ===> ' + str(r.font.highlight_color))
            try:
                run_dict[r.font.highlight_color].append(r)
            except KeyError:
                run_dict[r.font.highlight_color] = [r]

    logger.debug('Colors used and number of runs')
    for k, v in run_dict.items():
        logger.debug(str(k) + ': ' + str(len(v)))

    # For each color, make a temp copy of the document, clear text in runs
    # of different highlighting colors and write the output to a new file.
    all_colors = set(run_dict.keys())
    for keep_color in all_colors:
        save_file = doc_base + ' (' + str(keep_color).split()[0] + ').docx'
        shutil.copyfile(documentParam, save_file)
        new_document = Document(save_file)
        for remove_color in all_colors - set([keep_color]):
            remove_highlighted(new_document, remove_color)
        logger.info('Saving output file: %s' % save_file)
        new_document.save(save_file)

except (SystemExit, KeyboardInterrupt):
    raise
except Exception, e:
    logger.error('Program error', exc_info=True)
finally:
    logger.info('DONE.')
    print('Press any key to exit...')
    # Assign to a variable just to suppress output. Blocks until key press.
    junk = msvcrt.getch()
