#! /usr/bin/python

'''
USAGE:
make_line_relay_templates.py <Settings Dual SEL-421 Line Relay Master Standard.docx>
<Settings Dual SEL-421 Line Relay Master Standard.docx>:  MS Word docx document,
    master template for line relays.

ABOUT:
Creates individual Word document of a specific construction standard
based on the all-in-one master standard.

Progress information and errors are logged to the same directory as the program
is run from.

'''

from __future__ import print_function, unicode_literals

from WordHelpers import *

import re, sys, os.path, os, codecs, msvcrt, shutil

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

# See https://python-docx.readthedocs.io/en/latest/api/enum/WdColorIndex.html
from docx.enum.text import WD_COLOR_INDEX

# Definitions for how to modify the template for a specific application
# Adapted from the original MS Word macro.

# Assigned highlighting colors in this document:
#   wdTurquoise (3):    DCB (PLC) Settings
#   wdBrightGreen (4):  POTT (Mirrored Bits) Settings
#   wdGray25(16):       115/230 Settings (Non-piloted secondary, SEL-421-4, etc.)
#   wdRed (6):          DCUB Primary / POTT Secondary
#   wdTeal (10):        SEL-411L Primary / POTT Secondary
#   wdDarkYellow (14):  SEL-421 Primary
#   wdYellow (7):       Single-bus, single-breaker
#   wdPink (5):         Ring or breaker-and-one-half bus
#

# Some constants are used to make it the code more readable
# and allow the color to be tied to a logical use in one place
# so colors could be changed without editing a lot of code.
#
DCBPri = WD_COLOR_INDEX.TURQUOISE
POTTPri = WD_COLOR_INDEX.BRIGHT_GREEN
NonPilotSec = WD_COLOR_INDEX.GRAY_25
DCUBPriPOTTSec = WD_COLOR_INDEX.RED
SEL411LPriPOTTSec = WD_COLOR_INDEX.TEAL
SEL421Pri = WD_COLOR_INDEX.DARK_YELLOW
OneBkr = WD_COLOR_INDEX.YELLOW
TwoBkr = WD_COLOR_INDEX.PINK

all_colors = set([
    DCBPri,
    POTTPri,
    NonPilotSec,
    DCUBPriPOTTSec,
    SEL411LPriPOTTSec,
    SEL421Pri,
    OneBkr,
    TwoBkr])

# Colors to remove for each standard.
colors_to_remove = {
    #DCBPriNonPilotSecOneBreaker
    'PP115-230E1A3A': [
        TwoBkr,
        DCUBPriPOTTSec,
        POTTPri,
        SEL411LPriPOTTSec],

    #POTTPriNonPilotSecOneBreaker
    'PP115-230E1A3B': [
        TwoBkr,
        DCUBPriPOTTSec,
        DCBPri,
        SEL411LPriPOTTSec],

    #411LPriPOTTSecOneBreaker
    #ChangePri421to411L
    'PP115-230E1A3C': [
        TwoBkr,
        NonPilotSec,
        DCUBPriPOTTSec,
        DCBPri,
        POTTPri,
        SEL421Pri],

    #DCBPriNonPilotSecTwoBreaker
    'PP115-230E1B3A': [
        OneBkr,
        DCUBPriPOTTSec,
        POTTPri,
        SEL411LPriPOTTSec],

    #POTTPriNonPilotSecTwoBreaker
    'PP115-230E1B3B': [
        OneBkr,
        DCUBPriPOTTSec,
        DCBPri,
        SEL411LPriPOTTSec],
    
    #411LPriPOTTSecTwoBreaker
    #ChangePri421to411L
    'PP115-230E1B3C': [
        OneBkr,
        NonPilotSec,
        DCUBPriPOTTSec,
        DCBPri,
        POTTPri,
        SEL421Pri],

    #DCUBPriPOTTSecTwoBreaker
    'PP115-230E1B3A': [
        OneBkr,
        DCBPri,
        POTTPri,
        NonPilotSec,
        SEL411LPriPOTTSec]
}

# Colors to keep for each standard. (Inverse of colors to remove.)
all_colors = set([OneBkr, TwoBkr,
                  DCBPri, POTTPri, SEL411LPriPOTTSec, DCUBPriPOTTSec,
                  SEL421Pri, NonPilotSec])
                  
colors_to_keep = {
    'PP115-230E1A3A': set([SEL421Pri, DCBPri,  NonPilotSec, OneBkr]),
    'PP115-230E1A3B': set([SEL421Pri, POTTPri, NonPilotSec, OneBkr]),
    'PP115-230E1A3C': set([SEL411LPriPOTTSec,               OneBkr]),
    'PP115-230E1B3A': set([SEL421Pri, DCBPri,  NonPilotSec, TwoBkr]),
    'PP115-230E1B3B': set([SEL421Pri, POTTPri, NonPilotSec, TwoBkr]),
    'PP115-230E1B3C': set([SEL411LPriPOTTSec,               TwoBkr]),
    'PP345J0X5A':     set([SEL421Pri, DCBPri,  NonPilotSec, TwoBkr]),
    'TEST':           set()}
    
std_filenames = {
    'PP115-230E1A3A': 'Settings PP115-230E1A3A RFL9785-DCB 21P-02 11S-02 SEL-421-4',
    'PP115-230E1A3B': 'Settings PP115-230E1A3B MB-POTT 21P-02 11S-02 SEL-421-4',
    'PP115-230E1A3C': 'Settings PP115-230E1A3C 87P-02 11S-02 SEL-411L-421-4',
    'PP115-230E1B3A': 'Settings PP115-230E1B3A RFL9785-DCB 21P-L1098 21S-L1098 SEL-421-4',
    'PP115-230E1B3B': 'Settings PP115-230E1B3B MB-POTT 21P-L1369 21S-L1369 SEL-421-4',
    'PP115-230E1B3C': 'Settings PP115-230E1B3C 87P-L91A 21S-L91A SEL-411L-421-4',
    'PP345J0X5A':     'Settings PP345J0X5A 345kv line relays 21P-21S SEL-421-5',
    'TEST':           'Settings Dual SEL-421 Line Relay master Standard (macro test)'}
    
def change_pri_421_to_411L(document):
    replace_list = [('OUT2', 'OUT3'),
                    ('OUT1', 'OUT2'),
                    ('IN2',  'IN3'),
                    ('IN1',  'IN2'),
                    ('IAXM', 'IAXFM'),
                    ('IBXM', 'IBXFM'),
                    ('ICXM', 'ICXVM'),
                    ('51S1T', '51T01')]
    replace_start = get_bookmark_par_element(document, 'SecondarySettingsStart')
    for find_text, replace_text in replace_list:
        find_replace(document, find_text, replace_text, end=replace_start)
    find_replace(document, '87P-0X', '87P-LZZ')

def make_line_relay_templates(document, std):
    colors_to_remove = all_colors - colors_to_keep[std]
    for color in colors_to_remove:
        remove_highlighted(document, color, clean_logic_tables=True)
    if SEL411LPriPOTTSec in colors_to_keep[std]:
        change_pri_421_to_411L(document)

# Set up a logger so any errors can go to file to facilitate debugging
import logging
from logging.config import dictConfig

# By default, log to the same directory the program is run from    
if os.path.exists(os.path.dirname(sys.argv[0])):
    logfile = os.path.join(os.path.dirname(sys.argv[0]), 'make_line_relay_templates.log')
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

    # For testing, hard-code standard to use
    std = 'PP115-230E1A3A'
    for std in std_filenames.keys():
        
        try:
            file_rev = re.search(' Rev ?[0-9]+$', doc_base, flags=re.I).group(0)
        except AttributeError:
            file_rev = ''
            
        save_file = os.path.join(os.path.dirname(documentParam), std_filenames[std] + file_rev + '.docx')
        shutil.copyfile(documentParam, save_file)
        document = Document(save_file)
        
        make_line_relay_templates(document, std)

        
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