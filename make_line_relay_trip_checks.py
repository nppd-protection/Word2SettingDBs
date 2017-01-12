#! /usr/bin/python

'''
USAGE:
make_line_relay_trip_checks.py ...
    <Trip Checks Line Relay Master.docx>

<STrip Checks Line Relay Master.docx>:  MS Word docx document master template
    for line relay trip checks.

ABOUT:
Creates individual Word document of a specific construction standard
based on the all-in-one master standard.

Progress information and errors are logged to the same directory as the program
is run from.

'''

from __future__ import print_function, unicode_literals

from WordHelpers import find_replace, remove_highlighted, clear_highlighting

import re
import sys
import os.path
import os
import codecs
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

# See https://python-docx.readthedocs.io/en/latest/api/enum/WdColorIndex.html
from docx.enum.text import WD_COLOR_INDEX

# Definitions for how to modify the template for a specific application

# Assigned highlighting colors in this document:
#   wdTurquoise (3):    DCB (PLC) Settings
#   wdBrightGreen (4):  POTT (Mirrored Bits) Settings
#   wdGreen (11):       Panel with bus differential (single-bus only)
#   wdPink (5):         Ring or breaker-and-one-half bus
#   wdTeal (10):        Single-bus, single-breaker
#   wdYellow (7):       Automated RTU
#   wdGray25(16):       Non-automated RTU
#

# Some constants are used to make it the code more readable
# and allow the color to be tied to a logical use in one place
# so colors could be changed without editing a lot of code.
#
DCBPri = WD_COLOR_INDEX.TURQUOISE
POTTPri = WD_COLOR_INDEX.BRIGHT_GREEN
SEL411LPriPOTTSec = WD_COLOR_INDEX.DARK_YELLOW
BusDiff = WD_COLOR_INDEX.GREEN
OneBkr = WD_COLOR_INDEX.TEAL
TwoBkr = WD_COLOR_INDEX.PINK
Automated = WD_COLOR_INDEX.YELLOW
NonAutomated = WD_COLOR_INDEX.GRAY_25

all_colors = set([
    DCBPri,
    POTTPri,
    SEL411LPriPOTTSec,
    BusDiff,
    OneBkr,
    TwoBkr,
    Automated,
    NonAutomated])


# Colors to keep for each standard. (Inverse of colors to remove.)
colors_to_keep = {
    'PP115-230E1A3A': set([DCBPri,  OneBkr, BusDiff, Automated]),
    'PP115-230E1A3B': set([POTTPri, OneBkr, BusDiff, Automated]),
    'PP115-230E1A3C': set([SEL411LPriPOTTSec, OneBkr, BusDiff, Automated]),
    'PP115-230E1B3A': set([DCBPri,  TwoBkr, Automated]),
    'PP115-230E1B3B': set([POTTPri, TwoBkr, Automated]),
    'PP115-230E1B3C': set([SEL411LPriPOTTSec, TwoBkr, Automated]),
    'PP115-230E2A3A': set([DCBPri,  OneBkr, BusDiff, NonAutomated]),
    'PP115-230E2A3B': set([POTTPri, OneBkr, BusDiff, NonAutomated]),
    'PP115-230E2A3C': set([SEL411LPriPOTTSec, OneBkr, BusDiff, NonAutomated]),
    'PP115-230E2B3A': set([DCBPri,  TwoBkr, NonAutomated]),
    'PP115-230E2B3B': set([POTTPri, TwoBkr, NonAutomated]),
    'PP115-230E2B3C': set([SEL411LPriPOTTSec, TwoBkr, NonAutomated]),
    'PP115-230ExAxx': set([DCBPri, POTTPri, SEL411LPriPOTTSec, OneBkr, BusDiff, Automated, NonAutomated]),
    'PP115-230ExBxx': set([DCBPri, POTTPri, SEL411LPriPOTTSec, TwoBkr, Automated, NonAutomated]),
    'TEST':           set()}

highlights_to_keep = {
    'PP115-230E1A3A': set([BusDiff]),
    'PP115-230E1A3B': set([BusDiff]),
    'PP115-230E1A3C': set([BusDiff]),
    'PP115-230E1B3A': set([]),
    'PP115-230E1B3B': set([]),
    'PP115-230E1B3C': set([]),
    'PP115-230E2A3A': set([BusDiff]),
    'PP115-230E2A3B': set([BusDiff]),
    'PP115-230E2A3C': set([BusDiff]),
    'PP115-230E2B3A': set([]),
    'PP115-230E2B3B': set([]),
    'PP115-230E2B3C': set([]),
    'PP115-230ExAxx': set([DCBPri, POTTPri, SEL411LPriPOTTSec, BusDiff, Automated, NonAutomated]),
    'PP115-230ExBxx': set([DCBPri, POTTPri, SEL411LPriPOTTSec, Automated, NonAutomated]),
    'TEST':           set()}

std_filenames = {
    'PP115-230E1A3A': 'DCB Single-Bkr Automated',
    'PP115-230E1A3B': 'POTT Single-Bkr Automated',
    'PP115-230E1A3C': '411L-POTT Single-Bkr Automated',
    'PP115-230E1B3A': 'DCB Bkr-and-half or Ring Automated',
    'PP115-230E1B3B': 'POTT Bkr-and-half or Ring Automated',
    'PP115-230E1B3C': '411L-POTT Bkr-and-half or Ring Automated',
    'PP115-230E2A3A': 'DCB Single-Bkr Non-Automated',
    'PP115-230E2A3B': 'POTT Single-Bkr Non-Automated',
    'PP115-230E2A3C': '411L-POTT Single-Bkr Non-Automated',
    'PP115-230E2B3A': 'DCB Bkr-and-half or Ring Non-Automated',
    'PP115-230E2B3B': 'POTT Bkr-and-half or Ring Non-Automated',
    'PP115-230E2B3C': '411L-POTT Bkr-and-half or Ring Non-Automated',
    'PP115-230ExAxx': 'Single Breaker Trip Check Template',
    'PP115-230ExBxx': 'Ring Bkr-and-a-half Trip Check Template',
    'TEST':           'Script Test'}

# For testing only
#all_colors = set([OneBkr])
#std_filenames = {'TEST':  'Script Output'}

def make_line_relay_trip_checks(document, std):
    # Remove highlighted sections except colors to keep
    colors_to_remove = all_colors - colors_to_keep[std]
    for color in colors_to_remove:
        remove_highlighted(document, color)

    # Remove remaining highlighting except highlighting to keep
    highlights_to_remove = colors_to_keep[std] - highlights_to_keep[std]
    for color in highlights_to_remove:
        clear_highlighting(document, color)

    # Special find/replace for some automated/non-automated text
    if Automated in colors_to_keep[std] and NonAutomated not in colors_to_keep[std]:
        find_replace(document,
                     'HMI/annunciator/supervisory',
                     'HMI/supervisory')
    elif NonAutomated in colors_to_keep[std] and Automated not in colors_to_keep[std]:
        find_replace(document,
                     'HMI/supervisory',
                     'supervisory')
        find_replace(document,
                     'HMI/annunciator/supervisory',
                     'annunciator/supervisory')

# By default, log to the same directory the program is run from
if os.path.exists(os.path.dirname(sys.argv[0])):
    logfile = os.path.join(os.path.dirname(sys.argv[0]),
                           'make_line_relay_trip_checks.log')
else:
    logfile = 'make_line_relay_trip_checks.log'

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
                    'level': 'DEBUG'}
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

    for std in std_filenames.keys():

        try:
            file_rev = re.search(' Rev ?[0-9]+$', doc_base,
                                 flags=re.I).group(0)
        except AttributeError:
            file_rev = ''

        save_file = os.path.join(os.path.dirname(documentParam),
                                 std + ' ' + std_filenames[std] + \
                                 file_rev + '.docx')
        shutil.copyfile(documentParam, save_file)
        document = Document(save_file)

        make_line_relay_trip_checks(document, std)

        logger.info('Saving output file: %s' % save_file)
        document.save(save_file)

except (SystemExit, KeyboardInterrupt):
    raise
except Exception, e:
    logger.error('Program error', exc_info=True)
    raise
finally:
    logger.info('DONE.')
