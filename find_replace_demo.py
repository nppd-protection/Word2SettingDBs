#! /usr/bin/python

""" find_replace_demo.py

This script demonstrates and can be a template for doing find and replace on
MS Word docx files.
"""

from __future__ import print_function, unicode_literals

from docx import Document
from WordHelpers import find_replace
import re
import sys
import os.path
import shutil
from random import choice
from string import ascii_letters


def position_defs():
    bkr_panel = ['Panel %d' % n for n in (302, 303, 304, 305)]
    pri_line_panel = ['Panel %d' % n for n in (401, 403, 406, 408)]
    sec_line_panel = ['Panel %d' % n for n in (402, 404, 407, 409)]
    breakers = ['3304', '3306', '3308', '3310']
    bkr_relay = ['11-' + b[-2:] for b in breakers]
    bkr_lockout = ['86-' + b[-2:] for b in breakers]
    bkr_dc_pos = [b[-3:] + 'P' for b in breakers]
    bkr_dc_neg = [b[-3:] + 'N' for b in breakers]
    bkr_dc_trip = [b[-3:] + 'RT' for b in breakers]
    bkr_bf_pos = [b[-3:] + 'BFP' for b in breakers]
    bkr_bf_neg = [b[-3:] + 'BFN' for b in breakers]
    lines = ['3507', '3505B', '3505A', '3508']
    remotes = ['GGS', 'Grand Island', 'GGS', 'Axtell']
    ov_relay = ['59L'+line[2:] for line in lines]
    ov_relay[0] = '59L07/11R3401'
    line_relay = ['L'+line[2:] for line in lines]
    return [bkr_panel,
            pri_line_panel,
            sec_line_panel,
            breakers,
            bkr_relay,
            bkr_lockout,
            bkr_dc_pos,
            bkr_dc_neg,
            bkr_dc_trip,
            lines,
            remotes,
            ov_relay,
            line_relay]


def position_replacements(defs, rot_steps):
    num_positions = len(defs[0])
    rtn = []
    for i in range(num_positions):
        pos1 = i
        pos2 = (i + rot_steps) % num_positions
        rtn.extend([(p[pos1], p[pos2]) for p in defs])
    return rtn

pre_replace = [('21-14A',    '21P-L05B'),
               ('21-14B',    '21S-L05B'),
               ('TS-14',     'TS-L05B')]

replace_rot1 = position_replacements(position_defs(), 1)

def random_placeholder():
    # Could use uuid.uuid4().hex
    # Using letters only may be better since most find/replace text is
    # numbers and could have collisions.
    return ''.join(choice(ascii_letters) for _ in range(32))


def do_replace(document, replace_list, keep_list=None):
    #  Use temporary text as intermediate placeholder so that the order of the
    #  find/replace operations isn't critical unless find patterns overlap.
    from_txt, replace_txt = zip(*replace_list)
    tmp_txt = [random_placeholder() for _ in enumerate(replace_list)]
    # Replace text to keep with random string to keep it from being affected
    # by the find/replace operation.
    if keep_list is not None:
        keep_tmp = [random_placeholder() for _ in enumerate(keep_list)]
        for t1, t2 in zip(keep_list, keep_tmp):
            find_replace(document, t1, t2)
    for t1, t2 in zip(from_txt, tmp_txt):
        find_replace(document, t1, t2)
    for t1, t2 in zip(tmp_txt, replace_txt):
        find_replace(document, t1, t2)
    if keep_list is not None:
        for t1, t2 in zip(keep_tmp, keep_list):
            find_replace(document, t1, t2)

documentParam = sys.argv[1]

doc_base = re.match('(.*)\.doc[xm]$', documentParam, flags=re.I).group(1)

print('Input file: ' + documentParam)
print('Base filename:  ' + doc_base)

save_file = os.path.join(os.path.dirname(documentParam),
                         doc_base + ' (rev)' + '.docx')

shutil.copyfile(documentParam, save_file)
document = Document(save_file)

# Needed due to unusual relay names at Grand Island
do_replace(document, pre_replace, keep_list=['GGS RAS'])
do_replace(document, replace_rot1, keep_list=['GGS RAS'])

document.save(save_file)
