#!/usr/bin/python2.7
# -*- coding: utf-8 -*-

"""
url_expander.py

This program enables user to parse excel files generated from twitony
and to convert shortened urls in expended urls.
It doesn't modify the input files but create new files instead.

Update the User configuration below for your usage.
"""

from xlrd import open_workbook
from xlutils.copy import copy
import urllib2
import ast
import pickle
import signal
import sys
import os

# User configuration, should be updated
INPUT_EXCEL_DIR = './excels/'
OUTPUT_EXCEL_DIR = './expanded_urls_excels/'

# Program Configuration
VERSION = 'v2.0'
SOCKET_TIMEOUT = 30
DICT_PATH = 'urls.pkl'

# Functions
def load_already_known_urls():
    """
    Load urls that were saved during previous passes.

    Return:
        a dict of urls with key is the shortened url and value
        the expanded url
    """
    urls = dict()
    print 'Try to load list of already known URL from '+DICT_PATH+'.'
    try:
        urls_file = open(DICT_PATH, 'rb')
        urls = pickle.load(urls_file)
        urls_file.close()
        print 'Found '+str(len(urls.keys()))+' urls.'
    except IOError, error:
        print error
        print 'The '+DICT_PATH+' file doen\'t exists?'
        print 'It is normal the first start.'
    return urls

def save_url_list(urls):
    """
    Save the found urls in a pickle file
    """
    print 'Save urls in '+DICT_PATH+'.'
    try:
        urls_file = open(DICT_PATH, 'wb')
        pickle.dump(urls, urls_file)
        urls_file.close()
    except IOError, error:
        print error
        print 'Couldn\'t save url list for future usage!'

def signal_handler():
    """
    Catch the <CTRL>+<C> keyboard keys for saving results before quitting
    """
    print 'You pressed Ctrl+C!'
    save_url_list(urls_dict)
    sys.exit(0)

# -------------------------
# ----- Program Start -----
# -------------------------
print 'Welcome in url_expander.py '+VERSION
print 'Coded by Thomas Bores'

# Load already known urls
urls_dict = load_already_known_urls()

# Catch CTRL+C
signal.signal(signal.SIGINT, signal_handler)

# Check if INPUT_EXCEL_DIR exists
if os.path.exists(INPUT_EXCEL_DIR) == False:
    os.makedirs(INPUT_EXCEL_DIR)
# Check if OUTPUT_EXCEL_DIR exists
if os.path.exists(OUTPUT_EXCEL_DIR) == False:
    os.makedirs(OUTPUT_EXCEL_DIR)

# Open all exels files in INPUT_EXCEL_DIR and parse them
for f_xls in os.listdir(INPUT_EXCEL_DIR):
    if f_xls.endswith('.xls'):
        try:
            rb = open_workbook(INPUT_EXCEL_DIR+f_xls,
                               formatting_info=True)
            r_sheet = rb.sheet_by_index(0)

            wb = copy(rb)
            ws = wb.get_sheet(0)

            for c_row in range(r_sheet.nrows):
                content = r_sheet.cell(c_row, 3).value
                if 'http://' in content:
                    try:
                        print '--------------------'
                        print 'Find short url at row '+str(c_row)
                        pos1 = content.find('http://')
                        sub1 = content[pos1:]
                        pos2 = len(sub1)

                        if ' ' in sub1:
                            pos2 = sub1.find(' ')

                        url = sub1[:pos2]
                        print 'short url = \"'+url+'\"'

                        # Get real url
                        real_url = url
                        if url in urls_dict:
                            # Check if we already have it
                            print 'URL already in dict :)'
                            real_url = urls_dict[url]
                        else:
                            # Otherwise expand it with online expandurl service
                            print """URL is unknown :(,
we connect to expandurl service"""
                            try:
                                res = urllib2.urlopen(
                                            'http://expandurl.me/expand?url='+\
                                            url,
                                            timeout=SOCKET_TIMEOUT).read()
                                res_as_dict = ast.literal_eval(res)
                                real_url = res_as_dict['end_url']
                                urls_dict[url] = real_url
                            except IOError, error:
                                print error
                                print "Timeout error"
                        print 'real url = \"'+real_url+'\"'
                        new_content = content.replace(url, real_url)
                        ws.write(c_row, 3, new_content)

                    except UnicodeEncodeError, error:
                        print error
                        print 'Cannot read content of row '+str(c_row)+\
                              ' in file '+f_xls

            wb.save(OUTPUT_EXCEL_DIR+'expanded_url_'+f_xls)
        except IOError, error:
            print error
            print "Cannot read excel file "+f_xls

save_url_list(urls_dict)

