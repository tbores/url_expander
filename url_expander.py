#!/usr/bin/python2.7
from xlrd import open_workbook
from xlwt import easyxf
from xlutils.copy import copy
import urllib2
import ast
import pickle
import signal
import sys

# Configuration
VERSION = 'v1.0'
SOCKET_TIMEOUT = 30
DICT_PATH = 'urls.pkl'
INPUT_XLS_PATH = 'TEST twitonomy_studentenfutta_ceriseonthecake.xls'
OUTPUT_XLS_PATH = 'OUT_'+INPUT_XLS_PATH

# Functions
def load_already_known_urls():
    urls_dict = dict()
    print 'Try to load list of already known URL from '+DICT_PATH+'.'
    try:
        urls_file = open(DICT_PATH, 'rb')
        urls_dict = pickle.load(urls_file)
        urls_file.close()
	print 'Found '+str(len(urls_dict.keys()))+' urls.'
    except IOError, e:
        print e
        print 'The '+DICT_PATH+' file doen\'t exists?'
        print 'It is normal the first start.'
    return urls_dict

def save_url_list():
    print 'Save urls in '+DICT_PATH+'.' 
    try:
        urls_file = open(DICT_PATH, 'wb')
        pickle.dump(urls_dict, urls_file)
        urls_file.close()
    except IOError, e:
        print e
        print 'Couldn\'t save url list for future usage!'

def signal_handler(signal, frame):
    print 'You pressed Ctrl+C!'
    save_url_list()
    sys.exit(0)

# Start
print 'Welcome in url_expander.py '+VERSION
print 'Coded by Thomas Bores'
signal.signal(signal.SIGINT, signal_handler)
urls_dict = load_already_known_urls()
rb = open_workbook(INPUT_XLS_PATH,
                   formatting_info=True)
r_sheet = rb.sheet_by_index(0)

wb = copy(rb)
ws = wb.get_sheet(0)

for c_row in range(r_sheet.nrows):
    content = r_sheet.cell(c_row, 3).value
    if 'http://' in content:
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
            print 'URL is unknown :(, we connect to expandurl service'
            try:
                res = urllib2.urlopen('http://expandurl.me/expand?url='+url,
                                      timeout=SOCKET_TIMEOUT).read()
                res_as_dict = ast.literal_eval(res)
                real_url = res_as_dict['end_url']
                urls_dict[url] = real_url
            except IOError, e:
		print e
                print "Timeout error?"
        print 'real url = \"'+real_url+'\"'
        new_content = content.replace(url, real_url)
        ws.write(c_row, 3, new_content)

wb.save(OUTPUT_XLS_PATH)
save_url_list()

