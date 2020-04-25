#!/usr/bin/python

import sys
import os
import codecs
import datetime
import shutil
import re

day = datetime.date.today().strftime("%Y%m%d")
dir = os.listdir("C:/Users/TokuharM/Desktop/Python_Scripts/AppServerLogs")

key_word = sys.argv[1]
arg_num = len(sys.argv)
d_sum = []

if arg_num != 2:
    print ('Usage: python logsearch.py "search_word"')

#Functions
def fild_all_files(directory):
    for root, dirs, files in os.walk(directory):
        yield root
        for file in files:
            yield os.path.join(root, file)

#Main Procedure
for sub_dir in dir:
    d_sum = []
    for line in fild_all_files('C:/Users/TokuharM/Desktop/Python_Scripts/AppServerLogs/%s' %sub_dir):
        line = line.rstrip()
        if 'logfile' in line:
            path = line.replace('\\', '/')
            with codecs.open('%s' %path, 'r', 'utf-8') as logfile:
                dt = datetime.datetime.fromtimestamp(os.stat('%s' %path).st_mtime)
                dt = dt.strftime('%Y-%m-%d')
                count = 0
                for line in logfile:
                    line = line.rstrip()
                    if not re.match(r'^2015', line):
                        continue
                    if re.search(r'^2015.*', line) and key_word in line:
                        if count == 0:
                            time = line.split(' ')
                            s_time = time[1]
                            s_time = s_time.split(',', 1)
                        count += 1
                        time = line.split(' ')
                        e_time = time[1]
                        e_time = e_time.split(',', 1)
                print ('%s__%s_%s_occurred %s times' % (dt, sub_dir, key_word, count))
                if count != 0:
                    print ('Between %s and %s' % (s_time[0], e_time[0]))
            d_sum.append(count)
        t = sum(d_sum)
    print ('Total number of %s on %s is %s' % (key_word, sub_dir, t))
