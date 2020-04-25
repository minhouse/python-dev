#!/usr/bin/python

#Modules
import sys, os, codecs, datetime, shutil, re
import xlsxwriter
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from mpl_toolkits.axes_grid1 import Grid

#Variables
day = datetime.date.today().strftime("%Y%m%d")
dir = os.listdir("C:/Users/TokuharM/Desktop/Python_Scripts/GsDlmsCmdRsp_logs") #Change Here

#Reading Argument
key_word = sys.argv[1]
arg_num = len(sys.argv)

if arg_num != 2:
    print('Usage: python logsearch.py "search_word"')

#Writting Excel file
exceldata = "errordata.xlsx"
workbook = xlsxwriter.Workbook('C:/Users/TokuharM/Desktop/Python_Scripts/%s' % exceldata)
excelheader = ['date', 'server', 'errors']

# Functions
def fild_all_files(directory):
    for root, dirs, files in os.walk(directory):
        yield root
        for file in files:
            yield os.path.join(root, file)

def graph(excel, dflist):
    fig, axes = plt.subplots(nrows=4, ncols=3, figsize=(16, 9), sharex=False, sharey=False)
    #fig.subplots_adjust(wspace=0, hspace=0)
    fig.subplots_adjust(hspace=0.5, left=0.5, right=0.9, bottom=0.01, top=0.95)
    fig.autofmt_xdate()
    plt.setp(axes, xticklabels=[], yticks=[])

    for i, j in enumerate(dflist):
        df = pd.read_excel('C:/Users/TokuharM/Desktop/Python_Scripts/%s' %excel, sheetname ='%s' % j, na_values=['NA'],parse_dates=['date'])
        df.sort_index(by = ["date"])

        x = df[["date"]]
        y = df[["errors"]]

        ax = fig.add_subplot(4,3,i+1)
        #ax = fig.add_subplot(i/3+1,i%3+1,1)
        ax.plot(x, y, "-o") #axに対してラベル幅などの調整を行う

        if len(x) > 0:
            days = mdates.DayLocator()
            daysFmt = mdates.DateFormatter('%m-%d')
            ax.xaxis.set_major_locator(days)
            ax.xaxis.set_major_formatter(daysFmt)
            ax.xaxis.set_ticks(pd.date_range(x.iloc[1,0], x.iloc[-1,0], freq='7d'))
            #ax.set_xlabel('Date')
            ax.set_ylabel('Errors')
            #ax.set_xticklabels('off')
            #ax.set_yticklabels('off')
            ax.grid()
            plt.xticks(rotation=30, fontsize='8')
            plt.yticks(range(0,400,30), fontsize='8')
            plt.axis('tight')
            ax.set_title(j, fontsize='10')
    plt.tight_layout(pad=1.0, w_pad=1.0, h_pad=1.0)
    plt.show()

# Main Procedure
servers_name = []
for sub_dir in dir:
    d_sum = []
    cellnum = 2
    worksheet1 = workbook.add_worksheet(sub_dir)
    worksheet1.write('A1', excelheader[0])
    worksheet1.write('B1', excelheader[1])
    worksheet1.write('C1', excelheader[2])
    servers_name.append(sub_dir)
    for line in fild_all_files('C:/Users/TokuharM/Desktop/Python_Scripts/GsDlmsCmdRsp_logs/%s' % sub_dir): #Change Here
        line = line.rstrip()
        if 'logfile' in line:
            path = line.replace('\\', '/')
            with codecs.open('%s' % path, 'r', 'utf-8') as logfile:
                dt = datetime.datetime.fromtimestamp(os.stat('%s' % path).st_mtime)
                dt = dt.strftime('%Y-%m-%d')
                count = 0
                for line in logfile:
                    #print(line) #Uncomment this line to see actual error message
                    line = line.rstrip()
                    excellist = []
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
                print('%s__%s_%s_occurred %s times' % (dt, sub_dir, key_word, count))

                if count != 0:
                    print ('Between %s and %s' % (s_time[0], e_time[0]))
                    excellist.append(dt)
                    excellist.append(sub_dir)
                    excellist.append(count)
                    worksheet1.write('A' + str(cellnum), dt)
                    worksheet1.write('B' + str(cellnum), sub_dir)
                    worksheet1.write('C' + str(cellnum), count)
                    cellnum += 1
            d_sum.append(count)
        total = sum(d_sum)
    print('Total number of %s on %s is %s' % (key_word, sub_dir, total))

workbook.close()

#Displaying Graph
graph(exceldata, servers_name)




