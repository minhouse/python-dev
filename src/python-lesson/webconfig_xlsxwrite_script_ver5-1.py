#!/usr/bin/python
# coding:utf-8

#===============================================================================
#
#          FILE: webconfig_xlsxwrite_script_ver5-1.py
#
#         USAGE: ---
#
#   DESCRIPTION: This Python script is to extract <behaviors>, <bindings>, <appSettings> tags from web.config into Excel
# TARGET SYSTEM: System with Python ver3.x installed.
#       OPTIONS: ---
#  REQUIREMENTS: ---
#          BUGS: ---
#         NOTES: ---
#       CREATOR: Minho Tokuhara
#  ORGANIZATION: Landis+Gyr Japan
#       CREATED: 2015/07/13 ‏‎
#      REVISION: Code-Refactoring and minor change will be needed.
#===============================================================================

#Importing Modules
import os
import datetime
import os.path
import re
import codecs
import xlsxwriter

#Variables
day = datetime.date.today().strftime("%Y%m%d")
dir = os.listdir("C:/Users/TokuharM/Desktop/FT#11WebconfigFiles")

#Functions
def fild_all_files(directory):
    for root, dirs, files in os.walk(directory):
        yield root
        for file in files:
            yield os.path.join(root, file)

def add_header(sheet):
    global path
    global ValueCell_Format

    #Column Size
    sheet.set_column('A:A', 100)
    sheet.set_column('B:B', 25)
    sheet.set_column('C:C', 25)
    sheet.set_column('D:D', 25)

    #HeaderCell_Format
    fontsize = workbook.add_format({'bold': True})
    fontsize.set_font_size(18)
    HeaderCell_format = workbook.add_format({'bold': True})
    HeaderCell_format.set_pattern()
    HeaderCell_format.set_bg_color('#66CC33')
    HeaderCell_format.set_border(1)
    HeaderCell_format.set_border_color('black')

    #ValueCell_Format
    ValueCell_Format = workbook.add_format({'bold': False})
    ValueCell_Format.set_border(1)
    ValueCell_Format.set_border_color('black')
    ValueCell_Format.set_text_wrap()

    #Writing Header
    sheet.write('A1',  svname + "*.config (<behaviors>, <bindings>, <appSettings>)", fontsize)
    sheet.write('A3',  "Configuration ID No.", HeaderCell_format)
    sheet.write('A4',  svname + "-%s"% count, ValueCell_Format)
    sheet.write('A5',  "File Name and Path", HeaderCell_format)
    sheet.write('A6',  path, ValueCell_Format)
    sheet.write('A7',  "Key Name + Current Value", HeaderCell_format)
    sheet.write('B7',  "Recommended Value", HeaderCell_format)
    sheet.write('C7',  "Description", HeaderCell_format)
    sheet.write('D7',  "Timing to Change Value", HeaderCell_format)

#Main Procedure
for svname in dir:
    workbook = xlsxwriter.Workbook('C:/Users/TokuharM/Desktop/WebConfigList_FT11/%s_webconfigs_%s.xlsx' % (svname, day))

    count = 1
    for line in fild_all_files('C:/Users/TokuharM/Desktop/FT#11WebconfigFiles/%s' %svname):
        if 'web.conf' in line:
            path = line.replace('\\', '/')
            f_path = path.split('/')
            if len(f_path[-2]) > 31:
                ws_name = f_path[-2]
                ws_cut_string = ws_name[:30]
                worksheet1 = workbook.add_worksheet(ws_cut_string)
            else:
                ws_name = f_path[-2]
                worksheet1 = workbook.add_worksheet(ws_name)

            add_header(worksheet1)
            count += 1

            with codecs.open('%s' %path, 'r', 'utf_8') as webconf:
                cellnum = 8
                flag = 0
                for line in webconf:
                    line = line.rstrip()
                    if flag == 1:
                        worksheet1.write('A' + str(cellnum), line, ValueCell_Format)
                        worksheet1.write('B' + str(cellnum), "", ValueCell_Format)
                        worksheet1.write('C' + str(cellnum), "", ValueCell_Format)
                        worksheet1.write('D' + str(cellnum), "", ValueCell_Format)
                        cellnum += 1
                    if "<behaviors>" in line:
                        flag = 1
                        worksheet1.write('A' + str(cellnum), line, ValueCell_Format)
                        worksheet1.write('B' + str(cellnum), "", ValueCell_Format)
                        worksheet1.write('C' + str(cellnum), "", ValueCell_Format)
                        worksheet1.write('D' + str(cellnum), "", ValueCell_Format)
                        cellnum += 1
                    if "</behaviors>" in line:
                        cellnum2 = cellnum
                        break

            with codecs.open('%s' %path, 'r', 'utf_8') as webconf:
                flag = 0
                for line in webconf:
                    line = line.rstrip()
                    if flag == 1:
                        worksheet1.write('A' + str(cellnum2), line, ValueCell_Format)
                        worksheet1.write('B' + str(cellnum2), "", ValueCell_Format)
                        worksheet1.write('C' + str(cellnum2), "", ValueCell_Format)
                        worksheet1.write('D' + str(cellnum2), "", ValueCell_Format)
                        cellnum2 += 1
                    if "<bindings>" in line:
                        flag = 1
                        worksheet1.write('A' + str(cellnum2), line, ValueCell_Format)
                        worksheet1.write('B' + str(cellnum2), "", ValueCell_Format)
                        worksheet1.write('C' + str(cellnum2), "", ValueCell_Format)
                        worksheet1.write('D' + str(cellnum2), "", ValueCell_Format)
                        cellnum2 += 1
                    if "</bindings>" in line:
                        cellnum3 = cellnum2
                        break

            with codecs.open('%s' %path, 'r', 'utf_8') as webconf:
                flag = 0
                for line in webconf:
                    line = line.rstrip()
                    if flag == 1:
                        worksheet1.write('A' + str(cellnum3), line, ValueCell_Format)
                        worksheet1.write('B' + str(cellnum3), "", ValueCell_Format)
                        worksheet1.write('C' + str(cellnum3), "", ValueCell_Format)
                        worksheet1.write('D' + str(cellnum3), "", ValueCell_Format)
                        cellnum3 += 1
                    if "<appSettings>" in line:
                        flag = 1
                        worksheet1.write('A' + str(cellnum3), line, ValueCell_Format)
                        worksheet1.write('B' + str(cellnum3), "", ValueCell_Format)
                        worksheet1.write('C' + str(cellnum3), "", ValueCell_Format)
                        worksheet1.write('D' + str(cellnum3), "", ValueCell_Format)
                        cellnum3 += 1
                    if "</appSettings>" in line:
                        cellnum4 = cellnum3
                        break
    workbook.close()
