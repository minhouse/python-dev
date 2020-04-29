import sys, os, codecs, shutil, re, math, csv
import datetime as dt
import xlsxwriter
import pandas as pd
import numpy as np
import collections
import openpyxl
import seaborn as sns
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

#Please change file name to put _X number 'AllMeters_20181219_X.xlsx'
allmeter = 'ReadingPerf_20200420.csv'
crloc = 'Collector_Location20200224.xlsx'
full_smart_meters = pd.read_csv('C:/Users/TokuharM.AP/mywork/%s' % allmeter, sep=',',skiprows=1)
full_smart_meters.columns = ['meterno', 'serialnumber', 'endpointId', 'endpointtypeid', 'firmwareversion', 'endPointModelId', 'hwmodelid', 'date', 'initialDiscoveredDate', 'initialNormalDate', 'NoOfIntervals', 'name', 'abc_rank', 'DayEnd', 'meter_status', 'spuid', 'layer']

cr_list = pd.read_excel('C:/Users/TokuharM.AP/mywork/%s' % crloc, 'Sheet1', na_values=['NA'])
#cr_list = cr_list.drop(cr_list.columns[[4,5,6]], axis=1)

full_smart_meters.set_index('name').join(cr_list.set_index('CollectorNo'))
cr_list = full_smart_meters.join(cr_list.set_index('CollectorNo'), on='name', how='outer')
cr_list = cr_list.fillna({'Estates / Villages': 'Unlocated Area', 'BuildingType': 'Unknown BuildingType' })
cr_list = cr_list[cr_list['meterno'].notnull()]
cr_list = cr_list[~cr_list['abc_rank'].str.startswith('Load_DA')]
cr_village = cr_list[cr_list['BuildingType'].isin(['Village'])]
cr_highrise = cr_list[cr_list['BuildingType'].isin(['Highrise'])]
cr_unknownbuilding = cr_list[cr_list['BuildingType'].isin(['Unknown BuildingType'])]

#FW version performance
fw_avg = cr_list.pivot_table(values = ['NoOfIntervals'], index = ['firmwareversion'], aggfunc = {'NoOfIntervals': np.mean})
fw_std = cr_list.pivot_table(values = ['NoOfIntervals'], index = ['firmwareversion'], aggfunc = {'NoOfIntervals': np.std})
fw_perf = pd.concat([fw_avg, fw_std], axis=1, join_axes=[fw_avg.index])
fw_perf.columns = ['LP Average', 'LP Std Deviation']
fw_perf = fw_perf.round()

class District:
    def __init__(self, cr_list, attr):
        self.name = "{}".format(attr)
        self.district_meter = cr_list[cr_list['District'].str.contains(self.name, na=False)]
        self.district_meter_Count = self.district_meter['meterno'].count()
        self.district_meter_Full_48_LP_Interval = self.district_meter[self.district_meter['NoOfIntervals'] == 48]
        self.district_meter_Full_48_LP_Interval_Meter_Count = self.district_meter_Full_48_LP_Interval['meterno'].count()
        self.district_meter_Full_48_LP_Interval_Meter_Rate = round((self.district_meter_Full_48_LP_Interval_Meter_Count/self.district_meter_Count)*100,2)
        self.district_1468 = self.district_meter[self.district_meter['firmwareversion'].str.contains('-24.60', na=False)]
        self.district_1468_Count = self.district_1468['meterno'].count()
        self.district_1468_Rate = round((self.district_1468_Count/self.district_meter_Count)*100,2)
        self.district_meter_Normal_Meter = self.district_meter[self.district_meter['meter_status'] == 'Normal']
        self.district_meter_Normal_Meter_Count = self.district_meter_Normal_Meter['meterno'].count()
        self.district_meter_SecConfig_Meter = self.district_meter[self.district_meter['meter_status'] == 'SecConfig']
        self.district_meter_SecConfig_Meter_Count = self.district_meter_SecConfig_Meter['meterno'].count()
        self.district_meter_Discovered_Meter = self.district_meter[self.district_meter['meter_status'] == 'Discovered']
        self.district_meter_Discovered_Meter_Count = self.district_meter_Discovered_Meter['meterno'].count()
        self.district_meter_Config_Meter = self.district_meter[self.district_meter['meter_status'] == 'Configure']
        self.district_meter_Config_Meter_Count = self.district_meter_Config_Meter['meterno'].count()
        self.district_meter_Failed_Meter = self.district_meter[self.district_meter['meter_status'] == 'Failed']
        self.district_meter_Failed_Meter_Count = self.district_meter_Failed_Meter['meterno'].count()
        self.district_meter_Lost_Meter = self.district_meter[self.district_meter['meter_status'] == 'Lost']
        self.district_meter_Lost_Meter_Count = self.district_meter_Lost_Meter['meterno'].count()
        #LP-DayEnd-FULL_district Meter
        self.district_meter_LP_DayEnd_Full_Meter = self.district_meter[(self.district_meter['NoOfIntervals'] == 48) & (self.district_meter['DayEnd'] == 1)]
        self.district_meter_LP_DayEnd_Full_Meter_Count = self.district_meter_LP_DayEnd_Full_Meter['meterno'].count()
        self.district_meter_LP_DayEnd_Full_Meter_Rate = round((self.district_meter_LP_DayEnd_Full_Meter_Count/self.district_meter_Count)*100,2)
        self.district_meter_Missing_DayEnd_Reading = self.district_meter[self.district_meter['DayEnd'] != 1]
        self.district_meter_Missing_DayEnd_Reading_Meter_Count = self.district_meter_Missing_DayEnd_Reading['meterno'].count()
        self.Expected_district_meter_Total_LP_Count = ((self.district_meter_Count)*48)
        self.district_meter_Total_LP_Count = self.district_meter['NoOfIntervals'].sum()
        self.district_meter_Total_Dayend  = self.district_meter[self.district_meter['DayEnd'] == 1]
        self.district_meter_Total_Dayend_Count = self.district_meter_Total_Dayend['meterno'].count()
        self.district_meter_LP_Success_Rate = round((self.district_meter_Total_LP_Count/self.Expected_district_meter_Total_LP_Count)*100,2)
        self.district_meter_Dayend_Success_Rate  = round((self.district_meter_Total_Dayend_Count/self.district_meter_Count)*100,2)
        self.district_meter_Average_LP_Interval_Push_Count = self.district_meter['NoOfIntervals'].mean()
        self.district_meter_StdDev_LP_Interval_Push_Count = self.district_meter['NoOfIntervals'].std()
        #abc_rank
        self._CR_Rnk = self.district_meter.pivot_table(values = ['meter_status'], index = ['name'], columns = ['abc_rank'], aggfunc = 'count')
        #self._CR_Rnk.columns = self._CR_Rnk.columns.droplevel()
        self._CR_Rnk = self._CR_Rnk.loc[:,['P','A','B','C','D','E','F']]
        self._CR_Rnk = self._CR_Rnk.fillna(0)

    def get_dict(self):
        return collections.OrderedDict({
            '[ {} METERS SUMMARY ]'.format(self.name):'',
            '{} Meter Count'.format(self.name):self.district_meter_Count,
            '{} FW24.60 Meter Count'.format(self.name):self.district_1468_Count,
            '{} FW24.60 Meter(%)'.format(self.name):self.district_1468_Rate,
            '{} Meter LP Success(%)'.format(self.name):self.district_meter_LP_Success_Rate,
            '{} Meter Dayend Success(%)'.format(self.name):self.district_meter_Dayend_Success_Rate,
            '{} Average LP Push Count'.format(self.name):round(self.district_meter_Average_LP_Interval_Push_Count,2),
            '{} Std Deviation LP Push Count'.format(self.name):round(self.district_meter_StdDev_LP_Interval_Push_Count,2),
            '{} Meter LP-DayEnd-FULL Meter Count'.format(self.name):self.district_meter_LP_DayEnd_Full_Meter_Count,
            '{} Meter LP-DayEnd-FULL Meter(%)'.format(self.name):self.district_meter_LP_DayEnd_Full_Meter_Rate,
            '{} Meter Full 48 LP Interval Meter Count'.format(self.name):self.district_meter_Full_48_LP_Interval_Meter_Count,
            '{} Meter Full 48 LP Interval Meter(%)'.format(self.name):self.district_meter_Full_48_LP_Interval_Meter_Rate,
            '{} Meter Missing DayEnd Reading Meter Count'.format(self.name):self.district_meter_Missing_DayEnd_Reading_Meter_Count,
            '{} Meter Normal Meter Count'.format(self.name):self.district_meter_Normal_Meter_Count,
            '{} Meter SecConfig Meter Count'.format(self.name):self.district_meter_SecConfig_Meter_Count,
            '{} Meter Config Meter Count'.format(self.name):self.district_meter_Config_Meter_Count,
            '{} Meter Discovered Meter Count'.format(self.name):self.district_meter_Discovered_Meter_Count,
            '{} Meter Failed Meter Count'.format(self.name):self.district_meter_Failed_Meter_Count,
            '{} Meter Lost Meter Count'.format(self.name):self.district_meter_Lost_Meter_Count,
        })

class KeyPerformanceIndicator:
    def __init__(self, cr_list):
        self.Total_AllMeter_Count = cr_list['meterno'].count()
        self.RF_meter = cr_list[cr_list['endpointtypeid'] != 15]
        self.Collector_Count = self.RF_meter['name'].nunique()
        self.all_meter_1468 = cr_list[cr_list['firmwareversion'].str.contains('-24.60', na=False)]
        self.all_meter_1468_Count = self.all_meter_1468['meterno'].count()
        self.all_meter_1468_1468_Rate  = round((self.all_meter_1468_Count/self.Total_AllMeter_Count )*100,2)
        self.Total_HighRiseMeter_Count = cr_highrise['meterno'].count()
        self.Total_VillageMeter_Count = cr_village['meterno'].count()
        self.cell_meter = cr_list[cr_list['endpointtypeid'] == 15]
        self.onlycell_meter = self.cell_meter[~self.cell_meter['abc_rank'].str.startswith('Load_DA')]
        self.LDA_meter = self.cell_meter[self.cell_meter['abc_rank'].str.startswith('Load_DA')]
        self.Total_ALLCellMeter_Count = self.cell_meter['meterno'].count()  
        self.Total_LDAMeter_Count = self.LDA_meter['meterno'].count()
        self.Total_CellMeter_Count = self.Total_ALLCellMeter_Count - self.Total_LDAMeter_Count
        self.Total_CellMeter_Count = self.Total_ALLCellMeter_Count - self.Total_LDAMeter_Count
        self.unlocated_meter = cr_list[cr_list['Estates / Villages'] == 'Unlocated Area']
        self.unlocated_meter = self.unlocated_meter[self.unlocated_meter['meterno'].notnull()]
        self.unlocated_meter_Count = self.unlocated_meter['meterno'].count()
        self.unknownbuilding_Count = cr_unknownbuilding['meterno'].count()
        self.Expected_AllMeter_Total_LP_Count = (((self.Total_HighRiseMeter_Count + self.Total_VillageMeter_Count + self.unknownbuilding_Count) - self.Total_LDAMeter_Count)*48) + (self.Total_LDAMeter_Count*144)
        self.AllMeter_Total_LP_Count = cr_list['NoOfIntervals'].sum()
        self.AllMeter_Total_LP_SuccessRate = (self.AllMeter_Total_LP_Count/self.Expected_AllMeter_Total_LP_Count)*100
        self.Expected_AllMeter_Total_DayEnd_Reading_Count = (self.Total_HighRiseMeter_Count + self.Total_VillageMeter_Count + self.unknownbuilding_Count)
        self.AllMeter_Total_DayEnd_Reading_Count = cr_list['DayEnd'].sum()
        self.AllMeter_Total_DayEnd_Reading_SuccessRate = (self.AllMeter_Total_DayEnd_Reading_Count/self.Expected_AllMeter_Total_DayEnd_Reading_Count)*100
        self.Average_LP_Interval_Push_Count = cr_list['NoOfIntervals'].mean()
        self.StdDev_LP_Interval_Push_Count = cr_list['NoOfIntervals'].std()
        self.LP_DayEnd_Full_Meter = cr_list[(cr_list['NoOfIntervals'] == 48)&(cr_list['DayEnd'] == 1)]
        self.LP_DayEnd_Full_Meter_Count = self.LP_DayEnd_Full_Meter['meterno'].count()
        self.LP_DayEnd_Full_Meter_Rate = round((self.LP_DayEnd_Full_Meter_Count/self.Total_AllMeter_Count)*100,2)
        self.Full48_LP_Interval_AllMeter_Count = cr_list['NoOfIntervals'] == 48
        self.Full48_LP_Interval_AllMeter_Count = self.Full48_LP_Interval_AllMeter_Count.sum()
        self.Full48_LP_Interval_AllMeter_Rate = (self.Full48_LP_Interval_AllMeter_Count/self.Total_AllMeter_Count)*100
        self.Full48_LP_Interval_CellMeter_Count = self.onlycell_meter['NoOfIntervals'] == 48
        self.Full48_LP_Interval_CellMeter_Count = self.Full48_LP_Interval_CellMeter_Count.sum()
        self.Full48_LP_Interval_CellMeter_Count_Rate = round((self.Full48_LP_Interval_CellMeter_Count/self.Total_CellMeter_Count)*100,2)
        self.Missing_DayEnd_Reading_AllMeter_Count = self.Expected_AllMeter_Total_DayEnd_Reading_Count-self.AllMeter_Total_DayEnd_Reading_Count
        self.MissingDayEndReadingAllMeterRate = (self.Missing_DayEnd_Reading_AllMeter_Count/self.Expected_AllMeter_Total_DayEnd_Reading_Count)*100
        self.No_reading_meter = cr_list[cr_list['abc_rank'] == 'F']
        self.hexed_serial = pd.DataFrame(self.No_reading_meter['serialnumber'].astype(int))
        self.hexed_serial = self.hexed_serial.rename(columns={'serialnumber':'hex_serial'})
        self.hexed_serial = self.hexed_serial['hex_serial'].apply(lambda x:format(x, 'x'))
        self.No_reading_meter = pd.concat([self.No_reading_meter, self.hexed_serial], axis=1)
        self.No_reading_meter = self.No_reading_meter.reset_index(drop=True)
        self.No_LPandDayEnd_Reading_Meter_with_DayEnd = self.No_reading_meter[self.No_reading_meter['DayEnd'] == 0 ]
        self.No_LPandDayEnd_Reading_Meter_with_DayEnd_Count = self.No_LPandDayEnd_Reading_Meter_with_DayEnd['meterno'].count()
        self.No_LPandDayEnd_Reading_Meter_with_DayEnd_Rate = (self.No_LPandDayEnd_Reading_Meter_with_DayEnd_Count/self.Total_AllMeter_Count)*100
        self.No_Reading_Meter_Total_Count = self.No_reading_meter['abc_rank'].count()
        self.No_Reading_Meter_Rate = (self.No_Reading_Meter_Total_Count/self.Total_AllMeter_Count)*100
        self.No_Reading_Meter_with_DayEnd = self.No_reading_meter[self.No_reading_meter['DayEnd'] == 1 ]
        self.No_Reading_Meter_with_DayEnd_count = self.No_Reading_Meter_with_DayEnd['meterno'].count()
        self.NO_LPReading_ButWithDayEnd_Reading_Rate = (self.No_Reading_Meter_with_DayEnd_count/self.Total_AllMeter_Count)*100
        self.NO_DayEnd_Reading_but_with_LP_Reading_Meter = cr_list[cr_list['DayEnd'] == 0]
        self.NO_DayEnd_Reading_but_with_LP_Reading_Meter = self.NO_DayEnd_Reading_but_with_LP_Reading_Meter[self.NO_DayEnd_Reading_but_with_LP_Reading_Meter['NoOfIntervals'] != 0]
        self.NO_DayEnd_Reading_but_with_LP_Reading_Meter_Count = self.NO_DayEnd_Reading_but_with_LP_Reading_Meter['NoOfIntervals'].count()
        self.NO_DayEnd_Reading_but_with_LP_Reading_Meter_Rate = (self.NO_DayEnd_Reading_but_with_LP_Reading_Meter_Count/self.Total_AllMeter_Count)*100

    def get_dict_kpi(self):
        return collections.OrderedDict({
            '[ KEY PERFORMANCE INDICATOR ]':'',
            'Total Meter Count':self.Total_AllMeter_Count,
            'Total Collector Count':self.Collector_Count,
            'Total Meter FW24.60 Meter Count':self.all_meter_1468_Count,
            'Total Meter FW24.60 Meter(%)':self.all_meter_1468_1468_Rate,
            'All Meter LP Interval Push Success(%)':round(self.AllMeter_Total_LP_SuccessRate,2),
            'All Meter DayEnd Reading Push Success(%)':round(self.AllMeter_Total_DayEnd_Reading_SuccessRate,2),
            'Average LP Push Count':round(self.Average_LP_Interval_Push_Count,2),
            'Std Deviation LP Push Count':round(self.StdDev_LP_Interval_Push_Count,2),   
            'LP-DayEnd-FULL All Meter Count':self.LP_DayEnd_Full_Meter_Count,
            'LP-DayEnd-FULL All Meter(%)':round(self.LP_DayEnd_Full_Meter_Rate,2),
            'Full 48 LP Interval Meter Count':self.Full48_LP_Interval_AllMeter_Count,
            'Full 48 LP Interval Meter(%)':round(self.Full48_LP_Interval_AllMeter_Rate,2),
            'Full 48 LP Interval Cell Meter Count':self.Full48_LP_Interval_CellMeter_Count,
            'Full 48 LP Interval Cell Meter(%)':self.Full48_LP_Interval_CellMeter_Count_Rate,
            'NO DayEnd Reading All Meter Count':self.Missing_DayEnd_Reading_AllMeter_Count,
            'NO DayEnd Reading Meter(%)':round(self.MissingDayEndReadingAllMeterRate,2),
            'NO LP and DayEnd Reading Meter Count':self.No_LPandDayEnd_Reading_Meter_with_DayEnd_Count,
            'NO LP and DayEnd Reading Meter(%)':round(self.No_LPandDayEnd_Reading_Meter_with_DayEnd_Rate,2),
            'NO LP Reading Meter Count':self.No_Reading_Meter_Total_Count,
            'NO LP Reading Meter Total(%)':round(self.No_Reading_Meter_Rate,2),
            'NO LP Reading but with DayEnd Reading Meter Count':self.No_Reading_Meter_with_DayEnd_count,
            'NO LP Reading but with DayEnd_Reading Meter(%)':round(self.NO_LPReading_ButWithDayEnd_Reading_Rate,2),
            'NO DayEnd Reading but with LP Reading Meter Count':self.NO_DayEnd_Reading_but_with_LP_Reading_Meter_Count,
            'NO DayEnd Reading but with LP Reading Meter(%)':round(self.NO_DayEnd_Reading_but_with_LP_Reading_Meter_Rate,2),
        })

def main():
    district_list = list(
    [District(cr_list, 'district_a'),
    District(cr_list, 'district_b'),
    District(cr_list, 'district_c'),
    District(cr_list, 'district_d')])

    for district in district_list:
        print(district.get_dict())

    kpimetercount = KeyPerformanceIndicator(cr_list)
    print(kpimetercount.get_dict_kpi())

if __name__ == '__main__':
    main()