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

allmeter = 'ReadingPerf_20200420.csv'
crloc = 'Collector_Location20200224.xlsx'
full_smart_meters = pd.read_csv('C:/Users/TokuharM.AP/mywork/%s' % allmeter, sep=',',skiprows=1)
full_smart_meters.columns = ['meterno', 'serialnumber', 'endpointId', 'endpointtypeid', 'firmwareversion', 'endPointModelId', 'hwmodelid', 'date', 'initialDiscoveredDate', 'initialNormalDate', 'NoOfIntervals', 'name', 'abc_rank', 'DayEnd', 'meter_status', 'spuid', 'layer']

cr_list = pd.read_excel('C:/Users/TokuharM.AP/mywork/%s' % crloc, 'Sheet1', na_values=['NA'])
#cr_list = cr_list.drop(cr_list.columns[[4,5,6]], axis=1)

today_date = dt.date.today().strftime('%Y-%m-%d') 

full_smart_meters.set_index('name').join(cr_list.set_index('CollectorNo'))
cr_list = full_smart_meters.join(cr_list.set_index('CollectorNo'), on='name', how='outer')
cr_list = cr_list.fillna({'Estates / Villages': 'Unlocated Area', 'BuildingType': 'Unknown BuildingType' })
cr_list = cr_list[cr_list['meterno'].notnull()]
cr_list = cr_list[~cr_list['abc_rank'].str.startswith('Load_DA')]
cell_meter = cr_list[cr_list['endpointtypeid'] == 15]
cr_village = cr_list[cr_list['BuildingType'].isin(['Village'])]
cr_highrise = cr_list[cr_list['BuildingType'].isin(['Highrise'])]
onlycell_meter = cell_meter[~cell_meter['abc_rank'].str.startswith('Load_DA')]
LDA_meter = cell_meter[cell_meter['abc_rank'].str.startswith('Load_DA')]
unlocated_meter = cr_list[cr_list['Estates / Villages'] == 'Unlocated Area']
unlocated_meter = unlocated_meter[unlocated_meter['meterno'].notnull()]
cr_unknownbuilding = cr_list[cr_list['BuildingType'].isin(['Unknown BuildingType'])]
Normal_Meter = cr_list[cr_list['meter_status'] == 'Normal']
SecConfig_Meter = cr_list[cr_list['meter_status'] == 'SecConfig']
Discovered_Meter = cr_list[cr_list['meter_status'] == 'Discovered']
Config_Meter = cr_list[cr_list['meter_status'] == 'Configure']
Failed_Meter = cr_list[cr_list['meter_status'] == 'Failed']
Lost_Meter = cr_list[cr_list['meter_status'] == 'Lost']

#target_date = cr_list.iloc[0,7].strftime('%Y-%m-%d')

Total_AllMeter_Count = cr_list['meterno'].count()
Total_HighRiseMeter_Count = cr_highrise['meterno'].count()
Total_VillageMeter_Count = cr_village['meterno'].count()
Total_ALLCellMeter_Count = cell_meter['meterno'].count()
Total_LDAMeter_Count = LDA_meter['meterno'].count()
Total_CellMeter_Count = Total_ALLCellMeter_Count - Total_LDAMeter_Count
unlocated_meter_Count = unlocated_meter['meterno'].count()
unknownbuilding_Count = cr_unknownbuilding['meterno'].count()
#Meter Status Count
Normal_Meter_Count = Normal_Meter['meterno'].count()
SecConfig_Meter_Count = SecConfig_Meter['meterno'].count()
Config_Meter_Count = Config_Meter['meterno'].count()
Discovered_Meter_Count = Discovered_Meter['meterno'].count()
Failed_Meter_Count = Failed_Meter['meterno'].count()
Lost_Meter_Count = Lost_Meter['meterno'].count()
RF_meter = cr_list[cr_list['endpointtypeid'] != 15]
Collector_Count = RF_meter['name'].nunique()

No_reading_meter = cr_list[cr_list['abc_rank'] == 'F']
hexed_serial = pd.DataFrame(No_reading_meter['serialnumber'].astype(int))
hexed_serial = hexed_serial.rename(columns={'serialnumber':'hex_serial'})
hexed_serial = hexed_serial['hex_serial'].apply(lambda x:format(x, 'x'))
No_reading_meter = pd.concat([No_reading_meter, hexed_serial], axis=1)
No_reading_meter = No_reading_meter.reset_index(drop=True)

class DistrictPerformance:
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

        self.cr_rank = self.district_meter.pivot_table(values = ['meter_status'], index = ['name'], columns = ['abc_rank'], aggfunc = 'count')
        #self.cr_rank.columns = self.cr_rank.columns.droplevel()
        self.cr_rank = self.cr_rank.loc[:,['P','A','B','C','D','E','F']]
        self.cr_rank = self.cr_rank.fillna(0)

        self.cr_perf = cr_list[cr_list['name'].str.startswith('98020', na=False)]
        self.cr_perf_avg = self.cr_perf.groupby(['name'])['NoOfIntervals'].mean()
        self.cr_perf_std = self.cr_perf.groupby(['name'])['NoOfIntervals'].std()
        self.cr_perf = pd.concat([self.cr_perf_avg, self.cr_perf_std], axis=1, join_axes=[self.cr_perf_avg.index])
        self.cr_perf = self.cr_perf.round()
        self.cr_perf = self.cr_perf.fillna(0)
        self.cr_perf.columns = ['Average LP Count','Std LP Count']

        self.region_perf = cr_list.groupby(['Estates / Villages'])['NoOfIntervals'].mean()
        self.region_perf_std = cr_list.groupby(['Estates / Villages'])['NoOfIntervals'].std()
        self.region_perf = pd.concat([self.region_perf, self.region_perf_std], axis=1, join_axes=[self.region_perf.index])
        self.region_perf = self.region_perf.round()
        self.region_perf = self.region_perf.fillna(0)
        self.region_perf.columns = ['Average LP Count','Std LP Count']

        self.area_perf = cr_list.pivot_table(values = ['meter_status'], index = ['Estates / Villages'], columns = ['abc_rank'], aggfunc = 'count')
        self.area_perf.columns = self.area_perf.columns.droplevel()
        self.area_perf = self.area_perf.loc[:,['P','A','B','C','D','E','F']]
        self.area_perf = self.area_perf.fillna(0)

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

class CollectorPerformance(DistrictPerformance):
    def __init__(self, cr_list):
        super().__init__(cr_list, 'A') 

    def get_collector_statistics(self):
        return self.cr_perf

class KeyPerformanceIndicator:
    def __init__(self, Total_AllMeter_Count, Total_LDAMeter_Count, onlycell_meter, Total_HighRiseMeter_Count, Total_VillageMeter_Count, unknownbuilding_Count, Total_CellMeter_Count, cr_list):
        self.RF_meter = cr_list[cr_list['endpointtypeid'] != 15]
        self.Collector_Count = self.RF_meter['name'].nunique()
        self.all_meter_1468 = cr_list[cr_list['firmwareversion'].str.contains('-24.60', na=False)]
        self.all_meter_1468_Count = self.all_meter_1468['meterno'].count()
        self.all_meter_1468_1468_Rate  = round((self.all_meter_1468_Count/Total_AllMeter_Count)*100,2)
        self.Expected_AllMeter_Total_LP_Count = (((Total_HighRiseMeter_Count + Total_VillageMeter_Count + unknownbuilding_Count) - Total_LDAMeter_Count)*48) + (Total_LDAMeter_Count*144)
        self.AllMeter_Total_LP_Count = cr_list['NoOfIntervals'].sum()
        self.AllMeter_Total_LP_SuccessRate = (self.AllMeter_Total_LP_Count/self.Expected_AllMeter_Total_LP_Count)*100
        self.Expected_AllMeter_Total_DayEnd_Reading_Count = (Total_HighRiseMeter_Count + Total_VillageMeter_Count + unknownbuilding_Count)
        self.AllMeter_Total_DayEnd_Reading_Count = cr_list['DayEnd'].sum()
        self.AllMeter_Total_DayEnd_Reading_SuccessRate = (self.AllMeter_Total_DayEnd_Reading_Count/self.Expected_AllMeter_Total_DayEnd_Reading_Count)*100
        self.Average_LP_Interval_Push_Count = cr_list['NoOfIntervals'].mean()
        self.StdDev_LP_Interval_Push_Count = cr_list['NoOfIntervals'].std()
        self.LP_DayEnd_Full_Meter = cr_list[(cr_list['NoOfIntervals'] == 48)&(cr_list['DayEnd'] == 1)]
        self.LP_DayEnd_Full_Meter_Count = self.LP_DayEnd_Full_Meter['meterno'].count()
        self.LP_DayEnd_Full_Meter_Rate = round((self.LP_DayEnd_Full_Meter_Count/Total_AllMeter_Count)*100,2)
        self.Full48_LP_Interval_AllMeter_Count = cr_list['NoOfIntervals'] == 48
        self.Full48_LP_Interval_AllMeter_Count = self.Full48_LP_Interval_AllMeter_Count.sum()
        self.Full48_LP_Interval_AllMeter_Rate = (self.Full48_LP_Interval_AllMeter_Count/Total_AllMeter_Count)*100
        self.Full48_LP_Interval_CellMeter_Count = onlycell_meter['NoOfIntervals'] == 48
        self.Full48_LP_Interval_CellMeter_Count = self.Full48_LP_Interval_CellMeter_Count.sum()
        self.Full48_LP_Interval_CellMeter_Count_Rate = round((self.Full48_LP_Interval_CellMeter_Count/Total_CellMeter_Count)*100,2)
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
        self.No_LPandDayEnd_Reading_Meter_with_DayEnd_Rate = (self.No_LPandDayEnd_Reading_Meter_with_DayEnd_Count/Total_AllMeter_Count)*100
        self.No_Reading_Meter_Total_Count = self.No_reading_meter['abc_rank'].count()
        self.No_Reading_Meter_Rate = (self.No_Reading_Meter_Total_Count/Total_AllMeter_Count)*100
        self.No_Reading_Meter_with_DayEnd = self.No_reading_meter[self.No_reading_meter['DayEnd'] == 1 ]
        self.No_Reading_Meter_with_DayEnd_count = self.No_Reading_Meter_with_DayEnd['meterno'].count()
        self.NO_LPReading_ButWithDayEnd_Reading_Rate = (self.No_Reading_Meter_with_DayEnd_count/Total_AllMeter_Count)*100
        self.NO_DayEnd_Reading_but_with_LP_Reading_Meter = cr_list[cr_list['DayEnd'] == 0]
        self.NO_DayEnd_Reading_but_with_LP_Reading_Meter = self.NO_DayEnd_Reading_but_with_LP_Reading_Meter[self.NO_DayEnd_Reading_but_with_LP_Reading_Meter['NoOfIntervals'] != 0]
        self.NO_DayEnd_Reading_but_with_LP_Reading_Meter_Count = self.NO_DayEnd_Reading_but_with_LP_Reading_Meter['NoOfIntervals'].count()
        self.NO_DayEnd_Reading_but_with_LP_Reading_Meter_Rate = (self.NO_DayEnd_Reading_but_with_LP_Reading_Meter_Count/Total_AllMeter_Count)*100

    def get_dict_kpi(self):
        return collections.OrderedDict({
            '[ KEY PERFORMANCE INDICATOR ]':'',
            'Total Meter Count':Total_AllMeter_Count,
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

class SLAPerformance:
    def __init__(self, cr_list):
        self.today_date = dt.date.today().strftime('%Y-%m-%d') 
        cr_list['initialNormalDate'] = pd.to_datetime(cr_list['initialNormalDate'], format='%Y-%m-%d')
        cr_list['date'] = cr_list['date'].fillna(self.today_date)
        cr_list['date'] = pd.to_datetime(cr_list['date'], format='%Y-%m-%d')
        cr_list['7Days_After_Normal'] = (cr_list['initialNormalDate']  + dt.timedelta(days=7))
        cr_list['initialNormalDate'] = cr_list['initialNormalDate'].values.astype('datetime64[D]')
        cr_list['7Days_After_Normal'] = cr_list['7Days_After_Normal'].values.astype('datetime64[D]')
        cr_list['initialDiscoveredDate'] = cr_list['initialDiscoveredDate'].values.astype('datetime64[D]')
        cr_list['Difference'] = cr_list['date'] - cr_list['initialNormalDate']
        cr_list['DaysAfterDis'] = cr_list['date'] - cr_list['initialDiscoveredDate']
        cr_list['DisToNorm'] = cr_list['initialNormalDate'] - cr_list['initialDiscoveredDate']
        self.Effective_Meter = cr_list[cr_list['Difference'] >= '7 days']
        self.Effective_Meter = cr_list[(cr_list['meter_status'] == 'Normal')]
        self.Effective_Meter_Count = self.Effective_Meter['meterno'].count()
        self.EffectiveMeters_Full_48_LP_Interval = self.Effective_Meter[self.Effective_Meter['NoOfIntervals'] == 48]
        self.EffectiveMeters_Full_48_LP_Interval_Meter_Count = self.EffectiveMeters_Full_48_LP_Interval['meterno'].count()
        self.EffectiveMeters_Full_48_LP_Interval_Meter_Rate = round((self.EffectiveMeters_Full_48_LP_Interval_Meter_Count/self.Effective_Meter_Count)*100,2)
        self.LP_DayEnd_Full_Effective_Meter = self.Effective_Meter[(self.Effective_Meter['NoOfIntervals'] == 48)&(self.Effective_Meter['DayEnd'] == 1)]
        self.LP_DayEnd_Full_Effective_Meter_Count = self.LP_DayEnd_Full_Effective_Meter['meterno'].count()
        self.LP_DayEnd_Full_Effective_Meter_Rate = round((self.LP_DayEnd_Full_Effective_Meter_Count/self.Effective_Meter_Count)*100,2)
        self.EffectiveMeters_Missing_DayEnd_Reading = self.Effective_Meter[self.Effective_Meter['DayEnd'] != 1]
        self.EffectiveMeters_Missing_DayEnd_Reading_Meter_Count = self.EffectiveMeters_Missing_DayEnd_Reading['meterno'].count()
        self.EffectiveMeters_Missing_DayEnd_Reading_Meter_Rate = round((self.EffectiveMeters_Missing_DayEnd_Reading_Meter_Count/self.Effective_Meter_Count)*100,2)
        self.Expected_EffectiveMeter_Total_LP_Count = ((self.Effective_Meter_Count)*48)
        self.EffectiveMeter_Total_LP_Count = self.Effective_Meter['NoOfIntervals'].sum()
        self.EffectiveMeter_Total_Dayend  = self.Effective_Meter[self.Effective_Meter['DayEnd'] == 1]
        self.EffectiveMeter_Total_Dayend_Count = self.EffectiveMeter_Total_Dayend['meterno'].count()
        self.Effective_Meter_LP_Success_Rate = round((self.EffectiveMeter_Total_LP_Count/self.Expected_EffectiveMeter_Total_LP_Count)*100,2)
        self.Effective_Meter_Dayend_Success_Rate  = round((self.EffectiveMeter_Total_Dayend_Count/self.Effective_Meter_Count)*100,2)
        self.Effective_Meter_Average_LP_Interval_Push_Count = self.Effective_Meter['NoOfIntervals'].mean()
        self.Effective_Meter_StdDev_LP_Interval_Push_Count = self.Effective_Meter['NoOfIntervals'].std()
    
    def get_dict_sla(self):
        return collections.OrderedDict({
        '[ SLA METERS PERFORMANCE (NORMAL FOR OVER 7DAYS) ]':'',
            'SLA Meter Count':self.Effective_Meter_Count,
            'SLA Meter LP Success(%)':self.Effective_Meter_LP_Success_Rate,
            'SLA Meter Dayend Success(%)':self.Effective_Meter_Dayend_Success_Rate,
            'SLA Meter Average LP Push Count':round(self.Effective_Meter_Average_LP_Interval_Push_Count,2),
            'SLA Meter Std Deviation LP Push Count':round(self.Effective_Meter_StdDev_LP_Interval_Push_Count,2),
            'LP-DayEnd-FULL SLA Meter Count':self.LP_DayEnd_Full_Effective_Meter_Count,
            'LP-DayEnd-FULL SLA Meter(%)':self.LP_DayEnd_Full_Effective_Meter_Rate,
            'SLA Meters Full 48 LP Interval Meter Count':self.EffectiveMeters_Full_48_LP_Interval_Meter_Count,
            'SLA Meters Full 48 LP Interval Meter(%)':self.EffectiveMeters_Full_48_LP_Interval_Meter_Rate,
            'SLA Meters Missing DayEnd Reading Meter Count':self.EffectiveMeters_Missing_DayEnd_Reading_Meter_Count,
            'SLA Meters Missing DayEnd Reading Meter(%)':self.EffectiveMeters_Missing_DayEnd_Reading_Meter_Rate,
        })

class LatestMeterPerformance:
    def __init__(self, cr_list):
        self.Latest_Meters = cr_list[cr_list['DaysAfterDis']  < '30 days']
        self.Latest_Meters_Count = self.Latest_Meters['meterno'].count()
        self.Latest_Meters_Full_48_LP_Interval = self.Latest_Meters[self.Latest_Meters['NoOfIntervals'] == 48]
        self.Latest_Meters_Full_48_LP_Interval_Meter_Count = self.Latest_Meters_Full_48_LP_Interval['meterno'].count()
        self.Latest_Meters_Full_48_LP_Interval_Meter_Rate = round((self.Latest_Meters_Full_48_LP_Interval_Meter_Count/self.Latest_Meters_Count)*100,2)
        self.LP_DayEnd_Full_Latest_Meters = self.Latest_Meters[(self.Latest_Meters['NoOfIntervals'] == 48)&(self.Latest_Meters['DayEnd'] == 1)]
        self.LP_DayEnd_Full_Latest_Meters_Count = self.LP_DayEnd_Full_Latest_Meters['meterno'].count()
        self.LP_DayEnd_Full_Latest_Meters_Rate = round((self.LP_DayEnd_Full_Latest_Meters_Count/self.Latest_Meters_Count)*100,2)
        self.Latest_Meters_Missing_DayEnd_Reading = self.Latest_Meters[self.Latest_Meters['DayEnd'] != 1]
        self.Latest_Meters_Missing_DayEnd_Reading_Meter_Count = self.Latest_Meters_Missing_DayEnd_Reading['meterno'].count()
        self.Latest_Meters_Missing_DayEnd_Reading_Meter_Rate = round((self.Latest_Meters_Missing_DayEnd_Reading_Meter_Count/self.Latest_Meters_Count)*100,2)
        self.Expected_Latest_Meters_Count_Total_LP_Count = ((self.Latest_Meters_Count)*48)
        self.Latest_Meters_Total_LP_Count = self.Latest_Meters['NoOfIntervals'].sum()
        self.Latest_Meters_Total_Dayend  = self.Latest_Meters[self.Latest_Meters['DayEnd'] == 1]
        self.Latest_Meters_Total_Dayend_Count = self.Latest_Meters_Total_Dayend['meterno'].count()
        self.Latest_Meters_LP_Success_Rate = round((self.Latest_Meters_Total_LP_Count/self.Expected_Latest_Meters_Count_Total_LP_Count)*100,2)
        self.Latest_Meters_Dayend_Success_Rate  = round((self.Latest_Meters_Total_Dayend_Count/self.Latest_Meters_Count)*100,2)
        self.Latest_Meters_Average_LP_Interval_Push_Count = self.Latest_Meters['NoOfIntervals'].mean()
        self.Latest_Meters_StdDev_LP_Interval_Push_Count = self.Latest_Meters['NoOfIntervals'].std()

    def get_dict_latest(self):
        return collections.OrderedDict({
            '[ LATEST METERS PERFORMANCE (REGISTERED IN LAST 30DAYS) ]':'',
            'Latest Meter Count':self.Latest_Meters_Count,
            'Latest Meter LP Success(%)':self.Latest_Meters_LP_Success_Rate,
            'Latest Meter Dayend Success(%)':self.Latest_Meters_Dayend_Success_Rate,
            'Latest Meter Average LP Push Count':round(self.Latest_Meters_Average_LP_Interval_Push_Count,2),
            'Latest Meter Std Deviation LP Push Count':round(self.Latest_Meters_StdDev_LP_Interval_Push_Count,2),
            'Latest Meters Full 48 LP Interval Meter Count':self.Latest_Meters_Full_48_LP_Interval_Meter_Count,
            'Latest Meters Full 48 LP Interval Meter(%)':self.Latest_Meters_Full_48_LP_Interval_Meter_Rate,
            'Latest Meters Missing DayEnd Reading Meter Count':self.Latest_Meters_Missing_DayEnd_Reading_Meter_Count,
            'Latest Meters Missing DayEnd Reading Meter(%)':self.Latest_Meters_Missing_DayEnd_Reading_Meter_Rate,
        })

class AllMetersCount:
    def get_dict_allmetercount(self, Total_ALLCellMeter_Count, Total_HighRiseMeter_Count, Total_VillageMeter_Count, Total_CellMeter_Count, Total_LDAMeter_Count, unlocated_meter_Count):
        return collections.OrderedDict({
            '[ OVERALL METERS COUNT SUMMARY ]':'',
            'Total HighRise Meter Count':Total_HighRiseMeter_Count,
            'Total Village Meter Count':Total_VillageMeter_Count,
            'Total Cell Meter Count':Total_CellMeter_Count,
            'Total All Cell Type Meter Count':Total_ALLCellMeter_Count,
            'Total LDA Meter Count':Total_LDAMeter_Count,
            'Unlocated Meter Count':unlocated_meter_Count,
        })

class AllLPIntervalPushSuccessRate:
    def __init__(self, Total_HighRiseMeter_Count, Total_VillageMeter_Count, Total_CellMeter_Count, Total_LDAMeter_Count, cr_highrise, cr_village, onlycell_meter, LDA_meter):
        self.Expected_HighRiseMeter_Total_LP_Count = Total_HighRiseMeter_Count*48
        self.HighRiseMeter_Total_LP_Count = cr_highrise['NoOfIntervals'].sum()
        self.Expected_VillageMeter_Total_LP_Count = Total_VillageMeter_Count*48
        self.VillageMeter_Total_LP_Count = cr_village['NoOfIntervals'].sum()
        self.Expected_AllCellMeter_Total_LP_Count = (Total_CellMeter_Count*48+Total_LDAMeter_Count*144)
        self.AllCellMeter_Total_LP_Count = (onlycell_meter['NoOfIntervals'].sum()+LDA_meter['NoOfIntervals'].sum())
        self.Expected_CellMeter_Total_LP_Count = (Total_CellMeter_Count*48)
        self.CellMeter_Total_LP_Count = (onlycell_meter['NoOfIntervals'].sum())
        self.Expected_LDAMeter_Total_LP_Count = (Total_LDAMeter_Count*144)
        self.LDAMeter_Total_LP_Count = (LDA_meter['NoOfIntervals'].sum())
        self.HighRiseMeter_Total_LP_SuccessRate = (self.HighRiseMeter_Total_LP_Count/self.Expected_HighRiseMeter_Total_LP_Count)*100
        self.VillageMeter_Total_LP_SuccessRate = (self.VillageMeter_Total_LP_Count/self.Expected_VillageMeter_Total_LP_Count)*100
        self.AllCellMeter_Total_LP_SuccessRate = (self.AllCellMeter_Total_LP_Count/self.Expected_AllCellMeter_Total_LP_Count)*100
        self.CellMeter_Total_LP_SuccessRate = (self.CellMeter_Total_LP_Count/self.Expected_CellMeter_Total_LP_Count)*100
        self.LDAMeter_Total_LP_SuccessRate = (self.LDAMeter_Total_LP_Count/self.Expected_LDAMeter_Total_LP_Count)*100
    
    def get_dict_alllpintervalpushsuccessrate(self):
        return collections.OrderedDict({
            '[ OVERALL LP INTERVAL PUSH SUCCESS % ]':'',
            'HighRise Meter Total LP Interval Push Success(%)':round(self.HighRiseMeter_Total_LP_SuccessRate,2),
            'Village Meter Total LP Interval Push Success(%)':round(self.VillageMeter_Total_LP_SuccessRate,2),
            'All Cell Meter Total LP Interval Push Success(%)':round(self.AllCellMeter_Total_LP_SuccessRate,2),
            'Cell Meter Total LP Interval Push Success(%)':round(self.CellMeter_Total_LP_SuccessRate,2),
            'LDA Meter Total LP Interval Push Success(%)':round(self.LDAMeter_Total_LP_SuccessRate,2),
        })

class AllDayendPushSuccessRate:
    def __init__(self, Total_HighRiseMeter_Count, cr_highrise, Total_VillageMeter_Count, cr_village, Total_CellMeter_Count, Total_LDAMeter_Count, onlycell_meter, LDA_meter):
        self.Expected_HighRiseMeter_Total_DayEnd_Reading_Count = Total_HighRiseMeter_Count
        self.HighRiseMeter_Total_DayEnd_Reading_Count = cr_highrise['DayEnd'].sum()
        self.Expected_VillageMeter_Total_DayEnd_Reading_Count = Total_VillageMeter_Count
        self.VillageMeter_Total_DayEnd_Reading_Count = cr_village['DayEnd'].sum()
        self.Expected_AllCellMeter_Total_DayEnd_Reading_Count = (Total_CellMeter_Count+Total_LDAMeter_Count)
        self.AllCellMeter_Total_DayEnd_Reading_Count = (onlycell_meter['DayEnd'].sum() + LDA_meter['DayEnd'].sum())
        self.Expected_CellMeter_Total_DayEnd_Reading_Count = Total_CellMeter_Count
        self.CellMeter_Total_DayEnd_Reading_Count = (onlycell_meter['DayEnd'].sum())
        self.Expected_LDAMeter_Total_DayEnd_Reading_Count = (Total_LDAMeter_Count)
        self.LDAMeter_Total_DayEnd_Reading_Count = (LDA_meter['DayEnd'].sum())
        self.HighRiseMeter_Total_DayEnd_Reading_SuccessRate = (self.HighRiseMeter_Total_DayEnd_Reading_Count/self.Expected_HighRiseMeter_Total_DayEnd_Reading_Count)*100
        self.VillageMeter_Total_DayEnd_Reading_SuccessRate = (self.VillageMeter_Total_DayEnd_Reading_Count/self.Expected_VillageMeter_Total_DayEnd_Reading_Count)*100
        self.AllCellMeter_Total_DayEnd_Reading_SuccessRate = (self.AllCellMeter_Total_DayEnd_Reading_Count/self.Expected_AllCellMeter_Total_DayEnd_Reading_Count)*100
        self.CellMeter_Total_DayEnd_Reading_SuccessRate = (self.CellMeter_Total_DayEnd_Reading_Count/self.Expected_CellMeter_Total_DayEnd_Reading_Count)*100
        self.LDAMeter_Total_DayEnd_Reading_SuccessRate = (self.LDAMeter_Total_DayEnd_Reading_Count/self.Expected_LDAMeter_Total_DayEnd_Reading_Count)*100

    def get_dict_alldayendpushsuccessrate(self):
        return collections.OrderedDict({
            '[ OVERALL DAYEND READING PUSH SUCCESS % ]':'',
            'HighRise Meter Total DayEnd Reading Push Success(%)':round(self.HighRiseMeter_Total_DayEnd_Reading_SuccessRate,2),
            'Village Meter Total DayEnd Reading Push Success(%)':round(self.VillageMeter_Total_DayEnd_Reading_SuccessRate,2),
            'All Cell Meter Total DayEnd Reading Push Success(%)':round(self.AllCellMeter_Total_DayEnd_Reading_SuccessRate,2),
            'Cell Meter Total DayEnd Reading Push Success(%)':round(self.CellMeter_Total_DayEnd_Reading_SuccessRate,2),
            'LDA Meter Total DayEnd Reading Push Success(%)':round(self.LDAMeter_Total_DayEnd_Reading_SuccessRate,2),
        })

class NoLpPushMeterSummary():
    def __init__(self, Total_AllMeter_Count, cr_list):
        self.No_reading_meter = cr_list[cr_list['abc_rank'] == 'F']
        self.hexed_serial = pd.DataFrame(self.No_reading_meter['serialnumber'].astype(int))
        self.hexed_serial = self.hexed_serial.rename(columns={'serialnumber':'hex_serial'})
        self.hexed_serial = self.hexed_serial['hex_serial'].apply(lambda x:format(x, 'x'))
        self.No_reading_meter = pd.concat([self.No_reading_meter, self.hexed_serial], axis=1)
        self.No_reading_meter = self.No_reading_meter.reset_index(drop=True)
        self.No_Reading_RF_meter = self.No_reading_meter[self.No_reading_meter['endpointtypeid'] == 9]
        self.No_Reading_cell_meter = self.No_reading_meter[self.No_reading_meter['endpointtypeid'] == 15]
        self.No_Reading_Normal_Meter = self.No_reading_meter[self.No_reading_meter['meter_status'] == 'Normal']
        self.No_Reading_SecConfig_Meter = self.No_reading_meter[self.No_reading_meter['meter_status'] == 'SecConfig']
        self.No_Reading_Discovered_Meter = self.No_reading_meter[self.No_reading_meter['meter_status'] == 'Discovered']
        self.No_Reading_Config_Meter = self.No_reading_meter[self.No_reading_meter['meter_status'] == 'Configure']
        self.No_Reading_Failed_Meter = self.No_reading_meter[self.No_reading_meter['meter_status'] == 'Failed']
        self.No_Reading_Lost_Meter = self.No_reading_meter[self.No_reading_meter['meter_status'] == 'Lost']
        self.No_Reading_Meter_with_DayEnd = self.No_reading_meter[self.No_reading_meter['DayEnd'] == 1 ]
        self.No_LPandDayEnd_Reading_Meter_with_DayEnd = self.No_reading_meter[self.No_reading_meter['DayEnd'] == 0 ]
        self.No_reading_meter_Highrise = self.No_reading_meter[self.No_reading_meter['BuildingType'].isin(['Highrise'])]
        self.No_reading_meter_Village = self.No_reading_meter[self.No_reading_meter['BuildingType'].isin(['Village'])]
        self.No_reading_meter_Unlocated = self.No_reading_meter[self.No_reading_meter['BuildingType'].isin(['Unknown BuildingType'])]
        self.No_Reading_Meter_Total_Count = self.No_reading_meter['abc_rank'].count()
        self.No_Reading_RF_meter_Count = self.No_Reading_RF_meter['meterno'].count()
        self.No_Reading_Cell_meter_Count = self.No_Reading_cell_meter['meterno'].count()
        self.No_reading_Normal_meter_count = self.No_Reading_Normal_Meter['meterno'].count()
        self.No_reading_SecConfig_meter_count = self.No_Reading_SecConfig_Meter['meterno'].count()
        self.No_reading_Discovered_meter_count = self.No_Reading_Discovered_Meter['meterno'].count()
        self.No_reading_Config_meter_count = self.No_Reading_Config_Meter['meterno'].count()
        self.No_reading_Failed_meter_count = self.No_Reading_Failed_Meter['meterno'].count()
        self.No_reading_Lost_meter_count = self.No_Reading_Lost_Meter['meterno'].count()
        self.No_Reading_Meter_with_DayEnd_count = self.No_Reading_Meter_with_DayEnd['meterno'].count()
        self.No_LPandDayEnd_Reading_Meter_with_DayEnd_Count = self.No_LPandDayEnd_Reading_Meter_with_DayEnd['meterno'].count()
        self.No_reading_meter_Highrise_count = self.No_reading_meter_Highrise['abc_rank'].count()
        self.No_reading_meter_Village_count = self.No_reading_meter_Village['abc_rank'].count()
        self.No_reading_meter_Unlocated_count = self.No_reading_meter_Unlocated['abc_rank'].count()
        self.No_Reading_Meter_Rate = (self.No_Reading_Meter_Total_Count/Total_AllMeter_Count)*100
        self.No_Reading_Meter_Highrise_Rate = (self.No_reading_meter_Highrise_count/self.No_Reading_Meter_Total_Count)*100
        self.No_Reading_Meter_Village_Rate = (self.No_reading_meter_Village_count/self.No_Reading_Meter_Total_Count)*100
        self.No_reading_meter_Unlocated_Rate = (self.No_reading_meter_Unlocated_count/self.No_Reading_Meter_Total_Count)*100
        self.No_Reading_Normal_Meter_Rate = (self.No_reading_Normal_meter_count/self.No_Reading_Meter_Total_Count)*100
        self.No_Reading_SecConfig_Meter_Rate = (self.No_reading_SecConfig_meter_count/self.No_Reading_Meter_Total_Count)*100
        self.No_Reading_Config_Meter_Rate = (self.No_reading_Config_meter_count/self.No_Reading_Meter_Total_Count)*100
        self.No_Reading_Discovered_Meter_Rate = (self.No_reading_Discovered_meter_count/self.No_Reading_Meter_Total_Count)*100
        self.No_Reading_Failed_Meter_Rate = (self.No_reading_Failed_meter_count/self.No_Reading_Meter_Total_Count)*100
        self.No_Reading_Lost_Meter_Rate = (self.No_reading_Lost_meter_count/self.No_Reading_Meter_Total_Count)*100

    def get_dict_allnolppushmetersummary(self):
        return collections.OrderedDict({
            '[ OVERALL NO LP READING METERS SUMMARY ]':'',
            'NO LP Reading Highrise Meter Count':self.No_reading_meter_Highrise_count,
            'NO LP Reading Village Meter Count':self.No_reading_meter_Village_count,
            'NO LP Reading Unlocated Meter Count':self.No_reading_meter_Unlocated_count,
            'NO LP Reading RF Meter Count':self.No_Reading_RF_meter_Count,
            'NO LP Reading Cell Meter Count':self.No_Reading_Cell_meter_Count,
            'NO LP Reading Normal Meter Count':self.No_reading_Normal_meter_count,
            'NO LP Reading SecConfig Meter Count':self.No_reading_SecConfig_meter_count,
            'NO LP Reading Config Meter Count':self.No_reading_Config_meter_count,
            'NO LP Reading Discovered Meter Count':self.No_reading_Discovered_meter_count,
            'NO LP Reading Failed Meter Count':self.No_reading_Failed_meter_count,
            'NO LP Reading Lost Meter Count':self.No_reading_Lost_meter_count,
            '[ NO LP PUSH READING METER COMPOSITION RATE ]':'',
            'NO LP Reading HighRise Meter(%)':round(self.No_Reading_Meter_Highrise_Rate,2),
            'NO LP Reading Village Meter(%)':round(self.No_Reading_Meter_Village_Rate,2),
            'NO LP Reading Unlocated Meter(%)':round(self.No_reading_meter_Unlocated_Rate,2),
            'NO LP Reading Normal Meter(%)':round(self.No_Reading_Normal_Meter_Rate,2),
            'NO LP Reading SecConfig Meter(%)':round(self.No_Reading_SecConfig_Meter_Rate,2),
            'NO LP Reading Configure Meter(%)':round(self.No_Reading_Config_Meter_Rate,2),
            'NO LP Reading Discovered Meter(%)':round(self.No_Reading_Discovered_Meter_Rate,2),
            'NO LP Reading Failed Meter(%)':round(self.No_Reading_Failed_Meter_Rate,2),
            'NO LP Reading Lost Meter(%)':round(self.No_Reading_Lost_Meter_Rate,2)
        })

class MeterStatusCount:
    def __init__(self, cr_list):
        self.Normal_Meter = cr_list[cr_list['meter_status'] == 'Normal']
        self.SecConfig_Meter = cr_list[cr_list['meter_status'] == 'SecConfig']
        self.Discovered_Meter = cr_list[cr_list['meter_status'] == 'Discovered']
        self.Config_Meter = cr_list[cr_list['meter_status'] == 'Configure']
        self.Failed_Meter = cr_list[cr_list['meter_status'] == 'Failed']
        self.Lost_Meter = cr_list[cr_list['meter_status'] == 'Lost']
        self.Normal_Meter_Count = self.Normal_Meter['meterno'].count()
        self.SecConfig_Meter_Count = self.SecConfig_Meter['meterno'].count()
        self.Config_Meter_Count = self.Config_Meter['meterno'].count()
        self.Discovered_Meter_Count = self.Discovered_Meter['meterno'].count()
        self.Failed_Meter_Count = self.Failed_Meter['meterno'].count()
        self.Lost_Meter_Count = self.Lost_Meter['meterno'].count()

    def get_dict_meterstatuscount(self):
        return collections.OrderedDict({
            '[ METER STATUS COUNT WITH READINGS ]':'',
            'Normal Status Meter Count':self.Normal_Meter_Count,
            'SecConfig Status Meter Count':self.SecConfig_Meter_Count,
            'Configure Status Meter Count':self.Config_Meter_Count,
            'Discovered Status Meter Count':self.Discovered_Meter_Count,
            'Failed Status Meter Count':self.Failed_Meter_Count,
            'Lost Status Meter Count':self.Lost_Meter_Count
        })

class AllLpPushCountPerformance:
    def __init__(self, cr_list, Total_HighRiseMeter_Count, Total_VillageMeter_Count, unknownbuilding_Count, Total_LDAMeter_Count, cr_highrise, cr_village, Total_CellMeter_Count, onlycell_meter, LDA_meter):
        self.Expected_AllMeter_Total_LP_Count = (((Total_HighRiseMeter_Count+Total_VillageMeter_Count+unknownbuilding_Count)-Total_LDAMeter_Count)*48)+(Total_LDAMeter_Count*144)
        self.AllMeter_Total_LP_Count = cr_list['NoOfIntervals'].sum()
        self.Expected_HighRiseMeter_Total_LP_Count = Total_HighRiseMeter_Count*48
        self.HighRiseMeter_Total_LP_Count = cr_highrise['NoOfIntervals'].sum()
        self.Expected_VillageMeter_Total_LP_Count = Total_VillageMeter_Count*48
        self.VillageMeter_Total_LP_Count = cr_village['NoOfIntervals'].sum()
        self.Expected_AllCellMeter_Total_LP_Count = (Total_CellMeter_Count*48+Total_LDAMeter_Count*144)
        self.AllCellMeter_Total_LP_Count = (onlycell_meter['NoOfIntervals'].sum() + LDA_meter['NoOfIntervals'].sum())
        self.Expected_CellMeter_Total_LP_Count = (Total_CellMeter_Count*48)
        self.CellMeter_Total_LP_Count = (onlycell_meter['NoOfIntervals'].sum())
        self.Expected_LDAMeter_Total_LP_Count = (Total_LDAMeter_Count*144)
        self.LDAMeter_Total_LP_Count = (LDA_meter['NoOfIntervals'].sum())
        self.Full48_LP_Interval_AllMeter_Count = cr_list['NoOfIntervals'] == 48
        self.Average_LP_Interval_Push_Count = cr_list['NoOfIntervals'].mean()
        self.StdDev_LP_Interval_Push_Count = cr_list['NoOfIntervals'].std()
        self.Full48_LP_Interval_AllMeter_Count = self.Full48_LP_Interval_AllMeter_Count.sum()
        self.Full48_LP_Interval_HighRiseMeter_Count = cr_highrise['NoOfIntervals'] == 48
        self.Full48_LP_Interval_HIghRiseMeter_Count = self.Full48_LP_Interval_HighRiseMeter_Count.sum()
        self.Full48_LP_Interval_VillageMeter_Count = cr_village['NoOfIntervals'] == 48
        self.Full48_LP_Interval_VillageMeter_Count = self.Full48_LP_Interval_VillageMeter_Count.sum()
        self.Full144_LP_Interval_LDAMeter_Count = cr_list['NoOfIntervals'] == 144
        self.Full144_LP_Interval_LDAMeter_Count = self.Full144_LP_Interval_LDAMeter_Count.sum()
    
    def get_dict_alllppushcountperformance(self):
        return collections.OrderedDict({
        '[ OVERALL LP PUSH COUNT PERFORMANCE ]':'',
        'Expected All Meter Total LP Interval Push Count': self.Expected_AllMeter_Total_LP_Count,
        'All Meter Total LP Interval Push Count': self.AllMeter_Total_LP_Count,
        'Expected HighRise Meter Total LP Interval Push Count':self.Expected_HighRiseMeter_Total_LP_Count,
        'HighRise Meter Total LP Interval Push Count':self.HighRiseMeter_Total_LP_Count,
        'Expected Village Meter Total LP Interval Push Count':self.Expected_VillageMeter_Total_LP_Count,
        'Village Meter Total LP Interval Push Count':self.VillageMeter_Total_LP_Count,
        'Expected All Cell Meter Total LP Interval Push Count':self.Expected_AllCellMeter_Total_LP_Count,
        'All Cell Meter Total LP Interval Push Count':self.AllCellMeter_Total_LP_Count,
        'Expected Cell Meter Total LP Interval Push Count':self.Expected_CellMeter_Total_LP_Count,
        'Cell Meter Total LP Interval Push Count':self.CellMeter_Total_LP_Count,
        'Expected LDA Meter Total LP Interval Push Count':self.Expected_LDAMeter_Total_LP_Count,
        'LDA Meter Total LP Interval Push Count':self.LDAMeter_Total_LP_Count,
        'Full 48 LP Interval HighRise Meter Count':self.Full48_LP_Interval_HIghRiseMeter_Count,
        'Full 48 LP Interval Village Meter Count':self.Full48_LP_Interval_VillageMeter_Count,
        'Full 144 LP Interval LDA Meter Count':self.Full144_LP_Interval_LDAMeter_Count
        })

class AllDayendPushCountPerformance:
    def __init__(self, cr_list, Total_HighRiseMeter_Count, Total_VillageMeter_Count, unknownbuilding_Count, Total_LDAMeter_Count, cr_highrise, cr_village, Total_CellMeter_Count, onlycell_meter, LDA_meter):
        self.Expected_AllMeter_Total_DayEnd_Reading_Count = (Total_HighRiseMeter_Count + Total_VillageMeter_Count + unknownbuilding_Count)
        self.AllMeter_Total_DayEnd_Reading_Count = cr_list['DayEnd'].sum()
        self.Expected_HighRiseMeter_Total_DayEnd_Reading_Count = Total_HighRiseMeter_Count
        self.HighRiseMeter_Total_DayEnd_Reading_Count = cr_highrise['DayEnd'].sum()
        self.Expected_VillageMeter_Total_DayEnd_Reading_Count = Total_VillageMeter_Count
        self.VillageMeter_Total_DayEnd_Reading_Count = cr_village['DayEnd'].sum()
        self.Expected_AllCellMeter_Total_DayEnd_Reading_Count = (Total_CellMeter_Count+Total_LDAMeter_Count)
        self.AllCellMeter_Total_DayEnd_Reading_Count = (onlycell_meter['DayEnd'].sum() + LDA_meter['DayEnd'].sum())
        self.Expected_CellMeter_Total_DayEnd_Reading_Count = Total_CellMeter_Count
        self.CellMeter_Total_DayEnd_Reading_Count = (onlycell_meter['DayEnd'].sum())
        self.Expected_LDAMeter_Total_DayEnd_Reading_Count = (Total_LDAMeter_Count)
        self.LDAMeter_Total_DayEnd_Reading_Count = (LDA_meter['DayEnd'].sum())
        self.Missing_DayEnd_Reading_AllMeter_Count = self.Expected_AllMeter_Total_DayEnd_Reading_Count-self.AllMeter_Total_DayEnd_Reading_Count
        self.Missing_DayEnd_Reading_HighRiseMeter_Count = self.Expected_HighRiseMeter_Total_DayEnd_Reading_Count-self.HighRiseMeter_Total_DayEnd_Reading_Count
        self.Missing_DayEnd_Reading_VillageMeter_Count = self.Expected_VillageMeter_Total_DayEnd_Reading_Count-self.VillageMeter_Total_DayEnd_Reading_Count
        self.Missing_DayEnd_Reading_AllCellMeter_Count = self.Expected_AllCellMeter_Total_DayEnd_Reading_Count-self.AllCellMeter_Total_DayEnd_Reading_Count
        self.Missing_DayEnd_Reading_CellMeter_Count = self.Expected_CellMeter_Total_DayEnd_Reading_Count-self.CellMeter_Total_DayEnd_Reading_Count
        self.Missing_DayEnd_Reading_LDAMeter_Count = self.Expected_LDAMeter_Total_DayEnd_Reading_Count-self.LDAMeter_Total_DayEnd_Reading_Count

    def get_dict_alldayendpushcountperformance(self):
        return collections.OrderedDict({
        '[ OVERALL DAYEND READING PUSH COUNT PERFORMANCE ]':'',
        'Expected All Meter Total DayEnd Reading Push Count': self.Expected_AllMeter_Total_DayEnd_Reading_Count,
        'All Meter Total DayEnd Reading Push Count': self.AllMeter_Total_DayEnd_Reading_Count,
        'Expected HighRise Meter Total DayEnd Reading Push Count':self.Expected_HighRiseMeter_Total_DayEnd_Reading_Count,
        'HighRise Meter Total DayEnd Reading Push Count':self.HighRiseMeter_Total_DayEnd_Reading_Count,
        'Expected Village Meter Total DayEnd Reading Push Count':self.Expected_VillageMeter_Total_DayEnd_Reading_Count,
        'Village Meter Total DayEnd Reading Push Count':self.VillageMeter_Total_DayEnd_Reading_Count,
        'Expected All Cell Meter Total DayEnd Reading Push Count':self.Expected_AllCellMeter_Total_DayEnd_Reading_Count,
        'All Cell Meter Total DayEnd Reading Push Count':self.AllCellMeter_Total_DayEnd_Reading_Count,
        'Expected Cell Meter Total DayEnd Reading Push Count':self.Expected_CellMeter_Total_DayEnd_Reading_Count,
        'Cell Meter Total DayEnd Reading Push Count':self.CellMeter_Total_DayEnd_Reading_Count,
        'Expected LDA Meter Total DayEnd Reading Push Count':self.Expected_LDAMeter_Total_DayEnd_Reading_Count,
        'LDA Meter Total DayEnd Reading Push Count':self.LDAMeter_Total_DayEnd_Reading_Count,
        'Missing DayEnd Reading HighRise Meter Count':self.Missing_DayEnd_Reading_HighRiseMeter_Count,
        'Missing DayEnd Reading Village Meter Count':self.Missing_DayEnd_Reading_VillageMeter_Count,
        'Missing DayEnd Reading Cell Meter Count':self.Missing_DayEnd_Reading_AllCellMeter_Count,
        'Missing DayEnd Reading Cell Meter Count':self.Missing_DayEnd_Reading_CellMeter_Count,
        'Missing DayEnd Reading LDA Meter Count':self.Missing_DayEnd_Reading_LDAMeter_Count
        })

class MeterTypeCompositionRate:
    def __init__(self, Total_AllMeter_Count, Normal_Meter_Count, SecConfig_Meter_Count, Config_Meter_Count, Discovered_Meter_Count, Failed_Meter_Count, Lost_Meter_Count, 
    Total_ALLCellMeter_Count, Total_HighRiseMeter_Count, Total_VillageMeter_Count, unlocated_meter_Count, Total_LDAMeter_Count, Total_CellMeter_Count):
        self.HighRiseMeter_Rate = (Total_HighRiseMeter_Count/Total_AllMeter_Count)*100
        self.VillageMeter_Rate = (Total_VillageMeter_Count/Total_AllMeter_Count)*100
        self.AllCellMeter_Rate = (Total_ALLCellMeter_Count/Total_AllMeter_Count)*100
        self.CellMeter_Rate = (Total_CellMeter_Count/Total_AllMeter_Count)*100
        self.LDAMeter_Rate = (Total_LDAMeter_Count/Total_AllMeter_Count)*100
        self.UnlocatedMeter_Rate = (unlocated_meter_Count/Total_AllMeter_Count)*100
        self.Normal_Meter_Rate = (Normal_Meter_Count/Total_AllMeter_Count)*100
        self.SecConfig_Meter_Rate = (SecConfig_Meter_Count/Total_AllMeter_Count)*100
        self.Config_Meter_Rate = (Config_Meter_Count/Total_AllMeter_Count)*100
        self.Discovered_Meter_Rate = (Discovered_Meter_Count/Total_AllMeter_Count)*100
        self.Failed_Meter_Rate = (Failed_Meter_Count/Total_AllMeter_Count)*100
        self.Lost_Meter_Rate = (Lost_Meter_Count/Total_AllMeter_Count)*100

    def get_dict_metertypecompositionrate(self):
        return collections.OrderedDict({
        '[ METER TYPE COMPOSITION RATE ]':'',
        'HighRise Meter(%)':round(self.HighRiseMeter_Rate,2),
        'Village Meter(%)':round(self.VillageMeter_Rate,2),
        'All Cell Meter(%)':round(self.AllCellMeter_Rate,2),
        'Cell Meter(%)':round(self.CellMeter_Rate,2),
        'LDA Meter(%)':round(self.LDAMeter_Rate,2),
        'Unlocated Meter(%)':round(self.UnlocatedMeter_Rate,2),
        'Normal Status Meter(%)':round(self.Normal_Meter_Rate,2),
        'SecConfig Status Meter(%)':round(self.SecConfig_Meter_Rate,3),
        'Configure Status Meter(%)':round(self.Config_Meter_Rate,3),
        'Discovered Status Meter(%)':round(self.Discovered_Meter_Rate,3),
        'Failed Status Meter(%)':round(self.Failed_Meter_Rate,3),
        'Lost Status Meter(%)':round(self.Lost_Meter_Rate,3)
        })

class FirmwarePerformance():
    def __init__(self, cr_list):
        self.fw_avg = cr_list.pivot_table(values = ['NoOfIntervals'], index = ['firmwareversion'], aggfunc = {'NoOfIntervals': np.mean})
        self.fw_std = cr_list.pivot_table(values = ['NoOfIntervals'], index = ['firmwareversion'], aggfunc = {'NoOfIntervals': np.std})
        self.fw_perf = pd.concat([self.fw_avg, self.fw_std], axis=1, join_axes=[self.fw_avg.index])
        self.fw_perf.columns = ['LP Average', 'LP Std Deviation']
        self.fw_perf = self.fw_perf.round()
        
    def output_fw_stats(self):
        return self.fw_perf

class ToDataFrame():
    def __init__(self, param1, param2):
        self.title = "{}".format(param1)
        self.performance = pd.DataFrame(pd.io.json.json_normalize(param2).T)
        self.performance.columns = [self.title]

    def output_dataframe(self):
        return self.performance

class WriteToExcel(DistrictPerformance, LatestMeterPerformance):
    def __init__(self, dfname, sheetname, dfname2, sheetname2, sheetname3, sheetname4, sheetname5, sheetname6, todaydate):
        super().__init__(cr_list, 'A') 
        #super().__init__(cr_list)
        self.dir = 'C:/Users/TokuharM.AP/mywork/'
        self.writer = pd.ExcelWriter('%sReading_Performance_Report_%s.xlsx' % (self.dir, todaydate))

    def write_to_excel(self, dfname, sheetname, dfname2, sheetname2, sheetname3, sheetname4, sheetname5, sheetname6):
        dfname.to_excel(self.writer, sheetname)
        dfname2.to_excel(self.writer, sheetname2, index=False)
        self.cr_perf.to_excel(self.writer, sheetname3)
        #self.Latest_Meters.to_excel(self.writer, sheetname4)
        #self.fw_perf.to_excel(self.writer, sheetname5)
        self.cr_rank.to_excel(self.writer, sheetname4)
        self.region_perf.to_excel(self.writer, sheetname5)
        self.area_perf.to_excel(self.writer, sheetname6)

        self.workbook  = self.writer.book
        self.worksheet = self.writer.sheets[sheetname]
        self.cell_format = self.workbook.add_format()
        self.cell_format.set_align('right')
        self.worksheet.set_column(0,0,68, self.cell_format)
        self.worksheet.set_column(1,10,17)

        self.worksheet = self.writer.sheets[sheetname3]
        self.worksheet.set_column(0,0,27)
        self.worksheet.set_column(1,2,16)

        self.worksheet = self.writer.sheets[sheetname4]
        self.worksheet.set_column(0,0,11)
        self.worksheet.set_column(1,7,9)

        self.worksheet = self.writer.sheets[sheetname5]
        self.worksheet.set_column(0,0,27)
        self.worksheet.set_column(1,2,16)

        self.worksheet = self.writer.sheets[sheetname6]
        self.worksheet.set_column(0,0,37)
        self.worksheet.set_column(1,7,9)
        self.writer.save()

def main():

    '''
    district_list = list(
    [DistrictPerformance(cr_list, 'District A'),
    DistrictPerformance(cr_list, 'District B'),
    DistrictPerformance(cr_list, 'District C'),
    DistrictPerformance(cr_list, 'District D')])
    dict_district = []
    for district in district_list:
        dict_district.append(district.get_dict())
    '''

    district_a = DistrictPerformance(cr_list, 'District A')
    district_b = DistrictPerformance(cr_list, 'District B')
    district_c = DistrictPerformance(cr_list, 'District C')
    district_d = DistrictPerformance(cr_list, 'District D')

    kpimetercount = KeyPerformanceIndicator(
        Total_AllMeter_Count, 
        Total_LDAMeter_Count, 
        onlycell_meter, 
        Total_HighRiseMeter_Count, 
        Total_VillageMeter_Count, 
        unknownbuilding_Count, 
        Total_CellMeter_Count, 
        cr_list)

    slametercount = SLAPerformance(cr_list)

    latestmetercount = LatestMeterPerformance(cr_list)

    allmetermetercount = AllMetersCount()

    allllpintervalpushsuccessrate = AllLPIntervalPushSuccessRate(
        Total_HighRiseMeter_Count, 
        Total_VillageMeter_Count, 
        Total_CellMeter_Count, 
        Total_LDAMeter_Count, 
        cr_highrise, 
        cr_village, 
        onlycell_meter, 
        LDA_meter)

    allldayendpushsuccessrate = AllDayendPushSuccessRate(
        Total_HighRiseMeter_Count, 
        cr_highrise, 
        Total_VillageMeter_Count, 
        cr_village, 
        Total_CellMeter_Count, 
        Total_LDAMeter_Count, 
        onlycell_meter, 
        LDA_meter)

    nolppushmetersummary = NoLpPushMeterSummary(Total_AllMeter_Count, cr_list)

    metersstatuscount = MeterStatusCount(cr_list)

    alllppushcountperformance = AllLpPushCountPerformance(
        cr_list, 
        Total_HighRiseMeter_Count, 
        Total_VillageMeter_Count, 
        unknownbuilding_Count, 
        Total_LDAMeter_Count, 
        cr_highrise, 
        cr_village, 
        Total_CellMeter_Count, 
        onlycell_meter, 
        LDA_meter)

    alldayendpushcountperformance = AllDayendPushCountPerformance(
        cr_list, 
        Total_HighRiseMeter_Count, 
        Total_VillageMeter_Count, 
        unknownbuilding_Count, 
        Total_LDAMeter_Count, 
        cr_highrise, 
        cr_village, 
        Total_CellMeter_Count, 
        onlycell_meter, 
        LDA_meter)

    metercompositionrate = MeterTypeCompositionRate(
        Total_AllMeter_Count, 
        Normal_Meter_Count, 
        SecConfig_Meter_Count, 
        Config_Meter_Count, 
        Discovered_Meter_Count, 
        Failed_Meter_Count, 
        Lost_Meter_Count, 
        Total_ALLCellMeter_Count, 
        Total_HighRiseMeter_Count, 
        Total_VillageMeter_Count, 
        unlocated_meter_Count, 
        Total_LDAMeter_Count, 
        Total_CellMeter_Count)

    new_dict = collections.OrderedDict(
        **kpimetercount.get_dict_kpi(),
        **slametercount.get_dict_sla(),
        **latestmetercount.get_dict_latest(),
        **allmetermetercount.get_dict_allmetercount(Total_ALLCellMeter_Count, Total_HighRiseMeter_Count, Total_VillageMeter_Count, Total_CellMeter_Count, Total_LDAMeter_Count, unlocated_meter_Count),
        **allllpintervalpushsuccessrate.get_dict_alllpintervalpushsuccessrate(),
        **allldayendpushsuccessrate.get_dict_alldayendpushsuccessrate(),
        **nolppushmetersummary.get_dict_allnolppushmetersummary(),
        **metersstatuscount.get_dict_meterstatuscount(),
        **alllppushcountperformance.get_dict_alllppushcountperformance(),
        **alldayendpushcountperformance.get_dict_alldayendpushcountperformance(),
        **metercompositionrate.get_dict_metertypecompositionrate(),
        **district_a.get_dict(),
        **district_b.get_dict(),
        **district_c.get_dict(),
        **district_d.get_dict()
    )

    firmwareperformance = FirmwarePerformance(cr_list)
    firmwareperformance = firmwareperformance.output_fw_stats()

    df_performance = ToDataFrame('Performance Result', new_dict)
    df_performance = df_performance.output_dataframe()

    cr_performance = CollectorPerformance(cr_list)
    cr_performance = cr_performance.get_collector_statistics()

    perftoexcel = WriteToExcel(
        df_performance, 'Performance Report', 
        cr_list, 'Analyzed Individual Meters', 
        'CR Performance', 
        #'LatestMeters',
        #'FW Performance',
        'CR ABC Rank',
        "Region Performance",
        "Area Performance",
        today_date)

    perftoexcel.write_to_excel(
        df_performance, 'Performance Report', 
        cr_list, 'Analyzed Individual Meters', 
        'CR Performance',
        #'LatestMeters',
        #'FW Performance',
        'CR ABC Rank',
        "Region Performance",
        "Area Performance")

if __name__ == '__main__':
    main()