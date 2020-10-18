%matplotlib inline
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
full_smart_meters = pd.read_csv('C:/Users/Desktop/RF_Analysis/%s' % allmeter, sep=',',skiprows=1)
full_smart_meters.columns = ['meterno', 'serialnumber', 'endpointId', 'endpointtypeid', 'firmwareversion', 'endPointModelId', 'hwmodelid', 'date', 'initialDiscoveredDate', 'initialNormalDate', 'NoOfIntervals', 'name', 'abc_rank', 'DayEnd', 'meter_status', 'spuid', 'layer']

cr_list = pd.read_excel('C:/Users/Desktop/RF_Analysis/%s' % crloc, 'Sheet1', na_values=['NA'])
cr_list = cr_list.drop(cr_list.columns[[4,5,6]], axis=1)

full_smart_meters.set_index('name').join(cr_list.set_index('CollectorNo'))
cr_list = full_smart_meters.join(cr_list.set_index('CollectorNo'), on='name', how='outer')
cr_list = cr_list.fillna({'Estates / Villages': 'Unlocated Area', 'BuildingType': 'Unknown BuildingType' })
cr_list = cr_list[cr_list['meterno'].notnull()]

cr_village = cr_list[cr_list['BuildingType'].isin(['Village'])]
cr_highrise = cr_list[cr_list['BuildingType'].isin(['Highrise'])]
cr_unknownbuilding = cr_list[cr_list['BuildingType'].isin(['Unknown BuildingType'])]
cell_meter = cr_list[cr_list['endpointtypeid'] == 15]
onlycell_meter = cell_meter[~cell_meter['abc_rank'].str.startswith('Load_DA')]
LDA_meter = cell_meter[cell_meter['abc_rank'].str.startswith('Load_DA')]
cr_list = cr_list[~cr_list['abc_rank'].str.startswith('Load_DA')]

#FW version performance
fw_avg = cr_list.pivot_table(values = ['NoOfIntervals'], index = ['firmwareversion'], aggfunc = {'NoOfIntervals': np.mean})
fw_std = cr_list.pivot_table(values = ['NoOfIntervals'], index = ['firmwareversion'], aggfunc = {'NoOfIntervals': np.std})
fw_perf = pd.concat([fw_avg, fw_std], axis=1, join_axes=[fw_avg.index])
fw_perf.columns = ['LP Average', 'LP Std Deviation']
fw_perf = fw_perf.round()

class District:
    def __init__(self, cr_list, attr):
        self.name = "District {}".format(attr)
        self.district_meter = cr_list[cr_list['District'].str.contains(self.name, na=False)]
        self.district_meter_Count = district_meter['meterno'].count()
        self.district_meter_Full_48_LP_Interval = district_meter[district_meter['NoOfIntervals'] == 48]
        self.district_meter_Full_48_LP_Interval_Meter_Count = district_meter_Full_48_LP_Interval['meterno'].count()
        self.district_meter_Full_48_LP_Interval_Meter_Rate = round((district_meter_Full_48_LP_Interval_Meter_Count/district_meter_Count)*100,2)
        self.district_1468 = district_meter[district_meter['firmwareversion'].str.contains('-24.60', na=False)]
        self.district_1468_Count = district_1468['meterno'].count()
        self.district_1468_Rate = round((district_1468_Count/district_meter_Count)*100,2)
        self.district_meter_Normal_Meter = district_meter[district_meter['meter_status'] == 'Normal']
        self.district_meter_Normal_Meter_Count = district_meter_Normal_Meter['meterno'].count()
        self.district_meter_SecConfig_Meter = district_meter[district_meter['meter_status'] == 'SecConfig']
        self.district_meter_SecConfig_Meter_Count = district_meter_SecConfig_Meter['meterno'].count()
        self.district_meter_Discovered_Meter = district_meter[district_meter['meter_status'] == 'Discovered']
        self.district_meter_Discovered_Meter_Count = district_meter_Discovered_Meter['meterno'].count()
        self.district_meter_Config_Meter = district_meter[district_meter['meter_status'] == 'Configure']
        self.district_meter_Config_Meter_Count = district_meter_Config_Meter['meterno'].count()
        self.district_meter_Failed_Meter = district_meter[district_meter['meter_status'] == 'Failed']
        self.district_meter_Failed_Meter_Count = district_meter_Failed_Meter['meterno'].count()
        self.district_meter_Lost_Meter = district_meter[district_meter['meter_status'] == 'Lost']
        self.district_meter_Lost_Meter_Count = district_meter_Lost_Meter['meterno'].count()
        #LP-DayEnd-FULL_district Meter
        self.district_meter_LP_DayEnd_Full_Meter = district_meter[(district_meter['NoOfIntervals'] == 48) & (district_meter['DayEnd'] == 1)]
        self.district_meter_LP_DayEnd_Full_Meter_Count = district_meter_LP_DayEnd_Full_Meter['meterno'].count()
        self.district_meter_LP_DayEnd_Full_Meter_Rate = round((district_meter_LP_DayEnd_Full_Meter_Count/district_meter_Count)*100,2)
        self.district_meter_Missing_DayEnd_Reading = district_meter[district_meter['DayEnd'] != 1]
        self.district_meter_Missing_DayEnd_Reading_Meter_Count = district_meter_Missing_DayEnd_Reading['meterno'].count()
        self.Expected_district_meter_Total_LP_Count = ((district_meter_Count)*48)
        self.district_meter_Total_LP_Count = district_meter['NoOfIntervals'].sum()
        self.district_meter_Total_Dayend  = district_meter[district_meter['DayEnd'] == 1]
        self.district_meter_Total_Dayend_Count = district_meter_Total_Dayend['meterno'].count()
        self.district_meter_LP_Success_Rate = round((district_meter_Total_LP_Count/Expected_district_meter_Total_LP_Count)*100,2)
        self.district_meter_Dayend_Success_Rate  = round((district_meter_Total_Dayend_Count/district_meter_Count)*100,2)
        self.district_meter_Average_LP_Interval_Push_Count = district_meter['NoOfIntervals'].mean()
        self.district_meter_StdDev_LP_Interval_Push_Count = district_meter['NoOfIntervals'].std()
        #abc_rank
        self._CR_Rnk = district_meter.pivot_table(values = ['meter_status'], index = ['name'], columns = ['abc_rank'], aggfunc = 'count')
        self._CR_Rnk.columns = _CR_Rnk.columns.droplevel()
        self._CR_Rnk = _CR_Rnk.loc[:,['P','A','B','C','D','E','F']]
        self._CR_Rnk = _CR_Rnk.fillna(0)

    def get_dict(self):
        return collections.OrderedDict({
            '[ {} METERS SUMMARY ]'.format(self.name):'',
            '{} Meter Count'.format(self.name):district_meter_Count,
            '{} FW24.60 Meter Count'.format(self.name):district_1468_Count,
            '{} FW24.60 Meter(%)'.format(self.name):district_1468_Rate,
            '{} Meter LP Success(%)'.format(self.name):district_meter_LP_Success_Rate,
            '{} Meter Dayend Success(%)'.format(self.name):district_meter_Dayend_Success_Rate,
            '{} Average LP Push Count'.format(self.name):round(district_meter_Average_LP_Interval_Push_Count,2),
            '{} Std Deviation LP Push Count'.format(self.name):round(district_meter_StdDev_LP_Interval_Push_Count,2),
            '{} Meter LP-DayEnd-FULL Meter Count'.format(self.name):district_meter_LP_DayEnd_Full_Meter_Count,
            '{} Meter LP-DayEnd-FULL Meter(%)'.format(self.name):district_meter_LP_DayEnd_Full_Meter_Rate,
            '{} Meter Full 48 LP Interval Meter Count'.format(self.name):district_meter_Full_48_LP_Interval_Meter_Count,
            '{} Meter Full 48 LP Interval Meter(%)'.format(self.name):district_meter_Full_48_LP_Interval_Meter_Rate,
            '{} Meter Missing DayEnd Reading Meter Count'.format(self.name):district_meter_Missing_DayEnd_Reading_Meter_Count,
            '{} Meter Normal Meter Count'.format(self.name):district_meter_Normal_Meter_Count,
            '{} Meter SecConfig Meter Count'.format(self.name):district_meter_SecConfig_Meter_Count,
            '{} Meter Config Meter Count'.format(self.name):district_meter_Config_Meter_Count,
            '{} Meter Discovered Meter Count'.format(self.name):district_meter_Discovered_Meter_Count,
            '{} Meter Failed Meter Count'.format(self.name):district_meter_Failed_Meter_Count,
            '{} Meter Lost Meter Count'.format(self.name):district_meter_Lost_Meter_Count,
        })

district_a = District(cr_list, 'A')
district_b = District(cr_list, 'B')
district_c = District(cr_list, 'C')
district_d = District(cr_list, 'D')

No_reading_meter = cr_list[cr_list['abc_rank'] == 'F']
hexed_serial = pd.DataFrame(No_reading_meter['serialnumber'].astype(int))
hexed_serial = hexed_serial.rename(columns={'serialnumber':'hex_serial'})
hexed_serial = hexed_serial['hex_serial'].apply(lambda x:format(x, 'x'))
No_reading_meter = pd.concat([No_reading_meter, hexed_serial], axis=1)
No_reading_meter = No_reading_meter.reset_index(drop=True)

#No Reading Meter per Status
No_Reading_RF_meter = No_reading_meter[No_reading_meter['endpointtypeid'] == 9]
No_Reading_cell_meter = No_reading_meter[No_reading_meter['endpointtypeid'] == 15]
No_Reading_Normal_Meter = No_reading_meter[No_reading_meter['meter_status'] == 'Normal']
No_Reading_SecConfig_Meter = No_reading_meter[No_reading_meter['meter_status'] == 'SecConfig']
No_Reading_Discovered_Meter = No_reading_meter[No_reading_meter['meter_status'] == 'Discovered']
No_Reading_Config_Meter = No_reading_meter[No_reading_meter['meter_status'] == 'Configure']
No_Reading_Failed_Meter = No_reading_meter[No_reading_meter['meter_status'] == 'Failed']
No_Reading_Lost_Meter = No_reading_meter[No_reading_meter['meter_status'] == 'Lost']
No_Reading_Meter_with_DayEnd = No_reading_meter[No_reading_meter['DayEnd'] == 1 ]
No_LPandDayEnd_Reading_Meter_with_DayEnd = No_reading_meter[No_reading_meter['DayEnd'] == 0 ]
No_reading_meter_Highrise = No_reading_meter[No_reading_meter['BuildingType'].isin(['Highrise'])]
No_reading_meter_Village = No_reading_meter[No_reading_meter['BuildingType'].isin(['Village'])]
No_reading_meter_Unlocated = No_reading_meter[No_reading_meter['BuildingType'].isin(['Unknown BuildingType'])]

Normal_Meter = cr_list[cr_list['meter_status'] == 'Normal']
SecConfig_Meter = cr_list[cr_list['meter_status'] == 'SecConfig']
Discovered_Meter = cr_list[cr_list['meter_status'] == 'Discovered']
Config_Meter = cr_list[cr_list['meter_status'] == 'Configure']
Failed_Meter = cr_list[cr_list['meter_status'] == 'Failed']
Lost_Meter = cr_list[cr_list['meter_status'] == 'Lost']

today_date = dt.date.today().strftime('%Y-%m-%d') 

#Effective Meter Calculation(Only Normal meters (w/o LDA Meter) that passed more than 7 days from initial Normal date)
cr_list['initialNormalDate'] = pd.to_datetime(cr_list['initialNormalDate'], format='%Y-%m-%d')
cr_list['date'] = cr_list['date'].fillna(today_date)
cr_list['date'] = pd.to_datetime(cr_list['date'], format='%Y-%m-%d')
cr_list['7Days_After_Normal'] = (cr_list['initialNormalDate']  + dt.timedelta(days=7))

cr_list['initialNormalDate'] = cr_list['initialNormalDate'].values.astype('datetime64[D]')
cr_list['7Days_After_Normal'] = cr_list['7Days_After_Normal'].values.astype('datetime64[D]')
cr_list['initialDiscoveredDate'] = cr_list['initialDiscoveredDate'].values.astype('datetime64[D]')
cr_list['Difference'] = cr_list['date'] - cr_list['initialNormalDate']
cr_list['DaysAfterDis'] = cr_list['date'] - cr_list['initialDiscoveredDate']
cr_list['DisToNorm'] = cr_list['initialNormalDate'] - cr_list['initialDiscoveredDate']

#SLA-Meters
Effective_Meter = cr_list[cr_list['Difference'] >= '7 days']
Effective_Meter = cr_list[(cr_list['meter_status'] == 'Normal')]
Effective_Meter_Count = Effective_Meter['meterno'].count()

EffectiveMeters_Full_48_LP_Interval = Effective_Meter[Effective_Meter['NoOfIntervals'] == 48]
EffectiveMeters_Full_48_LP_Interval_Meter_Count = EffectiveMeters_Full_48_LP_Interval['meterno'].count()
EffectiveMeters_Full_48_LP_Interval_Meter_Rate = round((EffectiveMeters_Full_48_LP_Interval_Meter_Count/Effective_Meter_Count)*100,2)

#LP-DayEnd-FULL_SLA_Meter
LP_DayEnd_Full_Effective_Meter = Effective_Meter[(Effective_Meter['NoOfIntervals'] == 48)&(Effective_Meter['DayEnd'] == 1)]
LP_DayEnd_Full_Effective_Meter_Count = LP_DayEnd_Full_Effective_Meter['meterno'].count()
LP_DayEnd_Full_Effective_Meter_Rate = round((LP_DayEnd_Full_Effective_Meter_Count/Effective_Meter_Count)*100,2)

EffectiveMeters_Missing_DayEnd_Reading = Effective_Meter[Effective_Meter['DayEnd'] != 1]
EffectiveMeters_Missing_DayEnd_Reading_Meter_Count = EffectiveMeters_Missing_DayEnd_Reading['meterno'].count()
EffectiveMeters_Missing_DayEnd_Reading_Meter_Rate = round((EffectiveMeters_Missing_DayEnd_Reading_Meter_Count/Effective_Meter_Count)*100,2)

Expected_EffectiveMeter_Total_LP_Count = ((Effective_Meter_Count)*48)
EffectiveMeter_Total_LP_Count = Effective_Meter['NoOfIntervals'].sum()
EffectiveMeter_Total_Dayend  = Effective_Meter[Effective_Meter['DayEnd'] == 1]
EffectiveMeter_Total_Dayend_Count = EffectiveMeter_Total_Dayend['meterno'].count()

Effective_Meter_LP_Success_Rate = round((EffectiveMeter_Total_LP_Count/Expected_EffectiveMeter_Total_LP_Count)*100,2)
Effective_Meter_Dayend_Success_Rate  = round((EffectiveMeter_Total_Dayend_Count/Effective_Meter_Count)*100,2)
Effective_Meter_Average_LP_Interval_Push_Count = Effective_Meter['NoOfIntervals'].mean()
Effective_Meter_StdDev_LP_Interval_Push_Count = Effective_Meter['NoOfIntervals'].std()

#Latest-Meters(Registered in last 30days)
Latest_Meters = cr_list[cr_list['DaysAfterDis']  < '30 days']
Latest_Meters_Count = Latest_Meters['meterno'].count()

Latest_Meters_Full_48_LP_Interval = Latest_Meters[Latest_Meters['NoOfIntervals'] == 48]
Latest_Meters_Full_48_LP_Interval_Meter_Count = Latest_Meters_Full_48_LP_Interval['meterno'].count()
Latest_Meters_Full_48_LP_Interval_Meter_Rate = round((Latest_Meters_Full_48_LP_Interval_Meter_Count/Latest_Meters_Count)*100,2)

#LP-DayEnd-FULL_SLA_Meter
LP_DayEnd_Full_Latest_Meters = Latest_Meters[(Latest_Meters['NoOfIntervals'] == 48)&(Latest_Meters['DayEnd'] == 1)]
LP_DayEnd_Full_Latest_Meters_Count = LP_DayEnd_Full_Latest_Meters['meterno'].count()
LP_DayEnd_Full_Latest_Meters_Rate = round((LP_DayEnd_Full_Latest_Meters_Count/Latest_Meters_Count)*100,2)

Latest_Meters_Missing_DayEnd_Reading = Latest_Meters[Latest_Meters['DayEnd'] != 1]
Latest_Meters_Missing_DayEnd_Reading_Meter_Count = Latest_Meters_Missing_DayEnd_Reading['meterno'].count()
Latest_Meters_Missing_DayEnd_Reading_Meter_Rate = round((Latest_Meters_Missing_DayEnd_Reading_Meter_Count/Latest_Meters_Count)*100,2)

Expected_Latest_Meters_Count_Total_LP_Count = ((Latest_Meters_Count)*48)
Latest_Meters_Total_LP_Count = Latest_Meters['NoOfIntervals'].sum()
Latest_Meters_Total_Dayend  = Latest_Meters[Latest_Meters['DayEnd'] == 1]
Latest_Meters_Total_Dayend_Count = Latest_Meters_Total_Dayend['meterno'].count()

Latest_Meters_LP_Success_Rate = round((Latest_Meters_Total_LP_Count/Expected_Latest_Meters_Count_Total_LP_Count)*100,2)
Latest_Meters_Dayend_Success_Rate  = round((Latest_Meters_Total_Dayend_Count/Latest_Meters_Count)*100,2)
Latest_Meters_Average_LP_Interval_Push_Count = Latest_Meters['NoOfIntervals'].mean()
Latest_Meters_StdDev_LP_Interval_Push_Count = Latest_Meters['NoOfIntervals'].std()

unlocated_meter = cr_list[cr_list['Estates / Villages'] == 'Unlocated Area']
unlocated_meter = unlocated_meter[unlocated_meter['meterno'].notnull()]

#Number of Total Meters
Total_AllMeter_Count = cr_list['meterno'].count()
Total_HighRiseMeter_Count = cr_highrise['meterno'].count()
Total_VillageMeter_Count = cr_village['meterno'].count()
Total_ALLCellMeter_Count = cell_meter['meterno'].count()
Total_LDAMeter_Count = LDA_meter['meterno'].count()
Total_CellMeter_Count = Total_ALLCellMeter_Count - Total_LDAMeter_Count
unlocated_meter_Count = unlocated_meter['meterno'].count()
unknownbuilding_Count = cr_unknownbuilding['meterno'].count()

all_meter_1468 = cr_list[cr_list['firmwareversion'].str.contains('-24.60', na=False)]
all_meter_1468_Count = all_meter_1468['meterno'].count()
all_meter_1468_1468_Rate  = round((all_meter_1468_Count/Total_AllMeter_Count )*100,2)

Missing_Full_48_LP_Interval_Meters = cr_list[cr_list['NoOfIntervals'] < 48]
Missing_Full_48_LP_Interval_Meters = Missing_Full_48_LP_Interval_Meters.reset_index(drop=True)
Missing_DayEnd_Reading_All_Meters = cr_list[cr_list['DayEnd'] != 1]
Missing_DayEnd_Reading_All_Meters = Missing_DayEnd_Reading_All_Meters.reset_index(drop=True)
Success_DayEnd_Reading_All_Meters = cr_list[cr_list['DayEnd'] == 1]

#Number of No Reading Meter Status Count
No_Reading_Meter_Total_Count = No_reading_meter['abc_rank'].count()
No_Reading_RF_meter_Count = No_Reading_RF_meter['meterno'].count()
No_Reading_Cell_meter_Count = No_Reading_cell_meter['meterno'].count()
No_reading_Normal_meter_count = No_Reading_Normal_Meter['meterno'].count()
No_reading_SecConfig_meter_count = No_Reading_SecConfig_Meter['meterno'].count()
No_reading_Discovered_meter_count = No_Reading_Discovered_Meter['meterno'].count()
No_reading_Config_meter_count = No_Reading_Config_Meter['meterno'].count()
No_reading_Failed_meter_count = No_Reading_Failed_Meter['meterno'].count()
No_reading_Lost_meter_count = No_Reading_Lost_Meter['meterno'].count()
No_Reading_Meter_with_DayEnd_count = No_Reading_Meter_with_DayEnd['meterno'].count()
No_LPandDayEnd_Reading_Meter_with_DayEnd_Count = No_LPandDayEnd_Reading_Meter_with_DayEnd['meterno'].count()
No_reading_meter_Highrise_count = No_reading_meter_Highrise['abc_rank'].count()
No_reading_meter_Village_count = No_reading_meter_Village['abc_rank'].count()
No_reading_meter_Unlocated_count = No_reading_meter_Unlocated['abc_rank'].count()

#Meter Status Count
Normal_Meter_Count = Normal_Meter['meterno'].count()
SecConfig_Meter_Count = SecConfig_Meter['meterno'].count()
Config_Meter_Count = Config_Meter['meterno'].count()
Discovered_Meter_Count = Discovered_Meter['meterno'].count()
Failed_Meter_Count = Failed_Meter['meterno'].count()
Lost_Meter_Count = Lost_Meter['meterno'].count()
RF_meter = cr_list[cr_list['endpointtypeid'] != 15]
Collector_Count = RF_meter['name'].nunique()

#Performance per Areas
area_perf = cr_list.pivot_table(values = ['meter_status'], index = ['Estates / Villages'], columns = ['abc_rank'], aggfunc = 'count')
area_perf.columns = area_perf.columns.droplevel()
area_perf = area_perf.loc[:,['P','A','B','C','D','E','F']]
area_perf = area_perf.fillna(0)

#Performance per Areas
region_perf = cr_list.groupby(['Estates / Villages'])['NoOfIntervals'].mean()
region_perf_std = cr_list.groupby(['Estates / Villages'])['NoOfIntervals'].std()
region_perf = pd.concat([region_perf, region_perf_std], axis=1, join_axes=[region_perf.index])
region_perf = region_perf.round()
region_perf.columns = ['Average LP Count','Std LP Count']

Expected_AllMeter_Total_DayEnd_Reading_Count = (Total_HighRiseMeter_Count + Total_VillageMeter_Count + unknownbuilding_Count)

AllMeter_Total_DayEnd_Reading_Count = cr_list['DayEnd'].sum()
Expected_HighRiseMeter_Total_DayEnd_Reading_Count = Total_HighRiseMeter_Count
HighRiseMeter_Total_DayEnd_Reading_Count = cr_highrise['DayEnd'].sum()
Expected_VillageMeter_Total_DayEnd_Reading_Count = Total_VillageMeter_Count
VillageMeter_Total_DayEnd_Reading_Count = cr_village['DayEnd'].sum()
Expected_AllCellMeter_Total_DayEnd_Reading_Count = (Total_CellMeter_Count+Total_LDAMeter_Count)
AllCellMeter_Total_DayEnd_Reading_Count = (onlycell_meter['DayEnd'].sum() + LDA_meter['DayEnd'].sum())
Expected_CellMeter_Total_DayEnd_Reading_Count = Total_CellMeter_Count
CellMeter_Total_DayEnd_Reading_Count = (onlycell_meter['DayEnd'].sum())
Expected_LDAMeter_Total_DayEnd_Reading_Count = (Total_LDAMeter_Count)
LDAMeter_Total_DayEnd_Reading_Count = (LDA_meter['DayEnd'].sum())

Missing_DayEnd_Reading_AllMeter_Count = Expected_AllMeter_Total_DayEnd_Reading_Count-AllMeter_Total_DayEnd_Reading_Count
Missing_DayEnd_Reading_HighRiseMeter_Count = Expected_HighRiseMeter_Total_DayEnd_Reading_Count-HighRiseMeter_Total_DayEnd_Reading_Count
Missing_DayEnd_Reading_VillageMeter_Count = Expected_VillageMeter_Total_DayEnd_Reading_Count-VillageMeter_Total_DayEnd_Reading_Count
Missing_DayEnd_Reading_AllCellMeter_Count = Expected_AllCellMeter_Total_DayEnd_Reading_Count-AllCellMeter_Total_DayEnd_Reading_Count
Missing_DayEnd_Reading_CellMeter_Count = Expected_CellMeter_Total_DayEnd_Reading_Count-CellMeter_Total_DayEnd_Reading_Count
Missing_DayEnd_Reading_LDAMeter_Count = Expected_LDAMeter_Total_DayEnd_Reading_Count-LDAMeter_Total_DayEnd_Reading_Count

NO_DayEnd_Reading_but_with_LP_Reading_Meter = full_smart_meters[full_smart_meters['DayEnd'] == 0]
NO_DayEnd_Reading_but_with_LP_Reading_Meter = NO_DayEnd_Reading_but_with_LP_Reading_Meter[NO_DayEnd_Reading_but_with_LP_Reading_Meter['NoOfIntervals'] != 0]
NO_DayEnd_Reading_but_with_LP_Reading_Meter_Count = NO_DayEnd_Reading_but_with_LP_Reading_Meter['NoOfIntervals'].count()

#DayEnd Reading Push % Performance
AllMeter_Total_DayEnd_Reading_SuccessRate = (AllMeter_Total_DayEnd_Reading_Count/Expected_AllMeter_Total_DayEnd_Reading_Count)*100
MissingDayEndReadingAllMeterRate = (Missing_DayEnd_Reading_AllMeter_Count/Expected_AllMeter_Total_DayEnd_Reading_Count)*100

HighRiseMeter_Total_DayEnd_Reading_SuccessRate = (HighRiseMeter_Total_DayEnd_Reading_Count/Expected_HighRiseMeter_Total_DayEnd_Reading_Count)*100
VillageMeter_Total_DayEnd_Reading_SuccessRate = (VillageMeter_Total_DayEnd_Reading_Count/Expected_VillageMeter_Total_DayEnd_Reading_Count)*100
AllCellMeter_Total_DayEnd_Reading_SuccessRate = (AllCellMeter_Total_DayEnd_Reading_Count/Expected_AllCellMeter_Total_DayEnd_Reading_Count)*100
CellMeter_Total_DayEnd_Reading_SuccessRate = (CellMeter_Total_DayEnd_Reading_Count/Expected_CellMeter_Total_DayEnd_Reading_Count)*100
LDAMeter_Total_DayEnd_Reading_SuccessRate = (LDAMeter_Total_DayEnd_Reading_Count/Expected_LDAMeter_Total_DayEnd_Reading_Count)*100
NO_LPReading_ButWithDayEnd_Reading_Rate = (No_Reading_Meter_with_DayEnd_count/Total_AllMeter_Count)*100
No_LPandDayEnd_Reading_Meter_with_DayEnd_Rate = (No_LPandDayEnd_Reading_Meter_with_DayEnd_Count/Total_AllMeter_Count)*100
NO_DayEnd_Reading_but_with_LP_Reading_Meter_Rate = (NO_DayEnd_Reading_but_with_LP_Reading_Meter_Count/Total_AllMeter_Count)*100

#MeterType Composition Rate
HighRiseMeter_Rate = (Total_HighRiseMeter_Count/Total_AllMeter_Count)*100
VillageMeter_Rate = (Total_VillageMeter_Count/Total_AllMeter_Count)*100
AllCellMeter_Rate = (Total_ALLCellMeter_Count/Total_AllMeter_Count)*100
CellMeter_Rate = (Total_CellMeter_Count/Total_AllMeter_Count)*100
LDAMeter_Rate = (Total_LDAMeter_Count/Total_AllMeter_Count)*100
UnlocatedMeter_Rate = (unlocated_meter_Count/Total_AllMeter_Count)*100
No_Reading_Meter_Rate = (No_Reading_Meter_Total_Count/Total_AllMeter_Count)*100
No_Reading_Meter_Highrise_Rate = (No_reading_meter_Highrise_count/No_Reading_Meter_Total_Count)*100
No_Reading_Meter_Village_Rate = (No_reading_meter_Village_count/No_Reading_Meter_Total_Count)*100
No_reading_meter_Unlocated_Rate = (No_reading_meter_Unlocated_count/No_Reading_Meter_Total_Count)*100

#MeterStatus Composition Rate
Normal_Meter_Rate = (Normal_Meter_Count/Total_AllMeter_Count)*100
SecConfig_Meter_Rate = (SecConfig_Meter_Count/Total_AllMeter_Count)*100
Config_Meter_Rate = (Config_Meter_Count/Total_AllMeter_Count)*100
Discovered_Meter_Rate = (Discovered_Meter_Count/Total_AllMeter_Count)*100
Failed_Meter_Rate = (Failed_Meter_Count/Total_AllMeter_Count)*100
Lost_Meter_Rate = (Lost_Meter_Count/Total_AllMeter_Count)*100

#No Reading MeterStatus Composition Rate
No_Reading_Normal_Meter_Rate = (No_reading_Normal_meter_count/No_Reading_Meter_Total_Count)*100
No_Reading_SecConfig_Meter_Rate = (No_reading_SecConfig_meter_count/No_Reading_Meter_Total_Count)*100
No_Reading_Config_Meter_Rate = (No_reading_Config_meter_count/No_Reading_Meter_Total_Count)*100
No_Reading_Discovered_Meter_Rate = (No_reading_Discovered_meter_count/No_Reading_Meter_Total_Count)*100
No_Reading_Failed_Meter_Rate = (No_reading_Failed_meter_count/No_Reading_Meter_Total_Count)*100
No_Reading_Lost_Meter_Rate = (No_reading_Lost_meter_count/No_Reading_Meter_Total_Count)*100

#Overall LP Push Count Peformance
Expected_AllMeter_Total_LP_Count = (((Total_HighRiseMeter_Count+Total_VillageMeter_Count+unknownbuilding_Count)-Total_LDAMeter_Count)*48)+(Total_LDAMeter_Count*144)
AllMeter_Total_LP_Count = cr_list['NoOfIntervals'].sum()

Expected_HighRiseMeter_Total_LP_Count = Total_HighRiseMeter_Count*48
HighRiseMeter_Total_LP_Count = cr_highrise['NoOfIntervals'].sum()
Expected_VillageMeter_Total_LP_Count = Total_VillageMeter_Count*48
VillageMeter_Total_LP_Count = cr_village['NoOfIntervals'].sum()
Expected_AllCellMeter_Total_LP_Count = (Total_CellMeter_Count*48+Total_LDAMeter_Count*144)
AllCellMeter_Total_LP_Count = (onlycell_meter['NoOfIntervals'].sum() + LDA_meter['NoOfIntervals'].sum())
Expected_CellMeter_Total_LP_Count = (Total_CellMeter_Count*48)
CellMeter_Total_LP_Count = (onlycell_meter['NoOfIntervals'].sum())
Expected_LDAMeter_Total_LP_Count = (Total_LDAMeter_Count*144)
LDAMeter_Total_LP_Count = (LDA_meter['NoOfIntervals'].sum())
Full48_LP_Interval_AllMeter_Count = cr_list['NoOfIntervals'] == 48
Average_LP_Interval_Push_Count = cr_list['NoOfIntervals'].mean()
StdDev_LP_Interval_Push_Count = cr_list['NoOfIntervals'].std()
Full48_LP_Interval_AllMeter_Count = Full48_LP_Interval_AllMeter_Count.sum()
Full48_LP_Interval_HighRiseMeter_Count = cr_highrise['NoOfIntervals'] == 48
Full48_LP_Interval_HIghRiseMeter_Count = Full48_LP_Interval_HighRiseMeter_Count.sum()
Full48_LP_Interval_VillageMeter_Count = cr_village['NoOfIntervals'] == 48
Full48_LP_Interval_VillageMeter_Count = Full48_LP_Interval_VillageMeter_Count.sum()
Full144_LP_Interval_LDAMeter_Count = cr_list['NoOfIntervals'] == 144
Full144_LP_Interval_LDAMeter_Count = Full144_LP_Interval_LDAMeter_Count.sum()
Full48_LP_Interval_CellMeter_Count = onlycell_meter['NoOfIntervals'] == 48
Full48_LP_Interval_CellMeter_Count = Full48_LP_Interval_CellMeter_Count.sum()
Full48_LP_Interval_CellMeter_Count_Rate = round((Full48_LP_Interval_CellMeter_Count/Total_CellMeter_Count)*100,2)

Missing48_LP_unlocated_meter = unlocated_meter['NoOfIntervals'] != 48
Missing48_LP_unlocated_meter_Count = Missing48_LP_unlocated_meter.count()

#LP-DayEnd-FULL_Meter
LP_DayEnd_Full_Meter = cr_list[(cr_list['NoOfIntervals'] == 48)&(cr_list['DayEnd'] == 1)]
LP_DayEnd_Full_Meter_Count = LP_DayEnd_Full_Meter['meterno'].count()
LP_DayEnd_Full_Meter_Rate = round((LP_DayEnd_Full_Meter_Count/Total_AllMeter_Count)*100,2)

#Overall LP Push % Peformance
AllMeter_Total_LP_SuccessRate = (AllMeter_Total_LP_Count/Expected_AllMeter_Total_LP_Count)*100
Full48_LP_Interval_AllMeter_Rate = (Full48_LP_Interval_AllMeter_Count/Total_AllMeter_Count)*100
HighRiseMeter_Total_LP_SuccessRate = (HighRiseMeter_Total_LP_Count/Expected_HighRiseMeter_Total_LP_Count)*100
VillageMeter_Total_LP_SuccessRate = (VillageMeter_Total_LP_Count/Expected_VillageMeter_Total_LP_Count)*100
AllCellMeter_Total_LP_SuccessRate = (AllCellMeter_Total_LP_Count/Expected_AllCellMeter_Total_LP_Count)*100
CellMeter_Total_LP_SuccessRate = (CellMeter_Total_LP_Count/Expected_CellMeter_Total_LP_Count)*100
LDAMeter_Total_LP_SuccessRate = (LDAMeter_Total_LP_Count/Expected_LDAMeter_Total_LP_Count)*100

target_date = cr_list.iloc[0,7].strftime('%Y-%m-%d')

Performance = collections.OrderedDict({
'Execution Date': today_date,
'Target Date': target_date,
'[ KEY PERFORMANCE INDICATOR ]':'',
'Total Meter Count':Total_AllMeter_Count,
'Total Collector Count':Collector_Count,
'Total Meter FW24.60 Meter Count':all_meter_1468_Count,
'Total Meter FW24.60 Meter(%)':all_meter_1468_1468_Rate,
'All Meter LP Interval Push Success(%)':round(AllMeter_Total_LP_SuccessRate,2),
'All Meter DayEnd Reading Push Success(%)':round(AllMeter_Total_DayEnd_Reading_SuccessRate,2),
'Average LP Push Count':round(Average_LP_Interval_Push_Count,2),
'Std Deviation LP Push Count':round(StdDev_LP_Interval_Push_Count,2),   
'LP-DayEnd-FULL All Meter Count':LP_DayEnd_Full_Meter_Count,
'LP-DayEnd-FULL All Meter(%)':round(LP_DayEnd_Full_Meter_Rate,2),
'Full 48 LP Interval Meter Count':Full48_LP_Interval_AllMeter_Count,
'Full 48 LP Interval Meter(%)':round(Full48_LP_Interval_AllMeter_Rate,2),
'Full 48 LP Interval Cell Meter Count':Full48_LP_Interval_CellMeter_Count,
'Full 48 LP Interval Cell Meter(%)':Full48_LP_Interval_CellMeter_Count_Rate,
'NO DayEnd Reading All Meter Count':Missing_DayEnd_Reading_AllMeter_Count,
'NO DayEnd Reading Meter(%)':round(MissingDayEndReadingAllMeterRate,2),
'NO LP and DayEnd Reading Meter Count':No_LPandDayEnd_Reading_Meter_with_DayEnd_Count,
'NO LP and DayEnd Reading Meter(%)':round(No_LPandDayEnd_Reading_Meter_with_DayEnd_Rate,2),
'NO LP Reading Meter Count':No_Reading_Meter_Total_Count,
'NO LP Reading Meter Total(%)':round(No_Reading_Meter_Rate,2),
'NO LP Reading but with DayEnd Reading Meter Count':No_Reading_Meter_with_DayEnd_count,
'NO LP Reading but with DayEnd_Reading Meter(%)':round(NO_LPReading_ButWithDayEnd_Reading_Rate,2),
'NO DayEnd Reading but with LP Reading Meter Count':NO_DayEnd_Reading_but_with_LP_Reading_Meter_Count,
'NO DayEnd Reading but with LP Reading Meter(%)':round(NO_DayEnd_Reading_but_with_LP_Reading_Meter_Rate,2),
'[ SLA METERS PERFORMANCE (NORMAL FOR OVER 7DAYS) ]':'',
'SLA Meter Count':Effective_Meter_Count,
'SLA Meter LP Success(%)':Effective_Meter_LP_Success_Rate,
'SLA Meter Dayend Success(%)':Effective_Meter_Dayend_Success_Rate,
'SLA Meter Average LP Push Count':round(Effective_Meter_Average_LP_Interval_Push_Count,2),
'SLA Meter Std Deviation LP Push Count':round(Effective_Meter_StdDev_LP_Interval_Push_Count,2),
'LP-DayEnd-FULL SLA Meter Count':LP_DayEnd_Full_Effective_Meter_Count,
'LP-DayEnd-FULL SLA Meter(%)':LP_DayEnd_Full_Effective_Meter_Rate,
'SLA Meters Full 48 LP Interval Meter Count':EffectiveMeters_Full_48_LP_Interval_Meter_Count,
'SLA Meters Full 48 LP Interval Meter(%)':EffectiveMeters_Full_48_LP_Interval_Meter_Rate,
'SLA Meters Missing DayEnd Reading Meter Count':EffectiveMeters_Missing_DayEnd_Reading_Meter_Count,
'SLA Meters Missing DayEnd Reading Meter(%)':EffectiveMeters_Missing_DayEnd_Reading_Meter_Rate,
'[ LATEST METERS PERFORMANCE (REGISTERED IN LAST 30DAYS) ]':'',
'Latest Meter Count':Latest_Meters_Count,
'Latest Meter LP Success(%)':Latest_Meters_LP_Success_Rate,
'Latest Meter Dayend Success(%)':Latest_Meters_Dayend_Success_Rate,
'Latest Meter Average LP Push Count':round(Latest_Meters_Average_LP_Interval_Push_Count,2),
'Latest Meter Std Deviation LP Push Count':round(Latest_Meters_StdDev_LP_Interval_Push_Count,2),
'Latest Meters Full 48 LP Interval Meter Count':Latest_Meters_Full_48_LP_Interval_Meter_Count,
'Latest Meters Full 48 LP Interval Meter(%)':Latest_Meters_Full_48_LP_Interval_Meter_Rate,
'Latest Meters Missing DayEnd Reading Meter Count':Latest_Meters_Missing_DayEnd_Reading_Meter_Count,
'Latest Meters Missing DayEnd Reading Meter(%)':Latest_Meters_Missing_DayEnd_Reading_Meter_Rate,
'[ {} METERS SUMMARY ]':'',
'{} Meter Count':district_meter_Count,
'{} FW24.60 Meter Count':district_1468_Count,
'{} FW24.60 Meter(%)':district_1468_Rate,
'{} Meter LP Success(%)':district_meter_LP_Success_Rate,
'{} Meter Dayend Success(%)':district_meter_Dayend_Success_Rate,
'{} Average LP Push Count':round(district_meter_Average_LP_Interval_Push_Count,2),
'{} Std Deviation LP Push Count':round(district_meter_StdDev_LP_Interval_Push_Count,2),
'{} Meter LP-DayEnd-FULL Meter Count':district_meter_LP_DayEnd_Full_Meter_Count,
'{} Meter LP-DayEnd-FULL Meter(%)':district_meter_LP_DayEnd_Full_Meter_Rate,
'{} Meter Full 48 LP Interval Meter Count':district_meter_Full_48_LP_Interval_Meter_Count,
'{} Meter Full 48 LP Interval Meter(%)':district_meter_Full_48_LP_Interval_Meter_Rate,
'{} Meter Missing DayEnd Reading Meter Count':district_meter_Missing_DayEnd_Reading_Meter_Count,
'{} Meter Normal Meter Count':district_meter_Normal_Meter_Count,
'{} Meter SecConfig Meter Count':district_meter_SecConfig_Meter_Count,
'{} Meter Config Meter Count':district_meter_Config_Meter_Count,
'{} Meter Discovered Meter Count':district_meter_Discovered_Meter_Count,
'{} Meter Failed Meter Count':district_meter_Failed_Meter_Count,
'{} Meter Lost Meter Count':district_meter_Lost_Meter_Count,
'[ District B METERS SUMMARY ]':'',
'District B Meter Count':district_b_meter_Count,
'District B FW24.60 Meter Count':district_b_1468_Count,
'District B FW24.60 Meter(%)':district_b_1468_Rate,
'District B Meter LP Success(%)':district_b_meter_LP_Success_Rate,
'District B Meter Dayend Success(%)':district_b_meter_Dayend_Success_Rate,
'District B Average LP Push Count':round(district_b_meter_Average_LP_Interval_Push_Count,2),
'District B Std Deviation LP Push Count':round(district_b_meter_StdDev_LP_Interval_Push_Count,2),
'District B Meter LP-DayEnd-FULL Meter Count':district_b_meter_LP_DayEnd_Full_Meter_Count,
'District B Meter LP-DayEnd-FULL Meter(%)':district_b_meter_LP_DayEnd_Full_Meter_Rate,
'District B Meter Full 48 LP Interval Meter Count':district_b_meter_Full_48_LP_Interval_Meter_Count,
'District B Meter Full 48 LP Interval Meter(%)':district_b_meter_Full_48_LP_Interval_Meter_Rate,
'District B Meter Missing DayEnd Reading Meter Count':district_b_meter_Missing_DayEnd_Reading_Meter_Count,
'District B Meter Normal Meter Count':district_b_meter_Normal_Meter_Count,
'District B Meter SecConfig Meter Count':district_b_meter_SecConfig_Meter_Count,
'District B Meter Config Meter Count':district_b_meter_Config_Meter_Count,
'District B Meter Discovered Meter Count':district_b_meter_Discovered_Meter_Count,
'District B Meter Failed Meter Count':district_b_meter_Failed_Meter_Count,
'District B Meter Lost Meter Count':district_b_meter_Lost_Meter_Count,
'[ District C METERS SUMMARY ]':'',
'District C Meter Count':district_c_meter_Count,
'District C FW24.60 Meter Count':district_c_1468_Count,
'District C FW24.60 Meter(%)':district_c_1468_Rate,
'District C Meter LP Success(%)':district_c_meter_LP_Success_Rate,
'District C Meter Dayend Success(%)':district_c_meter_Dayend_Success_Rate,
'District C Average LP Push Count':round(district_c_meter_Average_LP_Interval_Push_Count,2),
'District C Std Deviation LP Push Count':round(district_c_meter_StdDev_LP_Interval_Push_Count,2),
'District C Meter LP-DayEnd-FULL Meter Count':district_c_meter_LP_DayEnd_Full_Meter_Count,
'District C Meter LP-DayEnd-FULL Meter(%)':district_c_meter_LP_DayEnd_Full_Meter_Rate,
'District C Meter Full 48 LP Interval Meter Count':district_c_meter_Full_48_LP_Interval_Meter_Count,
'District C Meter Full 48 LP Interval Meter(%)':district_c_meter_Full_48_LP_Interval_Meter_Rate,
'District C Meter Missing DayEnd Reading Meter Count':district_c_meter_Missing_DayEnd_Reading_Meter_Count,
'District C Meter Normal Meter Count':district_c_meter_Normal_Meter_Count,
'District C Meter SecConfig Meter Count':district_c_meter_SecConfig_Meter_Count,
'District C Meter Config Meter Count':district_c_meter_Config_Meter_Count,
'District C Meter Discovered Meter Count':district_c_meter_Discovered_Meter_Count,
'District C Meter Failed Meter Count':district_c_meter_Failed_Meter_Count,
'District C Meter Lost Meter Count':district_c_meter_Lost_Meter_Count,
'[ District D METERS SUMMARY ]':'',
'District D Meter Count':district_d_meter_Count,
'District D FW24.60 Meter Count':district_d_1468_Count,
'District D FW24.60 Meter(%)':district_d_1468_Rate,
'District D Meter LP Success(%)':district_d_meter_LP_Success_Rate,
'District D Meter Dayend Success(%)':district_d_meter_Dayend_Success_Rate,
'District D Average LP Push Count':round(district_d_meter_Average_LP_Interval_Push_Count,2),
'District D Std Deviation LP Push Count':round(district_d_meter_StdDev_LP_Interval_Push_Count,2),
'District D Meter LP-DayEnd-FULL Meter Count':district_d_meter_LP_DayEnd_Full_Meter_Count,
'District D Meter LP-DayEnd-FULL Meter(%)':district_d_meter_LP_DayEnd_Full_Meter_Rate,
'District D Meter Full 48 LP Interval Meter Count':district_d_meter_Full_48_LP_Interval_Meter_Count,
'District D Meter Full 48 LP Interval Meter(%)':district_d_meter_Full_48_LP_Interval_Meter_Rate,
'District D Meter Missing DayEnd Reading Meter Count':district_d_meter_Missing_DayEnd_Reading_Meter_Count,
'District D Meter Normal Meter Count':district_d_meter_Normal_Meter_Count,
'District D Meter SecConfig Meter Count':district_d_meter_SecConfig_Meter_Count,
'District D Meter Config Meter Count':district_d_meter_Config_Meter_Count,
'District D Meter Discovered Meter Count':district_d_meter_Discovered_Meter_Count,
'District D Meter Failed Meter Count':district_d_meter_Failed_Meter_Count,
'District D Meter Lost Meter Count':district_d_meter_Lost_Meter_Count,
'[ OVERALL METERS SUMMARY ]':'',
'Total HighRise Meter Count':Total_HighRiseMeter_Count,
'Total Village Meter Count':Total_VillageMeter_Count,
'Total All Cell Type Meter Count':Total_ALLCellMeter_Count,
'Total LDA Meter Count':Total_LDAMeter_Count,
'Total Cell Meter Count':Total_CellMeter_Count,
'Unlocated Meter Count':unlocated_meter_Count,
'[ OVERALL LP INTERVAL PUSH SUCCESS % ]':'',
'HighRise Meter Total LP Interval Push Success(%)':round(HighRiseMeter_Total_LP_SuccessRate,2),
'Village Meter Total LP Interval Push Success(%)':round(VillageMeter_Total_LP_SuccessRate,2),
'All Cell Meter Total LP Interval Push Success(%)':round(AllCellMeter_Total_LP_SuccessRate,2),
'Cell Meter Total LP Interval Push Success(%)':round(CellMeter_Total_LP_SuccessRate,2),
'LDA Meter Total LP Interval Push Success(%)':round(LDAMeter_Total_LP_SuccessRate,2),
'[ OVERALL DAYEND READING PUSH SUCCESS % ]':'',
'HighRise Meter Total DayEnd Reading Push Success(%)':round(HighRiseMeter_Total_DayEnd_Reading_SuccessRate,2),
'Village Meter Total DayEnd Reading Push Success(%)':round(VillageMeter_Total_DayEnd_Reading_SuccessRate,2),
'All Cell Meter Total DayEnd Reading Push Success(%)':round(AllCellMeter_Total_DayEnd_Reading_SuccessRate,2),
'Cell Meter Total DayEnd Reading Push Success(%)':round(CellMeter_Total_DayEnd_Reading_SuccessRate,2),
'LDA Meter Total DayEnd Reading Push Success(%)':round(LDAMeter_Total_DayEnd_Reading_SuccessRate,2),
'[ OVERALL NO LP READING METERS SUMMARY ]':'',
'NO LP Reading Highrise Meter Count':No_reading_meter_Highrise_count,
'NO LP Reading Village Meter Count':No_reading_meter_Village_count,
'NO LP Reading Unlocated Meter Count':No_reading_meter_Unlocated_count,
'NO LP Reading RF Meter Count':No_Reading_RF_meter_Count,
'NO LP Reading Cell Meter Count':No_Reading_Cell_meter_Count,
'NO LP Reading Normal Meter Count':No_reading_Normal_meter_count,
'NO LP Reading SecConfig Meter Count':No_reading_SecConfig_meter_count,
'NO LP Reading Config Meter Count':No_reading_Config_meter_count,
'NO LP Reading Discovered Meter Count':No_reading_Discovered_meter_count,
'NO LP Reading Failed Meter Count':No_reading_Failed_meter_count,
'NO LP Reading Lost Meter Count':No_reading_Lost_meter_count,
'[ NO LP PUSH READING METER COMPOSITION RATE ]':'',
'NO LP Reading HighRise Meter(%)':round(No_Reading_Meter_Highrise_Rate,2),
'NO LP Reading Village Meter(%)':round(No_Reading_Meter_Village_Rate,2),
'NO LP Reading Unlocated Meter(%)':round(No_reading_meter_Unlocated_Rate,2),
'NO LP Reading Normal Meter(%)':round(No_Reading_Normal_Meter_Rate,2),
'NO LP Reading SecConfig Meter(%)':round(No_Reading_SecConfig_Meter_Rate,2),
'NO LP Reading Configure Meter(%)':round(No_Reading_Config_Meter_Rate,2),
'NO LP Reading Discovered Meter(%)':round(No_Reading_Discovered_Meter_Rate,2),
'NO LP Reading Failed Meter(%)':round(No_Reading_Failed_Meter_Rate,2),
'NO LP Reading Lost Meter(%)':round(No_Reading_Lost_Meter_Rate,2),
'[ METER STATUS COUNT WITH READINGS ]':'',
'Normal Status Meter Count':Normal_Meter_Count,
'SecConfig Status Meter Count':SecConfig_Meter_Count,
'Configure Status Meter Count':Config_Meter_Count,
'Discovered Status Meter Count':Discovered_Meter_Count,
'Failed Status Meter Count':Failed_Meter_Count,
'Lost Status Meter Count':Lost_Meter_Count,
'[ OVERALL LP PUSH COUNT PERFORMANCE ]':'',
'Expected All Meter Total LP Interval Push Count': Expected_AllMeter_Total_LP_Count,
'All Meter Total LP Interval Push Count': AllMeter_Total_LP_Count,
'Expected HighRise Meter Total LP Interval Push Count':Expected_HighRiseMeter_Total_LP_Count,
'HighRise Meter Total LP Interval Push Count':HighRiseMeter_Total_LP_Count,
'Expected Village Meter Total LP Interval Push Count':Expected_VillageMeter_Total_LP_Count,
'Village Meter Total LP Interval Push Count':VillageMeter_Total_LP_Count,
'Expected All Cell Meter Total LP Interval Push Count':Expected_AllCellMeter_Total_LP_Count,
'All Cell Meter Total LP Interval Push Count':AllCellMeter_Total_LP_Count,
'Expected Cell Meter Total LP Interval Push Count':Expected_CellMeter_Total_LP_Count,
'Cell Meter Total LP Interval Push Count':CellMeter_Total_LP_Count,
'Expected LDA Meter Total LP Interval Push Count':Expected_LDAMeter_Total_LP_Count,
'LDA Meter Total LP Interval Push Count':LDAMeter_Total_LP_Count,
'Full 48 LP Interval HighRise Meter Count':Full48_LP_Interval_HIghRiseMeter_Count,
'Full 48 LP Interval Village Meter Count':Full48_LP_Interval_VillageMeter_Count,
'Full 144 LP Interval LDA Meter Count':Full144_LP_Interval_LDAMeter_Count,
'[ OVERALL DAYEND READING PUSH COUNT PERFORMANCE ]':'',
'Expected All Meter Total DayEnd Reading Push Count': Expected_AllMeter_Total_DayEnd_Reading_Count,
'All Meter Total DayEnd Reading Push Count': AllMeter_Total_DayEnd_Reading_Count,
'Expected HighRise Meter Total DayEnd Reading Push Count':Expected_HighRiseMeter_Total_DayEnd_Reading_Count,
'HighRise Meter Total DayEnd Reading Push Count':HighRiseMeter_Total_DayEnd_Reading_Count,
'Expected Village Meter Total DayEnd Reading Push Count':Expected_VillageMeter_Total_DayEnd_Reading_Count,
'Village Meter Total DayEnd Reading Push Count':VillageMeter_Total_DayEnd_Reading_Count,
'Expected All Cell Meter Total DayEnd Reading Push Count':Expected_AllCellMeter_Total_DayEnd_Reading_Count,
'All Cell Meter Total DayEnd Reading Push Count':AllCellMeter_Total_DayEnd_Reading_Count,
'Expected Cell Meter Total DayEnd Reading Push Count':Expected_CellMeter_Total_DayEnd_Reading_Count,
'Cell Meter Total DayEnd Reading Push Count':CellMeter_Total_DayEnd_Reading_Count,
'Expected LDA Meter Total DayEnd Reading Push Count':Expected_LDAMeter_Total_DayEnd_Reading_Count,
'LDA Meter Total DayEnd Reading Push Count':LDAMeter_Total_DayEnd_Reading_Count,
'Missing DayEnd Reading HighRise Meter Count':Missing_DayEnd_Reading_HighRiseMeter_Count,
'Missing DayEnd Reading Village Meter Count':Missing_DayEnd_Reading_VillageMeter_Count,
'Missing DayEnd Reading Cell Meter Count':Missing_DayEnd_Reading_AllCellMeter_Count,
'Missing DayEnd Reading Cell Meter Count':Missing_DayEnd_Reading_CellMeter_Count,
'Missing DayEnd Reading LDA Meter Count':Missing_DayEnd_Reading_LDAMeter_Count,
'[ METER TYPE COMPOSITION RATE ]':'',
'HighRise Meter(%)':round(HighRiseMeter_Rate,2),
'Village Meter(%)':round(VillageMeter_Rate,2),
'All Cell Meter(%)':round(AllCellMeter_Rate,2),
'Cell Meter(%)':round(CellMeter_Rate,2),
'LDA Meter(%)':round(LDAMeter_Rate,2),
'Unlocated Meter(%)':round(UnlocatedMeter_Rate,2),
'Normal Status Meter(%)':round(Normal_Meter_Rate,2),
'SecConfig Status Meter(%)':round(SecConfig_Meter_Rate,3),
'Configure Status Meter(%)':round(Config_Meter_Rate,3),
'Discovered Status Meter(%)':round(Discovered_Meter_Rate,3),
'Failed Status Meter(%)':round(Failed_Meter_Rate,3),
'Lost Status Meter(%)':round(Lost_Meter_Rate,3)})

df_performance = pd.DataFrame(pd.io.json.json_normalize(Performance).T)
df_performance.columns = ['Performance Result']

CR_perf = cr_list[cr_list['name'].str.startswith('8020', na=False)]
CR_perf = CR_perf.groupby(['name'])['NoOfIntervals'].mean()
CR_perf_std = cr_list.groupby(['name'])['NoOfIntervals'].std()
CR_perf = pd.concat([CR_perf, CR_perf_std], axis=1, join_axes=[CR_perf.index])
CR_perf = CR_perf.round()
CR_perf.columns = ['Average LP Count','Std LP Count']

CR_perf_district = district_meter[district_meter['name'].str.startswith('8020', na=False)]
CR_perf_district = CR_perf_district.groupby(['name'])['NoOfIntervals'].mean()
CR_perf_district_std = district_meter.groupby(['name'])['NoOfIntervals'].std()
CR_perf_district = pd.concat([CR_perf_district, CR_perf_district_std], axis=1, join_axes=[CR_perf_district.index])
CR_perf_district = CR_perf_district.round()
CR_perf_district.columns = ['Average LP Count','Std LP Count']

CR_perf_district_b = district_b_meter[district_b_meter['name'].str.startswith('8020', na=False)]
CR_perf_district_b = CR_perf_district_b.groupby(['name'])['NoOfIntervals'].mean()
CR_perf_district_b_std = district_b_meter.groupby(['name'])['NoOfIntervals'].std()
CR_perf_district_b = pd.concat([CR_perf_district_b, CR_perf_district_b_std], axis=1, join_axes=[CR_perf_district_b.index])
CR_perf_district_b = CR_perf_district_b.round()
CR_perf_district_b.columns = ['Average LP Count','Std LP Count']

CR_perf_district_c = district_c_meter[district_c_meter['name'].str.startswith('8020', na=False)]
CR_perf_district_c = CR_perf_district_c.groupby(['name'])['NoOfIntervals'].mean()
CR_perf_district_c_std = district_c_meter.groupby(['name'])['NoOfIntervals'].std()
CR_perf_district_c = pd.concat([CR_perf_district_c, CR_perf_district_c_std], axis=1, join_axes=[CR_perf_district_c.index])
CR_perf_district_c = CR_perf_district_c.round()
CR_perf_district_c.columns = ['Average LP Count','Std LP Count']

CR_perf_district_d = district_d_meter[district_d_meter['name'].str.startswith('8020', na=False)]
CR_perf_district_d = CR_perf_district_d.groupby(['name'])['NoOfIntervals'].mean()
CR_perf_district_d_std = district_d_meter.groupby(['name'])['NoOfIntervals'].std()
CR_perf_district_d = pd.concat([CR_perf_district_d, CR_perf_district_d_std], axis=1, join_axes=[CR_perf_district_d.index])
CR_perf_district_d = CR_perf_district_d.round()
CR_perf_district_d.columns = ['Average LP Count','Std LP Count']

dir = 'C:/Users/Desktop/RF_Analysis/SSR/'
writer = pd.ExcelWriter('%sReading_Performance_Report%s_%s.xlsx' % (dir,target_date,today_date))
df_performance.to_excel(writer, "Performance Report")
cr_list.to_excel(writer,"Analyzed Individual Meters", index=False)
CR_perf.to_excel(writer,"CR Performance")
CR_perf_district.to_excel(writer,"{} CR Performance")
CC_CR_Rnk.to_excel(writer,"{} CR ABC Rank")
CR_perf_district_b.to_excel(writer,"District B CR Performance")
TO_CR_Rnk.to_excel(writer,"District B CR ABC Rank")
CR_perf_district_c.to_excel(writer,"District C CR Performance")
TM_CR_Rnk.to_excel(writer,"District C CR ABC Rank")
CR_perf_district_d.to_excel(writer,"District D CR Performance")
TC_CR_Rnk.to_excel(writer,"District D CR ABC Rank")
Latest_Meters.to_excel(writer, "LatestMeters")
fw_perf.to_excel(writer, "FW Performance")
region_perf.to_excel(writer, "Region Performance")
area_perf.to_excel(writer, "Area Performance")

workbook  = writer.book
worksheet = writer.sheets['Performance Report']
cell_format = workbook.add_format()
cell_format.set_align('right')
worksheet.set_column(0,0,68, cell_format)
worksheet.set_column(1,10,17)
worksheet = writer.sheets['Region Performance']
worksheet.set_column(0,0,27)
writer.save()

df_performance
