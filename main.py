from createDocs import creatWord
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
# import time
import datetime
import csv
# import collections
# import gspread
# import collections, functools, operator
# from docx import Document
# from docx.shared import Inches
from collections import defaultdict
from itertools import groupby
from operator import itemgetter
from tkinter import *
import calendar
import os
# import xlsxwriter as exwriter


# ------user input box(need to update below info based on the clients)------------
# ====================================common data for SLA====================
sla_date = 5
sla_year = 2020
# client_name = "FJB"
# client_name = "VINX-WAN"
# client_name = "VINX-EMONEY"
client_name = "PMP"

# =======================================PMP SLA==========================
if "PMP" in client_name:
    RD = "Resources_folder\\File_resources-PMP"
    MasterSheet = RD + "\\PMP MASTER REPORT.csv"
    IncidentSheet = RD + "\\PMP Incident Report.csv"
    SLA = 99.97
    # 0---->N0  1---->Site Name  6---->TNS Router ID   8---->Wired ISP
    # 15---->SIM ISP  20---->SIM ISP
    MasterCo = [0, 1, 6, 8, 15, 20]
    # 0---->N0  1---->ticket   5---->site    9---->service impact start
    # 10---->service impact end  11---->service downtime  14---->ROOT CAUSE
    IncidentCol = [0, 1, 5, 9, 10, 11, 14]
# ================================FJB SLA================================
if "FJB" in client_name:
    RD = "Resources_folder\\File_resources-FJB"
    MasterSheet = RD + "\\FJ BENJAMIN MASTERSHEET.csv"
    IncidentSheet = RD + "\\FJB Incident Report.csv"
    SLA = 99.97
    # 0---->N0  1---->site code   2---->site name   3---->store type
    # 4---->region  12---->primary line   12----->secondary line
    MasterCo = [0, 1, 2, 3, 4, 12, 18]
    # 0---->N0  1---->ticket   4---->site    8---->service impact start
    # 9---->service impact end  10---->service downtime  12---->ROOT CAUSE
    IncidentCol = [0, 1, 4, 8, 9, 10, 12]
# =====================================VINX-Emoney SLA=========================
if "VINX-EMONEY" in client_name:
    RD = "Resources_folder\\File_resources-VINXEMONEY"
    MasterSheet = RD + "\\AEON (VINX) MASTERSHEET.csv"
    IncidentSheet = RD + "\\Vinx Incident Report.csv"
    SLA = 99.7
    # 0---->N0  1---->site code   2---->site name   3---->store type
    # 5---->region  13---->primary line   17----->secondary line   24---->secondary 3G
    MasterCo = [0, 1, 2, 3, 5, 13, 17, 24]
    # 0---->N0  1---->ticket   5---->site    9---->service impact start
    # 10---->service impact end  11---->service downtime  13---->ROOT CAUSE
    IncidentCol = [0, 1, 5, 9, 10, 11, 13]
# =====================================VINX-WAN===============================
if "VINX-WAN" in client_name:
    RD = "Resources_folder\\File_resources-VINXWAN"
    MasterSheet = RD + "\\AEON (VINX) MASTERSHEET.csv"
    IncidentSheet = RD + "\\Vinx Incident Report.csv"
    SLA = 99.7
    # 0---->N0  1---->site code   2---->site name   3---->store type
    # 6---->region  10---->primary line   16----->secondary line
    MasterCo = [0, 1, 2, 3, 6, 10, 16]
    # 0---->N0  1---->ticket   5---->site    9---->service impact start
    # 10---->service impact end  11---->service downtime  13---->ROOT CAUSE
    IncidentCol = [0, 1, 5, 9, 10, 11, 13]


# -------------------------------------sub functions---------------------
# this function open the input csv file return all the data in array list


def open_csv(file_text):
    with open(file_text) as csvfile:
        read_csv = csv.reader(csvfile, delimiter=',')
        datas = []
        for i, row1 in enumerate(read_csv):
            if i > 0:
                date = row1
                datas.append(date)

    return datas

# this function open the input csv file return the specific column data in array list


def open_csv_specifc_column(file_text, col_list):
    with open(file_text) as csvfile:
        read_csv = csv.reader(csvfile, delimiter=',')
        print(col_list)
        datas = []
        for i, row1 in enumerate(read_csv):
            if i >= 1:
                date = []
                for z in col_list:
                    date.append(row1[z])
                datas.append(date)
    return datas


def min_calculate(str1):
    res1 = str1.split()
    min_value = (int(res1[0])*1440)+(int(res1[2])*60)+int(res1[4])
    return min_value


def find_name(str1, cl):
    sn = None
    if "::" in str1:
        sitename = str1.split("::")
        sn = sitename[len(sitename)-1]
        if "VINX-WAN" in cl or "VINX-EMONEY" in cl:
            sp1 = sn.split("_")
            sn = sp1[len(sp1) - 1]
        elif "PMP" in cl:
            sn = str(sitename[1]).replace(" ", "")

    return sn


def set_colors(row1):
    c1 = []
    for it in row1:
        # print(it)
        if it >= SLA:
            c1.append('DodgerBlue')
        else:
            # c1.append('red')
            c1.append('DodgerBlue')

    return c1


def auto_label(ax, rects):
    """Attach a text label above each bar in *rects*, displaying its height."""
    for rect in rects:
        height = rect.get_height()
        ax.annotate('{}'.format(height),
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    # 3 points vertical offset
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom')


def make_multiline(list1):
    outputlist = []
    for st in list1:
        outputlist.append(st.replace('-', '\n'))
    tlist = tuple(outputlist)
    return tlist


def check_dict_exist(dict1, listd):
    output1 = False
    ind1 = None
    if len(listd) > 0:
        for ix, i in enumerate(listd):
            if dict1['SITE'] in i[0]:
                output1 = True
                ind1 = ix
    return [output1, ind1]


def combine_result(ld1, ld2):
    output = []
    for i in ld1:
        for j in ld2:
            if i['SITE'] == j['SITE']:
                dic_is_exist = check_dict_exist(i, output)
                if dic_is_exist[0] is False:
                    output.append([i['SITE'], i['value'], j['value']])
                elif dic_is_exist[0] is True:
                    print(dic_is_exist[1])
                    output[dic_is_exist[1]] = [i['SITE'], i['value'], j['value']]
            else:
                dic_is_exist = check_dict_exist(i, output)
                if dic_is_exist[0]==False:
                    output.append([i['SITE'], i['value'], 0])
                dic_is_exist = check_dict_exist(j, output)
                if dic_is_exist[0]==False:
                    output.append([j['SITE'], 0, j['value']])
    return output


def creatingFinalList(mdf,R1):
    FL = []
    for ind in mdf.index:
        Templist = []
        connectivity = ""
        if len(mdf.columns) == 8:
            if "N/A" in mdf['secondary'][ind]:
                connectivity = mdf['primary'][ind]
            else:
                if "3G" in mdf['secondary 3G'][ind]:
                    connectivity = mdf['primary'][ind] + "/" + mdf['secondary'][ind] + "/" + "3G"
                elif "4G" in mdf['secondary 3G'][ind]:
                    connectivity = mdf['primary'][ind] + "/" + mdf['secondary'][ind] + "/" + "4G"

        elif len(mdf.columns) == 7:
            if "N/A" in mdf['secondary'][ind]:
                connectivity = mdf['primary'][ind]
            else:
                connectivity = mdf['primary'][ind] + "/" + mdf['secondary'][ind]

        elif len(mdf.columns) == 6:
            if "N/A" in mdf['secondary 3G'][ind] :
                connectivity = mdf['primary'][ind] + "/" + mdf['secondary'][ind]
            else:
                connectivity = mdf['primary'][ind] + "/" + mdf['secondary'][ind]+ "/" + mdf['secondary 3G'][ind]


        if any(d1[0] in  mdf['Site Name'][ind] for i1, d1 in enumerate(R1)) :

            for i2, d3 in enumerate(R1):
               if d3[0] == mdf['Site Name'][ind]:
                    xx = i2
            no=mdf['NO'][ind]
            sn=mdf['Site Name'][ind]
            lastday = calendar.monthrange(sla_year, sla_date)[1]
            fdate='01/' + str(sla_date) + '/' + str(sla_year)
            tdate=str(lastday) + '/' + str(sla_date) + '/' + str(sla_year)
            totalmin = lastday * 24 * 60
            WANdowntime = float(R1[xx][1])
            CUSdowntime = float(R1[xx][2])
            uptimemin =totalmin - WANdowntime
            Totaldown=WANdowntime+CUSdowntime
            uptimePercentage = round((uptimemin / totalmin) * 100,3)

            # if "FJB" in client_name or "VINX-EMONEY" in client_name:
            Tempdict = [no, sn,connectivity,fdate,tdate, str(uptimemin), str(totalmin),str(WANdowntime),str(CUSdowntime),str(Totaldown),str(uptimePercentage)]
            FL.append(Tempdict)

        else:

            no = mdf['NO'][ind]
            sn = mdf['Site Name'][ind]
            lastday = calendar.monthrange(sla_year, sla_date)[1]
            fdate ='01/' + str(sla_date) + '/' + str(sla_year)
            tdate = str(lastday) + '/' + str(sla_date) + '/' + str(sla_year)

            totalmin = lastday * 24 * 60
            uptimemin = totalmin
            WANdowntime = 0
            CUSdowntime = 0
            Totaldown = WANdowntime + CUSdowntime
            uptimePercentage = round((uptimemin / totalmin) * 100, 3)

            # if "FJB" in client_name or "VINX-EMONEY" in client_name:
            Tempdict = [no, sn, connectivity, fdate, tdate, str(uptimemin), str(totalmin), str(WANdowntime),
                            str(CUSdowntime), str(Totaldown), str(uptimePercentage)]
            FL.append(Tempdict)
    return FL


# -----------------------------------------Main code------------------------------------
# --------------------------------------------------------------------------------------

Incident_Row_data = open_csv_specifc_column(IncidentSheet,IncidentCol)
Master_Row_data = open_csv_specifc_column(MasterSheet,MasterCo)
# client name may be "PMP", "FJB", "VINX"

if "VINX-EMONEY" in client_name:
  mdf = pd.DataFrame(Master_Row_data,
                     columns=['NO', 'Site Code', 'Site Name', 'Store Type', 'Region',
                              'primary', 'secondary', 'secondary 3G'])
elif "FJB" in client_name or "VINX-WAN" in client_name:
  mdf = pd.DataFrame(Master_Row_data,
                     columns=['NO', 'Site Code', 'Site Name', 'Store Type', 'Region', 'primary', 'secondary'])
elif "PMP" in client_name:
    mdf = pd.DataFrame(Master_Row_data,
                       columns=['NO', 'Site Name', 'Router ID', 'primary', 'secondary', 'secondary 3G'])
# df=pd.DataFrame(Row_data,columns=['NO','TICKET','WEEK','SEVERITY','SITE','SOURCE','INCIDENT START','INCIDENT END','SERVICE-IMPACT START (10AM-10PM)','SERVICE-IMPACT END (10AM-10PM)','SERVICE DOWNTIME','ONSITE','ROOT CAUSE','REMARKS','REVISED BY'])
idf = pd.DataFrame(Incident_Row_data,
                   columns=['NO', 'TICKET', 'SITE', 'SERVICE-IMPACT START',
                            'SERVICE-IMPACT END', 'SERVICE DOWNTIME', 'ROOT CAUSE'])


# fdf = pd.DataFrame(columns=['NO','TICKET','SITE','SERVICE-IMPACT START','SERVICE-IMPACT END','SERVICE DOWNTIME','ROOT CAUSE','Total DownTime_min']) #result after calculation of total min
incidentlist=[] #wan incident
incidentlistCUS=[] #customer incident
for x in idf.index:
    if  idf['SERVICE-IMPACT START'][x]!="" and idf['SERVICE-IMPACT END'][x]!="" :
        if  "WAN" in idf['ROOT CAUSE'][x] or "wan" in idf['ROOT CAUSE'][x]:
            totalmin= min_calculate(idf['SERVICE DOWNTIME'][x])
            print("{} >>>>{}-----------{}-----{}".format(x, idf['SITE'][x], idf['SERVICE DOWNTIME'][x], totalmin))
            incidentlist.append({'NO':idf['NO'][x],'TICKET':idf['TICKET'][x],'SITE':find_name(idf['SITE'][x],client_name),'SERVICE-IMPACT START':idf['SERVICE-IMPACT START'][x],'SERVICE-IMPACT END':idf['SERVICE-IMPACT END'][x], 'SERVICE DOWNTIME':idf['SERVICE DOWNTIME'][x],'ROOT CAUSE': idf['ROOT CAUSE'][x],'Total DownTime_min':totalmin})
        else :
            totalmin= min_calculate(idf['SERVICE DOWNTIME'][x])
            print("{} >>>>{}-----------{}-----{}".format(x, idf['SITE'][x], idf['SERVICE DOWNTIME'][x], totalmin))
            incidentlistCUS.append({'NO':idf['NO'][x],'TICKET':idf['TICKET'][x],'SITE':find_name(idf['SITE'][x],client_name),'SERVICE-IMPACT START':idf['SERVICE-IMPACT START'][x],'SERVICE-IMPACT END':idf['SERVICE-IMPACT END'][x], 'SERVICE DOWNTIME':idf['SERVICE DOWNTIME'][x],'ROOT CAUSE': idf['ROOT CAUSE'][x],'Total DownTime_min':totalmin})


print("-------------------------------")
for x in incidentlist:
     print(x)
for x in incidentlistCUS:
    print(x)

result = defaultdict(int)
get_name = itemgetter('SITE')
result = [{'SITE': name, 'value': str(sum(int(d['Total DownTime_min']) for d in dicts))}
           for name, dicts in groupby(sorted(incidentlist, key=get_name), key=get_name)]

resultCUS = defaultdict(int)
get_nameCUS = itemgetter('SITE')
resultCUS = [{'SITE': name, 'value': str(sum(int(d['Total DownTime_min']) for d in dicts))}
           for name, dicts in groupby(sorted(incidentlistCUS, key=get_nameCUS), key=get_nameCUS)]
finalresult=combine_result(result,resultCUS)


print("========================================================")
for x in finalresult:
    print(x)

print("=================================================")

final_list = []
final_list = creatingFinalList(mdf,finalresult)

print("final list has created")
for x in final_list:
    print(x)

path = os.getcwd()+"\Output_folder"
outputdirpath = path + ("\{}_{}_OutputFiles_{}").format(datetime.date(1900, sla_date, 1).strftime('%B'),client_name,datetime.datetime.now().strftime("%b%d%Y%H%M%S" ))
# if os.path.isdir(outputdirpath):
os.mkdir(outputdirpath)

if len(final_list)>0:
  count=0
  finaldfTable=pd.DataFrame(final_list,columns=['NO','Site Name','Connectivity','From Date','To Date','Exact Site Uptime (mins)','Total Availability Time-Business Hours (mins)','Total downtime_WAN','Customer Downtime (mins)','Total Downtime(wan+customer)(mins)','Uptime (%)'])
  print("befor list  pop============================ ",len(final_list))

  # for items in final_list:
  #     print(items )

  count = 1

  while len(final_list)> 0:
     row = ()
     colorlist=[]
     max1=100
     min1=0
     res=10
     Lables1 = ()
     print("after list  pop============================ ", len(final_list))
     sfl=len(final_list)
     if sfl > 10:
       print(("{} >>>>>>>>>>>>> 10").format(sfl))
       row = ((float(final_list[0][10]), float(final_list[1][10]), float(final_list[2][10]), float(final_list[3][10]),
            float(final_list[4][10]),float(final_list[5][10]), float(final_list[6][10]), float(final_list[7][10]), float(final_list[8][10]),
            float(final_list[9][10]),))
       # print(row)

       lables1= (final_list[0][1], final_list[1][1], final_list[2][1], final_list[3][1], final_list[4][1],
                     final_list[5][1], final_list[6][1], final_list[7][1], final_list[8][1], final_list[9][1])


       colorlist = set_colors(row)
       N = 10
       ii = 0
       max1=max(row)
       min1 = min(row)
       if max1 - min1 >= 10.0 or max1 - min1 ==0:
           res = 10

       elif max1 - min1 < 10.0 and max1 - min1 > 1:
           res = 1
       elif max1 - min1 < 1.0 and max1 - min1>0:
           res = 0.1
       if max1 - min1 ==0:
           min1 = 0
       while ii < 10:
          # final_list.pop(0)
          del final_list[0]
          ii += 1


     if sfl <=10:
        print(("{} <<<<<<<<<<< 10").format(sfl))
        N=len(final_list)
        ii = len(final_list)-1
        row = ()
        lables1 = ()
        while ii >=0:

            row = row + (float(final_list[ii][10]),)
            lables1 = lables1 + (final_list[ii][1],)
            colorlist = set_colors(row)
            # final_list.pop(ii)
            del final_list[ii]
            ii -= 1
        max1 = max(row)
        min1 = min(row)-0.05
        if max1-min1>=10.0 or max1 - min1 ==0:
            res=10

        elif max1-min1<10.0 and   max1-min1>1:
            res=1
        elif  max1-min1<1.0 :
            res=0.1
        if max1 - min1 == 0:
            min1 = 0

     ind = np.arange(N)
     print("min= ",min1,"----","max=",max1,"---------","res=",res)
     print(row)

     # ============================
     ind = np.arange(N)
     width = 0.60
     fig = plt.figure(figsize=(12, 4))
     print(">>>>>>>>>")
     lables2 = make_multiline(lables1)

     x = np.arange(len(lables1))
     s = fig.add_subplot(111)

     rects1=s.bar(ind,np.array(row), width,color=colorlist)
     auto_label(s,rects1)
     s.set_ylabel('Uptime %',fontsize=12)
     s.set_xticks(x)
     s.set_xticklabels( lables2,fontsize=12,rotation=90,ha='right',rotation_mode="anchor")
     s.set_title('SLA(%) GRAPH\n')
     if min1==0:
      # s.set_ylim(min1, max1 )
      s.set_ylim(95, 100)
     else:
      # s.set_ylim(min1-0.05, max1)
      s.set_ylim(95, 100)

     fig.savefig(('{}\{}.png').format(outputdirpath,count),bbox_inches = "tight",dpi=350)
     noOFgraph=count
     count+=1
  print("final============================ ", len(final_list))
  print("bar charts Created...")

  #creating excel file
  writer = pd.ExcelWriter(('{}\{}.xlsx').format(outputdirpath, client_name), engine='xlsxwriter')
  pd_incidentlist=pd.DataFrame(incidentlist)
  pd_incidentlistCUS = pd.DataFrame(incidentlistCUS)

  finaldfTable.to_excel(writer, 'main table')
  pd_incidentlist.to_excel(writer, 'incidentlist_wan')
  pd_incidentlistCUS.to_excel(writer, 'incidentlist_customer')
  writer.save()
  #==========================================================================
  # creating word Docs
  CW=creatWord(sla_date,sla_year,finaldfTable,outputdirpath,client_name,noOFgraph,SLA,RD)
  CW.creatingfunction_word()
  #=========================================================================
  print("output file created in word format... ")
