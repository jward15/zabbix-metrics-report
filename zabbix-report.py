#! /usr/bin/python
# -*- coding: utf-8 -*-
import sys
import MySQLdb
import time
import datetime
import calendar
import os
import os.path

from tempfile import TemporaryFile
from xlwt import Workbook,easyxf
from xlrd import open_workbook

#Today's Date
today = datetime.date.today()

#Contents of report
report_dir="/work/opt/zabbix-reports"
#MySQL Infomation
host='localhost'
port=3306
user = 'readonly'
password = 'Reporting123##'
database = 'zabbix'

#To generate executed KEY, save all of the info 
keys = ("cpuload","disk_usage","network_in","network_out")
#In addition to boolen type ( such as vip) monitoring items , the thresholds are set , the maximum value within the operating range of the report , showing a red background
#All no separate judgment key, which must define the threshold in this dictionary
#-------------------------------------------------------------
# NIC flow threshold : 50M / s 
#-------------------------------------------------------------
thre_dic = {"cpuload":15,"disk_usage":85,"network_in":409600}

#-------------------------------------------------------------
#Custom Report：
#-------------------------------------------------------------
def custom_report(startTime,endTime):
  
	sheetName =  time.strftime('%m%d_%H%M',startTime) + "_TO_" +time.strftime('%m%d_%H%M',endTime)
	customStart = time.mktime(startTime)
	customEnd = time.mktime(endTime)
	generate_excel(customStart,customEnd,0,sheetName)

#-------------------------------------------------------------
# Daily report :
# 	To execute the script and the current unix timestamp amd to extract statements
#	Script must be run before midnight
#-------------------------------------------------------------
def daily_report():
#	today = datetime.date.today() Get today's date
	dayStart = time.mktime(today.timetuple()) 
	dayEnd = time.time() #Get the current system unix timestamp
	sheetName = time.strftime('%Y%m%d',time.localtime(dayEnd))
    	generate_excel(dayStart,dayEnd,1,sheetName)    
#-------------------------------------------------------------
# Generate reports by week
#-------------------------------------------------------------

def weekly_report():
	lastMonday = today
#	lastMonday = datetime.date.today()#Get today's date
	#Grabs information on Monday
	while lastMonday.weekday() != calendar.MONDAY:
		lastMonday -= datetime.date.resolution
	
	weekStart = time.mktime(lastMonday.timetuple())#Will get data on Monday at midnight via unix timestamp
	weekEnd = time.time()#Get the current system unix timestamp
	#weekofmonth = (datetime.date.today().day+7-1)/7
	weekofmonth = (today.day+7-1)/7
	sheetName = "weekly_" + time.strftime('%Y%m',time.localtime(weekEnd)) + "_" + str(weekofmonth)
	generate_excel(weekStart,weekEnd,2,sheetName)
			
#-------------------------------------------------------------
# Generate reports by Month
#-------------------------------------------------------------

def monthly_repport():
##	firstDay =  datetime.date.today() #The first day of the current date
	firstDay =  today #The first day of the current date
	#The first day of the month of acquisition date
	while firstDay.day != 1:
		firstDay -= datetime.date.resolution
	monthStart = time.mktime(firstDay.timetuple()) #The first day of the month via unix timestamp
	monthEnd = time.time()	#Current unix timestamp
	sheetName = "monthly_" + time.strftime('%Y%m',time.localtime(monthEnd))
	generate_excel(monthStart,monthEnd,3,sheetName)
	

#-------------------------------------------------------------
#  Obtain MySQL Connection
#-------------------------------------------------------------
def getConnection():
       # print "Ready to connect to MySQL "
        try:
                connection=MySQLdb.connect(host=host,port=port,user=user,passwd=password,db=database,connect_timeout=1);
        except MySQLdb.Error, e:
                print "Error %d: %s" % (e.args[0], e.args[1])
                sys.exit(1)
	return connection

#-------------------------------------------------------------
# Back to all IP hosts and hostid, such as :( '192.168.10.62', 10113L, 0), which Role for the field to add , 1: M, 2: S, 3: N
#-------------------------------------------------------------
def getHosts():
	conn=getConnection()
	cursor = conn.cursor()
	command = cursor.execute("""select ip,hostid,Role from hosts where ip<>'127.0.0.1' and ip<>'' and status=0 order by ip;""");
	hosts = cursor.fetchall()
	cursor.close()
	conn.close()
	return hosts

#-------------------------------------------------------------
# Item Returns the specified host monitoring of itmeid,
#-------------------------------------------------------------
def getItemid(hostid):
	keys_str = "','".join(keys)
	conn=getConnection()
	cursor = conn.cursor()
	command = cursor.execute("""select itemid from items where hostid=%s and key_ in ('%s')""" %(hostid,keys_str));
	itemids =  cursor.fetchall()
	cursor.close()
	conn.close()
	return itemids
#-------------------------------------------------------------
# Returns None specified hostid reports the value of the host , only for digital history table
#-------------------------------------------------------------

def getReportById_1(hostid,start,end):
	keys_str = "','".join(keys)
        conn=getConnection()
        cursor = conn.cursor()
	command = cursor.execute("""select items.itemid , key_ as key_value ,units, max(history.value) as max,avg(history.value) as average ,min(history.value) as min  from history, items where items.hostid=%s and items.key_ in ('%s')and items.value_type=0  and history.itemid=items.itemid  and (clock>%s and clock<%s)  group by itemid, key_value;""" %(hostid,keys_str,start,end));	
	values =  cursor.fetchall()
        cursor.close()
	conn.close();
	return values

#-------------------------------------------------------------
# Returns None specified host hostid reports the value , only the needed unsigned history_uint table , items.value_type = 3
#-------------------------------------------------------------

def getReportById_2(hostid,start,end):
        keys_str = "','".join(keys)
        conn=getConnection()
        cursor = conn.cursor()
        command = cursor.execute("""select items.itemid , key_ as key_value ,units, max(history_uint.value) as max,avg(history_uint.value) as average ,min(history_uint.value) as min  from history_uint, items where items.hostid=%s and items.key_ in ('%s')and items.value_type=3  and history_uint.itemid=items.itemid and (clock>%s and clock<%s) group by itemid, key_value;""" %(hostid,keys_str,start,end));
        values =  cursor.fetchall()
        cursor.close()
        conn.close();
        return values
#--------------------------------------------------------------
# File : generate Excel reports
# Parameters , start: extract data start time , end: End point in time be able to get data
# ReportType: report type : 1 daily, 2 weekly, 3 monthly
#-----------------------------------------------------------------

def generate_excel(start,end,reportType,sheetName):
	book = Workbook(encoding='utf-8')
	sheet1 = book.add_sheet(sheetName)	
	merge_col = 1
	merge_col_step = 2

	title_col = 1
	title_col_step = 2
	
	hosts = getHosts()
	isFirstLoop=1
	host_row = 2 #host ip
	
	max_col = 1
	avg_col = 2
	
	#This is to format Excel
	normal_style = easyxf(
'borders: right thin,top thin,left thin, bottom thin;'
'align: vertical center, horizontal center;'
)
	abnormal_style = easyxf(
'borders: right thin, bottom thin,top thin,left thin;'
'pattern: pattern solid, fore_colour red;'
'align: vertical center, horizontal center;'
)


	sheet1.write_merge(0,1,0,0,"HOSTS")
	for ip,hostid,role in hosts:
		sheet1.row(host_row).set_style(normal_style)
		max_col = 1
	        avg_col = 2
		reports = getReportById_1(hostid,start,end) + getReportById_2(hostid,start,end)
		if(isFirstLoop==1): # The first time through the loop will write to the header
			sheet1.write(host_row,0,ip,normal_style)
			for report in reports:
				title = report[1] + " " + report[2]		
				sheet1.write_merge(0,0,merge_col,merge_col+1,title,normal_style)
				merge_col += merge_col_step
					
				sheet1.write(1,title_col,"MAX",normal_style)
				sheet1.write(1,title_col+1,"Average",normal_style)
				title_col += title_col_step
		
				# Writes data , determines whether the maximum value exceeds a specified threshold
                                # When the maximum value is greater than the specified threshold value , this is displayed in red
				if(report[3] >= thre_dic[report[1]]):
					sheet1.write(host_row,max_col,report[3],abnormal_style)
					sheet1.write(host_row,avg_col,report[4],normal_style)
				else:	# Does not exceed the threshold 
					sheet1.write(host_row,max_col,report[3],normal_style)
					sheet1.write(host_row,avg_col,report[4],normal_style)
				max_col = max_col + 2
				avg_col =avg_col+ 2
				isFirstLoop=0	
		else:
			sheet1.write(host_row,0,ip,normal_style)
			for report in reports:
				# When the maximum value is greater than the specified threshold value , this is displayed in red
                        	if(report[3] >= thre_dic[report[1]]):
                                	sheet1.write(host_row,max_col,report[3],abnormal_style)
                                        sheet1.write(host_row,avg_col,report[4],normal_style)
                        	 else: #Does not exceed the threshold then the normal display
                                 	sheet1.write(host_row,max_col,report[3],normal_style)
                                 	sheet1.write(host_row,avg_col,report[4],normal_style)
	
                        	max_col = max_col + 2
                                avg_col =avg_col+ 2

		host_row = host_row +1
	saveReport(reportType,book)

#----------------------------------------------------------------------
# Functions: Depending on the type of report , to implement different ways to save
# Parameters : reportType report types : 0 custom 1 daily, 2 weekly, 3 monthly
# WorkBook the current Excel workbook	
#---------------------------------------------------------------------
def saveReport(reportType,workBook):
	#Reports the directory exists , if it does not exist, then it creates new
	if(not (os.path.exists(report_dir))):
		os.makedirs(report_dir)
	#Switch 
	os.chdir(report_dir)
	#Reporting units per month stored in the directory 
	month_dir=time.strftime('%Y-%m',time.localtime(time.time()))
	if(not (os.path.exists(month_dir))):
		os.mkdir(month_dir)
	os.chdir(month_dir)
	#Custom Report	
	if(reportType == 0):
		excelName = "custom_report_"+ time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) + ".xls"		
	#Daily
	elif(reportType == 1):
		excelName = "daily_report_" + time.strftime('%Y%m%d',time.localtime(time.time())) + ".xls"
	#Weekly
	elif(reportType == 2):
		#weekofmonth = (datetime.date.today().day+7-1)/7			
		weekofmonth = (today.day+7-1)/7			
		excelName = "weekly_report_" +  time.strftime('%Y%m',time.localtime(time.time())) +"_" + str(weekofmonth) + ".xls"
	#Monthly
	else:
		monthName = time.strftime('%Y%m',time.localtime(time.time()))
		excelName = "monthly_report_" + monthName + ".xls"
#		currentDir = os.getcwd()
#		files = os.listdir(currentDir)# The default is the current directory , which is the month directory
#		for file in files:
#			wb = open_workbook(file)
	print excelName				
	workBook.save(excelName)

#----------------------------------------------------=-----------------
# Entry Function
#------------------------------------------------------------------------

def main():
	
	argvCount = len(sys.argv) #The number of parameters for judging is to generate custom reports or periodic reports	
	dateFormat = "%Y-%m-%d %H:%M:%S"
	today = datetime.date.today()
	if(argvCount == 2):
		
# Only pass a parameter to generate custom reports : 00 reports point to the current day time
# Time format for incoming tuple format
		startTime = today.timetuple()
		dateFormat = "%Y-%m-%d %H:%M:%S"
		endTime = time.strptime(sys.argv[1],dateFormat) 
		custom_report(startTime,endTime)
	
	elif(argvCount == 3):
		#Two arguments , generate custom reports for : the first parameter is the start time, the second parameter is the difference between the end time of the report
		startTime =  time.strptime(sys.argv[1],dateFormat)
		endTime =  time.strptime(sys.argv[2],dateFormat)
		custom_report(startTime,endTime)		
	elif(argvCount ==1):
		#No parameters are passed to generate recurring reports
		today = datetime.date.today()
		dayOfMonth = today.day #Made the same day for the first few days of the month
		
		year = int(time.strftime('%Y',time.localtime(time.time())))
		#Get the number of months
		month = int(time.strftime('%m',time.localtime(time.time())))
		#How many days of the month of acquisition
		lastDayOfMonth = calendar.monthrange(year,month)[1]	
		#generating daily
		daily_report()
		#Current Sunday , generate weekly
		if(today.weekday()==6): 
			weekly_report()
		#The last day of the month , generating monthly
		if(dayOfMonth == lastDayOfMonth):
			monthly_repport()
	else:
		#2 is greater than the number of parameters to be illegal , the print exception information and exit Report Builder
		usage()
def usage():
	print """
Script did not pass parameters, performs periodic reports ; parameter may be 1, 2 , pay attention to the time format mandatory： zabbix-report.py ['2012-09-01 01:12:00'] ['2012-09-01 01:12:00']"""

#Run the program
main()
