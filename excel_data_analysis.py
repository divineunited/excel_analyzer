# This OOP application takes an employees shift data from excel cells and analyzes their performance (can compare different shifts/excels and export to a final excel if needed)

import os
import pandas as pd
import numpy as np
import datetime as dt
import tools
import time
# import openpyxl


class Employee(object):
    '''takes an employee name and formatted excel file representing that employee's shift with columns as follows: 1:shift_id 2:date_worked 3:clock_in time (24hour) 4:clock_out time (24 hour) 5:hours digitized 6:tape format 7: employeeID'''

    employee_list = [] #class attribute that is a list of the entire instance of the employees created
    employee_DFS = [] #class attribute that is a list of all the DFs of employees created
    employee_efficiency = {} #class attribute dictionary: {employeeID : %time_digitized}


    @staticmethod
    def print_all():
        '''this static method prints all employees being compared'''
        if Employee.employee_list:
            print('\nThere are ' + str(len(Employee.employee_list)) + ' employees entered.\n')
            for employee, employeeDF in zip(Employee.employee_list, Employee.employee_DFS):
                print(employee.name, 'Data Preview:')
                print(employeeDF.head(n=7).to_string(index=False))
        else:
            print('There are 0 employees entered.\n')


    @staticmethod
    def employee_rank():
        '''this static method returns the efficiency ranking of all employees'''
        if Employee.employee_efficiency:
            sortedlist = [(employee, Employee.employee_efficiency[employee]) for employee in sorted(Employee.employee_efficiency, key=Employee.employee_efficiency.get, reverse=True)]
            print('\nEmployee & Efficiency - ranked from most efficient to least efficient:')
            print('(Efficiency is measured as hours digitized / hours worked)')
            for employee, efficiency in sortedlist:
                print(employee + ': ' + str(efficiency))
        else:
            print('There are 0 employees entered.\n')


    @staticmethod
    def analyze_all_date():
        '''this static method does a DATE analysis on all employees and sends to excel with each person being a sheet'''

        if Employee.employee_list:
            sheet_list = []
            dfs = []
            for employee in Employee.employee_list:
                #appending the name of the employee to sheet_list of titles for xcel sheets
                sheet_list.append(employee.name)
                #appending the DF of the DATE analysis
                dfs.append(Employee.date_analysis(employee))
            #sending these employee's analysis to 1 excel over seperate tabs/sheets
            tools.dfs_tabs_date(dfs, sheet_list, str(dt.date.today())+'_date_analysis.xlsx')
            print("\nOverall Date Analysis Excel Created!")

        else:
            print('There are 0 employees entered.\n')
    
    @staticmethod
    def analyze_all_format():
        '''this static method does a FORMAT analysis on all employees and sends to excel with each person being a sheet'''

        if Employee.employee_list:
            sheet_list = []
            dfs = []
            for employee in Employee.employee_list:
                #appending the name of the employee to sheet_list of titles for xcel sheets
                sheet_list.append(employee.name)
                #appending the DF of the FORMAT analysis
                dfs.append(Employee.format_analysis(employee))
            #sending these employee's analysis to 1 excel over seperate tabs/sheets
            tools.dfs_tabs_format(dfs, sheet_list, str(dt.date.today())+'_format_analysis.xlsx')
            print("\nOverall Format Analysis Excel Created!")

        else:
            print('There are 0 employees entered.\n')   



    def __init__(self, name, fullpath):
        self.name = name.replace(' ', '_') # Excel doesn't like spaces in sheet names:
        
        # extracting the directory and asking python to look there for the file
        path = os.path.dirname(fullpath)
        os.chdir(path)

        # getting the filename
        base = os.path.basename(fullpath)

        # setting the parser to know what the incoming date format looks like to correct it
        parser = lambda date: pd.datetime.strptime(date, '%Y/%m/%d')

        # this creates the dataframe from the excel and parses the date into a datetime object
        self.dataframe = pd.read_excel(base, parse_dates=['date_worked'], date_parser=parser)
        #removing the time section of the datetime
        self.dataframe['date_worked'] = pd.to_datetime(self.dataframe['date_worked']).dt.date

        # appending another column called hours_worked after the clock_in and clock_out column
        self.dataframe.insert(loc=4, column='hours_worked', value = np.nan)

        # inserting in the time differences into that column using a tools.py project that needs to be in same directory
        for i in range(len(self.dataframe)):
            self.dataframe.loc[i, 'hours_worked'] = tools.time_dif(self.dataframe.loc[i, 'clock_in'], self.dataframe.loc[i, 'clock_out'])

        # appending another column called efficiency after the hours_digitized column
        self.dataframe.insert(loc=6, column='efficiency', value = np.nan)

        #inserting in the efficiency of that day (defined as hours_digitized / hours_worked)
        for i in range(len(self.dataframe)):
            self.dataframe.loc[i, 'efficiency'] = round(self.dataframe.loc[i, 'hours_digitized'] / self.dataframe.loc[i, 'hours_worked'], 3)

        #adding this employee's instance & final dataframe to the class attributes
        Employee.employee_list.append(self)
        Employee.employee_DFS.append(self.dataframe)

        #adding overall efficiency to the class dictionary
        self.overall_efficiency = round(sum(self.dataframe.loc[:, 'hours_digitized']) / sum(self.dataframe.loc[:, 'hours_worked']), 3)
        Employee.employee_efficiency[self.name] = self.overall_efficiency

        # initialization done!
        print(self.name + "'s data successfully inputted and ready for analysis!")



    def __str__(self):
        '''this will return all the data of the employee and allows it to accept the print function'''
        reply = 'Employee: ' + self.name + ' | Employee ID: ' + str(self.dataframe.loc[0:1, 'employee_id']) + '\n'
        reply += '--------------------------------------\n'
        reply += str(self.dataframe)
        return reply


    def date_analysis(self):
        '''this function sorts by date and returns a dataframe that has dateworked, and efficiency of only the LAST 30 DAYS WORKED - can change this'''

        #make a copy of the dataframe sorted by date:
        datesorted_df = self.dataframe.sort_values(by='date_worked')

        #returning a new df with dates and efficiency for last 30 days worked:
        return datesorted_df.loc[:, ('date_worked', 'hours_digitized', 'hours_worked', 'efficiency')].tail(n=30)

        #this is just in case you wanted to get the middle 25-50th column for instance
        #return datesorted_df.loc[:, ('date_worked', 'efficiency')].iloc[25:50,:]


    def format_analysis(self):
        '''this function groups by format and then calculates average efficiency of each format and returns it as a dataframe'''
        ### the format efficiency is calculated by dividing the total number of hours digitized / hours worked for a specific format

        #getting a list of unique formats
        formats = set(self.dataframe.loc[:, 'format'])

        #getting a dataframe with sum of hours worked and hours digitized grouped by format
        format_df = self.dataframe.groupby(['format']).sum()

        #creating an empty dataframe to return result
        columns = ['Format', 'Format Efficiency', 'Sample Size']
        data = np.nan
        result_df = pd.DataFrame(data, index=formats, columns=columns)

        #filling the result_df with data:
        #this fills sample size column with the counts of all the different formats in our data
        result_df.loc[:, 'Sample Size'] = self.dataframe['format'].value_counts()

        #this fills average efficiency column by using the groupedby df and dividing hours worked by digitized of the specific format
        for format in formats:
            result_df.loc[format, 'Format'] = format
            result_df.loc[format, 'Format Efficiency'] = round(format_df.loc[format, 'hours_digitized'] / format_df.loc[format, 'hours_worked'], 3)

        #return in descending order from most efficient format to least efficient format
        return result_df.sort_values(by='Format Efficiency', ascending=[False])




#main

# creating dictionary which will hold names and fullpaths of xcel files from user referenced by a number index
namepaths = {}
counter = 1

while True:
    flag = 0 
    print('Welcome to the Employee Analysis Gateway.')
    name = input('Please enter name of employee. (or type STOP): ')
    # if user wants to break out
    if name.upper() == 'STOP':
        break
    # asking for directory path and filename
    else:
        while True:
            path = str(input('Please enter full path of employee excel data (directory path + filename) or STOP: '))
            if path.upper() == 'STOP':
                flag = 1
                break
            # making sure it was a valid input:
            try:
                directory = os.path.dirname(path)
                base = os.path.basename(path)
                # making sure directory is directory / file is a file / and it's an excel file. if not, force an error
                if not ((os.path.isdir(directory)) and (os.path.isfile(path)) and (base[-4:] == 'xlsx')):
                    base = 1/0
                else:
                    break
            except:
                print('That was not a valid input.')
    # appending to the dictionary, then increasing the counter for the next loop
    if flag == 0:
        namepaths[counter] = (name, path)
        print(counter, 'Initialized\n')
        counter += 1
    else:
        break


if namepaths:
    print('\nAnalyzing in...')
    print('3')
    time.sleep(1)
    print('2')
    time.sleep(1)
    print('1\n')
    time.sleep(1)
    for namepath in namepaths:
        # extracting the name and path from each dictionary key
        name, path = namepaths[namepath]
        # creating an instance of that employee to be analyzed and sending it the name and path
        namepath = Employee(name, path)
    # Running the application - printing sample data, printing the efficiency ranks of the employees entered, then creating the excel analyses
    Employee.print_all()
    Employee.employee_rank()
    Employee.analyze_all_date()
    Employee.analyze_all_format()
else:
    print('Goodbye.')

