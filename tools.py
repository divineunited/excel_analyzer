#--------------------------------------------------------------
#               Tools Module
#
#--------------------------------------------------------------

#---------------------------------------------------------------------
#To use this file, place it in the same directory as your python code,
#then add the line:
#               from tools import *
#to the top of your python code. You can then use these functions.
#
#As you write new functions, add them to this file! If you keep this
#up-to-date and organized, it will help you tremendously!
#---------------------------------------------------------------------

#------------------------------
# DataFrame to Excel functions:
#------------------------------

import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import datetime

def dfs_tabs_format(df_list, sheet_list, file_name):
    '''accepts a list of dfs, list of sheet names, and a file name - Puts multiple dataframes across MULTIPLE tabs/sheets in 1 excel'''
    ### Formatted specifically for the FORMAT ANALYSIS

    #this only works if all the dfs being sent are the same length/formatting
    number_rows = len(df_list[0].index)

    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, index=False, sheet_name=sheet, startrow=2 , startcol=0)

        # Get access to the workbook and sheet
        workbook = writer.book 
        worksheet = writer.sheets[sheet]

        # Creating some Formats in our workbook to assign later:
        right_fmt = workbook.add_format({'align': 'right'})
        title_fmt = workbook.add_format({'align': 'left', 'font_size':21,
                                        'bold': True})
        title_fmt.set_align('left_across')
        percent_fmt = workbook.add_format({'num_format': '0.0%', 'bold': True})

        # Setting the Column Width and Formatting
        worksheet.set_column('A:D', 20, right_fmt)
        worksheet.set_column('B:B', 20, percent_fmt)
        

        # Adding a Label at the top:
        worksheet.write_string(0, 0, sheet + " Analysis of Digitizing Format Efficiency", title_fmt)

        # Setting the Default Zoom
        worksheet.set_zoom(120)

        # Define our range for the color formatting
        color_range = "B4:B{}".format(number_rows+3)
       
        # 3 color scale from green = most efficient to red = least efficient
        worksheet.conditional_format(color_range, {'type': '3_color_scale'})

        # Create the COLUMN CHART:
        # --------------------------------------       
        column_chart = workbook.add_chart({'type': 'column'})

        # Configure the series of the chart from the dataframe data.
        column_chart.add_series({
            'categories': '='+sheet+'!$A$4:$A$15',
            'values': '='+sheet+'!$B$4:$B$15'
            })

        # Create a new line chart. This will be the secondary chart.
        line_chart = workbook.add_chart({'type': 'line'})

        # Add a series, on the secondary axis.
        line_chart.add_series({
            'categories': '='+sheet+'!$A$4:$A$15',
            'values': '='+sheet+'!$C$4:$C$15',
            'marker': {'type': 'automatic'},
            'y2_axis': True,
            'name': '# Digitized'
            })

        # Combine the charts.
        column_chart.combine(line_chart)

        # Add a chart title.
        column_chart.set_title({'name': sheet + " Format Efficiency"})

        # Turn off the chart legend.
        column_chart.set_legend({'position': 'none'})

        # Set the title of the Y axes
        column_chart.set_y_axis({'name': 'Efficiency %'})
            
        # Set the title of the Y2 axes - doesn't seem to work
        column_chart.set_y2_axis({'name': '# Digitized'})
        
        # Insert the chart into the worksheet.
        worksheet.insert_chart('E3', column_chart)


        #### Create the PIE Chart
        # --------------------------------------  
        chart_pie = workbook.add_chart({'type': 'pie'})

        # Configure the series. Note the use of the list syntax to define ranges:
        chart_pie.add_series({
            'name':       'Formats Digitized',
            'categories': '='+sheet+'!$A$4:$A$15',
            'values': '='+sheet+'!$C$4:$C$15',
        })

        # Add a title.
        chart_pie.set_title({'name': 'Formats Digitized'})

        # Set an Excel chart style. Colors with white outline and shadow.
        chart_pie.set_style(10)

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('B20', chart_pie)

    writer.save()


def dfs_tabs_date(df_list, sheet_list, file_name):
    '''accepts a list of dfs, list of sheet names, and a file name - Puts multiple dataframes across MULTIPLE tabs/sheets in 1 excel'''
    ### Formatted specifically for the DATE ANALYSIS

    #this only works if all the dfs being sent are the same length/formatting
    number_rows = len(df_list[0].index)

    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, index=False, sheet_name=sheet, startrow=2 , startcol=0)

        # Get access to the workbook and sheet
        workbook = writer.book 
        worksheet = writer.sheets[sheet]

        # Creating some Formats in our workbook to assign later:
        right_fmt = workbook.add_format({'align': 'right'})
        total_fmt = workbook.add_format({'align': 'right',
                                        'bold': True, 'bottom':6,
                                        'bg_color': '#85144B',
                                         'font_color': '#FFDC00'})
        percent_fmt = workbook.add_format({'num_format': '0.0%', 'bold': True})
        total_percent_fmt = workbook.add_format({'align': 'right', 'num_format': '0.0%',
                                         'bold': True, 'bottom':6,
                                         'bg_color': '#85144B',
                                         'font_color': '#FFDC00'})
        title_fmt = workbook.add_format({'align': 'left', 'font_size':21,
                                        'bold': True})
        title_fmt.set_align('left_across')

        # Add a format. Light red fill with dark red text.
        format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                    'font_color': '#9C0006'})

        # Add a format. Green fill with dark green text.
        format2 = workbook.add_format({'bg_color': '#C6EFCE',
                                    'font_color': '#006100'})

        # Setting the Column Width and Formatting
        worksheet.set_column('A:D', 20, right_fmt)
        worksheet.set_column('D:D', 20, percent_fmt)

        # Adding a Label at the top:
        worksheet.write_string(0, 0, sheet + " Analysis of Last 30 Days Worked", title_fmt)

        # Setting the Default Zoom
        worksheet.set_zoom(115)

        # Add total's formula at the end similar to VBA (visual basic application) - this is doing a for loop and creating a SUM for columns 1 and 2
        for column in range(1, 3): 
            # Determine which cell where we will place the 'excel formula' for each column
            cell_location = xl_rowcol_to_cell(number_rows+3, column)
            # Get the range to use for the sum formula
            start_range = xl_rowcol_to_cell(3, column) #start at 3rd row for that column (row 2 is title)
            end_range = xl_rowcol_to_cell(number_rows+2, column) #end at last row for that column
            # Construct and write the formula for each column
            formula = "=SUM({:s}:{:s})".format(start_range, end_range)
            worksheet.write_formula(cell_location, formula, total_fmt)
        
        # Add a total label
        worksheet.write_string(number_rows+3, 0, "Total:",total_fmt)

        # Add an average efficiency
        mean_formula = "=B{0}/C{0}".format(number_rows+4)
        worksheet.write_formula(number_rows+3, 3, mean_formula, total_percent_fmt)

        # Define our range for the color formatting
        color_range = "D4:D{}".format(number_rows+3)
        
        # Highlight the top values in Green
        worksheet.conditional_format(color_range, {'type': 'top',
                                                'value': '5',
                                                'format': format2})

        # Highlight the bottom values in Red
        worksheet.conditional_format(color_range, {'type': 'bottom',
                                                'value': '5',
                                                'format': format1})

        # Create the LINE CHART:
        # --------------------------------------
        # Create a new chart object.
        chart = workbook.add_chart({'type': 'line'})

        # Add a series to the chart along with a trendline.
        chart.add_series({
            'name':       'Efficiency Trends through Time',
            'categories': '='+sheet+'!$A$4:$A$33',
            'values': '='+sheet+'!$D$4:$D$33',
            'marker': {'type': 'diamond'},
            'trendline': {
                        'type': 'polynomial',
                        'name': 'Trend Line',
                        'order': 2,
                        'forward': 0.5,
                        'backward': 0.5,
                        'display_equation': False,
                        'line': {
                            'color': 'red',
                            'width': 1,
                            'dash_type': 'long_dash',
                        },
                    },
        })

        # Configure the X axis as a Date axis
        chart.set_x_axis({
            'date_axis': True,
            # 'min': datetime.date(2000, 1, 1), # this doesn't work
            # 'max': datetime.date.today(),
        })

        # Setting title and scale of Y axis
        chart.set_y_axis({
            'name': 'Efficiency',
            'min': 0,
            'max': 2.2
        })

        # Turn off the legend.
        chart.set_legend({'none': True})

        # Insert the chart into the worksheet.
        worksheet.insert_chart('F7', chart)

        
    writer.save()




def dfs_tabs(df_list, sheet_list, file_name):
    '''accepts a list of dfs, list of sheet names, and a file name - Puts multiple dataframes across MULTIPLE tabs/sheets in 1 excel'''
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, index=False, sheet_name=sheet, startrow=0 , startcol=0)
    writer.save()


def multiple_dfs(df_list, sheet, file_name, spaces):
    '''accepts a list of dfs, list of sheet names, and a file name - Puts multiple dataframes into ONE SINGLE sheet in 1 excel'''
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    row = 0
    for dataframe in df_list:
        dataframe.to_excel(writer, index=False, sheet_name=sheet, startrow=row, startcol=0)
        row = row + len(dataframe.index) + spaces + 1
    writer.save()




#------------------------------
# Date time functions:
#------------------------------

from datetime import datetime, timedelta

def time_dif(time1, time2):
    '''accepts time1 and time2 as strings HH:MM:SS with time2 being after time1 and returns difference'''
    format = '%H:%M:%S'
    tdelta = datetime.strptime(time2, format) - datetime.strptime(time1, format)
    # this if statement assumes time2 is always after time1 and just crosses midnight
    if tdelta.days < 0:
        tdelta = timedelta(days=0, seconds=tdelta.seconds, microseconds=tdelta.microseconds)
    # converting the tdelta into hours float with 1 decimal place
    h, m, s = str(tdelta).split(':')
    # seconds = int(h) * 3600 + int(m) * 60 + int(s)
    # minutes = int(h) * 60 + int(m) + (int(s) / 60)
    hours = int(h) + (int(m) / 60) + (int(s) / 3600)

    return round(hours, 1)




#------------------------------
# string manipulation functions
#------------------------------

#prints out a birthday greeting, demonstrates default values
def birthday(name="Joe", age=21):
    print("Happy Birthday", name, "! You're", age, "!")

#Places dashes around every character of the string and returns it
def dasher(string):
    return "-" + "-".join(list(string)) + "-"

#pads a string out to 20 characters with dashes
def dasher2(string):
    if len(string) > 20:
        string = "Error"
    dashes =(20 - len(string))

    half = int(dashes/2)

    formatted = "-" * half + string + "-" * half

    if (dashes % 2 == 1):
        formatted += "-"
    return formatted

#works like dasher2, but allows you to specify a length
#20 is the default
def dasher3(string, length = 20):
    if len(string) > length:
        string = "Error"
    dashes =(length - len(string))

    half = int(dashes/2)

    formatted = "-" * half + string + "-" * half

    if (dashes % 2 == 1):
        formatted += "-"
    return formatted



#------------------------------
# number manipulation functions
#------------------------------

#returns the sum and product of two numbers
def addmult(num1, num2):
    add = num1 + num2
    mult = num1 * num2
    return add, mult

#returns True if a number is odd, False otherwise
def odd(number):
    return number % 2 == 1

#returns True if the difference between two numbers is odd, False otherwise
def odd_diff(num1, num2):
    difference = num1 - num2
    return odd(difference)

#returns the sum of a list of numbers
def summation(numbers):
    total = 0
    for num in numbers:
	    total += num
    return total

#returns the average of a list of numbers
def mean(numbers):
    total = 0.0
    for num in numbers:
            total += num
    return total/len(numbers)



#------------------------------
# list manipulation functions
#------------------------------

def swap(a_list, x, y):
    """x and y should be ints that are valid index positions"""
    temp = a_list[y]
    a_list[y] = a_list[x]
    a_list[x] = temp
    #NO NEED TO RETURN ANYTHING. Lists pass by reference! It changes orig list!




#------------------------------
# printing functions (no return)
#------------------------------

#this function prints all numbers between low and high that are
#divisible by the factor given.
def print_range(low, high, factor):
    for num in range(low, high):
        if num % factor == 0:
            print(num, "is divisible by", factor)


# this function prints out data with headers in a nicely padded format
# headers should be a list or tuple
# data should be a list of tuples or lists
# padding is the maximum column width for all the columns.

def table_print(headers, data, padding):
    # We build up the output formatting string
    # It has this general look, but for any number of columns
    # output = "{0:>" + str(padding) + "} {1:>" + str(padding) + "}"
    output = []
    for i in range(len(headers)):
        output.append("{" + str(i) + ":>" + str(padding) + "}")
    output = " ".join(output)

    # Print the headers
    print(output.format(*headers))

    # Print as many dashes as there are columns
    # Times the padding value (plus 1 for each space)
    print(("-" * (padding) * len(headers)) + ("-" * (len(headers) - 1)))

    # Print out the data values
    for item in data:
        print(output.format(*item))
    print()

#------------------------------
# validation functions
#------------------------------



def valid_int(description):       #we are going to do this enough to warrant a function
    while True:
        try:
            valid = int(input("Please enter " + description + " (int): "))
        except:
            print("That's not an integer.")
        else:
            return valid


#-----------------------------
# test code goes here
#-----------------------------

if __name__ == "__main__":
    print("Testing dasher:", dasher("this is a test") == "-t-h-i-s- -i-s- -a- -t-e-s-t-")
    print("Testing odd_diff:", odd_diff(18,9))

    nums = [1,2,3,4]

    print("Testing summation:", summation(nums) == 10)
    print("Testing mean:", mean(nums) == 2.5)

    time1 = '9:48:28'
    time2 = '19:40:17'
    print('Testing time_dif:', str(time_dif(time1, time2)) == '9:51:49')
