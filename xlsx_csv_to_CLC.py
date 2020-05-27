#!/usr/bin/env python
# coding: utf-8

# Converts a selection of .csv, .xls, or .xlsx files containing historial data 
# to a .clc file for import into AspenTech DMCPlus/DMC3 Platforms.
# 
# Returns .clc data file(s) with names corresponding to the csv files.
# If there are errors, returns a text error file listing lines in the .clc
# file with tagnames and descriptions too long for importing. Files will be 
# located in the same folder as the .csv files selected.
# 
# CSV file format must generally follow sample .csv files provided.
# 
# Removes special characters from tagnames and double spaces from descriptions.

# In[110]:


import csv
import datetime
import re
import tkinter as tk
import xlrd

from dateutil.parser import parse
from tkinter import filedialog
from tkinter import messagebox


# In[111]:


def read_csv_as_nested_list(filename, separator, quote):
    """
    Inputs:
      filename  - Name of CSV file
      keyfield  - Field to use as key for rows
      separator - Character that separates fields
      quote     - Character used to optionally quote fields

    Output:
      Returns a list of lists where the outer list corresponds to a each 
      row in the CSV file. The inner lists contain each comma separated 
      value from the CSV file.
    """
    
    list1 = []
    with open(filename, newline='') as csvfile:
        csvreader = csv.reader(
            csvfile, delimiter=separator, quotechar=quote)
        for row in csvreader:
            list1.append(row)
            
    return list1


# In[112]:


def read_xlsx_as_nested_list(xlsx_file):
    """
    Inputs:
      filename  - Name of XLSX file
    Output:
      Returns a list of lists where the outer list corresponds to a each 
      row in the CSV file. The inner lists contain each comma separated 
      value from the CSV file.
    """
    workbook = xlrd.open_workbook(xlsx_file)
    sheet = workbook.sheet_by_index(0)

    list1 = []
    for rowx in range(sheet.nrows):
        values = sheet.row_values(rowx)
        list1.append(values)
    
    return list1


# In[113]:


def clc_tags_descriptions(csv_table):
    """
    Input: 
        Data list containing three header rows for tagname, description and
        units followed by timestamped data
    Actions:
        Removes any special characters from the tagname.
    Returns: 
        Tuple with list of strings with CLC-formatted tagnames, 
        descriptions and units, and a second list of length errors if there
        are any.
    """
    
    list1 = []
    list2 = []
    for idx in range(1, len(csv_table[0])):
        tag = re.sub('[^A-Za-z0-9]+','',csv_table[0][idx])
        
        if len(tag)>12:
            list2.append(
                "line{}: {} variable length too long".format(
                    idx + 8,tag))
        desc = re.sub(' +', ' ',csv_table[1][idx])
        
        if len(desc)>40:
            list2.append(
                "line{}: {} description length too long".format(
                    idx + 8,tag))
        
        units = re.sub(' +', ' ',csv_table[2][idx])
        clc_string = "{}~~~{}~~~{}~~~{}".format(tag, tag, desc, units)
        list1.append(clc_string)
    
    return (list1, list2)


# In[114]:


def determine_period_t0(csv_table, extension):
    """
    Input: 
        Data list containing three header rows for tagname, description and
        units followed by timestamped data
    Action: 
        Returns a two-element list containing datetime elements of the 
        period and first timestamp 
    """
    if extension == "csv":
        time_0 = parse(csv_table[3][0])
        time_1 = parse(csv_table[4][0])
        period = (time_1 - time_0)
    elif extension == "xlsx":
        time_0 = xlrd.xldate_as_tuple(csv_table[3][0], 0)
        time_1 = xlrd.xldate_as_tuple(csv_table[4][0], 0)
        time_0 = datetime.datetime(*time_0)
        time_1 = datetime.datetime(*time_1)
        period = (time_1 - time_0)
    
    return [period, time_0]


# In[115]:


def convert_to_date(csv_table, period_tup):
    """
    Input: 
        csv_table:
            Data list containing three header rows for tagname, description 
            and units followed by timestamped data
        period_tup: 
            2-element list containing datetime elements of the period and 
            first timestamp

    Returns: List of CLC-formatted timestamp strings  
    """
    datetime_list = []
    time_str = []
    
    period = period_tup[0]
    time_0 = period_tup[1]
    
    for ts in range(0,len(csv_table)-3):
        datetime_list.append(time_0 + (ts*period))

    for ts in datetime_list:
        ts_clc = datetime.datetime.strftime(ts, "%m-%d-%Y %H:%M:%S")
        time_str.append(ts_clc)       
        
    return time_str
    


# In[116]:


def get_timestamps(csv_table, extension):
    """
    Input: 
        CSV list containing three header rows for tagname, description and 
        units
    Returns: List of CLC-formatted timestamp strings  
    """
    
    period_t0 = determine_period_t0(csv_table, extension)
    timestamps = convert_to_date(csv_table, period_t0)
    
    return timestamps


# In[117]:


def format_data(csv_table):
    """
    Input: 
        CSV list containing three header rows for tagname, description and 
        units
    Returns: 
        List of CLC-formatted data with status strings. Data entries that 
        could not be converted to floating number will be flagged as bad 
        with "B" and replaced with integer value -9999 
    """
    data_row = []
    for row_idx in range(3, len(csv_table)):
        data_elements = []
        for data_idx in range(1, len(csv_table[row_idx])):
            try:
                data = (float(csv_table[row_idx][data_idx]))
                status = "G"
            except ValueError:
                data = -9999
                status = "B"
            data_elements.append(data)
            data_elements.append(status)
        data_row.append(data_elements)
    return data_row


# In[118]:


def data_section(timestamps, data):
    """
    Inputs:
        timestamps - Nested list CLC formatted timestamps
        data - Nested list of CLC formatted data
    Outputs:
        List of lists where each row in list corresponds to data rows
        in the CLC file
    """
    data_section = []
    for time_idx in range(len(timestamps)):
        data_row = []
        data_row.append(timestamps[time_idx])
        for idx in range(len(data[time_idx])):
            data_row.append(data[time_idx][idx])
        data_section.append(data_row)
    
    return data_section


# In[119]:


def header_section(csv_table, tag_list, timestamps, extension):
    """
    Inputs:
    Ouputs: 
        List of header entries per AspenTech documentation excerpt below:
    
        The first section of the .clc file contains header information in 
        the following order:
        Line
            1:filename of the Input file
            2:description from Line 1 of the Input file
            3:number of tags extracted
            4:number of tags per section
            5:beginning time of extraction (MM-DD-YYYY(blank)hh:mm:ss)
            6:sample period of extraction (in seconds)
            7:number of samples extracted
    """
    header = []
    
    [period, time_0] = determine_period_t0(csv_table, extension)
    period = int(period.total_seconds())
    
    header.append("CSV to CLC File Conversion")
    header.append("Developed by D.P. (AMT)")
    header.append(len(tag_list))
    header.append(len(tag_list))
    header.append(timestamps[0])
    header.append(period)
    header.append(len(timestamps))
    
    return header


# In[120]:


def create_clc_table(file_name, extension):
    """
    Input: 
        CSV list containing three header rows for tagname, description and 
        units
    Returns: Nested list of clc formatted rows and a list of errors
    """

    if extension == "csv":
        csv_table = read_csv_as_nested_list(file_name, ",", "'")
    elif extension == "xlsx":
        csv_table = read_xlsx_as_nested_list(file_name)
    
    tags = clc_tags_descriptions(csv_table)
    tag_list = tags[0]
    errors = tags [1]
    timestamps = get_timestamps(csv_table, extension)  
    data = format_data(csv_table)
    
    data_s = data_section(timestamps, data)
    header = header_section(csv_table, tag_list, timestamps, extension)
    
    write_list = []
    
    for h in header:
        write_list.append([h])
        
    write_list.append(["="*50])
    
    for tags in tag_list:
        write_list.append([tags])
    
    write_list.append(["="*50])
    
    for items in data_s:
        write_list.append(items)
        
    write_list.append(["="*50])
          
    return [write_list, errors]


# In[121]:


def write_csv_file(file_name):
    """
    Input: File path
    Action: Writes a CLC file using data from the input file name 
        and an error file if there were errors detected.
    """
    
    clc_str = file_name.split(".")
    clc_file_name = "{}.clc".format(clc_str[0])
    clc_error_file = "{}_errors.txt".format(clc_str[0])
    extension = clc_str[1]
       
    writes = create_clc_table(file_name, extension)
    table = writes[0]
    errors = writes[1]
    
    with open(clc_file_name, 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)

        for rows in table:
            writer.writerow(rows)

    if errors == []:
        pass
    else:
        with open(clc_error_file, 'w', newline='') as csv_file:
            writer = csv.writer(csv_file)

            for rows in errors:
                writer.writerow([rows])


# In[122]:


def multi_file_conversion():
    """
    Action: 
        Prompts user to select CSV files for conversion
        Converts selected CSV files to CLC files
    """
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilenames(title = "Select file", 
                                    filetypes = (("CSV Files","*.csv"),("Excel Files","*.xlsx")))

    messages = ""
    fail_msg = ""
    
    for files in file_path:
        try:
            write_csv_file(files)
            messages += "{} conversion complete\n".format(files)
        except:
            fail_msg += "{} conversion failed\n".format(files)
            continue
    
    messagebox.showinfo(title = "Conversion results", message = messages)
    
    if fail_msg != "":
        messagebox.showerror(title = "Conversion failed", message = fail_msg)


# In[136]:


multi_file_conversion()

