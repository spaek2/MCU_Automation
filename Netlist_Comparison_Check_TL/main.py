#!/usr/bin/env python3

import csv
import xlsxwriter

def intersection(lst1, lst2): # Intersection of two lists
    lst3 = [value for value in lst1 if value in lst2]
    return lst3


if __name__ == '__main__':

    MCU_str = 'U101_UP' # String identifier for nets connected to MCU only

    # Create various dictionaries and lists to keep track of Signals and which Ball they're connected to

    IO_dict = {}
    Netlist_dict = {}

    IO_BGA_List = []
    Netlist_BGA_List = []

    NC_dict = {}
    C_dict = {}
    NP_IO_dict = {}
    NP_NL_dict = {}

    file = open('IO Signal Table-Table 1.csv') # Open IO Spreadsheet

    IOSheet = csv.reader(file)

    for i in range(2):
        next(IOSheet) # skip first two rows

    for row in IOSheet:
        if row[21]: # If the application field is populated
            IO_dict[row[17]] = row[21] # add it to the IO dictionary
            IO_BGA_List.append(row[17]) # keep track of balls used on the IO spreadsheet

    file2 = open('dialcnet.dat') # open Netlist

    NetlistSheet = csv.reader(file2)

    for row in NetlistSheet:
        row_split = row[0].split() # Convert .dat format to usable list
        if MCU_str in row_split[1]: # If a signal is attached to MCU
            Netlist_dict[row_split[2]] = row_split[0] # Add it to netlist dictionary
            Netlist_BGA_List.append(row_split[2]) # Keep track of balls used on Netlist

    Intersection_BGA_List = intersection(IO_BGA_List, Netlist_BGA_List) # Find common balls used on IO List and Netlist
    NP_IO_BGA_list = list(set(Netlist_BGA_List) - set(Intersection_BGA_List)) # Find balls used only in Netlist
    NP_NL_BGA_list = list(set(IO_BGA_List) - set(Intersection_BGA_List)) # Find balls used only in IO List

    for BGA in Intersection_BGA_List:
        if IO_dict[BGA] == Netlist_dict[BGA]: # If signal names match between IO and Netlist
            C_dict[BGA] = IO_dict[BGA] # Add it to the checked signal list
        else:
            NC_dict[BGA] = [IO_dict[BGA], Netlist_dict[BGA]] # Otherwise add both signals names to the unchecked list

    for BGA in NP_IO_BGA_list: # For signals only in Netlist
        NP_IO_dict[BGA] = Netlist_dict[BGA] # Add it to netlist dictionary

    for BGA in NP_NL_BGA_list: # For signals only in IO List
        NP_NL_dict[BGA] = IO_dict[BGA] # Add it to the IO dictionary

    # Create Workbook/Worksheet of different categories of signals

    workbook = xlsxwriter.Workbook('MCU_Schematic_Check.xlsx')
    worksheet_nc = workbook.add_worksheet('Unchecked Signals')
    worksheet_io = workbook.add_worksheet('IO Sheet Only Signals')
    worksheet_nl = workbook.add_worksheet('Netlist Only Signals')
    worksheet_c = workbook.add_worksheet('Checked Signals')

    worksheet_nc.merge_range('A1:D1', 'Unchecked Signals [Pending Review]')
    worksheet_nc.write(1, 0, 'BGA')
    worksheet_nc.write(1, 1, 'Signal in IO Sheet')
    worksheet_nc.write(1, 2, 'Signal in Schematic Netlist')
    worksheet_nc.write(1, 3, 'Comment')

    row = 2
    for BGA in NC_dict:
        worksheet_nc.write(row, 0, BGA)
        worksheet_nc.write(row, 1, NC_dict[BGA][0])
        worksheet_nc.write(row, 2, NC_dict[BGA][1])
        row += 1

    worksheet_io.merge_range('A1:C1', 'Signals Missing from Netlist [Pending Review]')
    worksheet_io.write(1, 0, 'BGA')
    worksheet_io.write(1, 1, 'Signal in IO Sheet')
    worksheet_io.write(1, 2, 'Comment')

    row = 2
    for BGA in NP_NL_dict:
        worksheet_io.write(row, 0, BGA)
        worksheet_io.write(row, 1, NP_NL_dict[BGA])
        row += 1

    worksheet_nl.merge_range('A1:C1', 'Signals Missing from IO Sheet [Pending Review]')
    worksheet_nl.write(1, 0, 'BGA')
    worksheet_nl.write(1, 1, 'Signal in Netlist')
    worksheet_nl.write(1, 2, 'Comment')

    row = 2
    for BGA in NP_IO_dict:
        worksheet_nl.write(row, 0, BGA)
        worksheet_nl.write(row, 1, NP_IO_dict[BGA])
        row += 1

    worksheet_c.merge_range('A1:C1', 'Checked Signals')
    worksheet_c.write(1, 0, 'BGA')
    worksheet_c.write(1, 1, 'Signal in Netlist/IO Sheet')
    worksheet_c.write(1, 2, 'Comment')

    row = 2
    for BGA in C_dict:
        worksheet_c.write(row, 0, BGA)
        worksheet_c.write(row, 1, C_dict[BGA])
        row += 1

    workbook.close()

    print('Spreadsheet generated')