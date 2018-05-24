#!/usr/bin/python3

from __future__ import print_function
from os.path import join, dirname, abspath
import xlrd
from xlrd.sheet import ctype_text  
import sys

def get_first_sheet(filename):
    
    # Open the workbook
    xl_workbook = xlrd.open_workbook(filename)
    
    # Return first sheet
    return xl_workbook.sheet_by_index(0)

HELPTEXT = """
Prints a list of users and how many hours they have logged.
This is sorted by the number of hours logged, from least to most.
By default, 1 decimal place is used (e.g. 12.5).

Possible args:

a: use alphabetical order
d[int]: use [int] decimal places
f: use first names [Note that this will use the first name as a key, so multiple users with the same first name will be listed as one entry]
h: print this help text
n: include each user's ranking (a number; 1 -> n)
r: print in reverse sorted order
"""


"""
Prints the number of hours each user has done.
@param use_first_name if True, just displays first name, otherwise displays full name
@param sorting sets sorting method:
        "alpha": alphabetical by first name
        "hours": by number of hours (least to most) (default)
@param number_of_decimal_places how specific the hour count for each person should be
@param reverse_order if True, reverses order of list
@param rank if True, adds a ranking to each item in the list (1 -> n)
"""
def print_hours(xl_sheet, use_first_name = False, sorting = "hours", number_of_decimal_places = 1, reverse_order = False, rank = True):
    num_cols = xl_sheet.ncols   # Number of columns
    users = dict()
    
    # Generate hours-worked data
    
    # For each row (except header row):
    for row_i in range(1, xl_sheet.nrows):
        # For each column:
        for col_i in range(0, num_cols):  # Iterate through columns
            # Get cell value by row, col
            cell_value = xl_sheet.cell(row_i, col_i).value
            
            
            if col_i == 6: # User name
                if (use_first_name):
                    user = cell_value.split(" ")[0] #Note: this means people with the same first name are grouped together
                else:
                    user = cell_value
                    
                if user not in users:
                    ##print("Adding " + str(user))
                    users[user] = 0
                    
            elif col_i == 8: # Hours
                users[user] += float(cell_value)
                
    # Create a nice format
    
    max_length = max([len(user) for user in users])
    output_format = '{0:>%d}: {1:.%df}' % (max_length, number_of_decimal_places)
    rank_format = '{0:>%d}: ' % (len(str(len(users))))
    
    # Print it
    
    if (sorting == "alpha"):
        user_list = sorted(users.keys())
        
    else: ## if (sorting == "hours"):
        user_list = sorted(users, key=users.get)
        
    if reverse_order:
        user_list.reverse()
        
    i = 0
    for user in user_list:
        i += 1
        if (rank):
            print(rank_format.format(i), end = '')
        print(output_format.format(user, users[user])) 
    

def main(args):
    filename = join(dirname(abspath(__file__)), "agilefantTimesheet.xls")
    xl_sheet = get_first_sheet(filename)

    sorting = "hours"    
    number_of_decimal_places = 1
    use_first_name = False
    reverse_order = False
    rank = False
    
    if 'h' in args:
        print(HELPTEXT)
        
    else:
        if len(args) > 1:
            for arg in args:
                if arg == 'a':
                    sorting = "alpha"
                    
                if arg[0] == 'd':
                    decimal_places_str = arg[1:]
                    if (decimal_places_str.isdigit()):
                        number_of_decimal_places = int(decimal_places_str)
                        
                if arg == 'f':
                    use_first_name = True
                    
                if arg == 'n':
                    rank = True
                    
                if arg == 'r':
                    reverse_order = True
                
        print_hours(xl_sheet, use_first_name, sorting, number_of_decimal_places, reverse_order, rank)        
            
            
if __name__ == "__main__":
    main(sys.argv)
