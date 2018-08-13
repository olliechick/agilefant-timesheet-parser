#!/usr/bin/python3

from __future__ import print_function
from os.path import join, dirname, abspath
import xlrd
from xlrd.sheet import ctype_text  
import sys
import re

valid_tags = {"implement", "document", "test", "testmanual", "fix", "chore", "refactor"}
commitless_tags = {'commits', 'chore', 'document'}

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
d[int]: use [int] decimal places (e.g. d3)
f: use first names [Note that this will use the first name as a key, so multiple users with the same first name will be listed as one entry]
h: print this help text
n: include each user's ranking (a number; 1 -> n)
r: print in reverse sorted order
t: print all invalid log comments
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
def print_hours(xl_sheet, use_first_name = False, sorting = "hours", number_of_decimal_places = 1, reverse_order = False, rank = True, display_tag_errors = False):
    num_cols = xl_sheet.ncols   # Number of columns
    users = dict()
    
    # Generate hours-worked data
    
    # For each row (except header row):
    for row_i in range(1, xl_sheet.nrows):
        product = xl_sheet.cell(row_i, 0).value.strip()
        project = xl_sheet.cell(row_i, 1).value.strip()
        iteration = xl_sheet.cell(row_i, 2).value.strip()
        story = xl_sheet.cell(row_i, 3).value.strip() 
        task = xl_sheet.cell(row_i, 4).value.strip() 
        comment = xl_sheet.cell(row_i, 5).value.strip() 
        user = xl_sheet.cell(row_i, 6).value.strip() 
        date = xl_sheet.cell(row_i, 7).value
        spent_effort = xl_sheet.cell(row_i, 8).value

        if (use_first_name):
            user = user.split(" ")[0] #Note: this means people with the same first name are grouped together
        else:
            user = user
            
        if user not in users:
            ##print("Adding " + str(user))
            users[user] = 0   
            
        users[user] += float(spent_effort)        
            
        if display_tag_errors:
            words = re.split(" |\.|:|\[", comment)
            tags = set()
            for word in words:
                if len(word) > 0 and word[0] == "#":
                    tags.add(word[1:].lower())
            if story != '' and not (valid_tags.intersection(tags) and commitless_tags.intersection(tags)):
                # No valid tags, and not a task without story
                print(user + ": " + comment)
                
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
    display_tag_errors = False
    
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
                    
                if arg == 't':
                    display_tag_errors = True
                
        print_hours(xl_sheet, use_first_name, sorting, number_of_decimal_places, reverse_order, rank, display_tag_errors)        
            
            
if __name__ == "__main__":
    main(sys.argv)
