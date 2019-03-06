#!/usr/bin/python3

from __future__ import print_function

import re
import sys
from os.path import join, dirname, abspath
from collections import defaultdict

import xlrd

valid_tags = {"implement", "document", "test", "testmanual", "fix", "chore", "refactor"}
commitless_tags = {'commits', 'chore', 'document', 'testmanual'}
required_columns = {"Story", "Comment", "User", "Spent effort (hours)"}

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
ta: print all log comments
"""


def get_first_sheet(filename):
    # Open the workbook
    xl_workbook = xlrd.open_workbook(filename)

    # Return first sheet
    return xl_workbook.sheet_by_index(0)


def strip_if_string(x): return x.strip() if isinstance(x, str) else x


def get_commit_hashes(string):
    if '#commits' in string:
        trimmed_string = string[string.find('#commits') + len('#commits['):]
    else:
        return []
    return map(str.strip, trimmed_string[: trimmed_string.find(']')].split(','))


def process_entry(entry, users, existing_comments, use_first_name, display_all_tags, display_tag_errors):
    story = entry["Story"]
    comment = entry["Comment"]
    user = entry["User"]
    spent_effort = entry["Spent effort (hours)"]

    if use_first_name:
        user = user.split(" ")[0]  # Note: this means people with the same first name are grouped together

    users[user] += float(spent_effort)

    if display_all_tags:
        print(user + ": " + comment)

    if display_tag_errors:
        words = re.split(" |\.|:|\[", comment)
        tags = set()
        for word in words:
            if len(word) > 0 and word[0] == "#":
                tags.add(word[1:].lower())
        if story != '':
            if not (valid_tags.intersection(tags) and commitless_tags.intersection(tags)):
                # No valid tags, and not a task without story
                print(user + ": " + comment)

            else:
                for old_user, old_comment in existing_comments:
                    if old_comment == comment and 'pair' not in tags:
                        if old_user == user:
                            print(user + ': [warning: duplicate] ' + comment)
                        else:
                            print(user + ' & ' + old_user + ': [warning: untagged pair] ' + comment)

        commit_hashes = get_commit_hashes(comment)
        for commit_hash in commit_hashes:
            if len(commit_hash) not in [7, 8, 40]:
                print(user + ': [warning: commit SHA len] ' + comment)
                break
    return user, comment


def print_hours(xl_sheet, use_first_name=False, sorting="hours", number_of_decimal_places=1,
                reverse_order=False, rank=True, display_tag_errors=False, display_all_tags=False):
    """
    Prints the number of hours each user has done, and optionally other data

    :param xl_sheet: the sheet to get the data from
    :param use_first_name: if True, just displays first name, otherwise displays full name
    :param sorting: sets sorting method:
        "alpha": alphabetical by first name
        "hours": by number of hours (least to most) (default)
    :param number_of_decimal_places: how specific the hour count for each person should be
    :param reverse_order: if True, reverses order of list
    :param rank: if True, adds a ranking to each item in the list (1 -> n)
    :param display_tag_errors: if True, prints out each log that isn't tagged correctly
    :param display_all_tags: if True, prints out all logs
    """
    time_per_user = defaultdict(int)

    # Generate hours-worked data

    processed_results = set()

    headers = xl_sheet.row_values(0)
    if not set(headers).issuperset(required_columns):
        print("The input file does not contain all the required columns, expected: ", required_columns)
        print("Missing: ", required_columns.difference(headers))
        print("Got: ", headers)
        exit(1)

    # For each row (except header row):
    for row_i in range(1, xl_sheet.nrows):
        stripped_values = map(strip_if_string, xl_sheet.row_values(row_i))
        entry = dict(zip(headers, stripped_values))
        result = process_entry(entry, time_per_user, processed_results, use_first_name, display_all_tags, display_tag_errors)
        processed_results.add(result)

    max_length = max([len(user) for user in time_per_user])
    output_format = '{0:>%d}: {1:.%df}' % (max_length, number_of_decimal_places)
    rank_format = '{0:>%d}: ' % (len(str(len(time_per_user))))

    # Print it

    print()

    if sorting == "alpha":
        user_list = sorted(time_per_user.keys())

    else:  ## if (sorting == "hours"):
        user_list = sorted(time_per_user, key=time_per_user.get)

    if reverse_order:
        user_list.reverse()

    i = 0
    for user in user_list:
        i += 1
        if rank:
            print(rank_format.format(i), end='')
        print(output_format.format(user, time_per_user[user]))


def main(args):
    filename = join(dirname(abspath(__file__)), "agilefantTimesheet.xls")
    xl_sheet = get_first_sheet(filename)

    sorting = "hours"
    number_of_decimal_places = 1
    use_first_name = False
    reverse_order = False
    rank = False
    display_tag_errors = False
    display_all_tags = False

    if 'h' in args:
        print(HELPTEXT)

    else:
        if len(args) > 1:
            for arg in args:
                if arg == 'a':
                    sorting = "alpha"

                if arg[0] == 'd':
                    decimal_places_str = arg[1:]
                    if decimal_places_str.isdigit():
                        number_of_decimal_places = int(decimal_places_str)

                if arg == 'f':
                    use_first_name = True

                if arg == 'n':
                    rank = True

                if arg == 'r':
                    reverse_order = True

                if arg == 't':
                    display_tag_errors = True

                if arg == 'ta':
                    display_all_tags = True

        print_hours(xl_sheet, use_first_name, sorting, number_of_decimal_places,
                    reverse_order, rank, display_tag_errors, display_all_tags)


if __name__ == "__main__":
    main(sys.argv)
