# Agilefant timesheet parser
A parser for timesheets exported by the free version of Agilefant. I developed this to use for SENG302 in 2018. It has since been updated (by [@jackodsteel](https://github.com/jackodsteel)) to parse the headers (rather than assume a fixed layout) so it can be used for SENG302 in 2019 too, as the format has changed slightly (one less column is generated).

## Pre-requisites

Install python3 and xlrd.

## Running the program

Save the exported timesheet (`agilefantTimesheet.xls`) in the root directory of this repo, then run the command `./timesheet-parser.py`.

It accepts command-line arguments, in the form `./timesheet-parser.py <command>`, e.g. `./timesheet-parser.py a b c`. My personal favourite is `./timesheet-parser.py r n`, which produces a ranked list of who has logged the most hours.

### Available arguments

Argument |What it does
---------|------------
`a`      |Sort users in alphabetical order (by first name)
`d[int]` |Uses `int` decimal places (e.g. `d3`)
`f`      |Use first names (note that this means two users with the same first name get grouped together)
`h`      |Prints the help text (instead of running the main program)
`n`      |Include each user's ranking (a number; 1 -> n)
`r`      |Print in reverse sorted order
`t`      |Print all invalid log comments (for University of Canterbury course SENG302)
`ta`     |Print all log comments

#### When invalid log comments occur

Name | When it occurs
-|-
`error: missing tags` | A log doesn't have any of the seven tags (`implement`, `document`, `test`, `testmanual`, `fix`, `chore`, `refactor`)
`error: empty comment` | A log have an empty comment
`warning: no #commits tag` | A log has `implement`, `test`, `fix`,  or `refactor` without a `commits` tag
`warning: commit SHA len` | A `commits` tag has a commit SHA with a length that is not 7, 8, or 40
`warning: untagged pair` | There are two logs by different people with the same comment but no `pair` tag
`warning: duplicate` | There are two logs by the same person with the same comment
`warning: singleton #pair` | There is a log with a `pair` tag that doesn't have a matching comment with another log

 

## What it does
It prints a list of users and how many hours they have logged.

By default, this is sorted by the number of hours logged, from least to most, and 1 decimal place is used (e.g. `12.5`).

### Example output

```
$ ./timesheet-parser.py 

          Ollie Chick: 50.6
       Example person: 53.5
Verylong-named person: 63.5

$ ./timesheet-parser.py r n

1:        Example person: 63.5
2: Verylong-named person: 53.5
3:           Ollie Chick: 50.6
```
