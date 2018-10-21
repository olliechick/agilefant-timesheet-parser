# Agilefant timesheet parser
A parser for timesheets exported by the free version of Agilefant. I developed this to use for SENG302 in 2018.

# Pre-requisites

Install python3 and xlrd.

# Running the program

Run the command `./timesheet-parser.py`.

It accepts command-line arguments, in the form `./timesheet-parser.py <command>`, e.g. `./timesheet-parser.py a b c`. My personal favourite is `./timesheet-parser.py r n`, which produces a ranked list of who has logged the most hours.

## Available arguments

Argument |What it does
---------|------------
`a`      |Sort users in alphabetical order (by first name)
`d[int]` |Uses `int` decimal places (e.g. `d3`)
`f`      |Use first names (note that this means two users with the same first name get grouped together)
`h`      |Prints the help text (instead of running the main program)
`n`      |Include each user's ranking (a number; 1 -> n)
`r`      |Print in reverse sorted order
`t`      |Print all invalid log comments (for University of Canterbury course SENG302)

## What it does
It prints a list of users and how many hours they have logged.

By default, this is sorted by the number of hours logged, from least to most, and 1 decimal place is used (e.g. `12.5`).

## Example output

```
          Ollie Chick: 50.6
       Example person: 53.5
Verylong-named person: 63.5
```
