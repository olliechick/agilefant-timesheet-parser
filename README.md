# Agilefant timesheet parser
A parser for timesheets exported by the free version of Agilefant. I developed this to use for SENG302 in 2018.

# Running the program

Run the command python3 `timesheet-parser.py`.

It accepts command-line arguments, in the form `timesheet-parser.py <command>`, e.g. `timesheet-parser.py a b c`.

## Available arguments

Argument |What it does
---------|------------
`a`      |Sort users in alphabetical order (by first name)
`d[int]` |Uses `int` decimal places (e.g. `d3`)
`f`      |Use first names (note that this means two users with the same first name get grouped together)
`h`      |Print help text (instead of running the main program)
`n`      |Include each user's ranking (a number ;` -> n)
`r`      |Print in reverse sorted order
`t`      |Print all invalid log comments
