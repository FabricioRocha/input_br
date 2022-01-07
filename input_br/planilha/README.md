# Planilha (Spreadsheet)

Published in three parts, all of them included in the 4th volume of bound fascicles in the most popular INPUT binding scheme in Brazil. This is 100% BASIC code with no machine code embedded and no POKEs and alikes, and wow it´s a spreadsheet, so I thought it would be a good first program to type in and study.

The MSX BASIC code file here has 8954 characters in 229 lines &ndash; I have included 5 comment (REM) lines and some line breaks, so it is actually a bit larger than the original, print version. I still have not tested the code in an emulator.

It is a very simple spreadsheet program which makes heavy use of string manipulation functions in BASIC, especially the following:

- **VAL()**
- **STR()**
- **ASC()**
- **CHR$()**
- **MID$(st$, fr, nm)** and **MID$(st$, fr, to)=s2$** - in the first form the function returns a number _nm_ of characters from a string _st$_ starting from the _fr_ position; in the second form the function **replaces** part of the string _st$_ with part of another string _s2$_.

Most of the code is actually related to screen input/output.

## Operation

Instructions came with the third and last part of the article and code. The spreadsheet is of course very limited and peculiar in some ways, even for the 1980s.

- Values in the page and the _equations_ which led to them are shown separately, by pressing V and E keys.
- The cursor keys navigate through cells.
- To **I**nsert a value or formula in the cell where the cursor is, press I.
- Cell formulae are written in a kind of Reverse Polish Notation (RPN), like "A1A2+". They are restricted to two cells as operands (but they may refer to start and end of a range in some operations). A digit after the operator can be used to state the desired number of decimal places in the result.
- The following operators are available:
  - basic four operations: +, -, *, /
  - vertical range sum (through a number of lines): &
  - horizontal range sum (through a number of columns): $
  - percentage (%)
- Literal values are **not** allowed in formulae. To be used in a formula, as a constant, a literal value should be placed in the Z column to be used as a cell reference. Explanation follows.

There is a mechanism to allow formulae to be replicated to other cells, either as absolute (unchanged) duplicates or as _relative copies_ with their cell references automatically changed &ndash; so you can have A1B1* in C1 and have the formula copied to the following rows as A2B2*, A3B3*, etc. It works this way:

- Enter the **C**opy mode by pressing C;
- Then you will be asked to press a key to choose the copy mode: **R**elative or **A**bsolute;
- Place the cursor in the cell to be copied, then press L if the cell will be copied to different **L**ine(s) or press C if the cell willbe copied to different **C**olumn(s);
- Then you will be asked to type in the first and the last destination cell addresses.

Sheets are **S**aved or **L**oaded by pressing S or L anytime but when inserting values in a cell. Disk-drives were expensive and rare in Brazil back when INPUT was published, so the program is written to use a _data corder_ (cassette tape recorder), but this can be changed easily in lines 1260 and 1390. Apparently, data is saved in a CSV-like fashion (I still have to look closer at this).

To **Q**uit the program, press Q.

## Main variables

In old BASIC all variables are global, and names are quite commonly restricted to 2 or even 1 _significant_ character(s) (some computers like the MSX accepted variables with bigger names, but the third and following characters are just ignored). Subroutines act on these global variables, so I thought I should start by finding the most important variables first.

#### D$(26,30) and D(26,30)
Defined in line 30. These two-dimensional arrays store, respectively, formulae and their results for one sheet.

#### CC and CR
These numerical variables seem to be named after _**C**urrent **C**olumn_ and _**C**urrent **R**ow_.

#### A$
Apparently, this is a string variable used to catch input in various and not so related parts of the program.

#### OP$
From its declaration in line 20, this is a string which holds all the supported numerical **OP**erators.

## Subroutines (GOTOs and GOSUBs)
