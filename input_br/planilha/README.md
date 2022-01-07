# Planilha (Spreadsheet)

Published in three parts, all of them included in the 4th volume of bound fascicles in the most popular INPUT binding scheme in Brazil. This is 100% BASIC code with no machine code embedded and very few POKEs and PEEKs and alikes, and wow it´s a spreadsheet, so I thought it would be a good first program to type in and study.

The MSX BASIC code file here has 8954 characters in 229 lines &ndash; I have included 5 comment (REM) lines and some line breaks, so it is actually a bit larger than the original, print version. I still have not tested the code in an emulator.

It is a very simple spreadsheet program which makes heavy use of string manipulation functions in BASIC, especially the following:

- **VAL(st$)** - returns numerical value it can extract from a string _st$_
- **STR$(v)** - converts a numerical value _v_ into a string. The opposite of VAL().
- **ASC(c$)** - returns the ASCII value of a given character (or first character of) _c$_
- **CHR$(c)** - opposite of ASC(), returns the character corresponding to the one-byte code _c_ from the computer´s character table
- **MID$(st$, fr, nm)** and **MID$(st$, fr, to)=s2$** - in the first form the function returns a number _nm_ of characters from a string _st$_ starting from the _fr_ position; in the second form the function **replaces** part of the string _st$_ with part of another string _s2$_.

Most of the code is actually related to console input/output. The program uses MSX SCREEN 0, a 40x24 display of 6x8 pixel characters. Screen locations are 0 based.

## Operation

Instructions came with the third and last part of the article and code. The spreadsheet is of course very limited and peculiar in some ways, even for the 1980s.

- **V**alues in the page and the _**E**quations_ which resulted in some of them are shown separately, by pressing the V and E keys.
- The cursor keys navigate through cells.
- To **I**nsert a value or formula in the cell where the cursor is, press I.
- Cell formulae are written in a kind of Reverse Polish Notation (RPN), like "A1A2+". They are restricted to two cell addresses as operands (but they may refer to start and end of a range in some operations). A digit after the operator can be used to state the desired number of decimal places in the result.
- The following operators are available:
  - basic four operations: +, -, *, /
  - vertical range sum (through a number of lines): &
  - horizontal range sum (through a number of columns): $
  - percentage (%)
- Literal values are **not** allowed in formulae. To be used in a formula, as a constant, a literal value should be placed in the Z column to be used as a cell reference. This constraint is due to the copy mechanism explained below.

There is a mechanism to allow formulae to be replicated to other cells, either as absolute (unchanged) duplicates or as _relative copies_ with their cell references automatically changed &ndash; so you can have A1B1* in C1 and have the formula copied to the following rows as A2B2*, A3B3*, etc. Copies are done this way:

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

Supposing that 26 is the number of letters in the alphabet and columns are refered to by letters, I wonder why the declaration looks inverted, as a two-dimensional DIM is usually taken as DIM(rows,cols). Also D$ is initialized in line 40 with CHR$(128) characters, and I can´t figure out what exactly the character means in Brazilian MSX computers.


#### MO$(2) and MO
Related to the view/operation **MO**des of values or equations. Defined in line 20. The two-element array MO$ holds the strings for sheet view modes (values or equations). MO is initialized to 1 (equation) and seems to hold the current operation mode.

#### CC and CR
These numerical variables seem to be named after _**C**urrent **C**olumn_ and _**C**urrent **R**ow_.

#### A$
Apparently, this is a string variable used to catch input in various and not so related parts of the program.

#### OP$
From its declaration in line 20, this is a string which holds all the supported numerical **OP**erators.

### Literals and constants
MSX BASIC had no named constants, and literals abound in the code. Knowing beforehand what some of them mean might help.

**CHR$(91)** - Opening bracket \[, 5Bh
**CHR$(93)** - Closing bracket \], 5Dh
**CHR$(219)** - Graphic full block in MSX (DBh)
**CHR$() 28, 29, 30, 31** - Cursor keys, respectively right, left, up and down: 1Ch, 1Dh, 1Eh, 1Fh

## Subroutines (GOTOs and GOSUBs)

#### GOSUB 70 - 160 - Draw main screen
Called at lines 60, 270, 290, 400

Starts by printing a 3 graphic blocks at the third line of the screen

#### GOSUB 130
Called at line 320, looks like a dirty trick of branching to the _middle_ of a subroutine!

#### GOTO 170

#### GOTO 410 - Insert content into a cell


#### GOTO 600
From 490, 510, 530, 540

### GOTO 610

#### GOSUB 660
Called at 610

#### GOTO 720

#### GOSUB 730
Called at 890.

#### GOSUB 740

#### GOTO 920
You can get here from lines 1080 and 1120

#### GOSUB 1000, 1010, 1020, 1030, 1040
These one-line subroutines are called at 910 according to the operator found in a cell formula by the INSTR function.
Notice that for a 0 divisor the routine at 1030 returns 0.

#### GOTO 1050 - 1080

#### GOTO 1090

#### GOTO 1160

#### GOSUB 1230 - Save sheet file
Called at 300

#### GOSUB 1350 - Load sheet from file
Called at 310

#### GOSUB 1490 - Cell copy "dialog"
