# Planilha (Spreadsheet)

Published in three parts, all of them included in the 4th volume of bound fascicles in the most popular INPUT binding scheme in Brazil. This is 100% BASIC code with no machine code embedded and very few POKEs and PEEKs and alikes, and wow it´s a spreadsheet, so I thought it would be a good first program to type in and study.

The MSX BASIC code file here has 8954 characters in 229 lines &ndash; I have included 5 comment (REM) lines and some line breaks, so it is actually a bit larger than the original, print version.

I tried to run the program on [WebMSX](https://webmsx.org) and [MSXPen](https://msxpen.com/) (uses the same WebMSX engine) but had no luck. The only way to get the file into the emulator is adding it to a disk image, but probably because of MSX Disk BASIC I kept getting an "Out of memory in 30" error (while DIMensioning the two main arrays).

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
- Place the cursor in the cell to be copied, then press L if the cell will be copied to different **L**ine(s) or press C if the cell will be copied to different **C**olumn(s);
- Then you will be asked to type in the first and the last destination cell addresses.

Sheets are **S**aved or **L**oaded by pressing S or L anytime but when inserting values in a cell. Disk-drives were expensive and rare in Brazil back when INPUT was published, so the program is written to use a _data corder_ (cassette tape recorder), but this can be changed easily in lines 1260 and 1390. Apparently, data is saved in a CSV-like fashion (I still have to look closer at this).

To **Q**uit the program, press Q.

## Main variables

In old BASIC all variables are global, and names are quite commonly restricted to 2 or even 1 _significant_ character(s) (some computers like the MSX accepted variables with bigger names, but the third and following characters are just ignored). Subroutines act on these global variables, so I thought I should start by finding the most important variables first.

#### D$(26,30) and D(26,30)
Defined in line 30. These two-dimensional arrays store, respectively, formulae and their results for one sheet.

Supposing that 26 is the number of letters in the alphabet and columns are refered to by letters, I wonder why the declaration looks inverted, as a two-dimensional DIM is usually taken as DIM(rows,cols). Indeed, loops (like the one in lines 750 to 780) show an inversion of columns and rows between the array and the sheet view. Furthermore, D$ is initialized in line 40 with CHR$(128) characters, and I still can´t figure out why: in Brazilian MSX computers it would be the "Ç" character (80h).

#### MO$(2) and MO
Related to the view/operation **MO**des of values or equations. Defined in line 20. The two-element array MO$ holds the strings for sheet view modes (values or equations). MO is initialized to 1 (equation) and seems to hold the current operation mode.

#### RV$ and RV
I presume the names stand for **R**esult **V**alue. RV is set by all subroutines which perform the supported cell calculations. RV$

#### CC and CR
These numerical variables seem to be named after _**C**urrent **C**olumn_ and _**C**urrent **R**ow_.

#### A$
Apparently, this is a string variable used to catch input in various and not so related parts of the program.

#### OP$
From its declaration in line 20, this is a string which holds all the supported numerical **OP**erators.

### Literals and constants
MSX BASIC had no named constants. Instead, there are plenty literals and magic numbers all around the code. Knowing beforehand what some of them mean might help.

- **CHR$(91)** - Opening bracket \[, 5Bh
- **CHR$(93)** - Closing bracket \], 5Dh
- **CHR$(219)** - Graphic full block in MSX (DBh)
- **CHR$() 28, 29, 30, 31** - Cursor keys, respectively right, left, up and down: 1Ch, 1Dh, 1Eh, 1Fh
- As far as I could understand, the following "extended ASCII" characters are used as prefixes in cell contents stored in D$ in order to classify them for proper processing:
  - **CHR$(128)** - 80h, _Ç_ in most MSXs, seems to be used as a placeholder for empty cells, in lines 40, 420, 680 (from 670), 1130, 1170 and 1290.
  - **CHR$(129)** - 81h is the _ü_ character in most MSX machines. The value is refered in lines 590, 690 and 770. Identifies values.
  - **CHR$(130)** - 82h, character _é_. Used in line 600, 690 (from 670) and 1170. Identifies labels (textual content).
  - **CHR$(131)** - 83h, character _â_. Mentioned in lines 550, 820 and 1550. Identifies formulae.

## Subroutines (GOTOs and GOSUBs)

#### GOSUB 70 - 160 - Draw main screen
Called at lines 60, 270, 290, 400

Starts by printing a 3 graphic blocks at the third line of the screen.

Line 90 contains some mysterious calculation and screen output in a FOR...NEXT loop. Let´s open it:
```BASIC
FOR I=0 TO 15
  C1=INT((RS+I)/10)+48
  C2=(RS+I)-((C1-48)*10)+48
  LOCATE 0,I+1
  PRINT CHR$(C1);CHR$(C2):
NEXT
```
RS was declared and initialized in line 20 with value 1.

#### GOSUB 130 - 160 - Update the "status line" at screen top
Called at line 320, this target is still part of the main screen drawing routine started at line 70 &ndash; the jump looks like a dirty trick of branching to the _middle_ of a subroutine!

#### GOTO 170 - 320 - Command keys input
We get here right after initialization, from line 60. This section puts the program in "waiting command" state, and branches processing when any of the cursor or command keys is pressed.

Things start with some nice mathemagics and the only PEEKs and POKEs in this code. The graphic full block character (219, or DBh)


#### GOSUB 410 - 650 - Cell content insertion
Called at line 260.



### GOTO 610

#### GOSUB 660
Called at lines 120, 610

#### GOTO 720

#### GOSUB 730 - Get value from cell address
Called at lines 880 and 890, it is a one-line subroutine which "parses" a cell address like A19 stored in variable Z$, gets its correspondent value from the D array and puts it in variable V.

#### GOSUB 740 - 1220 - Recalculate the whole sheet
Called only in line 100 if the view mode is set to show values.

This subroutine contains an enormous nested loop, ranging from lines 790 to 1210, with subroutine calls and declarations, internal and external jumps, screen operations and everything.

#### GOTO 920 - 990 - Prepare cell output
You can get here only from lines 1080 and 1120, as a final move of both subroutines for vertical and horizontal range sums. I can´t guess why this was not made into a subroutine.

Variable DP (**D**ecimal **P**laces?) has its value set right before the jump and its content is the optional digit which determines how many decimal places should be exhibited in the result of a formula.

PU$ might stand for *P*rint *U*sing as it is set to a formatting string for PRINT USING. I have a feeling that a dot is missing at line 930.

#### GOSUB 1000, 1010, 1020, 1030, 1040
Those one-line subroutines are called at 910 according to the operator found in a cell formula by the INSTR function.
Notice that for a 0 divisor the routine at 1030 returns 0.

#### GOTO 1050 - 1080 - Vertical range sum
Flow comes here from line 850.

#### GOTO 1090 - 1120 - Horizontal range sum
Branched to by line 860.

#### GOTO 1130
One gets here from line 820.

Entangled logic here. Apparently, lines 1140 and 1150 could be eliminated if line 1130 had been written as:

```BASIC
1130 IF ASC(D$(I,J))<>128 THEN RV$=MID$(D$(I,J),2) ELSE RV$=STRING$(7,32)
```
Which makes the same: in the end, RV$ has 7 blank spaces or the row number from a cell address.

#### GOTO 1160

#### GOSUB 1230 - Save sheet file
Called at 300

#### GOSUB 1350 - Load sheet from file
Called at 310

#### GOSUB 1490 - Cell copy "dialog"
