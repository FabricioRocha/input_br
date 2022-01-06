# Planilha (Spreadsheet)

Published in three parts, all of them included in the 4th volume of bound fascicles in the most popular INPUT binding scheme in Brazil. The MSX BASIC code file here has 8954 characters in 229 lines &ndash; I have included 5 comment (REM) lines and some line breaks, so it is actually a bit larger than the original, print version.

I still have not tested the code in an emulator.

It is a very simple spreadsheet program which makes heavy use of string manipulation functions in BASIC, especially the following:

- **VAL()**
- **STR()**
- **ASC()**
- **CHR$()**
- **MID$(st$, fr, nm)** and **MID$(st$, fr, to)=s2$** - in the first form the function returns a number _nm_ of characters from a string _st$_ starting from the _fr_ position; in the second form the function **replaces** part of the string _st$_ with part of another string _s2$_.

Most of the code is actually related to screen input/output.

## Operation

According to the third part of the article which brought it, this spreadsheet uses an unusual format for its cell formulae, in a kind of Reverse Polish Notation (RPN), like "A1A2+".

Literal values are **not** allowed.


## Main variables

In old BASIC all variables are global, and names are quite commonly restricted to 2 or even 1 _significant_ character(s) (some computers like the MSX accepted variables with bigger names, but the third and following characters are just ignored).

#### D$(26,30) and D(26,30)
Defined in line 30. These two-dimensional arrays store, respectively, formulae and their results for one sheet.

#### CC and CR
These numerical variables seem to be named after _**C**urrent **C**olumn_ and _**C**urrent **R**ow_.

#### A$
Apparently, this is a string variable used to catch input in various and not so related parts of the program.

#### OP$
From its declaration in line 20, this is a string which holds all the supported numerical **OP**erators.
