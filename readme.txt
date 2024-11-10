M2000 Interpreter and Environment

Version 12 Revision 44 active-X
1. A bug removed about the @() special function for Print statement.
move 5000, 3000
polygon 3, ANGLE PI/4,3000,-PI/4,3000,PI/4,-3000,-PI/4,-3000
Print @(10, 10, 10+tab(1), 10+1, 5, 15);"George"

2. Use of a string expression using string variables without $ for:

statement Legend.

3. Hanlde for decimal point based on locale for expression with strings and numeric variables.
locale 1033
a=12.4
? "... "+a    ' ... 12.4
? a+" ..."    ' 12.4 ...
locale 1032
a=12.4
? "... "+a    ' ... 12,4
? a+" ..."    ' 12,4 ...

For using numeric expression we have to put parenthesis:
? "width "+(round(sqrt(2), 2))+" mm ("+locale+")"
width 1.41 mm (1033)
     or
width 1,41 mm (1032)
     base on locale

4. Udpated Info file. So load it again using Dir Appdir$: Load Info
Then press F1 to overwrite the old one on user directory.

George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (384 tasks)

ExportM2000 all files with executables (you can get the ca.crt):
https://drive.google.com/drive/folders/1IbYgPtwaWpWC5pXLRqEaTaSoky37iK16

only source, with old revisions and a wiki, for executables see releases
https://github.com/M2000Interpreter/Environment

M2000language.exe (Chrome can't scan, say it is a virus - heuristic choice)
All exe/dll files are signed
https://github.com/M2000Interpreter/Environment/releases

M2000 paper (305 pages). Included in M2000language.exe
M2000 Greek Small Manual (488 pages). Included in M2000language.exe

                                                             