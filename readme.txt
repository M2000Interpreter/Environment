M2000 Interpreter and Environment

Version 13 revision 2 active-X
1. Remove of stop commands...(used for debugging)
2. Fix some bugs
This bug comes from combined ccode to produce the Complex number using parentesis:
k=((100,),("aaaa",),(3,))
? k#val(1)#val$(0)="aaaa"
3. Added some features
Printer.Margins    ' reset margins
Printer.Margins LeftInTwips
Printer.Margins LeftInTwips, TopInTwips
Because Lmargin shift page left, the Right margin computed from the original width of the page, so we have to add the left margin. This is the same for bottom margin. These two place a white strip to exclude printing area, they don't resize the page.
Printer.Margins LeftInTwips, TopInTwips, RightInTwips
Printer.Margins LeftInTwips, TopInTwips, RightInTwips, BottomInTwips

Dates based on locale. Str$() based on Locale 1033 always.
Date and Complex now works fine using initial type/values

Test Code:
DATE A=45000, B="27/10/1966"
LIST
CLEAR
VAR DATE A=45000, B="27/10/1966"
LIST
CLEAR
DEF DATE A=45000, B="27/10/1966"
LIST
CLEAR
LOCAL DATE A=45000, B="27/10/1966"
LIST
CLEAR
GLOBAL DATE A=45000, B="27/10/1966"
LIST
CLEAR
VAR A AS DATE=45000, B AS DATE="27/10/1966"
LIST
CLEAR
DEF A AS DATE=45000, B AS DATE="27/10/1966"
LIST
CLEAR
LOCAL A AS DATE=45000, B AS DATE="27/10/1966"
LIST
CLEAR
GLOBAL A AS DATE=45000, B AS DATE="27/10/1966"
LIST
CLEAR
COMPLEX A = (1,-2I)
LIST
CLEAR
VAR COMPLEX A = (1,-2I)
LIST
CLEAR
DEF COMPLEX A = (1,-2I)
LIST
CLEAR
LOCAL COMPLEX A = (1,-2I)
LIST
CLEAR
GLOBAL COMPLEX A = (1,-2I)
LIST
CLEAR
VAR A AS COMPLEX = (1,-2I)
LIST
CLEAR
DEF A AS COMPLEX = (1,-2I)
LIST
CLEAR
LOCAL A AS COMPLEX = (1,-2I)
LIST
CLEAR
GLOBAL A AS COMPLEX = (1,-2I)
LIST
CLEAR
LOCALE 1033
ALFA(45000)
BETA()
LOCALE 1032  ' GREEK
ALFA(45000)
BETA()

SUB BETA(A AS DATE=45000)
	? A
END SUB
SUB ALFA(A AS DATE)
	? A
END SUB



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

                                                             