M2000 Interpreter and Environment

Version 12 Revision 51 active-X
1) using strings without $ and with indexes from M2000 command line interpreter (CLI)

string a="0k"
string b[30]
b[4]=a
? a=b[4], a, b[4]
string a=1234.5
? a="1234.5"

If we give that: a=1234.5 from CLI we don't get conversion to string, but we change the type of a to numeric. So the statement ? a="1234.5" raise error (Missing number expression (after =)). This isn't a fault. From CLI we create variables as variants always (except the identifiers with $ which are strings or documents (which return string value), and % for numeric integer values - with any type under, just round to half when we assign values)

2) Updates for Speech, Insert and Overwrite. For overwrite we need to place & before a string variable say b, becuase &b is string for "ahead status" the method of M2000 to inspect ahead for an expression before actuall interpret it. There is no always check ahead, so we can use b as string (when ahead status return no string). For Insert and Overwrite also the help file changed.

// speech return the number of voices
// speech$(i) return the name of voice i (from 1 to speech)
// statement speech play using the voice we want
a="Hello George; How are you?"
for i=1 to speech
	print speech$(i)	
	speech a, i
next
// upgrade to array with square brackets
// indexes from 0 to 4
string a[4]
a[1]="Hello George; How are you?"
for i=1 to speech
	print speech$(i)
	speech a[1], i
next
Print Len(a)=5
b="This is a boy"
speech "This is a boy" ! ,1  ' speak like a boy
speech "This is a boy" # ,1  ' spelling
// using variable
speech b ! ,1
speech a[1] ! ,1


3) We can convert numbers to string using locale. We didn't do that for string variables with $, because a pupil (M2000 is for education) can do that fault, and Interpreter need to raise the error. Strings without using $ are for more advanced users.

locale 1032
string a=123.56
? a="123,56"
try ok {
	a$=123.56  ' convet to string not for string variables with $
}
if not ok then print error$  ' no value for variable a$

locale 1032   ' set Greek locale so type of decimal point is ","
string a, b[20]
a=123.56
b[2]=123.54
? a="123,56", b[2]="123,54"
try ok {
	a$=123.56
}
if not ok then print error$  ' no value for variable a$




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

                                                             