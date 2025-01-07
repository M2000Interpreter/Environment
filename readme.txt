M2000 Interpreter and Environment

Version 12 Revision 62 active-X
1. Fixed an error from rev 58: long a = 100: a|div 1000: ? a ' normal 0 (I do a bad refactoring so this was 0 - only for long type variables), I found it when I run Magic4 module in Info.gsb.

2. Arrays with square brackets now have |Div |Mod |Div# |Mod# (I found that for these I did nothing until now).

3. Format$() fixed so the fault of reading the width for formating a boolean value on BTABLE module (previous revision 60 worked fine). Also this continue to what rev. 60 did, to display fix dot for BigInteger, to render like a real value with fix fractional part, with or without a total width for rendering.


locale 1033
' no need u at the end for var xx as biginteger=....
var a as biginteger=93923809803980918919831083180312
? a
? format$("{0:10}", a)  ' without a 
? format$("{0:10:-40}", a)  ' -40 width 40 chars Right Justify
var b as long long=12123131231
? format$("{0:20}", b)  '  for long long is width 20 left justify
? format$("{0:-20}", b)  ' for long long is width 20 right justify
var double c=23234.34232, d=1234.2234e-124, e=3.e45, f=100
' change locale to 1032 to see the change of "dot" char
? format$("{0:-20}", d)  ' 1.E-121 (NOT GOOD)  for double width 20 left justify
' better:
? format$("{0:3:-20}", d)  ' 1.234E-121  for double width 20 left justify
' same as this:
? format$("{0:-20}", str$(d,"0.000E+###"))  ' 1.234E-121  for double width 20 left justify
? format$("{0:2:-20}", c)  ' 23234.34 for double width 20 right justify
? format$("{0:2:-20}", e)  ' 3.00E+45  for double width 20 right justify
? format$("{0:2:-20}", f)  ' 100.00 for double width 20 right justify
' strings:
? format$("{0:-20} {0:20} {1}", "stringA", "stringb")
' format with only one string procsses specoal chars.
print #-2, format$("hello\r\nSecond Line \u234")   ' u234 is 0x234
print chrcode$(0x234)=format$("\u234")  ' true



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

                                                             