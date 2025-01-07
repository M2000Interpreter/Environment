M2000 Interpreter and Environment

Version 12 Revision 60 active-X

1. Fix error who prevent module compiler in info to work properly. (the bug came from pervious revision.
2. Add Var or Variable or Variables to define variablew in a line.
There are Static, Local, Global and Def and now there is Var.
The difference is:
- Local and Global always make new variables/arrays.
- Static are variables bound to execution object, not to list of variables. Local variables and Static variables can't exist with the same name. Next execution from the same branch will find the variable preserving the value.
-Def raise error it variable/array exist. If not make local only
-Var make local variables, only if they didn't exist.

Some examples:

var a=100, b, c  ' b and c take 0 double type
var d as integer=100, z as biginteger=13923913912739173912737127
var m=d+z  ' m get the type from z because z has higher priority from imteger
Print type$(m)="BigInteger"
var k="hello", p$="George"
var byte w[100], ww[20][40]=100
print ww[0][2]=100
ww[0][2]++
print ww[0][2]=101
var row0=ww[0]  ' get copy of row
print row0[2]=101
row0[2]++
print row0[2]<>ww[0][2]
' arrays with parenthesis are different from previous type
' inside are flat (one dimension)
var gg(2, 8) as long=10
gg(1,3)=500
List
var gg(20, 8) ' zero only new items
list
gg(1,3)=500
var gg(20, 8)=0 ' zero all items
var p()  ' variant type can be zero item
p()=(1,2,3)
print p(1)=2
list


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

                                                             