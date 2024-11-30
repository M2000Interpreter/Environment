M2000 Interpreter and Environment

Version 12 Revision 50 active-X
1. Bug fixed (problem was on push/data not on array())
string s="ok"
flush
push s, s
k=array([]) ' current stack to array
? k   ' now print ok ok  (before nothing)
data s, s
k=array([])
? k  ' now print ok ok  (before nothing)

2. Bug Fixed (problem was on read statement when value was string and variable has to be variant type, but get the string type and not changed after that). Now is ok.
module WasOk {
	push 100
	read z as variant
	z="string value"
	? type$(z)="String"
	z=100
	? type$(z)="Double"
}
WasOk
module WasNotOk {
	push "string value"
	read z as variant
	? type$(z)="String"
	z=100
	? type$(z)="Double"
}
WasNotOk

3. Fix LET and variant variable to get a string value. Works now with and without Global clause
global variant s=100
? s
let s="12345"
? s
List



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

                                                             