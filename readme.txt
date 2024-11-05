M2000 Interpreter and Environment

Version 12 Revision 42 active-X
1. A Bug removed. Passing Enum value with minus sing as parameter to a numeric variable now has the proper sign.
enum aa {one=1, two=2}
def foo(x as long)=x
print foo(-one)=-1  ' previous was 1
2. Using READ and LET for arrays with square brackets
Single a[4]
// the Let evaluate first the expression right from assign symbol =
// then evaluate the index(es) of the array 
Let a[4]=1.343 ' not supported before
// without Let evaluation of index(es) of array and then expression after =
a[5]=1.343  ' had no problem
Print a[4], a[5]
data 1.4545
k=random(0, 4)
// Read not supported before for array items of this type of array
read a[k]
print k, a[k]
// still the by reference pass of an array item of this type not supported



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

                                                             