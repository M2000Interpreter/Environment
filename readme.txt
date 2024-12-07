M2000 Interpreter and Environment

Version 12 Revision 53 active-X
1. Fix Do Until and Do When to be used with/without boolean values. A zero is false and a non zero is true.
a=0
do
	? "ok"
	if inkey$<>"" then exit
until a
a=100
do
	? "ok"
	if inkey$<>"" then exit
when a
2. Arrays with brackets can be used for moving bytes from/to binary files (like Buffer)
	Arrays of type variant, string and object can't be used for this.
double a[4]=123.23, b[4]
a[3]=121.e23
open "datapass.dat" for output as #f
	put #f, a, 1
close #f
? filelen("datapass.dat")
open "datapass.dat" for input as #f
	get #f, b[], 1
close #f
for i=0 to 4
	? a[i], b[i]
next

double a[4]=123.23, b[4], c[1]
a[3]=121.e23
open "datapass1.dat" for output as #f
	put #f, a[2], 1, 8*2
close #f
? filelen("datapass1.dat")
open "datapass1.dat" for input as #f
	get #f, b, 1
close #f
for i=0 to 4
? a[i], b[i]
next
open "datapass1.dat" for output as #f
	put #f, a[], 1
close #f
N=2  ' from
L=2 ' many
W=8 ' byte per item
open "datapass1.dat" for input as #f
	? records(#f)  ' same as filelen() because record width is one byte
	get #f, c[0], W*N+1, W*L
close #f
for i=0 to 1
	print i, c[i]
next
a[1]=123456
open "datapass1.dat" for append as #f
	N=4  ' try 5
	put #f, a[1], W*N+1, w
close #f
close 0
double a[5]
open "datapass1.dat" for input as #f
	get #f, a, 1
close #f
Print array(a)
	
3. Conversion from array with breckets to a tuple/array, and the reverse:
dim a(10) as long=10
a(3)=500
double a[0]
a=a()
? a[3]=500, len(a)=10
// we can put a variant type
a=("aaa", 1, 20.2232)
? len(a), a[0]="aaa", a[1]=1, a[2]=20.2232
// array() return a tuple from array with brackets
? array(a)#sum()=21.2232


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

                                                             