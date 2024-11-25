M2000 Interpreter and Environment

Version 12 Revision 48 active-X
Updated info file.

1. Fix function Compare() for strings without $ (and arrays with strings without $)
(not work for arrays with brackets. use <=> )
2. Fix statement Swap for strings without $
a$="aaaa" :b="klmn"
swap b, a$ : print a$, b
swap a$, b : print a$, b
3. Fix Swap for arrays with brackets: A[]
long b[3][4]=100
object c[3]
c[0]=b[]  // get a copy of b
let c[0][2][1]=40
? c[0][2][1]<> b[2][1]
c[1]=b   // get a pointer to b
let c[1][2][1]=400
? c[1][2][1]=b[2][1]

gen=lambda k=0 ->{
	k++
	=k
}
? c[1][2][1]=400, c[0][2][1]=40
// gen() ret 1 then ret 2
// swap c[1][2][1], c[0][2][1]
swap c[gen()][gen()][1], c[0][2][1]
? c[1][2][1]=40, c[0][2][1]=400

4. Fix Max.Data() and Min.Data() for using String Expressions too. The first expression define the type of the other expressions (numeric or string).

5. Fix boolean to string conversion using Locale (for English and Greek people):
dim b$(10)
link b$() to b()
boolean t=true
Locale 1033
b(2)=t
b(3)=122.23
? b$(2)="True", b$(3)="122.23"
boolean z
string s
s = z
Print "FalseFalseFalse"=z+s+z ' True
clear  'erase all variables
dim b$(10)
link b$() to b()
boolean t=true
Locale 1032
b(2)=t
b(3)=122.23
? b$(2)="Αληθές", b$(3)="122,23"
boolean z
string s
s = z
Print "ΨευδέςΨευδέςΨευδές"=z+s+z  ' Αληθές΅
exit
This not work: z+s+z="ΨευδέςΨευδέςΨευδές"
because z is not string
This work: ""+z+s+z="ΨευδέςΨευδέςΨευδές"



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

                                                             