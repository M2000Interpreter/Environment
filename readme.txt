M2000 Interpreter and Environment

Version 12 Revision 64 active-X
1. More on Structures:
==Test Program==
structure alfa {
	{ i as integer *50
	}
	x as byte *100
	b as double
}
alfa a[100]
return a, 4!x!9:=1235 as integer, 4!i!10:=2334
? a[4]|x[9]=211
? a[4]|x[10]=4
? 4*256+211=1235
b=a[4]
? b|i[10]
a[4]|i[10]="hello"
a|b=12312.12312
? a|b
locale 1033
? "["+a|b+"]"="[12312.12312]"
locale 1032
? "["+a|b+"]"="[12312,12312]"
locale 1033
? "hello"=a[4]|i[10, 10]
a[4]|i[0]="hello"
? "hello"=a[4, 10]
? a[4, 10]="hello" and a[4]|i[0, 10]="hello"

==Another Example==
structure beta {
	k as byte *100
}
structure alfa {
	z as integer*50
	delta as beta*10
}
alfa m[30]
beta mm
m[0]|delta[3 as byte]=202
? eval(m, 100+100*3 as byte)=202
? m[0]|delta[3 as byte]=202
mm[0]=m[0]|delta[3]
? mm|k[0]=202
? len(alfa)=1100

2. A bug removed:
The bug was on the reading of weak references at SUBS. The problem not happen to Module, because module has own name space. A sub has the modules name space. So when we call again ALFA() in ALFA() and push the weak reference of B as &B for reading as &A first the interpreter create A (shadowing all A) and then as the &A pop from stack of values, the weak reference attach to B, but which weak reference? The problem: Tho older revision get the newer A which refer to old B, so the new B also refer to old B (so non has the A value).
To correct the problem: I mark the number of currently defined variables, before start to read the stack of values and then on a weak reference I get the variable which has index (at the variables list) less or equal to Mark. So when A created from B, the B get reference not from the last A (index>Mark) but from the previous one. The List is a hash table and all the "same" variables are in the same linked list. Modules not have this problem because its new call to same name a new namespace created so for that name space we didn't have weak references to have problem.

===Test Program---
MODULE ALFA (&A, &B, DEP) {
	IF DEP<1 THEN EXIT
	CALL ALFA &B, &A, DEP-1
	B+=2*A
}
var M=3, A=10, B=5
? A, B
ALFA(&A, &B, M)
? A=60, B=145
var A=10, B=5
? A, B
CALL ALFA &A, &B, M
? A=60, B=145
SUB ALFA(&A, &B, DEP)
	IF DEP<1 THEN EXIT SUB
	ALFA(&B, &A, DEP-1)
	B+=2*A
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

                                                             