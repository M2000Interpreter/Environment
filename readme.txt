M2000 Interpreter and Environment

Version 12 Revision 65 active-X
1. More for Structures:
Now we can use the fields of multiple structures in structures
Also we can get part of a struncture and copy back, easy.
Rember structures used for buffers. Buffers are memory buffers attatch to an object "the Buffer". So we handle this object to read/write at memory buffer. We use buffers for binary files and for passing to c functions, or assembly code.
Because buffer is an object we can place it in arrays or passing to functions or return from functions. Also we can make user objects who hold buffers.
Also now we have bound check on arrays of fields. We can use unions to have access anywhere or we can use Eval$() which also can read from anywhere from a buffer. For all readings/writings there is another check at the level of offsets and sizes.

==Example 1==
structure zeta {
	delta as integer * 10
}
structure alfa {
	kappa as double*10
	beta as integer*5
	epsilon as zeta * 10
}
alfa k[10]
Print len(alfa)
k[2]|epsilon[1]|delta[2]=100

Print k[2]|epsilon[1]|delta[2]
k[2]|epsilon[5]=k[2]|epsilon[1]
zeta m
m[0]=k[2]|epsilon[5]
m|delta[2]=m|delta[2]*3
return m, 0!delta!4:=500
m|delta[5]=600
k[2]|epsilon[5]=m[0]
for i=2 to 5
print i, k[2]|epsilon[5]|delta[i]
next
list
==Example 2==
' long long here is Unsigned Long Long (has type Decimal, but values as unsigned long long)
structure epsilon { 'this is union
	{
	betaLong as long*20
	}	
	beta as long long*10
}
structure delta {
	something as byte*100
	beta as epsilon*10
}
buffer clear kappa as long long*5 ' no structure
delta alfa 
epsilon z
? (z(0),alfa(0), kappa(0))#str$(" ")
? len(z)=80
alfa|beta[4]|beta[2]=300
? alfa|beta[4]|beta[2]=300
z=alfa|beta[]
? z[4]|beta[2]=300
? len(Z)=800
zz=z[4]
? zz|beta[2]=300
? zz|betaLong[4]=300  ' low byte
' kappa has no fields, just 10 long long
old=kappa(0)
'return kappa, 0:=zz|beta[0, 80] ' same pointer
'return kappa, 0:=eval$(zz|beta[]) ' same pointer
' kappa[0]=zz|beta[]  ' same pointer but not fit the new one
? len(kappa)=40
kappa=zz|beta[]  ' new pointer to k - expanded
? len(kappa)=80
Print Old+" "+(kappa(0))
Print kappa[2]=300
Print Eval(Kappa, 8*2 as long)=300 ' unsigned long low
Print Eval(Kappa, 8*2+4 as long) =0 ' unsigned long High
with delta, "done", true, "index", 1, "StructMany" as Many, "StructOffset" as Offset
? Many=10, Offset=100
Print "Only alfa has same address, all other changed"
? (z(0),alfa(0), kappa(0))#str$(" ")
list

2. We can define arrays with parenthesis and types as empty arrays. (new from this revision)
==Example 3==
dim a() as long, b() as integer
Print len(a()), len(b())
dim a(10), b(5)
Print len(a()), len(b())
? type$(a(3)), type$(b(4))

3.We can define arrays with square brackets and types as empty arrays. (new from this revision).
==Example 4==
long a[], b[]
Print len(a)=0, len(b)=0
a[10]=100
b[4]=50
z=b
Print z is b = true
Print len(a)=11, len(b)=5
' change pointer to new empty arrays.
long a[], b[]
Print z is b = false
Print len(a)=0, len(b)=0, len(z)=5
Print z[4]=50
List

4. Fix the Format$() which now work as expected
==example 5==
a=1221.1212
Print format$("{0}", a)
Print format$("{0:-10}", a)
Print format$("{0:-10}", ""+a) 
Print format$("{0:10}", a)
Print format$("{0:10}", ""+a)
Print format$("{0:3:-10}", a) 
Print format$("{0::-10}", a)
Print format$("{0:3:10}", a)
Print format$("{0::10}", a)
5. Fix Inner While when the inner one uses multiple iterators.
--Example 6==
Module Iterators (f as long=-2) {
	NameOfDay=("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
	NameOfColor=Stack:="Red","Orange","Yellow","Green","Blue","Purple"
	Positions=list:="first":=1,"fourth":=4,"fifth":=5
	
	Print #f, "All days:"
	m=each(NameOfDay)
	while m    ' ok
		print #f, (m^+1)+". "+array$(m)
	end while
	Print #f
	Print #f, "All colors:"
	m=each(NameOfColor)
	while m
		print #f, (m^+1)+". "+StackItem$(m)
	end while
	Print #f
	Print #f, "Print from positions from the first:"
	a=each(Positions)
	while a   
		m=each(NameOfDay, eval(a))
		n=each(NameOfColor, eval(a))
		while m, n
			Print #f, format$("{0:10} {1:10} {2:10}", eval$(a!), array$(m), stackitem$(n))
			exit
		end while
	end while
	Print #f
	Print #f, "Print from positions from the last:"
	a=each(Positions)
	while a
		m=each(NameOfDay, -eval(a))
		n=each(NameOfColor, -eval(a))
		while m, n
			Print #f, format$("{0:10} {1:10} {2:10}", eval$(a!), array$(m), stackitem$(n))
			exit
		end while		
	end while
}
open "" for wide output as #c
Iterators c
close #c
exit
open "output1.txt" for wide output as #c
Iterators c
close #c
win dir$+"output1.txt"


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

                                                             