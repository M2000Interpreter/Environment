M2000 Interpreter and Environment

Version 12 Revision 56 active-X
1. Addition for BigInteger Class:
- IsProbablyPrime( BigNumber, Iterations as integer)
- IntSqr()
- modpow(expBigNumber, ModulusBigNumber)
2.
// modpow example
function ToString(x as *BigInteger) {
	with x,"toString" as ret
	=ret	
}
Function PowerTen(x as integer) {
	a=Biginteger("10")
	method a, "intPower", biginteger(str$(x,"")) as a
	=a
}
a=bigInteger("2988348162058574136915891421498819466320163312926952423791023078876139")
b=biginteger("2351399303373464486466122544523690094744975233415544072992656881240319")
profiler
	method a, "modpow", b, PowerTen(40) as result
print timecount
Print "A=";ToString(a)
Print "B=";ToString(b)
Print "R=";ToString(result)

3.  isProbablyPrime example
// finding ULTRA-USEFUL primes of 2^(2^n)-k, where k is the smaller number where the formula give a prime number.
// you can go more than 6th but you need time for the process.
minusOne=biginteger("-1")	
one=biginteger("1")
two=biginteger("2")
num=biginteger("0")
k=num
with num,"tostring" as numS
with k,"tostring" as kS
for n= 1 to 6
	k=minusOne
	kk=-1
	n1=biginteger(str$(n,""))
	method two,"intpower", n1 as n1
	do
		method k,"add", two as k
		kk+=2
		method two, "intpower", n1 as num
		method num, "subtract", k as num
		method num,"isProbablyPrime", 10 as ret
		if ret then
			Print n, kk : Refresh
			exit
		end if
	always
next

4. Using IntSqr inside isPrime (isProbablyPrime is faster).
IsPrime=lambda -> {
	boolean T=true, F
	zero=BigInteger("0")
	two=BigInteger("2")
	three=BigInteger("3")
	four=BigInteger("4")
	five=BigInteger("5")
	=lambda T,F,Known1, IntSqrt, zero, two, three, four, five (x as *BigInteger) -> {
		=F
		with x, "toString" as xs
		method x, "compare", five as c
		if c<1 then   ' x<=5
			if c=0 Then =T: break
			method x, "compare", two as c
			if c=0 then =T: break
			method x, "compare", three as c
			if c=0 then =T: break
		end if
		method x,"modulus", two as frac
		method frac,"compare", zero as c
		if c=0 then Exit
		method x,"modulus", three as frac
		method frac,"compare", zero as c
		if c=0 then Exit
		method x,"IntSqr" as x1
		d = five
		with d,"tostring" as ds
		with frac,"tostring" as fracS
		do
			method x,"modulus", d as frac
			method frac,"compare", zero as c
			if c=0 then Exit
			method d, "add", two as d
			method d, "compare", x1 as c
			if c=1 then =T : exit
			method x,"modulus", d as frac
			method frac,"compare", zero as c
			if c=0 then Exit
			method d, "add", four as d
			method d, "compare", x1 as c				
			if c=1 then =T: exit
		Always
	} 
}() ' execute, so IsPrime get the inner lambda

a=BigInteger("5400349")
profiler
? IsPrime(a)
? timecount
profiler
method a,"isProbablyPrime", 5 as ret
? ret
? timecount
5. isProbablyPrime using M2000 code
// place here the Isprime
function isProbablyPrime(n as *BigInteger, k as long) {
	boolean T=true, F=false
	=F
	Zero=BigInteger("0")
	One=BigInteger("1")
	Two=BigInteger("2")
	Method n, "compare", Two as c1
	Method n, "modulus", Two as m2
	method m2,"compare", zero as C2
	if c1=0 or c2=0 then exit
	with n, "toString" as ns$
	s=0
	Method n, "subtract", one as nn
	d=nn
	
	with d, "tostring" as dstr$
	do
		method d,"modulus", two as m2
		method m2,"compare", zero as C
		if c else exit
		s++
		method d, "divide", two as d
	Always
	z=len(ns$)
	a=one
	with a, "toString" as astr$
	x=a
	=T
	for i=1 to k {
		do
			zs=""
			for j=1 to len(ns$)
				zs+=chr$(47+random(1,10))
			next
			a=bigInteger(zs)
			method nn,"compare", a as C
			method a,"compare", one as c1
		until c=1 and c1>-1
		method a, "modpow", d, n as x
		method x,"compare", one as c1
		if c1 else continue
		method x,"compare", nn as c1
		if c1 else continue
		for r=1 to s	{
			method x, "modpow", two, n  as x
			method x,"compare", one as c1
			if c1 else =F : break
			method x,"compare", nn as c1
			if c1 else exit
		}
		if c1 then =F: break
	}
}
a=BigInteger("5400349")
profiler
? IsPrime(a)
? timecount
profiler
ret=isProbablyPrime(a, 5)
? ret
? timecount


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

                                                             