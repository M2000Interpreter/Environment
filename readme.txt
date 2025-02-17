M2000 Interpreter and Environment

Version 13 revision 11 active-X
1. Sort for tuple now work fine for Biginteger:
function prepareArray {
	dim a(4) as BIGINTEGER=-32u
	a(0)=120u
	a(1)=15u
	=a()   ' this is mArray type
}
// cons()  comes form lisp, but here add a list of arrays or and tuple and return a tuple.
for this { // this is a block for temporary definitions
	a=cons(prepareArray())#expanse(10)
	print type$(a)="tuple"
	return a, 2:=23244u
	def type(x)=type$(x)
	Print type(a#val(0))
	Print a#sort()
}
// so now we have no variables
for this {
	a=prepareArray()
	print type$(a)="mArray"
	return a, 2:=23244u
	def type(x)=type$(x)
	Print type(a#val(0))
	try ok {
		Print a#sort()
	}
	If error or not ok Then Print Error$
	// this say: Invalid item for sorting. I found object
	// BigInteger is an object.
}


2. #MAT() fixed for string values:
a=(,)
a=a#expanse(120)#mat("=", "Hello")#mat("+=", " World")
aa=each(A)
while aa
	print format$("{0:-3}",aa^);")"+array(aa)
	aa=each(A, aa^+10)  ' this is the Step 10
end while


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (534 tasks)
                 