M2000 Interpreter and Environment

Version 13 revision 9 active-X
1. A small fix. Now module C run as expected.

2.  Const for Groups, BigIntegers, Tuple
Also using the class CopyArray() we can place a group which has value, so  when we use the constant the object return a copy of the internal saved array (tuple). When a group return a value (has Value part) then the const sub system execute the value before pass the value (so const never pass a group which have a value, but the group's value.


const a=(1,2i)
const b=12&
const c=lambda (x)->x**2
const d=12129371897398173981739871128371237821u
const e=exp(1)
class CopyArray {
	m
	value {
		=cons(.m)
	}
class:
	module CopyArray(a as array) {
		.m<=cons(a)
	}
}
const f=CopyArray((1,2,3,4))
Print a=(1,2i), a|r=1, a|i=2
Print d*100
Print e
Print f#sum()
z=f  // pointers to arrays
z1=f
print z is z1 ' false: in't same
Print z 
Print z1
return z, 2:=100
' we can change z items, but no f items
Print f#val(2)=3, z#val(2)=100
try {	' return find f is a group so can't go further and raise error
	return f, 2:=100
}
Print Error$  ' Wrong Use of Return
list

George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement Settings to change font/language/colors and size of console letters.

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

                                                             