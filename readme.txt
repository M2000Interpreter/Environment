M2000 Interpreter and Environment
Version 14 Revision 38


1) fix an error which occurs from revision 36 (for 36 to 37).
variant z
long k[10]
z=(1,2,3)
z=k  ' this raise error (revision 36 and 37)


2) Variables for Complex numbers do not have ++, --, += ... operators. For tuple the operator repeated all items but skipped complex items. For arrays with parenthesis now return error message "wrong operator". Before was an exception on the code and returned "type mismatch".
a=((1,-1i), (2,3i))
a++  ' no message - for complex ++ not work
link a to a()
a(0)++  ' wrong operator ' before was type mismatch a VB6 error.

3) Fix Variant z=<Expr ret object>, old workout was to split to two statements: variant z: z=<Expr ret object>. Example shows variant z to get a group. Because z is variant when we pass a copy of a pointer to group if we use =, so first time variant z=alfa() gets a copy of a pointer. This example shows when each object is destroyed and count the number of objects destroyed.
 
global counter=1
module tst {
	class alfa {
		x=10
		remove {
			print "deleted ", counter
			counter++
		}
	}
	dim a(10)<<alfa() '' 10 objects
	variant z=alfa() ' now work - before need variant z: z=alfa()
	' 11 objects
	print z=>x=10  ' z is a pointer to alfa
	z->a(4)  'get a pointer of a(4) ' z=a(4) get a pointer of a copy
	z=>x+=100
	? z=>x=110
	? a(4).x=110
	' z is variant - can't hold a named group
	' 12 objects
	z=a(5)  ' get a copy as a pointer to group
	z=>x+=100
	? z=>x=110
	? a(5).x=10
	b=alfa()  ' b is a named group
	' 13 objects = z has a new object
	z=b ' get a pointer to a copy of b
	b.x+=1000
	? b.x=1010
	? z=>x=10
	' z get a pointer to b (as a weak reference not a real pointer)
	z->b
	b.x+=1000   ' b.x = 2010
	? z=>x=2010
	clear b  ' to execute remove part
}
tst



George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and then open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 