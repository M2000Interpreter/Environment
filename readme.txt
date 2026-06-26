M2000 Interpreter and Environment
Version 14 Revision 53

1) Latest revisions use a() of an object as a call to default property of object "a" using no parameters (or a number of parameters)
(for the time being we can't use named parametes here). We get a named conflict if we try to use DIM a() if a is an object (local to current module). This is not bad, but what if the "a" object is a BigNumber (treated as numeric value, but is an object). So this revision do not raise error if "a" is biginteger (there is no default property for biginteger).
1.1 Test Code:
	biginteger a=1000u
	//	a=(1,2,3,4) ' is OK
	dim a(10) as long ' now this pass
	list

1.2 Test Code which raise error (we have local object):
	a=getobject("","m2000.VarItem") 
	//	a=list ' the same if a is a list (because a() and a$() exist for list)
	//	buffer clear a as Integer*100 ' memory buffer a(0) is the address at offset 0. 
	buffer clear a as Integer*100 ' memory buffer a(0) is the address at offset 0. 
	dim a(10) as long ' we get error becaue "a" is an object.
	list

1.3 Test Code which not raise error because local shadow global:
	global a=getobject("","m2000.VarItem")
	dim a(10) as long ' now this pass 
	list

2) Swap Array items in any combination if array item is type of Biginteger
Was not problem when array was variant. Now fixed for all other combinations.




George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
THEN press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 