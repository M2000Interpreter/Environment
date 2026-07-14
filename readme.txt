M2000 Interpreter and Environment
Version 15 Revision 4

1) Objects for arguments when we use object1=>function(object2)
 
BIGINTEGER A=100U
' A is object of type BigInteger
' Function ADD() return BigInteger
' But need BigInteger as parameter
' I fix the simple format =>functionOrProperty()
' to get Objects as parameters (I forgot to do that)
PRINT A=>ADD(1u)^=A+1
' This is the old way which always works nice
' and can be used to pass named parameters and by reference too.
Method A,"Add", 12u as B
Print B=112u

2) Fix a bug which is rare to occur. 
First the Inner module need to get the pointer and produce a copy as a named Group, when the Group is type ONE.(We use: a as *One to get the pointer to a, but we have to use a=>X to read the value).

The problem was at the second call to Inner when the first one use a pointer of type weak reference (not a real pointer, means the object life not depend from the weak reference), and the second one use the real pointer - from a copy - (like in this example). M2000 not offer the real pointer from a group which are "Static" like alfa (so always the static objects are deleted at the exit of the code, module's or function's where they defined) 

GROUP alfa {
	Type: One	
	Χ=10
}
Module Inner (a as One){		
	List
	Print a.Χ
}
' Pointer(alfa) is a weak reference to alfa 
Inner Pointer(alfa)
' Pointer((alfa)) is a real pointer to a copy of alfa
Inner Pointer((alfa))


  
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