M2000 Interpreter and Environment
Version 14 Revision 54

' Fixed the Read buffer item like Read a[2] when a is a Buffer object.
' test program
buffer a as long*20  ' long is not signed here
? "Buffer a start at address:";a(0);" for ";len(a);"bytes"
a[0]=100
a[1]=-100  ' only unsigned values
? a[1]=0 ' is zero
a[1]=Uint(-100) ' you have to change to unsinged value
? Sint(a[1])=-100 ' you have to change to signed value
// a[2]="George" put the bytes from string from offset 2*item_length
Let	a[2]="George"  ' this is new addition (it is a Push "George":Read a[2])
print a[2,6*2]="George"
Hex a[2], a[3], a[4] ' two characters per long
' same as this
push "George"
read a[5]  ' this is new  addition
print a[5,6*2]="George"
Hex a[5], a[6], a[7] ' two characters per long
' this is the old working feature, which also works
a[8]="George" ' this work  
print a[8,6*2]="George"
Hex a[8], a[9], a[10] ' two characters per long



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