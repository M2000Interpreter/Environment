M2000 Interpreter and Environment
Version 13 revision 34 active-X
New #command as #Stuff() this get value from tuple as is. The difference from #Val() is that the later return arithmetic value or object, and for string values convert them to number ("12.23" returned as 12.23 but "12,23" returned as 0). See the example bellow.
Also the Array([]) which convert as stack object (here the [] is the current stack object) to array, now convert it to tuple, not mArray. From version 13 there is a tuple object, which is a lighter mArray object.

CLASS ALFA {
	X=10
}
DIM A(10)
Z=POINTER(ALFA())
A(3):="OK", 100, (1,2,3,4), Z
M=A()
? M#STUFF(6).X
Z=>X++
? M#STUFF(6).X
ZZ=M#STUFF(6)
? Z IS ZZ
? A(3)
? M#STUFF(3)="OK", M#VAL(3)=0, M#VAL$(3)="OK"
? M#STUFF(4)=100, M#VAL(4)=100, M#VAL$(4)="100"
? M#STUFF(5)#MAX()
' second change:
FLUSH
DATA 1,2,3
m=ARRAY([])
? TYPE$(m)="tuple"



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
                 