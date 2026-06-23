M2000 Interpreter and Environment
Version 14 Revision 51

1) Remove a bug (two letters supposed to be Greek but was English) for IF (the Greek equivalent). This isn't something for English statements.

2) New limit for pool of object's for stack items. 
Flush Garbage ' this exist from previous revisions - reset the two pools, of varitem (stack item object) and mstiva (stack object).
Flush Garbage 100000 reset pools and set limit to 100 thousand varitems. So if we use more, that more will destroyed and not recycled to pool
Flush Garbage 100000; set only the limit.
So if we get a lot of stack items, we can return to the lower limit. Default limit is the 200 thousands objects.


3) New read only variable SYMBOL which return the Varitem, the object under each item on a stack object. So Push 123 wrap 123 in a Varitem and place it to an item in mstiva object (the stack of values). A statement A=Symbol pop the VarItem, not the value which carry. A statement Push A, A, A places only three pointers of the same object, not in three varitems, but in place we use for varitems.
In M2000 stack objects using Symbol we can say "all is an object".

def Symbol()=Symbol
def sVal(a)=a
object A=Symbol("hello"), B=Symbol(123456780000000000000000000000000001u)
object C=Symbol((1,2,3,4,5))
? sVal(A), sVal(B), sVal(C)
Push A, A, A  
Stack
Stack New {
	Push A, A, A
	Print Array([])#Str$()="Hello Hello Hello"
}
A=>mItem="Other"
Stack  ' print stack items: Other Other Other
Print Array([])#Str$()="Other Other Other"
let m=a  ' push a : read m ' read unbox value from a
var k=a  ' set k = a
Print type$(m)="String", k is a



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