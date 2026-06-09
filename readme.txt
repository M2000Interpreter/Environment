M2000 Interpreter and Environment
Version 14 Revision 44

I fixed an error from revision 40. I forgot to fix Interpreter for constant objects and a lambda function. The Interpreter finds the constant (is an object of type constant) and then uses this object to pass arguments x, y to the default property. That is fault. So, I changed it, so now Interpret skips this object and then find the function L() and we get the result as expected. Previous revisions are ok because they did nothing for any object, except the Inventory object.

This is the code:
Const L=Lambda (x, y)->x^y
Print L(2,3)=8  ' her was the error

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