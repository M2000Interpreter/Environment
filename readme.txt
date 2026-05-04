M2000 Interpreter and Environment
Version 14 revision 25 active-X
George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com
1. Fix a mistake at the break function, which produce an overflow.
2. At the textbox for watching variables in the test form, the enumerator type with string value not displayed using the string format for the purpose of the form. Although was not wrong the value shown as is, not bad but not what we want. This fault produced by the change of the evaluator, which now return string from the part which previous works only for numbers/objects


The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and then open it again.

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

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 