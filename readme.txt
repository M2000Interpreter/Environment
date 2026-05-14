M2000 Interpreter and Environment
Version 14 Revision 32

1) Fix for Greek functions of type Function/end function for names with accents (normal functions registered in uppercase so any name first convert to uppercase with no accents for Greek names, but simple functions use the name as is with accents, we can call with uppercase name but we have to use the accents as those used on the definition of the function)

2) I give more speed with removing a upper case conversion which not need any more. 



George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

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