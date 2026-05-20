M2000 Interpreter and Environment
Version 14 Revision 35

1) fix a type problem in read for basic like programs (when we use BASIC statement)
2) NEW: ON X RESTORE 100, 200, ALFA, 500, BETA
When we use BASIC the RESTORE move the hidden pointer to choose the next DATA statement (DATA inside blocks { } skipped)
So now there are 3 variations for ON (ON x GOTO....|ON x GOSUB....|ON x RESTORE....)
- also help file updated.

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