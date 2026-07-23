M2000 Interpreter and Environment
Version 15 Revision 11

July 23, 2026,

1. I found a difficult fault in Document class, for colouring code. So the fix is good for TextViewer Class and GuiEditBox class.
2. Update FORMLABEL and GRADIENT statement for handle string expressions according version's 14 changes (forgot to update)
3. Added two new modules in INFO file and update the Compiler module to Version 15 (although the previous code was good)

  
George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do this:

dir appdir$
load info

then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 