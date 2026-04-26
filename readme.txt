M2000 Interpreter and Environment
Version 14 revision 21 active-X

1. Fix the borken ctrl+z/ctrl+y from revision 20 (was a mistake)
2. Fix the Fill @ ... version of Fill statement (Now KB module show the Notes as letters - the fault was by the new IsExp() function which return string, but only if we pass a variant, and this old function ProcFill use the X as Double, so internal was variant, get the string but convert to number so X returned the 0,which then converted to string and we see the 0 for string value.)



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

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 