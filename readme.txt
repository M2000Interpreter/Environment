M2000 Interpreter and Environment
Version 14 Revision 28

1. For the M2000 structures I remove one UDT (contents) and split the FastCollection to two classes, the FastCollection and the StructCollection. Is better to have a dedicated class and not one for everything. 
2. I made a manifest for the current dll (you can find it in the Release section), for the SxS (side by side) use of the dll without refistration. From version 12 revision 9 I start to use UDT which break the regfree use of M2000.dll (But some day I found the way to use the udt as is and make the SxS a working history).
3. The M2000.exe now can call external regsvr32  through Powershell (asking for elevated execution - only for the regsvr23 command).

 

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