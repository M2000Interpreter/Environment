M2000 Interpreter and Environment

Version 12 Revision 45 active-X
1. Fix #slice() special function for tuples:
? (1,2,3)#slice(0,0)   ' print 1 before nothing...
? (1,2,3)#slice(1,2)   ' print 2  3 
2. Updated Info.gsb



The M2000language.exe (the setup exe) google say is a virus (but it isn't) due to the use of VB6. The file is signed see bellow.
This is the MD5 file checksum: c324d1b358c62e6366a1735a54212ace



George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (384 tasks)

ExportM2000 all files with executables (you can get the ca.crt):
https://drive.google.com/drive/folders/1IbYgPtwaWpWC5pXLRqEaTaSoky37iK16

only source, with old revisions and a wiki, for executables see releases
https://github.com/M2000Interpreter/Environment

M2000language.exe (Chrome can't scan, say it is a virus - heuristic choice)
All exe/dll files are signed
https://github.com/M2000Interpreter/Environment/releases

M2000 paper (305 pages). Included in M2000language.exe
M2000 Greek Small Manual (488 pages). Included in M2000language.exe

                                                             