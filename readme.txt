M2000 Interpreter and Environment

Version 13 revision 1 active-X
Two things:
1 If we place white backcolor and black forecolor now the drop down list on M2000 console (open using Menu statement) now change the colors enough for good work of the XOR at the menu item (which always display white as the "inverted" color).
This work with any combination when A xor B equal 0xFFFFFF:

This is the test program:
PEN #dd00FF
CLS #22FF00

MENU "OK","AA","BB"
? MENU

2. Some changes for printing to printers (I found a fault but not on my computer so I expand a precious soloution for another part of specific print code). I would like to change a lot, so this is an entry change.



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

                                                             