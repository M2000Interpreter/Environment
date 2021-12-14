M2000 Interpreter and Environment

Version 10 revision 47 active-X
1. UPGRADE the file functions/statements, to handle over 2GB files:
OPEN, CLOSE, LINE INPUT, INPUT, PRINT, WRITE, SEEK, PUT, GET
SEEK(), RECORDS(), EOF()  Now the file pointer is Currency.
2. Correction for WRITEWITH example (in info file), now work perfect. (found in the file support upgrade phase)
3. Some minor fixes.

George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The fist time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)


http://georgekarras.blogspot.gr/

ExportM2000 all files with executables:
https://drive.google.com/drive/folders/1IbYgPtwaWpWC5pXLRqEaTaSoky37iK16

only source without executables (something going wrong with GitHub)
https://github.com/M2000Interpreter/Environment

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             