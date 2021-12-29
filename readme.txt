M2000 Interpreter and Environment

Version 10 revision 51 active-X
Fix a bug (from 48 revision), when interpreter decide to call a subroutine without using Gosub statement. The problem was in Echelon example in Info file, in a line A(r,c)/=div1 (because / used for remark, but /= isn't a remark). So instead performing division for array element, decide to call the A(), and that is an error.



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

                                                             