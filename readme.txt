M2000 Interpreter and Environment

Version 10 revision 48 active-X
Two additions in Interpreter, for errors and one fix in savr file dialog form.
1. Remove check for read only folder in Save.As statement (which prevents the new filename entry).
2. Add check for the known "Dim 12, 1" error (was fatal error, but now return error message)
3. Add check for a ";" after a string expression in Report statement (shows Syntax Error)



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

                                                             