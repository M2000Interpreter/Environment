M2000 Interpreter and Environment

Version 10 revision 36 active-X
1)A broken statement:
1. Fix the STOP statement. Now M1 module in info works fine.

2-4)Improvements:
2. Point now return 0x7FFFFFFF when we get color out of layer.
3. Gradient statement to Player layers preserve colour in non displayed pixels (transparent by a region on the window). So We can test if mouse pointer on a sprite layer return 0x7FFFFFFF (out) or the transparent color which we use at Player statement.
4. Module Sprites in info now has the new definitions (2 & 3) to handle better the displayed sprites using mouse.

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

                                                             