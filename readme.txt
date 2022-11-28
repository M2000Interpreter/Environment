M2000 Interpreter and Environment

Version 11 Revision 19 active-X
1. A very rare bug founded and removed. This bug come from using operator for comparison for Groups (objects), when we use a pointer to group before the operator.

2. New info examples: Priority, and Priority1 (for priority queues). Showkeys which open a user form now close using any key press (excluding special keys like arrows).

3. New Functionality for User Forms: Using Windows key and up arrow or down arrow we can handle the maximize - restore to original size, for user forms which can change size. We must press the arrow key and quick press the window key, then press the appropriate arrow key again. The user forms in M2000 simulate the title bar, so this style normally didn't get the "maximize" command from windows, using Windows Key and arrows (there are many combinations, and yet I implement a fraction of it). Check it with info examples, press F8 to open mEditor, or run CS (c sharp editor) and htmlEditor.

George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The fist time you run the interpreter do this in M2000 console:
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

M2000language.exe (Chrome can't scan, say it is a virus - heuristic choise)
All exe files are signed
https://drive.google.com/u/0/uc?id=1hjEO6XvAu-l7TTwPYmPEkZZrxAXPtA41

M2000 paper (305 pages). Included in M2000language.exe
https://drive.google.com/file/d/1pHBjLVeaGkyMhyyfvXyvh42cJ3njY7wa

M2000 Greek Small Manual (488 pages). Included in M2000language.exe
https://drive.google.com/file/d/0BwSrrDW66vvvS2lzQzhvZWJ0RVE
                                                             