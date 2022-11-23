M2000 Interpreter and Environment

Version 11 Revision 18 active-X
- Bug fixes: Val() statement now works fine, MINESWEEPER in info works fine too.
- I found a piece of old code which removed, but prevent to call a sub when a function had same name, like function alfa() { } and Sub alfa() ... End Sub. The Gosub alfa() has no problem, but if we call alfa() without Gosub the current block terminate execution like executing exit statement, without any notice. So now this bug removed.

- Also I found this: If we have a global array A() and we have a Dim A() in a module we change the global array, we don't make a local one. To prevent this we have to do this: Dim New A(), because we know it is a global array A(), or for a library (where some other place a global A() before execution of the module) is better to use Dim New (but if this line executed again before the end of module execution a new local array A() shadow the last local one, so it is better to use this only one time in a module's block (and then for redim we can use DIM A(), because local array has priority for picking). If a module read an array to A() and exist a global A() there is no problem because the read statement (which get the array from stack of values) never read to a global array, except we execute the READ in console's command line, or using SET READ where Set execute a line of code like it is a manual entered, from M2000 console. Those places have namespace empty. All other places have the namespace from module or function. For DIM this happen because a DIM define arrays, and can redifene arrays (works as REDIM PRESERVE for already defined arrays).


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
                                                             