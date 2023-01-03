M2000 Interpreter and Environment

Version 12 active-X
This is the first revision of Version 12.
This version an advanced expression evaluator, similar to VB6. Now -2^2 return -4, (Vesion 11 and lower works like Excel, so 0-2^2 return -4 but -2^2 return 4).
Also has an advanced system for error control.
We may have strings with id names like numeric.we can define like this a="string". Some work needed for checking all embeded functions.
So Mid$(a$,1,3) works with names as numeric variables Mid$(a, 2, 3).
There is a new type, the Variant. So we can use local variables which may change type. This needed for passing by reference for invoking methods for external objects, which have by reference Variant. (old versions, use variants but for by reference call pass the type, so there was nothing from interpreter to know if a variable (which was a variant) can pass as variant or as the type. So now Interpreter having the type Variant, knows that the by reference pass have type Variant.



This isn't a finished version, although I do a lot of debugging.



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
                                                             