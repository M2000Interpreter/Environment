M2000 Interpreter and Environment

Version 11 Revision 13 active-X
Fix three bugs, and one addition:
1) Enum alfa {a=1, b=2}: Push a: Print Number. This was error because a is an object in stack. Now interpreter take the value of enum a and return from Number (which also pop the value).
2) Print 10 div 2 //remark using //
This also was an error, because 10 div 2/2 is 10 div (2/2).
So now 10 div 2//any char 
is 10 div 2 (the // reject characters until the start of next line)
3)Statement codepage was a fault, because Win32 function (under the hood) return 1 or 0, and was Not retvalue, but in VB6 Not 0 is -1 (ok that), and Not 1 is -2, which is no zero so it is true again. So this change to revalue=0, for 0 give True, for non zero give false.
4)Structure alfa {x as integer, y as integer} by design not work with comma between field definitions, now work with comma and also comment type "//" work good too. Also a structure definition into a structure definition is posible, with or without *multiplier and a new symbol ";" can be used to break the union (which interpreter apply by default).

**** M2000Paper Updated (see Structures with new examples).
**** Updated the links in this txt file.


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
                                                             