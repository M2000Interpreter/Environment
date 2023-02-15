M2000 Interpreter and Environment

Version 12 Revision 14 active-X
Added logic on making types with type at the left of variable (long a=100). So now we didn't change a global variable with same name if that global variable was made before the current module/function. For previous versions only the Long exist as a modifier to long type. To modify a global we have to use Set before the type modifier so the interpreter at that moment turn to global scope and do the modification. This difference breaks compatibility with older versions, but is a small one. There are some other differences with old version like the A[123] which for older versions is an identifier, but for version 12 is A (has to be a pointer to a RefArray) with index 123. A name may have [ ] characters if has the first one or the first character after dot. So the hidden variables of properties for groups which have these characters are ok. So a proprty A for group B has a This.[A] or .[A] or [A] identifier for the private variable which make the interpreter automatic.
Alsp two more modules added in info file: Arch (Archemides Spiral) and MapRange.


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
                                                             