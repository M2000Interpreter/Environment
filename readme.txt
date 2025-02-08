M2000 Interpreter and Environment

Version 13 revision 7 active-X
1. The user.name$ return the actual user name of the user of OS, or the M2000 user if we set active one. (Previous the actual user name conflict with OneDrive because the name was extracted from folder name, but now extracted from an OS function, GetUserName from advapi32.dll).
2. Eval$() and Eval() for pointers of groups which return values.
3. Event object add/drop funtion now work fine from different modules.
4. Fix the Roots example in Info (because of changes on overflow on complex expressions)
5. There are three more modules in Info. The FUNCTOR show the Eval$()/Eval() and the EVENTNEW show the Event object, and COMPARE_LIST.
 


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement Settings to change font/language/colors and size of console letters.

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

                                                             