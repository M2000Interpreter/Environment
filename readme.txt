M2000 Interpreter and Environment

Version 13 revision 22 active-X
Last fix.
The lambda$ break from version 12 revision 50
Because the second form (lambda vs lambda$) used, which had no problem. Before some versions, a string expression had only strings literals or and string functions with $ suffix like Mid$(). The last versions can use string variables and functions without $ suffix (but a name with $ suffix is different with the same name without the $ suffix).

Dim A$(3)
A$(1)=lambda$ (a$,wd)->field$(a$, wd)
? "["+A$(1)("hello", 10)+"]"
link A$() to A()
' this works on revision 21
A(2)=lambda (a$,wd)->field$(a$, wd)
? "["+(A(2)("hello", 10))+"]"

George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (534 tasks)
                 