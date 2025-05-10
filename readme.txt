M2000 Interpreter and Environment
Version 13 revision 49 active-X

1. Final fix for fonts which have break character number 13 (like Seqoe UI). Now we can give full justification (left and right) using a left margin and a length (so this is like a right margin).

2. Expand DECLARE for using C and return values like LONG LONG and DOUBLE (8bytes). There is a new module SQLITE3, a demo for the use of SQLITE3, download from internet, unzip it, and use of c functions directly from M2000. We make a database with one table, we insert rows and we read the table and display it. Last we unload the library (sqlite3.dll). (See the info.gsb file for the new code).


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

https://rosettacode.org/wiki/Category:M2000_Interpreter (544 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 