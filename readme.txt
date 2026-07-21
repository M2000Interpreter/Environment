M2000 Interpreter and Environment
Version 15 Revision 10

July 21, 2026,
I am working for the new M2000paper.pdf. When I found something not working I fix it.

1. The Text statement example now working as expected. Also upgraded to support biginteger

Const programname$="MyProgram"
biginteger alfa=1234567890123456789012345678901234567890u
Text UTF-16 logging.txt {##STR$(TODAY+NOW,"YYYYMMDDHHNNSS")##
	This is a BigInteger ##alfa##
	} 
For var1=1 to 10
	Text UTF-16 logging.txt + {This is a line for LOG ##var1## for ##programname$##
	}
Next var1
\\ win temporary$+"logging.txt"' we can open the file in notepad
\\ or we can open using OPEN (not for UTF-8)
Open temporary$+"logging.txt" for wide input as #k
	Try {Seek #k, 3}  ' SKIP BOM
	While not EOF(#k)
		Line Input #k, aLine$
		Print aLine$
	End While
Close #k
\\ delete the log file
Text logging.txt


2. Valid() now return boolean type for false (was integer 0) 
  
George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do this:

dir appdir$
load info

then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 