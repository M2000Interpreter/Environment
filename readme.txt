M2000 Interpreter and Environment
Version 15 Revision 17

July 27, 2026,
Very fast new revision. The programmer of Regex (in M2000 the ASF_RegexEngine.cls, code from https://github.com/ECP-Solutions) fix the issue 27 (which I found): Now Replace("Myer, Ken") return the same value as VBscript.

declare RegEx "M2000.Regex"
RegEx=>Pattern = "(\S+), (\S+)"
Print RegEx=>Replace("Myer, Ken", "$2 $1")="Ken Myer"
Clear
declare RegEx "VBscript.RegExp"
RegEx=>Pattern = "(\S+), (\S+)"
Print RegEx=>Replace("Myer, Ken", "$2 $1")="Ken Myer"





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

English old paper for M2000
https://github.com/M2000Interpreter/Environment/releases/download/ver13rev44/M2000paper.pdf

Greek Book for learning programming
https://github.com/M2000Interpreter/Environment/releases/download/version15revision10/GreekBookM2000.pdf

Greek Manual (a work in progress)
https://github.com/M2000Interpreter/Environment/releases/download/version15revision10/GreekManualVersion15_preview.pdf

Greek Book About OOP in M2000
https://github.com/M2000Interpreter/Environment/releases/download/version14revision51/OOP_M2000_2026.pdf

http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 