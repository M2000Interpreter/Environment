M2000 Interpreter and Environment
Version 14 revision 5 (updated - info file) active-X

1. frac(expression!) return fraction of the result of expression. frac(expression) return fraction of the result of the rounded number ( at the 13th digit).
2. Writable() now work without raising error for drives which didn't exist. Another way to check if a path may exist then ShortDir$(path$) return the short dir or if not exist return empty string. So shortdir$("a:")="" means drive a not exist.
3. Declare a small fix for ..., which now we can use it in a multiline declaration, using optional remarks using ' or // or \\:
mybuf$=String$(Chr$(0), 1000)
Declare Global MyPrint Lib C "msvcrt.swprintf" {
	&sBuf$,  ' byref pass - pass address of first byte of string
	sFmt$, ' by value pass
	... ' variadic parameters  - always as last parameter
	} 
A=MyPrint(&myBuf$, "P1=%s, P2=%d, P3=%.4f, P4=%s", "ABC", 123&, 1.23456, "xyz")
Print Left$(myBuf$,A)  ' A has the number of characters (words - 2 bytes per character)


 
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

https://rosettacode.miraheze.org/wiki/M2000_Interpreter (544 tasks)
Old (not working rosettacode.org)
https://rosettacode.org/wiki/Category:M2000_Interpreter (544 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 