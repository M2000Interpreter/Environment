M2000 Interpreter and Environment

Version 13 revision 12 active-X
1. Input statement now can used for BigIntegers (for console input)
2. Numbers including BigNumbers can use undersocre on code (now fixed for all situations).
3. I put the new formula for converting numbers to BigIntegers (based of the formula used in sort for tuples). The new formula use this VB6 functions: format(int(p),"0") so with this we expand the exponent e.g 1.e30 is 1000000000000000000000000000000, so now we can convert it to biginteger (previous was Cstr(Int(p)) and the 1.e30 converted to "1.e30" which can't convert to biginteger.
This is the equivalent function of M2000 (which internal call format())
Print str$(1e30, "0")="1000000000000000000000000000000"

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
                 