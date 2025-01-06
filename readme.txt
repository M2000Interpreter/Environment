M2000 Interpreter and Environment

Version 12 Revision 60 active-X

1. Modulus of -25u div -7u return -4
This is same as M2000 and VB6. I adjust the Rebecca Gabriella's String Math Module to the same as of M2000 mod function. Also I made the Euclid div# and mod# to work with bigintegers.
2. Function mod(a) return the last modulus who stored to a:
We get all prints (? is alias for Print statement)

biginteger a=-25u div 7
? a=-3, mod(a)=-4
? 7*a+mod(a)=-25
? type$(a)="BigInteger"
m=mod(a)
? type$(m)="BigInteger"
biginteger a=-25u div# 7
? a=-4, mod(a)=3
? 7*a+mod(a)=-25
? type$(a)="BigInteger"
m=mod(a)
? type$(m)="BigInteger"

3. Operators: <, >, >=, >=, <>, =, ==, <=> works now for BigInteger

4. For now we can do 2u^3u (internal 2u turn to decimal, and 2u turn to decimal too, and then perform the power operation, which return type double). Use modpow(a,b,c) for a^b mod c function on BigInteger.

5. Sqrt() now work for BigInteger (Return integer square).

6. Sgn() and Abs() also works for BigInteger too.

See BigInt example in info file.
 


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time you run the interpreter do this in M2000 console:
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

M2000language.exe (Chrome can't scan, say it is a virus - heuristic choice)
All exe/dll files are signed
https://github.com/M2000Interpreter/Environment/releases

M2000 paper (305 pages). Included in M2000language.exe
M2000 Greek Small Manual (488 pages). Included in M2000language.exe

                                                             