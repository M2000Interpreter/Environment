M2000 Interpreter and Environment

Version 12 Revision 59 active-X

Update of the BigInteger
Many things about BigInteger.
We can make arrays.
We can place them in expressions
Any other numeric type converted to biginteger if a biginteger exist in expression (biginteger has priority over decimal).
+-*/ div mod  (div and / work the same for bigintegers).

 
1) ModPow()
a=848093284092840932840386876876876876892u
b=3432402938402820840923809880009089u
c=10000000000000000u
? modpow(a,b,c)
2) Mod()  return the reminder.
k=10203223230u
a=k/4734u
r=mod(a)
? a*4734u+r
? K, a, r
3) Arrays/Declarations:
c=121212u  ' so now c get literal bigInteger
BigInteger a[10]=100, b=100
Dim A(10, 30) as BigInteger

4) Format$("{0:4:-40}", a) place decimal point so if a = 12345u we get 1.2345 in a field of 40 characters, right justification.

5) Str$(c) return UTF16LE encoding string of the value of biginteger

6) Using BigInteger(string, base) we can comvert from a base (2 to 36)

7) Using With a, "outputbase" as a.base we can handle the output base. So a a.base=2 make the Print a to print a binary number.
Also "characters"+a+"characters" can be done (automatic placed the toString value of bigInteger)
8) We don't need to use ToString property because Print handle it.
9) A=056u: A=100 (now numeric 100 convert to bigInteger)
10) there is no ++ -- -= += *= /= (but some day....)
11) Use compare function for a and b through:
Method a,"compare", b as result
Wait some days to find time to make the compare operators <,>,= etc.


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

                                                             