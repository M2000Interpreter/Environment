M2000 Interpreter and Environment

Version 13 active-X
This is the 13th version of M2000. 

Expression Evaluator is capable to use Complex numeric ang BigIntegers, among other types.

One small thing breaks the compatibility with previous versions: Parentheses over a string now return string. Previous versions return a single item array. So (“alfa”) was like (1,) as one item array. This change to (“alfa”,) now. There is a module in Info file called compiler which uses the old format, and now this is changed to new standard format.
Earlier versions of M2000 uses two evaluators, the  logic/arithmetic evaluator and the string evaluator, and a string always had to return from the string evaluator. Because for the use of string variables without the postfix $ the evaluators expanded to do many things: So the arithmetic evaluator may find a string, and the string evaluator may find a number. The rule here is that string comes first. So, if we mix numbers and strings, we get strings.
The new two types: Complex and BigInteger can be mixed with other numeric types. Complex after BigInteger raise error, but BigInteger after Complex convert BigInteger to double value. Complex can be used with all other numeric values. If we use BigInteger without using Complex type all other numeric values are converted to BigInteger. If a String is founded on expression then we get a string expression. Complex and BigInteger values converted to strings automatically (Complex use decimal point from the current Locale id). Operators like Div and Mod can’t be used with Complex (we get errors). We can use +, -, *, / and power ** or ^ and unary -/+. There are new functions Mod() (used from BigInteger for returning modulus from division and for Complex type which return the complex Modulus). There are other like Conjugate(), Phase(), Arg(), Abs() (like Mod()), Cos(), Sin(), Exp(), Tan(), Atn(), Str$() for convert to string using formatting string, Round() for rounding both imaginary and real values. We can read imaginary and real values using |, so for variable a, the a|r and the a|I are the two parts of complex number (they are doubles). These are read/write properties (internal is a UDT, user defined type from math2 class). There are more functions which we can use from this class. We can define arrays of complex types, and of BigInteger type. BigInteger internal is an object.
There are examples of complex and BigInteger types in info file. (See about the setup of M2000 Interpreter)




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

                                                             