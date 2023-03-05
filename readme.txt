M2000 Interpreter and Environment

Version 12 Revision 19 active-X
1. Sign + now used as unary + (identity) as unary - (negation). Also evaluator fixed now, so (-2)^2 return 4, and (-2)^3 return -8. There is a new module in INFO file the Evaluator which test using code written in M2000 to produce RPN and execute it, with identical results with interval evaluator, which use through Eval(). ? -2^+3^-4==-1.008594091577 return true (== used for rounding to 13th and then check equality). ? -2^3^-4==-0.000244140625 (look seems  -2^+3^-4 is -2^3^-4 but it isn't). Use Evaluator module to find the RPN (Reverse Polish Notation) for these expressions.

2. bug removed: integer x=1, y (now y isn't 1 but 0)

3.Fixed some statements to use strings with variable name without suffix $:
open, copy, clipboard, Input (this work continue, so check next revisions)
For input statement we can use also arrays and now interpreter check if it has numeric or string value. Prior this not happen, because Input use the suffix to find if we want to input string or number. Now also look the suffix, but without suffix look for type too. So old programs can be used as is.

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

M2000language.exe (Chrome can't scan, say it is a virus - heuristic choise)
All exe files are signed
https://drive.google.com/u/0/uc?id=1hjEO6XvAu-l7TTwPYmPEkZZrxAXPtA41

M2000 paper (305 pages). Included in M2000language.exe
https://drive.google.com/file/d/1pHBjLVeaGkyMhyyfvXyvh42cJ3njY7wa

M2000 Greek Small Manual (488 pages). Included in M2000language.exe
https://drive.google.com/file/d/0BwSrrDW66vvvS2lzQzhvZWJ0RVE
                                                             