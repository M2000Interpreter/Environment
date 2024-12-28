M2000 Interpreter and Environment

Version 12 Revision 56 active-X
1. Fix the "paste problem". A not careful clear cliboard produce later a problem when using ctrl + V (paste) at M2000 command line.

2. Fixing global/local for properties for objects. See VarPtr in info.gsb. There is a a.ArrPtr() property of a, a RefArray object, which defined as global. So we defined in Check3 and Check4 but from a bug was no new object, so we get an ArrPtr() from the global a, not the new one at check3 and check4. Now works fine as expected.

3. Handle of UDT type of variants from objects MATH and MATH2 (new library). We can use UDT values from external DLL (we have to get one as return value).
We can use UDT in variables and arrays (both () and [] types)

MATH library 
Public Type VecType
    X As Double
    Y As Double
    Z As Double
End Type

Public Type SegType
    Origin As VecType
    AxisF  As VecType   ' Forward, West,  AxisX
    AxisL  As VecType   ' Left,    North, AxisY
    AxisU  As VecType   ' Up,      Up,    AxisZ
End Type

Public Type LineType
    point1 As VecType
    point2 As VecType
End Type

Public Type QuatType                      
    W As Double
    X As Double
    Y As Double
    Z As Double
End Type

MATH2 Library

Public Type cxComplex
    r   As Double
    I   As Double
End Type

Public Type Matrix
    Col As Long                 ' Number of columns
    Row As Long                 ' Number of rows
    D() As Double
End Type

4. BigInteger (Credit to Rebecca Gabriella's String Math Module)
// methods
// MULTIPLY, DIVIDE, SUBTRACT, ADD, ANYBASEINPUT, MODULUS, INTPOWER
a=bigInteger("6864797660130609714981900799081393217269435300143305409394463459185543183397656052122559640661454554977296311391480858037121987999716643812574028291115057151")
b=bigInteger("162259276829213363391578010288127987979798")
method a, "multiply", b as c
with c, "tostring" as ret1  
Print "C=";ret1
method c, "add", biginteger("500") as c
Print "C=";ret1
method c, "modulus", b as c
Print "C=";ret1  '500
' use this to get the result to clipboard: clipboard ret1

==========Output=========
C=1113877103911668754551067286547922604225256302130205294556686724395491813360805268911760151336381665870098050947868749061640471753019685645335076862009313458611655962295771807497288683885535803435498
C=1113877103911668754551067286547922604225256302130205294556686724395491813360805268911760151336381665870098050947868749061640471753019685645335076862009313458611655962295771807497288683885535803435998
C=500

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

                                                             