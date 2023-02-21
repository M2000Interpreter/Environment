M2000 Interpreter and Environment

Version 12 Revision 16 active-X
1. The Date type included to the set default of types. Suffix for number to convert to date ud, so 44978ud return date. For string to date we use this: date("21/2/2023").
2. Byte suffix is ub, so 0x1Aub, 0xFF_ub and 200ub return byte type.
3. Advanced functionality on Enumerations
4. Select Case/End Select has a variant Select Enum/End Select where in each case we can place member names and the comparison use the names not the values of them (but any other type in cases traited by value).

This example show (3) and (4).

// we can place at run time values to enums. For this revision this happens at enum definition
// values may have numeric types or can be strings. We can't change the values.
// we can make variables of delta and of alfa, and we can assign a member of delta/alfa
b="this"
enum delta {
	  y, k=500, L=b
}
enum alfa {
	z="ok", x, delta, w  // delta place the three members on the internal list of alfa
}
// enum alfa has 5 items, but define only z and x (the other defined as delta type, so we can't change type)
// in the definition of inner the enum alfa is not in scope (we can make it global, but not for this example)
// How Interpreter knows about alfa? from the b. Also Interpreter keep at local storage the last enum to find "missing" identifiers, like the z, x, y here. So the key to execute this function is the clause "as alfa".
// If the b pass the "as alfa" the function execute the body.
function inner(b as alfa) {
	// using select case we get error
	// z,x,y aren't in same scope
	// but select enum use the names of members of b (and values if no names used)
	//  Len() return the index of member (from 0), eval$() return the name of member
	? Len(b), eval$(b), b, 
	select enum b
	case z
		? "  I am z"
	case x
		? "  I am x"
	case y
		? "  I am y"
	case 500
		? "  I am k"
	case "this"
		? "  I am L"
	end select
}
def p as Alfa
// we can't pass L as is, because isn't alfa
try {call inner(L)} : ? Error$
? type$(L)="delta"
p=z : call inner(p)
p=x : call inner(p)
p=y : call inner(p)
p=k : call inner(p)
// because alfa has L, we can pass the L to p (which p is alfa, and L is delta, but member of alfa)
p=L : call inner(p)
p=w : call inner(p)


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The fist time you run the interpreter do this in M2000 console:
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
                                                             