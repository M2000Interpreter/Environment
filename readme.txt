M2000 Interpreter and Environment

Version 12 Revision 20 active-X
1. Fixed: an error in M2000 expression evaluator when occur open parenthesis after exponation operator.(was a mistake on code, introduced in version 12, from early revisions). Update Evaluator module (infix to RPN) at INFO file (look how you load the Info file at the end of this text). New module Eval2, produce evaluation to a line of M2000 statements and execute this using INLINE statement (the Eval2 based on a previous Evaluator, but gives same numeric results as Evaluator module.
2. Fixed: += for string variables/arrays without suffix $ 
3. Fixed: EnumStringValue used in Return ListObjectPointer, EnumStringValue:=value
4. R2 module in info, has a Local integer=... which can't parse from version 12, because local integer is statement which require an identifier least (the old module use the simple form Local identifier_name, which also works in current version, but no for names like double, integer etc).
5. Fixed the descent parse of M2000 (which immediate execute parsing fragments) for these situations:
Using keys as strings for lists which retrieve the value from an enum constant with string value. There is also two more examples which use stack object, and pointer to tuple (array). Just do this: Edit A then  press <enter> then copy these and press <Esc>, and write A and press <enter>. Module check1 redefined two times more, this is normal for modules in M2000.

module check1{
	enum alfa {
		s="yes"
		k="no"
	}
	def z as alfa=s
	a=list:=1:="ok",s:=z,2:=3
	print a
	z++
	return a, s:=z
	print a, a("yes")
	
}
check1
module check1{
	enum alfa {
		s="yes"
		k="no"
	}
	def z as alfa=s
	a=Stack:="ok",z,3
	print a
	z++
	return a, 2:=z
	print a, stackitem(a, 2)
}
check1
module check1{
	enum alfa {
		s="yes"
		k="no"
	}
	def z as alfa=s
	a=("ok",z,3)
	print a
	z++
	return a, 1:=z
	print a, a#val(1), array(a,1)="no"
}
check1

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
                                                             