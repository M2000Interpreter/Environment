M2000 Interpreter and Environment

Version 12 Revision 40 active-X
Fix a bug for using special arrays like a[] and a[][] for string expressions:
string a[10]="ok1234"
Print a[5] ' works as before
Print left$(a[5],3)="ok1"  ' now works
Print a[2]+a[3]="ok1234ok1234"

string b[10][4]="ok1234"
Print b[5][1] ' works as before
Print left$(b[5][1],3)="ok1"  ' now works
Print b[2][1]+b[3][1]="ok1234ok1234"

b[12][6]="expand"
? b[12][6]

these special arrays have indexes from 0 always and expand length simply by using a bigger index. For two dimension arrays, each sub array may have different length.

For these arrays we want to pass pointer to c functions:
Look module TestArr, we work with pointers. PathAddBackslashW need an address to first character of string. We constract the Declare using Long parameter.
The array a[] has 4 bytes per string to point to string (or maybe is null)
So to find the string at index 1 we have to multiply by 4 and add to address offset of the first item.

The second module show the construction of declaration using by reference string. The first two arrays, a$() and a() are type of variant, but the third is type of string (has inside an object same as the a[] array, but is an marray object who handle max 10 dimensions)

module testArr {
	Declare PathAddBackslash Lib "Shlwapi.PathAddBackslashW" { long string_pointer}
	Declare global GetMem4 lib "msvbvm60.GetMem4" {Long addr, Long retValue}
	
	string a[2]="ok1234"
	a[1] = "C:"+String$(Chr$(0), 250)
	function ArrPtr(a, x) {
		long ret
		With a, "ArrPtr" as a.ArrPtr()
		x=GetMem4(a.ArrPtr(0)+4*x, varptr(ret))
		=ret
	}
	if len(a[1])=0 then exit
	m = PathAddBackslash(Arrptr(a, 1))
	Print LeftPart$(a[1],0)
}
testArr

module testArr2 {
	Declare PathAddBackslash Lib "Shlwapi.PathAddBackslashW" { &Path$ }
	dim a$(4)
	a$(3) = "C:"+String$(Chr$(0), 250)
	A = PathAddBackslash(&a$(3))
	Print LeftPart$(a$(3),0)
	
	dim a(4)
	a(3) = "C:"+String$(Chr$(0), 250)
	A = PathAddBackslash(&a(3))
	Print LeftPart$(a(3),0)
	
	dim s(4, 2, 3) as string
	s(3, 1, 2) = "C:"+String$(Chr$(0), 250)
	A = PathAddBackslash(&s(3,1,2))
	Print LeftPart$(s(3,1,2),0)
}
testArr2



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

                                                             