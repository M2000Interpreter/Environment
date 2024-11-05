M2000 Interpreter and Environment

Version 12 Revision 42 active-X
1. #SORT() FOR BYTE ARRAYS NOW WORKS.
Dim a(3) as byte
a(0)=255,12,18
Print a()#sort()

2. date$(a, language_ID) works as expected
LOCALE 1033
date a="2024-12-25"
? date$(a, 1032)="12/25/2024"
? date$(a)<>"25/12/2024" 
LOCALE 1032
date b="2024-12-23"
? date$(b, 1033)="12/23/2024"
? date$(b)<>"23/12/2024" 
Long d=a-b
? d=2

3. RefArray type of arrays (extension)
- using [] not ()
- each "dimension" may have different length
- adjust size by using index: byte b[0]: b[7]=255:Print Len(b)=8
- Was 1 to 2 dimension, now can be more, making object arrays to hold 1 or 2 dimension arrays.

3.1 Multi assign values
variant z[10]
z[0]="George", 12122, 12&, "hello", 0.0001212312312312312@, 0.0001212312312312312
for i=0 to 10:print z[i]: next

3.2 Making 4 dimension array (using a function to pack arrays inside objects)
// Integer is 16 bit (2 bytes), Object is 4 bytes (32 bit pointer to object).
function multidim_integer() {
	if stack.size=1 then
		integer a[number]
		=a 
	else.if stack.size=2 then
		integer a[number][number]
		=a
	else.if stack.size>2 then
		read k
		bb=[]
		object b[0]
		for j=0 to k
			b[j]=lambda(!stack(bb))
		next
		=b
	end if
}
// 0 to 2, 0 to 4, 0 to 2, 0 to 3
// 3*5*3*4=180 items
k=multidim_integer(2, 4, 2, 3)
k[1][4][2][3]=100
k[1][4][2][3]++
Print k[1][4][2][3]=101
k[1][4][1][0]=1, 2, 3, 4
for i=0 to 3: Print k[1][4][1][i],:Next:Print

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

                                                             