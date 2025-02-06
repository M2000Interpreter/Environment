M2000 Interpreter and Environment

Version 13 revision 5 active-X
1. Addition of operators for BigIntegers:
++ -- -! /= *= += -=

a=2390493043902493249032840932840923809842304820398412u
print a
a/=12237239479238749823472893792837493400000u
print mod(a)
print a*12237239479238749823472893792837493400000u+mod(a)
a++
print a
a--
print a
a-!
?a
a+=100924804093480u
print a

2. Operators for biginteger on arrays with parenthesis (we can use max 10 dimensions, internal is flat one dimension)
++ -- -! /= *= += -=   plus = for multiple values (see example)
we can mix normal numbers and bigintegers

dim a(10) as biginteger=10U
a(3)++
a(3)+=21798379832794873298479823798237000000u
a(4)=20000, 2999u, 10009797983743892749823789237489237489237489u
print a(4)
print a(5)
print a(6)

3. Operators for bigintegers on one dimension arrays with square parenthesis.
bigInteger a[0] '=9080938409240924832849028380239489023u
a[3]=1U,23242U,3
a[4]++
a[2]+=1000000u
for i=0 to len(a)-1
	print a[i], type$(a[i])
next
? "len=";LEN(a)
exit

4. Operators for bigintegers on two dimension arrays with square parenthesis. First dimension is object with pointer to second dimension, which is a one dimension type of bigintegers. Each array can be extended as we write new items. Here we start with one at 0 index. So the total items aren't in a flat memory block (like the array with parenthesis).

Biginteger a[2][0] '=9080938409240924832849028380239489023u
a[1][3]=5u,23242U,3u
a[1][3]*=10000000000u
for i=0 to len(a[1])-1
	print a[1][i], type$(a[1][i])
next
? LEN(a[1])

5. Three dimensions (we can make more dimensions):
function three(z, b, c, v) {
	function arr(b, c, v){
		Biginteger a[b][c]=v
		=a
	}
	object r[z]
	for i=0 to z
		r[i]=arr(b, c, v)
	next
	=r
}
kk=234234u
a=three(3,2,4, 11212u)
for j=0 to len(a)-1
	for k=0 to len(a[j])-1
		for l=0 to len(a[j][k])-1
			a[j][k][l]+=kk
			kk+=234344u
			a[j][k][l]--
			? (j,k,l)#str$(), a[j][k][l], type$(a[j][k][l])
		next
	next
next

6. Read and Let now works for fields of structures:
structure alfa {
	x as double
	y as double
}
alfa z
let z|x = 100, z|y =200
print z|x, z|y
push 500, 300
read z|x, z|y
print z|x=300, z|y=500

7. Fix the Print statement when we print to file and use the $() internal function for setting rounding to numbers (per layer, not per file). Now we can put the $() inside the Print to channel (Print #), more than one time. If we place a name for file we get the file. Without name the f is -2, the channel for the current layer. 

locale 1033
open "" for wide output as #f
print #f, $(" 0.000e-00"), 100000000000000000000,2,3,$(" 0.000"),4,5,6
close #f
print $("")


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement Settings to change font/language/colors and size of console letters.

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

                                                             