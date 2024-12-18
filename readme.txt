M2000 Interpreter and Environment

Version 12 Revision 55 active-X
1)fix #val$() to return decimal based on locale, if item is numeric
#val() to string work ok (we have to place parenthesis in string expression)
locale 1033
a=(1.333,)
? a#val$(0), "["+(a#val(0))+"]"
locale 1032
? a#val$(0), "["+(a#val(0))+"]"

2)fix lambda function as final member after a copy. Here we get a copy from class to m.
group alfa {
	final k=lambda ->500
}
m=alfa
try {? m.k()}  ' now print 500 before raise error k() not found.

3)fix copy of long long value after copy of group.
group alfa {
	long long k=12345
	k1=12345&&  ' it is a long long too
}
m=alfa
? type$(m.k), m.k=12345
? type$(m.k1), m.k1=12345

4) Fix reading long long value
structure cc {
	d as integer
	c as long long  ' usigned long long (8 bytes)
}
buffer clear mem as cc*100
a=0xFFFF_FFFF_FFFF_FFFF
' unsigned long long are Decimal types
Print type$(a)="Decimal"
' but we fit it on 8 bytes (the first 64bit)
return mem, 10!c:=0xFFFF_FFFF_FFFF_FFFF
z=eval(mem, 10!c)
Print type$(z)="Decimal", z=18446744073709551615@
k=sint(z)  'same bits as long long (signed) )is the number -1
Print type$(k)="Long Long", k=-1  ' so k is long long signed

5. Underscore for all numbers (except Html Colors):
? 0x6000_0000_0000_0000&& ' long long 64bit
? 9_223_372_036_854_775_807&& ' max long long
? 0x0000_0000& ' long 32 bit
? 2_147_483_647& ' max long
? 0x0000%  ' integer 16bit
? 32_767% ' max integer
? 0x00ub ' byte 8bit unsigned only (0 to 255)
? 2_55ub ' max byte
? 0x8000_0000_0000_0000   ' decimal type (hold values as unsigned long long)
? 18_446_744_073_709_551_615@ ' max usnigned long long (decimal type)
? 123_456_789.876_543_210@   ' decimal type 27 digits
? #FFFFFF, #FFFF00=-65535  ' HTML COLORS - NO underscore
? 46_000ud=date("9/12/2025")
Def t(x)=type$(x)
? t(12_456.123_4#)="Currency"
? t(1_012.232e-10~)="Single"

6. Fix syntax color for 0XFF.. and &HFF..



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

                                                             