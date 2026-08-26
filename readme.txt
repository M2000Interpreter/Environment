M2000 Interpreter and Environment
Version 15 Revision 31

Athens, August 26, 2026

Upgrade/fix Select case using Enumeratons, by index using Select Emum EnumVar1, or by value Select Case EnumVar1.
Test program has 4 modules. There are enums with values as number and enums with value as string. Each set has same values.
Module enum1/enum2 use Select Case, works using values from enums.
Module enum3/enum4 use Select Enum works using index from enums. 

module enum1 {
	print module.name$
	enum abc {
		a=1,b=2,c=1,d=2
	}
	z=b
	do
		o=z^ ' index
		select case z
		case a to c
			print "a..c"
		case d
			print "d"
		end select
		z++
	until z^=o' is equal if z can't advanced to next index 
}
enum1
module enum2 {
	print module.name$
	enum abc {
		a="hello",b="hi",c="hello",d="hi"
	}
	z=b
	do
		o=z^
		select case z
		case a to c
			print "a..c"
		case d
			print "d"
		end select
		z++
	until z^=o
}
enum2
module enum3 {
	print module.name$
	enum abc {
		a=1,b=2,c=1,d=2
	}
	z=b
	do
		o=z^ ' index
		select enum z
		case a to c
			print "a..c"
		case d
			print "d"
		end select
		z++
	until z^=o' is equal if z can't advanced to next index 
}
enum3
module enum4 {
	print module.name$
	enum abc {
		a="hello",b="hi",c="hello",d="hi"
	}
	z=b
	do
		o=z^
		select enum z
		case a to c
			print "a..c"
		case d
			print "d"
		end select
		z++
	until z^=o
}
enum4




George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do this:

dir appdir$
load info

then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).

English old paper for M2000
https://github.com/M2000Interpreter/Environment/releases/download/ver13rev44/M2000paper.pdf

Greek Book for learning programming
https://github.com/M2000Interpreter/Environment/releases/download/version15revision10/GreekBookM2000.pdf

Greek Manual (a work in progress)
https://github.com/M2000Interpreter/Environment/releases/download/version15revision10/GreekManualVersion15_preview.pdf

Greek Book About OOP in M2000
https://github.com/M2000Interpreter/Environment/releases/download/version14revision51/OOP_M2000_2026.pdf

http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 