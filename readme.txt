eM2000 Interpreter and Environment
Version 15 Revision 42

1. Now String() can be used as String$()
2. Upgrade Interpreter:
 2.1functions for structures
 2.2 We can use name of structure to define memory buffers in a class
 2.3 We can apply a series of functions to all items of a buffer when we use the struct_name buffer_name statement
 The exampe bellow has three functions for structure alfa.
In statement alfa kappa[200]#new(30,40) the value of new(30,40) applied to all items
in  print beta.kappa[3]#str()  we pass in alfa  kappa[3]  as a copy (one item only)
All structures funcrions works for one item only. 
 
structure alfa {
	x as double,
	y as double
	function mul(a) {
		alfa|x=alfa|x*a
		alfa|y=alfa|y*a
		=alfa
	}
	function new(a=100, b=200) {
		alfa|x=a
		alfa|y=b
		=alfa
	}
	function str() {
		="("+(alfa|x)+", "+(alfa|y)+")"
	}
}
alfa kappa[20]#new(30,40)#mul(3)
print kappa[3]#str()="(90, 120)", kappa=>items=20, len(kappa)=320 ' bytes
alfa kappa[200]#new(31,41)#mul(3) ' append more items, applied to that items only
print kappa[19]#str()="(90, 120)"
print kappa[20]#str()="(93, 123)", kappa=>items=200, len(kappa)=3200 ' bytes

class beta {
	structure alfa {
		x as double,
		y as double
		function mul(a) {
			alfa|x=alfa|x*a
			alfa|y=alfa|y*a
			=alfa
		}
		function new(a=100, b=200) {
			alfa|x=a
			alfa|y=b
			=alfa
		}
		function str() {
			="("+(alfa|x)+", "+(alfa|y)+")"
		}
	}
	{read many}
	alfa kappa[many]#new(30,40)#mul(3)
}
beta=beta(10)
print beta.kappa[3]#str()="(90, 120)", beta.kappa=>items=10, len(beta.kappa)=160 ' bytes



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