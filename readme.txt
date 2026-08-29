M2000 Interpreter and Environment
Version 15 Revision 36

Athens, August 29, 2026

Expand symbols to work decoupled from caller. In module Ver15rev36 symbols defined inside module alfa.
A name for decoupled symbol can't be a html color like this #a0ccff (this is a number), but can be this #a0ccf or this #a0ccffa
Works with subs, functions, lambda too.

Module Ver15rev35{
	symbol beta, epsilon, quality, best
	module alfa (a!,b!, v="", c!) {
		print "module alfa - Ver15 rev35"	
		if a=beta then print a
		if b=epsilon then print b
		print v
		' c has operator = only
		' for equal
		if c="" else print c, c="BEST"
		
	}
	' using external symbols as qualifiers
	alfa "a value"
	alfa beta "a value"
	alfa beta epsilon "a value" best
}
Ver15rev35
Module Ver15rev36{
	module alfa (a!,b!, v="", c!) {
		print "module alfa - Ver15 rev36"	
		symbol beta, epsilon, quality
		if a=beta then print a
		if b=epsilon then print b
		print v
		' c has operator = only
		' for equal
		if c="" else print c, c="BEST"
		
	}
	' using internal symbols as qualifiers
	alfa "a value"
	alfa #beta "a value"
	alfa#beta#epsilon "a value"#best
}
Ver15rev36


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