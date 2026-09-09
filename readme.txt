
eM2000 Interpreter and Environment
Version 15 Revision 43

1. Update cZipArchive
We use cZipArchive through ZipTool (which return buffers)
We can use directly cZipArchive

' Compressor (ZipTool) drive cZipArchive
declare WithEvents simpleZip Compressor
//	declare simpleZip Compressor
//	simpleZip=GetObject("","m2000.ZipTool")
declare WithEvents Zip "m2000.cZipArchive"
//	declare Zip "m2000.cZipArchive"  ' without events
//	Zip=GetObject("","m2000.cZipArchive") ' without events
Print Type(SimpleZip)="ZipTool"
Print Type(Zip)="cZipArchive"
Print Zip=>SemVersion="0.3.2"
Print Zip=>ThunkBuildDate="12.1.2018 17:15:52"
// TotalSize event write by me
function Zip_TotalSize() {
	Print "Stack of values:"
	stack	
	read new &zipsize
	print zipsize
}
Zip=>AddFile Dir$ + "info.gsb"
Zip=>CompressArchive Dir$ + "test.zip"
print filelen("test.zip")
function simpleZip_TotalSize() {
	Print "Stack of values:"
	stack	
	read new zipsize
	print zipsize
}
simpleZip=>AddFile Dir$ + "info.gsb"
simpleZip=>CreateZipFile Dir$ + "test1.zip"
print filelen("test1.zip")


2. Update functions for structures.
Now if we get an error we get right message and using shift+F1 open the specific source with the cursor at the error point.
Try this code (Write Edit A then Paste this code Press Esc Write A and Press enter)
structure alfa {
	x as double
	y as double
	function inc {
		alfa|x++
		z = 1/0 ' division by zero
		alfa|y++
		=alfa
	}
}
alfa something[20]#inc()


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
https://github.com/M2000Interpreter/Environment/releases/download/version15revision39/GreekManualVersion15.pdf

Greek Book About OOP in M2000
https://github.com/M2000Interpreter/Environment/releases/download/version14revision51/OOP_M2000_2026.pdf

http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 