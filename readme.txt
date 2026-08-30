M2000 Interpreter and Environment
Version 15 Revision 37

Athens, August 31, 2026

1) Added EXPORT for PNG in IMAGE statement
The old way was to make the PNG through the IMAGE() function. See example.

REFRESH 2000
CLS
SMOOTH ON
GRADIENT 1, 5
MOVE SCALE.X/2, SCALE.Y/2
R=MIN.DATA(SCALE.X, SCALE.Y)/3
CIRCLE R
R*=1.2
WIDTH 2, 4 {
	STEP -R
	DRAW R*2
	STEP -R, -R
	DRAW 0, R*2
	STEP 0, -R
}
STEP -R,-R
VAR A : COPY R*2, R*2 TO A
IMAGE A EXPORT "TARGET.PNG"
IMAGE A EXPORT "TARGET.BMP"
IMAGE A EXPORT "TARGET.JPG", 100 ' 80% quality
MOVE 0,0
IMAGE "TARGET.PNG"
PRINT TYPE(A)="String", LEN(A)*2 ' BYTES
PRINT "file length TARGET.BMP : ";FILELEN("TARGET.BMP")
PRINT "file length TARGET.PNG : ";FILELEN("TARGET.PNG")
PRINT "file length TARGET.JPG : ";FILELEN("TARGET.JPG")
' MAKE FILES INSIDE
MEMFILE=IMAGE(IMAGE(A) AS PNG)
PRINT LEN(MEMFILE)=FILELEN("TARGET.PNG")
MEMFILE=IMAGE(IMAGE(A) AS JPG 100)
' ITS NOT THE SAME ALGORITHM (IMAGE USE GDI+, EXPORT USE INNER VB6 JPG ENCODER)
PRINT LEN(MEMFILE)<>FILELEN("TARGET.JPG")
MOVE SCALE.X-R*2/3, 0
IMAGE MEMFILE, R*2/3
MOVE SCALE.X-R/3, R
IMAGE MEMFILE, R*2/3,,45 ' DEGREE USE HOTSPOT AT CENTER OF IMAGE.
DRAWING {
	MOVE 0,0
	IMAGE MEMFILE
	PEN 15
	PRINT "THIS IS AN IMAGE"	
} AS ALFA ' EMF FILE
MOVE SCALE.X/2, SCALE.Y/2
IMAGE ALFA, R,,45 
REFRESH 25

2) Update Help: see: Help SYMBOL_Definition

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