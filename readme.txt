M2000 Interpreter and Environment
Version 13 revision 36 active-X
1. Fix a problem when we pass an enumerator value and we wish to hold the value non the object, and the parameter has non enumerator type and has an optional value.  So now this work:

MODULE tst {
	MODULE SOMETHING (VAL1, VAL2=0&, VAL3="OK") {
		LIST ' show variables
		? VAL2=10, TYPE$(VAL2)="Long"	
	}
	ENUM ALFA {X=10}
	'call module SOMETHING
	SOMETHING 0, X
	'call sub SOMETHING()
	SOMETHING(0, X)
	
	SUB SOMETHING(VAL1, VAL2=0&, VAL3="OK")
		LIST ' show variables
		? VAL2=10, TYPE$(VAL2)="Long"
	END SUB
}
tst

2. Non close lines for curves, circles and polygons using last character the semicolon.
SMOOTH ON  // WORKS WITH SMOOTH OFF
MODULE ALFA (N) {
	MOVE 5000, 8000
	PEN ,N{  ' N IS OPACITY (255-100%, 0 -0%)
		WIDTH 10 {CIRCLE FILL 10, 3000,2/1,, PI/2, PI;}
		WIDTH 10{COLOR 10{CURVE 0,2000,1000,-1000,3000,-3000;}}
		WIDTH 10{CURVE 0,2000,1000,-1000,3000,-3000;}
		WIDTH 10 {POLYGON 1, 1000,0,0,1000,-1000,0;}
		WIDTH 10 {POLYGON 1, 1000,0,0,1000,-1000,0}
		WIDTH 10 {POLYGON 1, 1000,0,0,1000,-1000,0;}
		COLOR 0,1{WIDTH 10 {POLYGON 1, 1000,0,0,1000,-1000,0;}}
		WIDTH 10 {POLYGON 1, 1000,0,0,1000,-1000,0;}
				
	}
	WIDTH 10 {POLYGON 1, 1000,0,0,1000,-1000,0, 0,-1000}
}
CLS, 0
DRAWING {
	PEN , 100 {GRADIENT 1, 2}
	ALFA 120
} AS SOMETHING
MOVE SCALE.X DIV 2, SCALE.Y DIV 2
ΕΙΚΟΝΑ SOMETHING,SCALE.X DIV 2,,30
ALFA 100


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (534 tasks)
                 