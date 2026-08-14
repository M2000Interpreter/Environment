M2000 Interpreter and Environment
Version 15 Revision 24

August 15, 2026,
Upgrade the of the buffer object.
1. Now I made some properties for reading/writing to addresses various types using these properties. This bypass the structure or and basic item size, so it is more difficult for a niewbe to handle it properly. So it is not recommented, but has the power to do some extra things...

2. We can import data using a byte array. We can define the starting index, the number of bytes and the final offset. This is done using a new method as declared in VB6 code (of M2000)):
Sub ImportFromByteArray(a() As Byte, Optional ByVal index As Long = 0, Optional ByVal items As Long = -1, Optional ByVal Offset As Long = 0)

DIM a(1 TO 30) AS BYTE
FOR I=1 TO 30:a(I)=I:  NEXT
BUFFER CLEAR b AS BYTE*10
b=>importfrombytearray a(), 26, 3, 2
FOR I=0 TO LEN(b)-1: PRINT b[I],:NEXT
PRINT
b=>fillbyte 0
b=>importfrombytearray a(), , , 5
FOR I=0 TO LEN(b)-1: PRINT b[I],:NEXT
PRINT

3. Now passing parameters no need always comma. here is a situation when we want to pass two lists (literals). Comma can't be used, so before this revision we have to place the first one in parenthesis. Now this not used any more.

module TestThis {
	// these are optional
	// we use it if we want this code to be used
	// from a bigger program which may have
	// same names as global modules/functions
	
	static sub delta
	static function gamma
		
	module alfa (a as list, b as list) {
		print a
		print b
	}
	alfa list:=1,2,3 list:=4,5,6
	function beta (a as list, b as list) {
		print a
		print b
	}
	call beta(list:=1,2,3 list:=4,5,6)
	thisscope=100
	a=gamma(list:=1,2,3 list:=4,5,6)
	a=@gamma(list:=1,2,3 list:=4,5,6) ' old wat
	delta list:=1,2,3 list:=4,5,6 
	delta(list:=1,2,3 list:=4,5,6) ' old way 
	// these are static members - and are in the same scope as
	// the code here
	// Except Modules/Functions with { } which they have own namespace.
	function gamma(a as list, b as list)
		print a
		print b
		print thisscope
		=0
	end function
	sub delta(a as list, b as list)
		print a
		print b
		print thisscope
	end sub
}
TestThis

4. The changes on the passing of parameters give more speed to M2000 Interpreter about 3 to 5%. Also now we can use named arguments from any type of call:

module alfa (a, b=100, c="Hello") {
	Print a, b, c, Type(b)
}
alfa %c="Good", %a=500, %b=10  ' 500 10Good Double
alfa 100, %c="M200"  ' 100 100M2000 Double
alfa 30 ' 30 100Hello Double
// we can pass constant values:
alfa b:=20, c:="Hi", 2000 ' 2000, 20Hi Constant
// without using comma (we don't have unary operators to broke the call)
alfa b:=20 c:="Hi" 2000 ' 2000, 20Hi Constant
// using Call we can get recursion (without it modules can't call itself)
call alfa %c="Good", %a=500, %b=10  ' 500 10Good Double
call alfa 100, %c="M200"  ' 100 100M2000 Double
call alfa 30 ' 30 100Hello Double
// we can pass constant values:
call alfa b:=20, c:="Hi", 2000 ' 2000, 20Hi Constant
// without using comma (we don't have unary operators to broke the call)
call alfa b:=20 c:="Hi" 2000 ' 2000, 20Hi Constant

// using call local, we place the caller's namespase
// so these are like called like subs, although the are not the same
// subs use different return stack and are lighter than modules
// call local need to use New clause to make it local
// and then we have to use Local (like in subs) 
module alfa (new a, b=100, c="Hello") {
	Print a, b, c, Type(b)
}
call local alfa %c="Good", %a=500, %b=10  ' 500 10Good Double
call local alfa 100, %c="M200"  ' 100 100M2000 Double
call local alfa 30 ' 30 100Hello Double
// we can pass constant values:
call local alfa b:=20, c:="Hi", 2000 ' 2000, 20Hi Constant
// without using comma (we don't have unary operators to broke the call)
call local alfa b:=20 c:="Hi" 2000 ' 2000, 20Hi Constant

5. A new definition called Symbol.

Module ExampleSymbol1 {
	FLUSH ' empty stack of values
	' test variables. We want these unchanged
	VAR SECOND="OK...", X=12345, Y=12345, A=100
	' we have two subs called without parenthesis
	STATIC SUB DRAWPLAYER, DRAW
	
	' these are the Symbols:
	SYMBOL TO, ANGLE, COLOR, PIXEL, TWIPS
	
	' using %name=value we place a named parameter
	' using name:=value we place a constant named parameter
	
	' USING SYMBOLS AS QUALIFIERS
	DRAWPLAYER PIXEL %Y=200, %X=100
	DRAWPLAYER TWIPS 200, 300
	DRAWPLAYER 200, 300
	DRAWPLAYER %Y=500
	
	' VARIABLE NUMBER OF ARGUMENTS PER SYMBOL
	DRAW Y:=5000, X:=1000
	DRAW ANGLE 30, 100, 200 COLOR 14
	DRAW 100, 200
	DRAW 100, 200 COLOR #FF4422
	DRAW
	' check the remaining variables:
	LIST
	
	' we use A! to get a symboll or an empty symbol, if no symbol found.
	
	SUB DRAWPLAYER(A!, X=0, Y=0)	
		IF A=PIXEL THEN
			PRINT "DRAWPLAYER PIXEL "+X+", "+Y
		ELSE.IF A=TWIPS OR A="" THEN
			PRINT "DRAWPLAYER TWIPS "+X+", "+Y 
		END IF
	END SUB
	
	' this is more advanced
	' the Read Local read also the named arguments which we didn't process
	' in the first place (see we check only WHAT at the begining)
	
	SUB DRAW(WHAT!)
		SELECT CASE WHAT
		CASE ANGLE
			READ LOCAL ANG, X, Y
			PRINT "DRAW ANGLE "+ANG+", "+X+", "+Y
		CASE TO
			READ LOCAL X, Y			
			PRINT "DRAW TO "+X+", "+Y
		CASE ELSE
				IF EMPTY THEN ? "MISSING DRAW PARAMETERS": EXIT SUB
				READ LOCAL X, Y
				PRINT "DRAW "+X+", "+Y	
		END SELECT
		READ LOCAL SECOND!
		' Number pop a number from stack of values
		' HTML COLOR IS RGB HEXVALUE, ACTUAL IS AN BGR NUNBER
		' SO #FF4422 DDBB00
		LOCAL COL
		IF SECOND=COLOR THEN
		READ COL
		IF COL<0 THEN COL=UINT(COL-1)
			HEX @(10),"COLOR=";COL
		END IF
	END SUB
}
ExampleSymbol1
MODULE ExampleSymbol2 {
	SYMBOL TO -> DRAW, ANGLE -> DRAW
	SYMBOL COLOR  ' SYMBOL MAY HAVE A TYPE VALUE
	PRINT TO|VALUE="DRAW", TO|CLASS="TO", TYPE(TO)="DRAW"
	PRINT COLOR|VALUE="COLOR", COLOR|CLASS="COLOR", TYPE(TO)="COLOR"
	
	' modules have own name space.
	' we didn't use Local or Read Local
	MODULE DRAW (A!){
		MODULE CHECK(T, A!) {
			PUSH A=T
		}
		IF TYPE(A)="DRAW" THEN
			SELECT CASE A
			CASE ANGLE
				READ ANG, X, Y
				PRINT "DRAW ANGLE "+ANG+", "+X+", "+Y, TYPE$(ANG)
			CASE TO
				READ X, Y			
				PRINT "DRAW TO "+X+", "+Y
			CASE ELSE
				PRINT "MISSING DRAW PARAMETERS"
			END SELECT
		ELSE.IF EMPTY THEN
			PRINT "NOTHING TO DRAW"
		ELSE
			READ X, Y
			PRINT "DRAW "+X+", "+Y 
		END IF
		CHECK COLOR
		IF NUMBER THEN PRINT @(10),"COLOR=";NUMBER
	}
	FLUSH
	DRAW ANGLE 30, 100, 200 COLOR 14
	DRAW TO 100, 200
	DRAW 100, 200 COLOR #FF4422
	DRAW
	' PASSING NAMED ARGUMENTS AS CONSTANTS
	DRAW ANGLE X:=100, ANG:=30, Y:=200 COLOR 14
	' PASSING NAMED ARGUMENTS (AS IS)
	DRAW TO %Y=200, %X=100 COLOR #FF00FF
	' PASSING NAMED ARGUMENTS AS CONSTANTS
	DRAW Y:=200, X:=100 COLOR #FF4422
	DRAW
}

6. Just some calculations for the value of an Html Color
M2000 use negative number for HTML colors, 0 to 15 for windows color adn &8000_00XX for gui windows colors.
HTMLCOLOR_IN_M2000=#FF0000
PRINT "M2000 RGB VALUE (DECIMAL):";HTMLCOLOR_IN_M2000
PRINT "M2000 RGB VALUE (255,0,0)(DECIMAL):";COLOR(255,0,0)
A$=HEX$(BINARY.NOT(UINT(HTMLCOLOR_IN_M2000-1)),3)
PRINT "REAL BGR VALUE (HEX):";A$
RGBVALUE=VAL("0X"+RIGHT$(A$,2)+MID$(A$,3,2)+LEFT$(A$,2))
PRINT "REAL RGB VALUE (AS DECIMAL):";RGBVALUE
PRINT "HTML COLOR: #"+HEX$(RGBVALUE, 3)

HTMLCOLOR_IN_M2000=#22AAFF
PRINT "M2000 RGB VALUE (DECIMAL):";HTMLCOLOR_IN_M2000
PRINT "M2000 RGB VALUE (0x22,0xAA,0xFF)(DECIMAL):";COLOR(0x22,0xAA,0xFF)
A$=HEX$(BINARY.NOT(UINT(HTMLCOLOR_IN_M2000-1)),3)
PRINT "REAL BGR VALUE (HEX):";A$
RGBVALUE=VAL("0X"+RIGHT$(A$,2)+MID$(A$,3,2)+LEFT$(A$,2))
PRINT "REAL RGB VALUE (AS DECIMAL):";RGBVALUE
PRINT "HTML COLOR: #"+HEX$(RGBVALUE, 3)

7. Fix a fault in Select Case when we leace an empty line before End Select.
8. Fix the Exit Sub through a Case from a Select Case.
MODULE TestMe {
	B(1)
	PRINT "OK"
	B(2)
	PRINT "OK TOO"
	
	SUB B(N)
		LOCAL OK
		SELECT CASE N
		CASE 1
			OK=TRUE
		CASE 2
			EXIT SUB
		END SELECT
		IF OK THEN EXIT SUB
		PRINT "FINALE"
	END SUB
}
TestMe


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