M2000 Interpreter and Environment
Version 14 Revision 52

1) Remove a bug in Eval() - Rat module now run in info.gsb
2) Stock statement found not workig with tuple array (tuple object), only with mArray. Fixed now for tuple also.
PRINT "Part One Using an mArray object"
DIM A(7)  ' this is an mArray object
PRINT Type$(A())="mArray"
A(0)=1,2,3,4,5,6,7
A=A()
SamePart()
clear
PRINT "Part Two Using an mArray object"
A=(1,2,3,4,5,6,7)
PRINT Type$(A)="tuple"
LINK A TO A()
Try ok {
	SamePart()
}
IF ERROR THEN
	PRINT Error$=" Type mismatch" ' this is message from VB6
	PRINT "Version>12 and (Version<14 or ( Version=14 and Revision<52)):";
	PRINT Version>12 and (Version<14 or ( Version=14 and Revision<52))
END IF
SUB SamePart()
	PRINT A(0), " Type of A():";Type$(A())
	LOCAL A1, B1, C1
	' Stock transfer values to variables
	STOCK A(0) OUT A1, B1, C1
	PRINT A1, B1, C1
	A1++
	B1+=10
	C1*=100
	' Stock transfer values from variables
	STOCK A(0) IN A1, B1, C1
	PRINT A
	' Stock copy items to same or other array.
	STOCK A(0)KEEP 2, A(2)
	PRINT A
END SUB

3) Fix the mArray inside tuple or mArray always copied if we get a copy of mArray or tuple.
A=(1,2,3)
PRINT A#START((4,5,6)) ' 4,5,6,1,2,3
PRINT A#END((4,5,6)) ' 1,2,3,4,5,6
PRINT A#END((4,5,6))#START((0,)) ' 0,1,2,3,4,5,6
CLEAR
DIM A(3)
A(0)=1,2,3
POINTERA=A()
T=(1,2,A())
T=T#END(T) ' WORKS ALSO THIS T=T#START(T)
A(2)+=100
PRINT T#VAL(2)#VAL(2)=3 ' WAS 103
PRINT T#VAL(2+3)#VAL(2)=103
T=T#END(T) ' WORKS ALSO THIS T=T#START(T)

A(2)+=100
PRINT T#VAL(2)#VAL(2)=3 ' WAS 203
PRINT T#VAL(2+3)#VAL(2)=103 ' WAS 203
PRINT T#VAL(2+6)#VAL(2)=3 ' WAS 203
PRINT T#VAL(2+9)#VAL(2)=203 ' ONLY THE LAST IS THE SAME AS A()
CLEAR ' CLEAR ALL VARIABLES
DIM A(3)
A(0)=1,2,3
POINTERA=A() ' POINTER IS AN MHANDLER OBJECT WHICH HAVE THE ACTUAL POINTER TO A()
T=(1,2,POINTERA) ' USING POINTER NOW
T=T#START(T) ' WORKS ALSO THIS T=T#END(T)
A(2)+=100
PRINT T#VAL(2)#VAL(2)=103
PRINT T#VAL(2+3)#VAL(2)=103
T=T#START(T) ' WORKS ALSO THIS T=T#END(T)
A(2)+=100
PRINT T#VAL(2)#VAL(2)=203
PRINT T#VAL(2+3)#VAL(2)=203
PRINT T#VAL(2+6)#VAL(2)=203
PRINT T#VAL(2+9)#VAL(2)=203

4) Function Array() now works for tuple without second parameter, we get a copy
	4.1 Copy a tuple:
		A=(1,2,3,4)
		B=ARRAY(A)
		PRINT B IS A  ' false
		PRINT A ' 1,2,3,4
		PRINT B ' 1,2,3,4
  4.2 Access items directly (is as before, no change)
		A=("alfa", 1200, (12,3i), 1234567890123456789012345678901234567890u)
		PRINT ARRAY(A,0)="alfa" ' string
		PRINT ARRAY(A,1)=1200 ' double
		PRINT ARRAY(A,2)=(12,3i) ' complex
		PRINT ARRAY(A,3)=1234567890123456789012345678901234567890u ' BigInteger  
	4.3 Access Iterator object (is an mHandler with a property UseIterator to True)
    A=("alfa", 1200, (12,3i), 1234567890123456789012345678901234567890u)
    M=EACH(a, -1, 1) 'reverse from last to first
    PRINT "INDEX", "VALUE"
    WHILE M ' ITERATE FROM 3 TO 0
     	PRINT "(";M^;") ";ARRAY(M)
    END WHILE
    
  So the 4.1 works because UseIterator property is false and has no second parameter.

George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
THEN press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 