M2000 Interpreter and Environment
Version 14 revision 9 active-X

1) BASIC statement to make the current module/function to use BASIC specific style:
1.1 DATA, READ and RESTORE as of BBC BASIC
	WE CAN USE STANDARD READ USING @READ
FROM PAGE 122 OF BBC MICRO MANUAL (LINE 0 ADDED FOR M2000)
  0 BASIC
  5 REPEAT
 10 PRINT "GIVE THE MONTH AS A NUMBER"
 20   INPUT M
 30   UNTIL M>0 AND M<13
 40 FOR X=1 TO M
 50   READ A$
 60   NEXT X
 70 PRINT "THE MONTH IS ";A$
100 DATA JANUARY, FEBRUARY, MARCH, APRIL
110 DATA MAY, JUNE, JULY, AUGUST, SEPTEMBER
120 DATA OCTOMBER, NOVEMBER, DECEMBER	


DATA may have a list of literals (1, 3&, 2~ and all the symbols)

2) Advanced OpenGl support.

3) Exclude a user control for server. We can use M2000 code with objects to do the same job.

4)Now we can make properties for objects which can shadow names from original read only variables (we can continue read these by using the @ before the name for current module/function, or we can use it as is from other modules/functions (so the version in the example stay as the original in a call to an inner module). The idea is to make modules/functions without conflicts with M2000 reserved names.

? version=14
version=100
? version=100, @version=14
Module Inner {
	Print version=14
}
inner

5) Some minor tunings...for speed.

1.2 DIMENSION OF ARRRAYS PLUS ONE (A(4) HAS 5 ITEMS)
1.3 FOR NEXT SAME AS BASIC. SO NOW WE CAN SKIP THE FOR NEXT IF WE HAVE NO LOOP BASED ON STEP AND START END END VALUES.
	FOR I=10 TO 1
	' SKIP THE STRUCTURE BECAUSE I>1
	' NUT BY DEFAULT M2000'S "FOR" GOES FROM 10 TO 1
	NEXT I


 
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

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 