M2000 Interpreter and Environment
Version 14 Revision 31

Extend the error recovery without error raising but using the error value for an enum.
This now work for parameters for Modules/Functions/Subs
Also work with the Let statement and normal assign.

CONST RECOVER_ERROR=RANDOM(-1,0)
GLOBAL BAD_NUMBER=-999
IF RECOVER_ERROR THEN
	PRINT "RECOVER_ERROR = TRUE"
	GLOBAL ENUM FLAGS_B {
	    Hold, Right, Left
	ERROR:
		    ERB=BAD_NUMBER
	}
ELSE
	PRINT "RECOVER_ERROR = FALSE"
	GLOBAL ENUM FLAGS_B {Hold, Right, Left}
END IF
GLOBAL ENUM FLAGS_A {
    In=-2, Out, FLAGS_B, Up, Down, Back, Front
ERROR:
    ERA=BAD_NUMBER
}
MODULE ALFA (M AS FLAGS_B) {
    IF M=BAD_NUMBER THEN PRINT "NOT GOOD VALUE":EXIT
    PRINT "OK - I DO THE JOB"
    PRINT "INDEX:";M^
    PRINT "VALUE:";M
    PRINT "NAME:";EVAL$(M)
    PRINT "TYPE:";TYPE$(M)
}
VAR Z AS FLAGS_A=Front
TRY OK {
	ALFA Z
}
IF ERROR OR NOT OK THEN PRINT "ERROR"+ERROR$
Z=Left
PRINT "INDEX:";Z^
PRINT "VALUE:";Z
PRINT "NAME:";EVAL$(Z)
PRINT "TYPE:";TYPE$(Z)
ALFA Z ' this pass ok




George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and then open it again.

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