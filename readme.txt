M2000 Interpreter and Environment
Version 15 Revision 30

Athens, August 25, 2026

1. FIX #HAVE(), #NOTHAVE() AND #POS() FOR ARRAYS AND LISTS/QUEUES (WHICH ARE ARRAYS WITH HASH TABLE FOR KEYS)
	' QUEUE CAN GET SAME KEYS
	' KEYS ARE VALUES IF WE DIDN'T PROVIDE VALUES
	A=QUEUE:="2":=100,"3","3","4","3","1","1":="VALUE","2"
	' USING APPEND WE CAN APPEND MORE SAME KEYS OR KEY/VALUES
	APPEND A,"2":=500,"3":=1000
	' SERIAL SEARCH  - THIS FIXED
	PRINT A#POS(5->"2","3")=8
	FINDTHIS(A, "2")
	
	' SORTING QUEUE LIST KEEP THE ORDER OF SAME KEYS.
	' USE INSERTION SORT ' WE CAN CHANGE IT ALTERING PROPERTY STABLE
	' LIST USE QUICKSORT
	B=A=>COPY() ' COPY OF QUEUE
	PRINT B=>ISQUEUE = TRUE
	PRINT B=>STABLE = TRUE
	PRINT COPY.ARR(B!)#STR$(",") 'KEYS
	PRINT COPY.ARR(B)#STR$(",") 'VALUES
	
	SORT A AS NUMBER
	PRINT "A STABLE SORT"
	PRINT COPY.ARR(A!)#STR$(",")
	FINDTHIS(A, "2")
	B=>STABLE=FALSE
	SORT B AS NUMBER
	PRINT "B NOT STABLE SORT"
	PRINT COPY.ARR(B!)#STR$(",")
	FINDTHIS(B, "2")
	
	
	
	SUB FINDTHIS(A AS QUEUE, S AS STRING)
		' USING HASH FUNCTION
		IF EXIST(A, S) THEN
			PRINT "SEARCH FOR "+S
			' USING HASH FUNCTION
			' SAME KEYS ARE IN SAME LINKED LIST	
			LOCAL TIMES=EXIST(A, S, 0), I
			FOR I=1 TO TIMES
				' VERY FAST HASH FOR FIRST ITEM AND
				' SERIAL SEARCH FOR KEYS WITH SAME HASH
				IF EXIST(A, S, I) THEN
					PRINT "POSITION:";EVAL(A!), "VALUE:";EVAL(A)
				END IF 
			NEXT
		END IF
	END SUB
2. FIX ASM2 EXAMPLE IN INFO FILE (USING GEMINI AI)
2.1 A FAULT USE OF STACK CORRUPTED IT. NOW IS OK.
2.2 NOW THE EXE PROGRAM CLOSED NICE, BECAUSE PROCESS THE DESTROY MESSAGE IN WINDOWPROC.

3. FIX THE M2000.DLL SO WE CAN OPEN THE M2000 INTERPRETER WITH MINIMUM CODE:
' THIS IS FOR VB6 BUT WORK WITH OTHER LANGUAGES TOO
' NEED A REFERENCE TO M2000.DLL
Sub Main()
    Dim m As New M2000.callback
    m.Reset
    m.ShowGui = True
    m.Show  ' WE CAN USE LOAD {OURPROGRAM}, (THISKEY)  TO DECRYPT AT LOADING.
    m.Run "cls 5,0:pen 14:form 80,32:dir appdir$:load {info}", False
    m.Cli "" ' SO NOW OPEN THE IMMEDIATE MODE - NOT NEED FOR A PROGRAM
    m.Hide
    m.ShowGui = False
    m.Shutdown 0
End Sub

USING M2000.EXE IS WAY BETTER, BUT THIS TINY CODE TO THE JOB FOR 98% OF PROGRAMS.



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