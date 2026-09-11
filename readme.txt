M2000 Interpreter and Environment
Version 15 Revision 44

1. Upgrade Assembler.
Look Asm2 in Info file we make an exe file. Now we place Icon and a manifest (resources),


2. I found a way to take account the END IF error, from a goto (in module when we didn't use exit to exit the module). Goto to back the code block of last "running" part of an IF flow structure, checked if go inside or outside, so if go outside erase the "open" if counter. We can't do that for a jump forward because for IF THEN multiline and without using block {} we not check the "end if" mark until found it as the execution of code advance to next statements. But we can use a block { } and put inside the "if structure" with forward goto (out of the structure). Although using labels are not for every day programming, M2000 have these for using it by pupils to find out why are difficult to program with that.


 
1.1 The example bellow has a double loop inside a case of a select case and another try without select case. It works.

//	TEST CODE
	SELECT CASE 10
	CASE 10
			B=1
ONE:		
		IF B<10 THEN
			A=1
ALFA:
			IF A<10 THEN
				PRINT A,
				A++
				GOTO ALFA
			ELSE
				PRINT
				B++	
				GOTO ONE	
			END IF	
		END IF
	END SELECT
	
	B=1
TWO:
	IF B<10 THEN
		A=1
BETA:
		IF A<10 THEN
			PRINT A,
			A++
			GOTO BETA
		ELSE
			PRINT
			B++	
			GOTO TWO	
		END IF	
	END IF

1.2 This example has four parts. Part one has a forward jump and we use a block. The second part has a While End While (this has a hidden block so forward jump is ok). The last two parts are for backward jump, one inside a While End While.
 
' forward jump passing end if need a block
' or a loop which have hidden block like While and Do/Repeat
{
	if true then
		if true then
			if true then
				goto 5000
			end if
		end if
	end if
}
5000 ? "ok"
//	end ' normal exit - check END IF
//	' using EXIT no check for END IF
while true
	if true then
		if true then
			goto 5020
		end if
	end if
end while
5020 ? "ok"

a=1
goto 5040
alfa:
5040 ' no need for block
	if a=1 then
		if true then
			a=0
			? "goto alfa"
			goto alfa ' back jump, now M2000 check if is inside code block.
		end if
	end if
print "ok"
a=1
while a<>0
5050
	if a=1 then
		if true then
			a=0
			? "goto 5050"
			goto 5050
		end if
	end if
end while
print "ok"

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