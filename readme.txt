M2000 Interpreter and Environment
Version 15 Revision 23

August 11, 2026,

Another mistake removed. When I place functions like Type$() as Type() I break without knowing the Evaluator module in Info file. This file has over 500 modules, so it is hard to execute all, but I do selected checks, so I found it and fix it. The problem fixed for any array and those links to arrays who have the same name with any predefined function (for user functions was not a problem).
Using @ before the name we call the original function. I upgrade the interprerer to use static functions/subs with names like inner statements (Subs could use any name if we call it with parenthseis, but this not worked before this revision when we call the sub without parenthesis and have name like a m2000 command). See Example3 the use of Type() and Move (these exist in M2000 vocabulary)
 
Module Example1 {
	Module Inner {
		Print Type(1)="Double"
	}
	Dim Type(10)=12345
	Print Type(1)=12345
	Type(1)++
	Print Type(1)=12346
	Print @Type(1)="Double", Type$(1)="Double"
	Inner
}
Example1
Module Example2 {
	Module Inner {
		Print Type(1)="Double"
	}
	a=(12345,12345,12345)
	Link a to Type()
	Print Type(1)=12345
	Type(1)++
	Print Type(1)=12346
	Print @Type(1)="Double", Type$(1)="Double"
	Inner
}
Example2

Check this too. Now we can use an internal function name for a static function. But we can't use @name to call it, because that call the original function. We have to declare the function using Static Function Type or Type()
Also now we can do the same for SUB, for calling it without parenthesis like statements.

Module Example3 {
	Static Function Type() ' same name as Type()
	Static Sub Move  ' same name as Move
	Module Inner {
		Static Sub Move  ' same name as Move
		Print Type(1)="Double"
		' USE THE PARENT'S CODE
		' LIKE IT IS HERE
		move 100, 100 
	}
	move 100, 100
	Print Type(1)=12345
	Print @Type(1)="Double", Type$(1)="Double"
	Inner
	Function Type(x)
		=12345
	End Function
	Sub Move(X, Y) 'pixel
		' CALL THE ORIGINAL STATEMENT
		@move X div twipsX, y div twipsY
		print "move graphic cursor to "+x+","+y
	End Sub
}
Example3

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