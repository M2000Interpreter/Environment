M2000 Interpreter and Environment
Version 14 Revision 41

1) Fix a hidden bug:
This example now work without error in module ALFA.
MODULE ALFA {
	TRY {
		PROBLEM("GEORGE", 12123)
	}
	PRINT ERROR$
	// PROBLEM START FROM THE READ AT SUB PARAMETER LIST
	SUB PROBLEM(THAT AS VARIANT="")
		PRINT THAT
		READ THAT  ' HERE WAS THE PROBLEM
		PRINT THAT
	END SUB
}
ALFA
MODULE ALFA1 {
	PROBLEM("GEORGE", 12123)
	SUB PROBLEM()
		' NO PROBLEM BECAUSE THIS IS THE GENERIC READ
		READ THAT AS VARIANT=""
		PRINT THAT
		READ THAT  ' HERE WAS THE PROBLEM
		PRINT THAT
	END SUB
}
ALFA1

2) Fix a problem with Continue inside then else without block - worked as restart so never go to actual block for/next 
aa=0
for i=1 to 100
	if i<3 then continue else aa++
next
print aa=98

3) Now we get true from the last print (?)
a=("aa",2,3,"hello",5)
link a to a$()
try {
	a$(0)="???"
}
? a$(0)="???"  ' true


4) When we use Ctrl+C (copy to clipboard) in EditBox and Editor (as text editors) then we get visual message that the copy to clipboard done (sometimes I press light the C so I was thing that the copy done, but that not happen, so I insert the message to take knowledge about my difficulty...Also I have a lot  this year I become 60 years old).

5) We can use STATIC SUB to inform Interpreter for subs before we use them. This it the code from INFO, TSTSUB module:

global const many=30000
form 80
print $(,10) ' 10 characters column
print "iterations per test:";many
module TestThis (UseStatic=True){
	if UseStatic then static sub TestMe,TestMe2
	aa=100
	print "TestMe - New Call without Parentheses"
	profiler
	for i=1 to many
		TestMe
	next
	print timecount
	print "TestMe() - Standard"
	profiler
	for i=1 to many
		TestMe()
	next
	print timecount
	print "TestMe() - Standard using GOSUB"
	profiler
	for i=1 to many
		gosub TestMe()
	next
	print timecount	
	print {TestMe2 "ok", 1000  - New Call without Parentheses}
	profiler
	for i=1 to many
		TestMe2 "ok", 1000
	next
	print timecount
	print {TestMe2("ok", 1000) - Standard}
	profiler
	for i=1 to many
		TestMe2("ok", 1000)	
	next
	print timecount
	print {TestMe2("ok", 1000) - Standard using GOSUB}
	profiler
	for i=1 to many
		gosub TestMe2("ok", 1000)	
	next
	print timecount
}
pen 11 {print "Using Static Sub Definition"}
TestThis
pen 11 {print "Skip Static Sub Definition"}
TestThis False
Report {
Using Static Sub Definition is faster. Also we shadowing any global module with same name.
Names like print can be used for subs (e.g. SUB print()) but to call it we have to use:
1 - print()
2- gosub print()
if we use Static Sub Print we can use Call
3 - Call Print
and (1)  and (2)

To use our Print we have to use Module Print { } as local module
This make Print as a call to local module
We can use @Print  to call the original Print in the same module where we define local module print
Inside module print statement print is the original one. 
}

sub TestMe()
	aa++
end sub
sub TestMe2(Alfa, k)
	aa+=k
end sub



6) Use of Quaternion (we can make new UDT of Quaternions - exist in M2000.dll - this type (among other udt) made M2000.dll the need to be registered). 
Module English_Quaternion {
	locale 1033
	Static Function DispQuat
	Declare Math Math
	q=Math=>NewQuatType(10, 30, -50, 30)
	print type$(q)="QuatType"
	Print q|w=10, q|x=30, q|y=-50, q|z=30
	Print Math=>QuatStringQuat(q)
	Print Math=>QuatStringQuat(q, "0.000")
	Print "q=";DispQuat(q)  ' (10+30i-50j+30k)
	Print "Norm of Quaternion |q|:=";Math=>NormQuat(q)
	Print "Conjugate Quaternion:=";DispQuat(Math=>NegQuat(q))  ' (10-30i+50j-30k)
	Function DispQuat(q as QuatType)
		' break for displayed better - M2000 handle any line length
		local ret="("+(q|w)+if$(q|x<0->"","+")+(q|x)+if$(q|y<0->"i","i+")
		=ret+(q|y)+if$(q|z<0->"j","j+")+(q|z)+"k)"
	End Function
}



George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and then open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 