M2000 Interpreter and Environment
Version 15 Revision 21

August 10, 2026,

1. Restore compatibility  for the array() function.

The issue start from revision 52 version 14, on Array() funtion.
So an new function born, the Copy.Arr() for copies of array types including lists (which now have a secondary interface common to secondary interface of mArray and tuple). Array([]) return array from a current stack object, placing an empty one, and Array(s)  when s is a stack object works. 



this was the test code:
A=(1,2,3,4)  
B=ARRAY(A)  ' this now is the same as Array(A, 0) return the first item or raise error if no item exist
PRINT B IS A  ' false
PRINT A ' 1,2,3,4
PRINT B ' 1,2,3,4

now change to:
A=(1,2,3,4)
B=COPY.ARR(A)
PRINT B IS A  ' false
PRINT A ' 1,2,3,4
PRINT B ' 1,2,3,4

2. I found a way to make M2000 to play without lag the KB module (keyboard for music), on my laptop. My desktop has no noticable lag, because has a real synth inside. My laptop is another story. So I found VirtualMIDISynth #1 (Google Gemini told me about it), which we can set buffer to 0ms. (has 250ms preset). So now you play music with zero lang on a cheep laptop...by setting the default midi out

So I defined the midi.out() (code to enumerate midi output from Gemini) to return either an array of zero or more strings, and when we place an numeric argument then return the string (the name of the midi output)

So to change output we use "Play to" as variant of "Play". Remember there is Play 0 to stop all scores... The play statement set an instrument on a score (and optional the stacatto per sent. by default we play legato, ewual to 100 for the paremeter of stacatto).  

m=midi.out()
if len(m)>0 then
	menu
	for i=0 to len(m)-1
		menu + midi.out(i)
	next
	print "select:";
	menu !
	if menu=0 then
		midi_number=0
		print menu$(1)
	else
		print menu$(menu)
		midi_number=menu-1	
	end if
else
	print "no midi found"
	break
end if

Play to midi_number  ' disable midi (erase all music threads), set the midi number, enable midi

or 
Play to midi_number, ... mormal parameters (also disable/ser/enable midi).



Form M2000 console load Info and do this:

Play to 0  ' VirtualMIDISynth get the 0
music_notes ' now you hear all the notes instantly.
then do again:
Play to 1
music_notes ' now you use the Microsoft synth with a lag, so you beleave that you hear the D#9 and above...notes.



3. Modules on Info: DD6, HU, RADIAL now works fine. We can get a copy of a list (copy properly the Group type of objects) using the Copy() new function directly (shallow copy, not shown here):
A=list:=1:=1000,2,3:=500,4
B=A=>Copy()
Print A, B

C=Array(A) 'same as Array(A,0)
? C=1000
D=Copy.Arr(A)
Print Type(D)="tuple"
Print D#Str$(",")="1000,2,500,4"
Keys=Copy.Arr(A!)
Print Keys#Str$(",")="1,2,3,4"
' A is a list but for Version 15 can be used like an array too
' using #functions
Print A#Str$(",")="1000,2,500,4"
Higher100=lambda (x)->x>100
' no need to copy A first...
Print A#Filter(Higher100)#sort()#str$(",")="500,1000"


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