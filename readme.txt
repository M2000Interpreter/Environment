M2000 Interpreter and Environment
Not Finish: Version 15 Revision 27

August 23.
Not finish 27
1. I do some changes so when I press Shift F1 after an error stop e program automatic open the fault module at the point of error.
I have a little job to finish this (Works 95% of situations)

2. Upgrade Control Form with Step Over/Step In/Step out (for one level for this revision). Also Control Form Title change background color to dark gray when form missing focus. Also, pressing ctrl F1 or Shift F1 a sheet of functions keys displayed in help form, Using StepOver F9 in a module then F5 step on that module statements and not going in functions or modules which called from that module. Using StepIn F10 you can go in and pressing F11 (Step Out) quick executed code and stop and wait for f5 returning to module which we press the F9 (step Over).  I am thinking about make this functionallity with more than one level..

3. Fix a problem with arrays when we feed them from stack of values, from tuples
' first using poiners to arrays (with no parenthesis) - no problem on old revisions
push (1,2)
read z
m=z  ' point to z
push (3,4)
read z ' z point to (3,4)
print m ' m point to (1,2)
print m#str$()="1 2"

' second we get a pointer from array. -  has problem on old revisions
clear ' erase variables for this module
dim z() 
push (1,2)
read z() 
print type(z())="mArray"
m=z() ' m points to (1,2)
push (3,4)
read z() ' z() get a new object
print type(z()) = "mArray"
' here was the fault: m points to new z() value
print m ' show 1 2 m points to old one - this was by design and restored
z(0)+=100  ' fault show 103 2 on m
print m ' right value show 1 2 m, points to old one - this was by design and restored

print m#str$()="1 2"  ' this show 
' third - we make a link (a reference) to an array()  - no problem on old revisions 
clear '
dim z() 
push (1,2)
read z() 
link z() to m  ' now m is z()
push (3,4)
read z() ' z() get a new object
print m  ' now m show 3, 4
print type(z()) = "mArray"
print m#str$()="3 4"
z(0)+=100  '  show 103 2
print m#str$()="103 4"


No binaries yet fot Revision 27



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