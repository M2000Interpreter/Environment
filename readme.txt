M2000 Interpreter and Environment
Version 14 Revision 47

Two faults fixed - was difficult to found because altering fontname may not produced becasue of different font metrics:

1. Print type 4, for printing on output layer, use proportional text. If text need more space than the column have we get more columns. The problem was when column advance to same position as the width of screen (text coordinates). An exit from a sub without altering a by reference value do the fault, leaving 0 - so we get same line second time and overwrite it. This is the test program:

Font "Courier New"
Form 120, 32
? $(4),  ' use proportional text
a=Random(!12345)
c$=string$("0", 124)
for i=1 to 100
	print i+string$("*", Random(80, 130)),
	aa=key$ ' press any key to continue
next
print c$,
print c$,
Print

2. Print type 0 (the default), for printing on output layer using non proportional text. If text need more space than the column have we  get more columns and we get columns from other lines too (in no free line exist we get a free line scrolling up the "scrolling part" of screen). Here the problem was for an internal ruller which stayed at position width, at newer prints, so no print happen until the program execute a new line and reset the position to 0 (or any other statement which move the cursor). Older revisions stopped after 26.

Flush
Font "Courier New"
Form 120, 32
Print $(0),  ' this is the default one
a=Random(!12345)
c$=string$("0", 124)
Z=FALSE
for i=1 to 100 ' fault after 26 for older revisions
	IF POS>119 THEN Z=TRUE
	IF Z THEN PUSH I, POS
	print i+string$("*", Random(10, 300)+118*1.5)+"-",
	aaa$=key$	' press a key to print next
	if not empty then exit for
next
print c$,
Print c$,
Print
Stack ' show stack values
Flush ' delete stack values





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