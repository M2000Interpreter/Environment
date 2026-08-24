M2000 Interpreter and Environment
Version 15 Revision 29

August 24

' The Bomb Situation, or you need a lifetime or more to understand how VB6 works
' I think this error can't be found from AI

Cast =lambda (that$)->{
	Interface that, that$ {dummy}
	=lambda that (t as *that)->t
}
get_iUnknown = Cast("{00000000-0000-0000-C000-000000000046}")
' buffer object has functions to read memory everywhere;
' (checking for bad address first)
buffer inspect as long

declare form1 form
declare a type "ctxninebutton" form form1
testme = get_iUnknown(a)
? "These have different VTables"
vtable_a=inspect=>peekint32(varptr(a))
vtable_testme=inspect=>peekint32(varptr(testme))
? "Same Objects: ";testme is a 
? "Different VTables: "; vtable_a<>vtable_testme
if version<15 or (version=15 and revision<29) then "do not this - program hang": exit
? type(testme)
' why ? old one hang? Who knows..
' how overcame this problem?
' This problem was for the ExtControl class (see ExtControl.cls), the real class behind external controls.
' Type() didn't return RxtControl but go deeper and get the value property.
' This value property is the ctxninebutton (the usectxninebutton.ctl)
' So when we use Type(a) M2000 get the value of object and return ctxninebutton
' When we get the iUnkown interface, we get different Vbtable.
' That is not bad as idea, but for this control the use of value property hang the program.
' The solution was to get the iDispatch interface and then use on that the value property. 

declare form1 nothing


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