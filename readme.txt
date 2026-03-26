M2000 Interpreter and Environment
Version 14 revision 10 active-X

1)BigInteger issues:
1.1 str$(2u,"") now return 2 (before an empty string).
1.2 str$(1212313u, "#,###") not supporeted so we get the number (no error raised).
1.3 Problem on expression evaluator:
' problem for <10 revision. Now the two parts are equal.
var a=1000234, b=34450
?  (a+1)*b
biginteger z=a, zz=b, one=1
? z, one
?  (z+one)'*zz
?  (one+z)'*zz
? zz* (z+one)'*zz
?  zz*(one+z)'*zz
? (z+one)*zz, "???"
? (one+z)*zz
? (z-one)*zz, "???"
? (one-z)*zz, "good"
? (-z*one-one)*zz, "???-"
? (-one*z-one)*zz
? (z*one-one)*zz, "???+"
? (one*z-one)*zz
?  z*zz+one*zz
?  one*zz+z*zz
clear
var a=1000234, b=34450
? "----------long long no problem----------------"
?  (a+1)*b
long long z=a, zz=b, one=1
? z, one
?  (z+one)'*zz
?  (one+z)'*zz
? zz* (z+one)'*zz
?  zz*(one+z)'*zz
? (z+one)*zz, "???"
? (one+z)*zz
? (z-one)*zz, "???"
? (one-z)*zz, "good"
? (-z*one-one)*zz, "???-"
? (-one*z-one)*zz
? (z*one-one)*zz, "???+"
? (one*z-one)*zz
?  z*zz+one*zz
?  one*zz+z*zz



2) Advanced OpenGl support - now can load images/textures2d more than background.

3) Report statement has a hold function waiting user input a key or a mouse click when print 3/4 of output lines (except on printer and when run in a thread or an event service function). So now if the output is not visible then no hold happen. Also this works for Users Forms when the form wait for user input (only mouse) we can move the window without the holding function restart printing. Also if we minimize the form then automatic restrart printing.

// Example on form:
document a$
for i=1 to 1000
a$=str$(i,"0000")+" "+string$(chrcode$(random(65, 90)), random(10, 30))+{
}
next
a$=string$(chrcode$(random(65, 90)), random(10, 30))
declare form1 form
declare image1 Image Form Form1
with form1, "visible" as visible, "titleheight" as th, "width" as w, "height" as h
method image1, "move", 0, th, w,  h-th
once=false
function image1.click {
	if once then exit
	once=true
	layer image1 {
		report a$
		cls
	}
	after 200 {
		try {once=false}
	}
}
method form1, "show", 1
wait 500
threads erase
declare Form1 Nothing


//example printing on M2000 console's background (we hide foreground)
// try to minimize the console, (you have to do this from the context menu from the taskbar, on Windows 10 and 11 you have to press Esc when the a popup menu first pop, and then you get a small window where on the title bar you do a right mouse click to open the context menu).
document a$
for i=1 to 1000
a$=str$(i,"0000")+" "+string$(chrcode$(random(65, 90)), random(10, 30))+{
}
next
a$=string$(chrcode$(random(65, 90)), random(10, 30))
HIDE
k=10
BACK {
	report A$
	cls
	k-- : if k=0 then exit
	if not keypress(32) then loop
}
SHOW
 

//example printing on M2000 console
// try to minimize the console (you have to do this from the context menu from the taskbar).
document a$
for i=1 to 1000
a$=str$(i,"0000")+" "+string$(chrcode$(random(65, 90)), random(10, 30))+{
}
next
a$=string$(chrcode$(random(65, 90)), random(10, 30))
k=10
' just remove HIDE/SHOW and BACK from the previus example
{
	report A$
	cls
	k-- : if k=0 then exit
	if not keypress(32) then loop
}

//Also we can use one of 32 layers above console's layer:
document a$
for i=1 to 1000
a$=str$(i,"0000")+" "+string$(chrcode$(random(65, 90)), random(10, 30))+{
}
next
a$=string$(chrcode$(random(65, 90)), random(10, 30))
k=10
layer 1 {
	window 15, 10000,8000;
	show
}
Layer 1 {
	report A$
	cls
	k-- : if k=0 then exit
	if not keypress(32) then loop
}
Layer 1 {hide}

4)Fix the BASIC/DATA/RESTORE for Greek language (BASIC always in English,is the name of the language).


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