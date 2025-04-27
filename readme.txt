M2000 Interpreter and Environment
Version 13 revision 43 active-X

Now we can use each() iterator for static variables of type tuple, list and stack:
'the each function has two types:using comma after object like a function, and anohter one using only space to pass START or END or number then TO then START or END or Number. For each type a negative number -1 means last item, and -2 one before the last.


Rem : clear ' this clear can reset all static variables by erase them
module inner {
	static a=(10,20,30,40,50), b=(LIST:=1,2,3,4,5), c=(STACK:=-1,-2,-3,-4,-5)
	' need () because if c exist as static interpreter need to skip expression
	' so we need to  use (  ) because the := isn't part of expression but is part of STACK
	k=each(a 1 to  3)
	while k
		print array(k),
	end while
	print
	k=each(b, -1,  -3)
	while k
		print eval(k),
	end while
	print
	k=each(c start to  -2)
	while k
		print stackitem(k),
	end while
	print
}
inner




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

https://rosettacode.org/wiki/Category:M2000_Interpreter (544 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 