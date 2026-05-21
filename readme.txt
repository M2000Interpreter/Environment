M2000 Interpreter and Environment
Version 14 Revision 36

Fix Eval(), now Works for Standard types like numeric types. So in this example Eval() return number in any case. So found the pointer to group execute the Value part of Class alfa.

class alfa {
	{Read z as double=10}
	x=z
	id=int(rnd*100)
	value {
		=.X*100
	}
	remove {
		? "deleted "+.id
	}
	' if we comment these lines
	' then we get a uinque pointer
	' so array(b) return a copy of actual object
	{ ' return a pointer (object may have multiple pointers)
		' standard group has a unique pointer.
		->group(alfa)
		break
	}
}
// numeric, pointer to group, numeric, pointer to group, numeric
a=(1,alfa(5),3,alfa(15),5)
link a to a()
b=each(a)
variant z
while b	
	' fixed eval to work with nornal numeric values too
	' we need eval() to run value part of class alfa
		print eval(array(b)), b^   ' return a copy - but we have copy on pointer
		print eval(a(b^)), b^  ' a() always return the item
end while
clear a
print len(a)=0


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and then open it again.

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