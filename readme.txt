M2000 Interpreter and Environment

Version 13 revision 15 active-X
1. Fix ReadMe for optional parameters
Function TestMe$ {
	Read a$, b$, pp=10u, rd as boolean=false,  z as complex=(1,2i)
	Print TYPE$(PP) ' BigInteger
	Print type$(rd)
	Print z
	LIST
}
K=TestMe$("AA", "BB", 100,  true, (2,-5i))
K=TestMe$("AA", "BB", 100)
K=TestMe$("AA", "BB")
2. New addition. We can pass a value as a constant which isn't parameter for subs and simple functions. This not used for modules, functions, lambda functions.
Puropose for this: We can use values as properties for subs and simple functions without feeding the stack of values. These properties are not changed inside subs and simple functions, they are constant.
The site on parameter list can be anywhere, so here we have b:=500 as first item and at the next line as second item.

print @alfa(b:=500, 300)=500
print @alfa(300, b:=500)=500
print @alfa(300)=100
alfa(b:=500, 300)
alfa(300, b:=500)
alfa(300)
function alfa(x)
	const b=100
	=b
end function
Sub alfa(x)
	const b=100
	print b, x
end sub

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

https://rosettacode.org/wiki/Category:M2000_Interpreter (534 tasks)
                 