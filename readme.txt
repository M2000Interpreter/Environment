M2000 Interpreter and Environment
Version 14 Revision 29

- I found a vb6 compiler fault
In revision 28 I have a call to a friend function of an object passing a FastCollection object but the signature of function want StructCollection and this fault compiler not found, and leave the program compiled and not run at the specific call.
So I change it and I do some other additions. So now when we make a Struct and after we add members to struct, we can use the buffers which made from these structures and also we can alter the size and we can pass to a function checking the older type

In this example we have structure alfa in two versions. The buffers which made with first version can be used, and we can make new with the same old version, either by using a copy of buffer, or by using a second name which not only save the structure (the old pointer, before the Append statement), we can use as statement to define new buffers.


test program:

structure alfa {
	x as long
	y as long
}
buffer kappa as alfa*10
old_alfa=alfa  ' now old_alfa is also a statement for new variables
Structure alfa append {
	z as long
}
buffer delta as alfa*10

? old_alfa is alfa = false ' true, old_alfa hold previous alfa
d=kappa[2]
z=delta[3]
? len(d)=8, len(z)=12

? type$(d)="Buffer", type$(z)="Buffer"
list errors clear
// ? valid(z|z)=true, valid(d|z)=false
// we get the error on the list (from second valid() function)
/ errors just written to list but skipped through valid()
List errors

buffer d as alfa*100  ' now works  ' previous alfa
buffer z as alfa*100  ' now works
if rnd>0.5 then
	buffer dd as old_alfa*10   ' dd is [alfa:10] the previous alfa
else
	' using upgrade 
	old_alfa zz[10]	' now this works too and make olf Alfa
end if
? len(d)=800, len(z)=1200
d1=buffer(d) ' get a copy of d
? d is d1 = false ' true
function ab(d2 as old_alfa){ ' new now take the type from actuall pointer
	=d2
}
d2=ab(d1)
? d2 is d1  ' true we get the pointer to d1
? "ok", len(d1), len(d2)
list ' this is for listing variables.
Modules ?




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