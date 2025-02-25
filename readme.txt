M2000 Interpreter and Environment

Version 13 revision 15 active-X

1. Assign string value from a constant of type enum to a string global value
enum aaa {
	k="hello", m="other"
}
global a$
a$<=k
Print a$="hello"
2. Same for array
dim a(10) as string
link a() to b$()
enum aaa {
	k="hello",
	m="hello there"
}
a(3)=m
? a(3)
? type$(a(3)), a(3)
b$(4)=m
? b$(4)
? type$(b$(4)), b$(4)

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
                 