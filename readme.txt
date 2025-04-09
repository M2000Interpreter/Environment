M2000 Interpreter and Environment
Version 13 revision 41 active-X

Fix the second, when was array when we read the read only array, we expect to get a copy and not the actual object whis is a read only:
' first
const alfa=(,)
push alfa
read z as array  'we say array so we get a copy from constant
print len(z)=0 ' ok
push (1,2,3)
read z ' now is ok
print len(z)=3
clear
' second
const alfa=(,)
push alfa
z=(1,2,3)
read z
print len(z)=0 ' ok
push (1,2,3)
read z ' now is ok
print len(z)=3
clear
' third
const alfa=(,)
push alfa
read z          ' z get first value, so get the constant
print type$(z)="tuple" ' but is a read only tuple.
print len(z)=0 ' ok
push (1,2,3)
try {
	read z ' now is ok
	print len(z)=3
}
Print Error$
clear
' forth
const alfa=(,)
push alfa
read z as const
Print type$(z)="Constant"
Print type$((z))="tuple"  ' but read only
' let m=z ' this get the Constant
m=z ' this get a copy
print m is z = false
push (1,2,3)
read m ' ok


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
                 