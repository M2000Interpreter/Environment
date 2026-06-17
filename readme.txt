M2000 Interpreter and Environment
Version 14 Revision 48


1) Add outputbase value (optional) in BigInteger() function
2) we can add string and biginteger to get get string 

a=biginteger("1AC3FF4FAFE", 16, 2) ' two optional inputbase and outputbase
biginteger b="127398127398172837128937812738127381273812738"
biginteger c=127398127398172837128937812738127381273812737u
b=>outputbase=16
print a ' binary
print b ' hex
c++
print b=c
hex a
print a+"value" ' add string return string
print "0b"+a  ' add string return string 
print "(0b"+a+")" ' add string return string


3) Fix Sprite moving (using hardware sprites through Player) when the player move I found thisL I hade to hide, then to move, then to show (this was too fast to see the hide state, but the moving is instant, otherwise we get a non sync rendering). I found it looking the sprite rotation which not have problem so I think maybe the problem can be fixed with hide/show effect and I it is that.
Execute Sprites module on Info file (see below how to load it).

4) Some minor additions like error messages for For loop when we use BigInteger for control value (not allowed), or complex type (also not allowed).



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