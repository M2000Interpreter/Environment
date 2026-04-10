M2000 Interpreter and Environment
Version 14 revision 16 active-X
Many Improvements

Timezones have different names based on language per PC. The names are written in Registry so now M2000 read the registry. We get a List from time() without parameters.
There is a zones module which place the zones in a drop down menu for selection.

zones=time()
k=each(zones)
while k
	Report eval$(k!)+" "+(eval(k))
end while

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