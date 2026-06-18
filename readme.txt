M2000 Interpreter and Environment
Version 14 Revision 49

A bug removed (Was there for at least tow revisions I suspect)
The bug was in While End while. Repeat Until, Repeat When, Repeat Always
When we call a sub which is not in the same module but in the parent module, and only on the first call the read function not executed  so the i variable not exist;
This fault not happen if the sub is in the module inner
Also this fault not happen if we use For loop or a standard block { }
(It is one statement error, which I forgot to change when I change the read function for better results from a monolithic old function)
So now this test display true the i exist.

module inner {
	m=true
	while m {
		sub1 100
		m=false
	}
}
inner

sub sub1(i)
	? valid(i)
end sub



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