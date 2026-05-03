M2000 Interpreter and Environment
Version 14 revision 24 active-X
Fix an issue from Revision 20 to 23
The third test now show result, no error
This was not hard to fix but hard to find out as an issue:
because the issue was after the Or (or similar operator) and when the first value was an enumarator type which return a string value.
As you see the second example has "."=z so this works for Revision 20 to 23
The fix was the addition of AND noand  (where noand is a flag, and used from evaluator to stop before AND/OR/XOR, so a noand=false means we have these and we continue. So after AND/OR/XOR a string value can be placed (and this was known from evaluator, see the two tests, first and second), but for the Enumarator type has not fixed as the first operand after AND/OR/XOR (only for these revisions). Previous revisions work different so they don't have this issue. 

module testme (a as boolean, changeit as boolean=false) {
	if a then
		enum aa {
			a_dot="."
			a_nothing=""
			a_underscore="_"
		}
		z=a_dot
	else
		z="."
	end if
	boolean dot
	if changeit then
		? z="" or "."=z
		? (z="." and not dot) or ""=z
	else
		? z="" or z="."
		? (z="." and not dot) or z=""
	end if
}
Try {
	testme false
}
Try {
	testme true, true
}
Try {
	testme true
}


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