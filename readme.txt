M2000 Interpreter and Environment
Version 13 revision 30 active-X

Continue the work of revision 29 on select case
This example for case 4 we get
1.
2.
3. 
First "select" find case 4. But we have Break.
A break in the block on a block of a case means continue execute all cases until one block of a case execute Continue.
So "select" execute case 5 and case 6 which have continue, so skip case 7 (case else always if one case match the case)
A Break or a Continue or an Exit not using in a block on a case, goes to block which the select/end select exist..
So for this example if we place a REM before Continuw we get also the case 7
If we place a REM before break in case 4 we get only the 1. and no other case
if we place a second block or a statement and then a block we get error. After case we may have lines or a block. If you have lines but not in a block do not place loops For/While/Do. Place these on blocks only. The if statement can be used as one line statement or multilines using blocks: If then { multiline block } else { multiline block} etc.

select case 4
case 1
	? "not this"
case 2
	{
		' nothing
	}
case 4
	{
		? "1."
		break
	}
case 5
	? "2."
case 6
	{
		? "3."
		continue
	}
case 7
	? "not this too (1)"
case else
	? "not this too (2)"
end select


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
                 