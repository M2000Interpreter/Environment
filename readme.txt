M2000 Interpreter and Environment
Version 13 revision 29 active-X

1. Fix select case for strings.
2. Update Select case. Now we can use multiline strings literals. Before this release between two case statements we may have a line (using : to seperate statements) or a block. Now we can use more lines but not for loops or nested select cases, we have to use block of statements.

We get the second case...
select case {aa
}
case "k" to "z"
	print "no"
case >{a
}
	print "ok"
	print "------------"
case else
	print "line1"
	print "line2"
end select

previous (from version 10):
a$={aa
}
z$={a
}
select case a$
case "k" to "z"
	print "no"
case >z$   ' two statements in one line
	print "ok": print "------------"
case else
{ ' or using a block. Also here we may have nested select/case
	print "line1"
	print "line2"
}
end select

3. Update GuiM20000 form. I have a mistake from revision 26, which not use enable property, so all controls get visible=true (which is not true for every case). So now fixed.
4. English and the same for Greek language now are coloured as a statement (I have forgot to fix this).




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
                 