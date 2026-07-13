M2000 Interpreter and Environment
Version 15 Revision 3

1) Remove a fault from a literal value for biginteger (broken from last 2-3 revisions):
	local a as biginteger=10
2) Fix BASIC switch to work nice when we call a module from a module which is BASIC enabled. So the READ inside inner is the norma read as designed to be.
module alfa {
	BASIC
	module inner {
		read x
		if x>0 then print x, else print
	}
	do
		read a
		inner a
	until a=0	
	data 1,2,3,4,5,0
}
alfa

3) Upgrade the colorized code procedure:
3.1 Now we get the right color for this: 
	Print IF(1>0->{alfa_String},{beta_string})
	Explain: (the {alfa_String} get colour for string, before upgrade the -> works for lambda so the colorize code procedure interpreted this as code which is not)

3.2 Handle of parenthesis for more than one paragraph. Before upgrade there was a check for balanced parenthesis which works only for line boundaries. Now the check extended if the reason for extension is a lambda function as a parameter spliting in more than one line, or the string literal which use curly brackets (which used for strings with paragraphs - spliting with CRLF) or both as a combination. Also the handling retain the color of the parenthesis. There are three different colors, one for known functions like LEN(), one for user functions and the bare parenthesis. So now the colors for the close parenthesis adjust to the color of the open parenthesis, for more than one paragraph.
This uprgrade done for the TextViewer (the class which is the editor of M2000 code) and the EditBox (which have two editors inside, one for M2000 code and a universal one - see CS and HtmlEditor module in info for editing c# code and html code).

4) Fix an error which prevent the reading of an array passed by VBScript back to M2000 in an example VbScript (now included at INFO file see below). Also I put three modules to show how we can use python from M2000.




  
George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
THEN press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 