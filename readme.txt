M2000 Interpreter and Environment
Version 14 Revision 37

1) A fault index send value a(2)+a(1) to a(1) not a(3). Now fixed
a=list:= 0:=0, 1:=1, 2:=2
a(3)=a(2)+a(1)
? a(3)=3

2) fix the algorithm for finding the type of text file. A problem was about unicode without BOM UTF16 (LE or BE) vs ANSI.
The problem was for ANSI file which the system open using Edit or Load.Doc and some time read it as UTF16LE.
There is a new module in INFO.gsb called  CheckDoc which make all possible combinations for UTF8/UTF16LE/UTF16BE/ANSI with line enconding CRLF, LF, CR with BOM/No BOM (ANSI always without BOM). Also this show the algorithm works. And also show how we can program to save various combinations (reading is automatic, unless we suggest the type, although that isn't always a command because for some types algorithm are sure about it).
You can open any text file with the inner editor of M2000 console (press Esc to save it or Shift F12 to not save it):
Edit "TestDocUTF8.txt"
You can open this file using the defaul application for txt files:
Win "TestDocUTF8.txt"
or
Win TestDocUTF8.txt
Some statements may get filenames/paths without quotes, but not Edit because without "" make/edit modules  

3) fix a problem with a comment using ' after calling a sub...in another sub only;
Module TestModule {
	alfa(10)' no problem here
	End	
	sub alfa(m)
		beta(m*2)'here was the problem
		Print m
	End Sub
	Sub beta(x)
		Print x, "inner"
	End Sub
}
TestModule


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