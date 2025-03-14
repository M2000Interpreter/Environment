M2000 Interpreter and Environment

Version 13 revision 26 active-X

1. List com to A ' export ProgID/CLSID for objects
2. Enum with comments as memebers of Groups (added use of comments)
3. DropList from control box on user form, now can

DECLARE Form1 FORM
METHOD Form1, "MakeInfo", -10
METHOD Form1, "MenuItem", "First", TRUE, idD:=500, acc:="A", ctrl:=TRUE
METHOD Form1, "MenuItem", ""
METHOD Form1, "MenuItem", "Second", TRUE, IdD:=1000, acc:="F1"
WITH Form1, "id" AS Who()
FUNCTION Form1.infoClick(NEW V) {
	SELECT CASE VAL(Who(V))
	CASE 500
		ΤΥΠΩΣΕ "ΠΡΩΤΟ", V
	CASE 1000
		ΤΥΠΩΣΕ "ΔΕΥΤΕΡΟ", V
	END SELECT
	REFRESH
}
FUNCTION Form1.MouseDown {
	READ NEW key, shift, x, y
	IF key=2 AND shift=0 THEN		
		METHOD Form1, "OpenInfoAt",x, y	
	END IF
}
METHOD Form1,"SHOW" , 1
DECLARE Form1 NOTHING


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
                 