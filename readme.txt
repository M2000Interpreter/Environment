M2000 Interpreter and Environment
Version 13 revision 28 active-X

This was not an error but a lack of design:

To call a sub or a simple function before revision 28 the language for the definition at the first search was inferenced from the language of the name (checking the first letter). That works but some time you want to use f5 in editor to change names, so change a Greek name to an English one you will expect that code run as before, but not if the name you just replace is a sub's or simple function's name, because you have to alter the definition according to specific language of its identifier. So now this isn't a problem anymore. You can mix Greek/English.

After first search, the actual position recorded for subsequent call.

For modules/functions/lambda functions this was not a problem because the definition statement must be interpreted before the call. See here module testme {} is the definition and after this the testme is the call.


MODULE TESTME {
	ΑΛΦΑ()
	IF VERSION>13 OR (VERSION=13 AND REVISION>27) THEN ALFA()
	IF VERSION>13 OR (VERSION=13 AND REVISION>27) THEN ΑΛΦΑ1()
	ALFA1()
	? @ΑΛΦΑ()
	IF VERSION>13 OR (VERSION=13 AND REVISION>27) THEN ? @ALFA()
	IF VERSION>13 OR (VERSION=13 AND REVISION>27) THEN ? @ΑΛΦΑ1()
	? @ALFA1()
	
	' ΡΟΥΤΙΝΑ/ΡΟΥΤΊΝΑ/ΣΥΝΑΡΤΗΣΗ/ΣΥΝΆΡΤΗΣΗ/SUB/FUNCTION ALIAS END ..
	' IF ENCOUNTERED AS NEXT STATEMENT BY INTERPRETER.
	ΡΟΥΤΙΝΑ ΑΛΦΑ()
		? 100
	ΤΕΛΟΣ ΡΟΥΤΙΝΑΣ
	ΡΟΥΤΙΝΑ ALFA()  ' NOT FOUND <28 GREEK/ENGLISH
		? 100
	ΤΕΛΟΣ ΡΟΥΤΙΝΑΣ
	SUB ΑΛΦΑ1()  ' NOT FOUND <28 ENGLISH/GREEK
		? 100
	END SUB
	SUB ALFA1()
		? 100
	END SUB
	ΣΥΝΑΡΤΗΣΗ ΑΛΦΑ()
		=101
	ΤΕΛΟΣ ΣΥΝΑΡΤΗΣΗΣ
	ΣΥΝΑΡΤΗΣΗ ALFA()  ' NOT FOUND <28 GREEK/ENGLISH
		=101
	ΤΕΛΟΣ ΣΥΝΑΡΤΗΣΗΣ
	FUNCTION ΑΛΦΑ1()  ' NOT FOUND <28 ENGLISH/GREEK
		=101
	END FUNCTION
	FUNCTION ALFA1()
		=101
	END FUNCTION
}
TESTME



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
                 