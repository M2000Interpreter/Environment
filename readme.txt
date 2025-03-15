M2000 Interpreter and Environment
1. FIX THE INPUT STATEMENT FOR GREEK SETTINGS (NOW INPUT FOR DECIMAL NUMBERS WORK FINE.
WE PRESS 12.12 AND WE SEE 12,12 AND THAT RETURN TO 12.12 AND INSERT TO A
This was broken from revision 17 
There was no problem for locale 1033 (Default English)
The Locale 1032 statememt set the Greek Locale, for M2000 the keyboard key "." return "," at the Input process for Input statement.

LOCALE 1032 
KEYBOARD "12.12", 13
INPUT A
PRINT A=12.12

The other input variants have no problem:

LOCALE 1032
A=12.12
' THIS INPUT VARIANT NOT USE THE SAME KEYBOARD BUFFER
' THIS USE A CONTROL OVER THE CONSOLE TO MIMIC THAT IS THE CONSOLE
' WE CAN'T PASS KEYS USING KEYBOARD STATEMENT BUT WE CAN STOP THE INPUT
' AFTER MAKE A THREAD WHICH START 200msec after and STOP the input.
' IF WE GIVE MORE TIME (2000 FOR 2 SECONDS) WE CAN CHANGE THE VALUE.
AFTER 200 {INPUT END}
INPUT ! A, 10  ' 10 IS THE WITDH IN CHARACTERS FOR THE CONTROL
PRINT A
PRINT A=12.12

For User forms the input control for numbers also have no problem.


2. FIX FOR FORMLABEL STATEMENT - ALSO THE HELP FILE UPDATED FOR FORMLABEL.


Version 13 revision 27 active-X

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
                 