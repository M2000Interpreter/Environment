M2000 Interpreter and Environment
Version 14 Revision 43

Fix English Case and Greek Με when use semi boolean function, like Case <50
Interpreter see < after Case and believe is not a command at Flow Execution level (as in Select Case, for cases)
So now I fix it by intercept the error and reroute it. Also I found another bug in Select Case and now is ok.  
This is the test code:

tuple=(100, 200, 77, 500, 1005, -20, 75, 90)
Case=100  ' this is a variable
for i=1 to 100
    x=tuple#Val(Random(0,7))
    LIST
    Print "x=";x
    Select Case x
    Case 100, 200
        Print "One line1"
        Print "One line2"
    Case 1000, 500
    		x+=10
    Case <50
    {
        x++
    }
    Case 70 έως 80
  			Select Case x
  			Case 73
  	      	x/=10
  	    Case Else
  	    		x|Υπόλ 10
  	    End Select
    Case Else
        x-=50
    End Select
    Print "Export x=";x, Case=100, i
next 

// This is the Greek Version:

Π=(100, 200, 77, 500, 1005, -20, 75, 90)
Για ι=1 εως 100
  	Με=100
  	Χ=Π#Τιμή(Τυχαίος(0,7))
  	Λίστα	
  	Τύπωσε "Χ=";Χ
  	Επίλεξε Με Χ
  	Με 100, 200
        Τύπωσε "Σε μια γραμμή1"
        Τύπωσε "Σε μια γραμμή2"
  	Με 1000, 500
  			Χ+=10
  	Με <50
    {
        Χ++
    }
  	Με 70 έως 80
  			Επίλεξε Με Χ
  			Με 73
  	        Χ/=10
  	    Με Αλλιώς
  	        Χ|Υπόλ 10
  	    Τέλος Επιλογής
  	Με Αλλιώς
  	    Χ-=50
  	Τέλος Επιλογής
  	Τύπωσε "Εξαγωγή Χ=";Χ, Με=100, ι
Επόμενο


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