M2000 Interpreter and Environment
Version 14 revision 7 active-X

The read only variable internet call a function which use the 142.250.187.100:80 to check internet connection. But for some reason this ip not repsonsed.
So I put two more. This fix also fix the read only variable internet$ which return the ip of the current pc from outside (which also check the same function for checking the internet connection).

We can write M2000 code for the old Internet read only variable:
// the CLIENT object is the cTlsClient object (cTlsClient1.cls)
FUNCTION check_internet {
	DECLARE CLIENT CLIENT
	WITH CLIENT, "NoError", TRUE
	METHOD CLIENT, "SetTimeouts", 100, 300, 200, 300
	// change 100 to 0 or 10
	METHOD CLIENT, "Connect","142.250.187.100", 80 AS CONNECT
	=CONNECT
}
We can write M2000 code for internet$

// HTTPS.REQUEST is object clsHttpsRequest (HttpsRequest.cls)
FUNCTION  get_ip$ {	
	DECLARE HttpsRequest HTTPS.REQUEST
	WITH HttpsRequest,"BodyFistLine" AS RESP$
	METHOD  HttpsRequest, "HttpsRequest", "HTTPS://ifconfig.co/ip" AS OK
	IF OK THEN =RESP$
}
PRINT get_ip$()




 
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