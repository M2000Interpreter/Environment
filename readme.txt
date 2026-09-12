M2000 Interpreter and Environment
Version 15 Revision 45

1. Upgrade Asm2  example
Now we can make Version Resource using M2000 code 

So now we can make an exe with resources: manifest, ico and Version info.

2. I found two things and fixit:

2.1 A missing functionallity in Print statement. Normally if we pass code 0 to 31 we get unicode characters which show something lik FF (for 12). When a string was bigger than the width of console only the last part show these characters (the other just move the cursor). Now I fix it to work ok. Note that this not work using  Report (which process the CR+LF), or whaen we use PRINT to file (and PRINT #-2, "print to screen as file" which also CR + LF working as in any console).

2.2 If buf1 is a buffer object and also buf2 is another one we can do this buf1[0]=buf2 to copy bytes drom buf2 to buf1 at offset 0.  The problem was inside M2000 code the object buf1 copied in pad object and that pad was not cleared (until a numeric expression starts). So a Print "alfa" found the object and return "Buffer" and not "alfa". Now a fix it (in an interanl TakeOffset function).

 
George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do this:

dir appdir$
load info

then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).

English old paper for M2000
https://github.com/M2000Interpreter/Environment/releases/download/ver13rev44/M2000paper.pdf

Greek Book for learning programming
https://github.com/M2000Interpreter/Environment/releases/download/version15revision10/GreekBookM2000.pdf

Greek Manual (a work in progress)
https://github.com/M2000Interpreter/Environment/releases/download/version15revision39/GreekManualVersion15.pdf

Greek Book About OOP in M2000
https://github.com/M2000Interpreter/Environment/releases/download/version14revision51/OOP_M2000_2026.pdf

http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 