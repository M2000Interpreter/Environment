M2000 Interpreter and Environment
Version 14 Revision 46

A big fault when using after AND/OR/XOR/NOT two or more strings and then comparison operator. Now is ok.
Test Example:

a$="aa" ' same without suffix $
b$="bb" ' same without suffix $

' these was ok:
if a$+b$="aabb" then print "ok 1"
if "aabb"=a$+b$ and a$<>b$ then print "ok 2"
if a$+b$="aabb" and a$<>b$ then print "ok 3"
if a$<>b$ and "aabb"=a$+b$ then print "ok 4"

' this was wrong:
if a$<>b$ and a$+b$="aabb" then print "ok 5"

' these was ok:

if (a$+b$="aabb") then print "ok 1-1"
if ("aabb"=a$+b$) and a$<>b$ then print "ok 2-1"
if (a$+b$="aabb") and a$<>b$ then print "ok 3-1"
if a$<>b$ and ("aabb"=a$+b$) then print "ok 4-1"

' this was wrong:
if a$<>b$ and (a$+b$="aabb") then print "ok 5-1"

' this was wrong:
if not a$+b$="aabb" then print "ok 1-2"

' these was ok:
if "aabb"=a$+b$ and not a$=b$ then print "ok 2-2"
if a$+b$="aabb" and not a$=b$ then print "ok 3-2"
if a$<>b$ and not "aabb"<>a$+b$ then print "ok 4-2"

' this was wrong:
if a$<>b$ and not a$+b$<>"aabb" then print "ok 5-2"

? "Without suffix $ - these are different variables"
a="aa"
b="bb"

' these was ok:
if a+b="aabb" then print "ok 1"
if "aabb"=a+b and a<>b then print "ok 2"
if a+b="aabb" and a<>b then print "ok 3"
if a<>b and "aabb"=a+b then print "ok 4"

' this was wrong:
if a<>b and a+b="aabb" then print "ok 5"

' these was ok:

if (a+b="aabb") then print "ok 1-1"
if ("aabb"=a+b) and a<>b then print "ok 2-1"
if (a+b="aabb") and a<>b then print "ok 3-1"
if a<>b and ("aabb"=a+b) then print "ok 4-1"

' this was wrong:
if a<>b and (a+b="aabb") then print "ok 5-1"

' this was wrong:
if not a+b="aabb" then print "ok 1-2"

' these was ok:
if "aabb"=a+b and not a=b then print "ok 2-2"
if a+b="aabb" and not a=b then print "ok 3-2"
if a<>b and not "aabb"<>a+b then print "ok 4-2"

' this was wrong:
if a<>b and not a+b<>"aabb" then print "ok 5-2"



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