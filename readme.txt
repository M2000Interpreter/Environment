M2000 Interpreter and Environment

Version 13 revision 6 active-X
1. fix send to clipboard (missing last character)
2. fix reading ansi txt with odd number of characters:
a$=str$("abcde") ' return 5 bytes - convert to ANSI based on LOCALE
clipboard a$
z$=clipboard$  ' we get the actual 5 bytes
? len(z$)=2.5 ' true  2.5x2=5 bytes  ' unit Word (2 bytes)
? len(a$)=2.5 ' true
? z$=a$ ' true
3. fix Str$(), Date$(), Time$(), to use both vb6 format, and Windows format.
So we can display: Based on host computer locale, based on standard 1033, and user selected locale id.
Look th sdate module on INFO file (explain how to read/interpret as string using locale id which have a date and a time part).
 
This is the test code:

cls,0
date d="2025.2.6"  ' JUN 2
date d="2025-2-6"  ' FEB 6
date d="2025/2/6"  ' ' FEB 6
date t="20:05:12"
d+=t
greek
'latin
Print Locale
print str$(d, locale), str$(d,"yyyy mmm dd")
print str$(d,"short date")  ' short day/long day using 1033 locale
print str$(d,"ddddd")  ' using host locale
print str$(d,"yyyy MMM dddd d")  ' using 1033 Locale (has M inside)
print str$(d,"yyyy mmm dddd d")  ' using host Locale (has M inside)
print date$(d,,"yyyy MMM dddd d")  ' using Locale (has M inside)
Print str$(t)="20:05:12", str$(t, 1033)="8:05 PM", str$(t, 1032)="8:05 μμ"
Print str$(t, "hh:nn:ss")="20:05:12" ' host locale
Print str$(t, "ttttt") ' host locale long time
Print str$(t, "short time")="8:05:12"  ' has no am/pm
Print str$(t, "long time")="8:05 PM"  ' has am/pm but no secondes
Print time$(t, "long time") ' using locale
Print time$(t, "hh:nn:ss") ' has :n (or n:) so use host locale
Print time$(t, "hh:mm:ss t") ' using locale a/p
Print time$(t, "hh:mm:ss tt") ' using locale am/pm
Print time$(t, "HH:mm:ss") ' using locale 24 hour
Print time$(t, 1033, "long time") ' using 1032 locale
latin
Print Locale
print str$(d, locale), str$(d,"yyyy mmm dd")
print str$(d,"short date")  ' short day/long day using 1033 locale
print str$(d,"ddddd")  ' using host locale
print str$(d,"yyyy MMM dddd d")  ' using 1033 Locale (has M inside)
print str$(d,"yyyy mmm dddd d")  ' using host Locale (has M inside)
print date$(d,,"yyyy MMM dddd d")  ' using Locale (has M inside)
Print str$(t)="20:05:12", str$(t, 1033)="8:05 PM", str$(t, 1032)="8:05 μμ"
Print str$(t, "hh:nn:ss")="20:05:12" ' host locale
Print str$(t, "ttttt") ' host locale long time
Print str$(t, "short time")="8:05:12"  ' has no am/pm 
Print str$(t, "long time")="8:05 PM"  ' has am/pm but no secondes
Print time$(t, "long time") ' using locale
Print time$(t, "hh:nn:ss") ' has :n (or n:) so use host locale
Print time$(t, "hh:mm:ss t") ' using locale a/p
Print time$(t, "hh:mm:ss tt") ' using locale am/pm
Print time$(t, "HH:mm:ss") ' using locale 24 hour
Print time$(t, 1032, "long time") ' using 1032 locale




George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement Settings to change font/language/colors and size of console letters.

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (384 tasks)

ExportM2000 all files with executables (you can get the ca.crt):
https://drive.google.com/drive/folders/1IbYgPtwaWpWC5pXLRqEaTaSoky37iK16

only source, with old revisions and a wiki, for executables see releases
https://github.com/M2000Interpreter/Environment

M2000language.exe (Chrome can't scan, say it is a virus - heuristic choice)
All exe/dll files are signed
https://github.com/M2000Interpreter/Environment/releases

M2000 paper (305 pages). Included in M2000language.exe
M2000 Greek Small Manual (488 pages). Included in M2000language.exe

                                                             