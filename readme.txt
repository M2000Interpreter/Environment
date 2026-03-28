M2000 Interpreter and Environment
Version 14 revision 11 active-X

1) FORMALABEL variant using ! for non antialiasing text rendering. See help FormLabel
2) IMAGE() now works for a bitmap in a string using the same as for PNG except the transparent color. See help Image()
also fixed on the resize to not leave a line in right and potom edge.
3)Report now has another clause: SCROLL for the specific variants who use selected lines to display. So without Scroll we get the lines only which we see without scroll (same in a loop). Using Scroll we get the same like the Report without limit lines (the standard Report). So the SCROLL enable the scroll hold system for page view.
4)fixed in OpenGl the problem with background not alignment as excpected (removed a multiplication 1.01 which alter the size of the texture)



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