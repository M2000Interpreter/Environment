M2000 Interpreter and Environment
Version 14 revision 20 active-X

1. Fix some bugs from the previous revisions 18 & 19
2. I found a way to call a M2000 program from another M2000 program as a child window which is always infront of the caller but for the caller own windows are like siblings, they share the same group of windows. In Info there is now the clock, which open infront of console. We can run Form44 module which make 3 windows (or more if we change a number). These windows may selected an stand infront of clock.We can put the clock between these windows, but never goes behind the "parent", the console (which is the caller). Clock has transparency and irreqular shape as window.
3. More than 20 changes/additions..
4. Its faster about 15-17% (less time to run tests).
I check it against Info modules. I continue check for breaking code.
5. More modules on INFO file.


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