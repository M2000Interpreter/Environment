M2000 Interpreter and Environment

Version 13 revision 13 active-X
1. Help path$() now works without raising error (was an error from the right parenthesis, because help expectd path$( only, now works with or without right parenthesis).
2. Test Form has a line where we can print something if the prompt is ? or we can execute statement if has prompt > (these prompts tongle when we press backspace, erasing the input line). Now when we send a statement that statement's error now dropped (so we didn't produce errors at the return of execution from the test form).


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
                 