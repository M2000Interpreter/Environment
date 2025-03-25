M2000 Interpreter and Environment
Version 13 revision 32 active-X
1. fix a mistake for function VAL() with negative values with exponent part.
2. Some fine tuning for M2000 Editor and EditBox control when we click on parts of names with multiple parts separeted with dot (like aaa.bb.ccc.dd)
3. Change the M2000.exe to start the M2000.dll if found in the same folder, if there is no file with name M2000.vbp in the same folder, or no M2000.dll then m2000.exe expect the m2000.dll exist as a COM object on registry of Windows.

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
                 