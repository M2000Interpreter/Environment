M2000 Interpreter and Environment

Version 10 revision 26 active-X
1. Correction in #Sort() clause for tuple
2. new m2000.exe:
   if a m2000.gsb exist in same directory then m2000.dll (or lib.bin as alternate name) loaded (without using windows registry) and the m2000.gsb executed.
   We can change m2000.exe name to something else like myApp.exe, then we can change name to m2000.dll as lib.bin and the auto file as myApp.gsb
   If we use console then we have to include a Show statement to open console
   If we want myApp.gsb to saved encrypted we can save App from console:
	*Change to user directory: c:\users\USERNAME\appdata\roaming\m2000\
		Save myApp @, {Dir User: ModuleA : End}
        *Start in AppDir$ (application directory)
		Save myApp @, {ModuleA : End}
    For Application MyApp minimum files to export:
		MyApp.exe  (M2000.exe rename to MyApp.exe)
		MyApp.gsb  (use one of the above variation of save statement)
		lib.bin (M2000.dll rename to lib.bin)
		(if an old M2000 installation exist, the new one used)
    For M2000 standard environment minimum files, without dll registration:
		M2000.exe
		help2000utf8.dat
		M2000.dll
		Optional: M2000.gsb (empty file or with a SHOW statement)
		(if an old M2000 installation exist, then that old dll used)
		(AppDir$ show the directory of m2000.dll)
		(Use an empty M2000.gsb in same directory to select the local m2000.dll)


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The fist time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)


http://georgekarras.blogspot.gr/

ExportM2000 all files with executables:
https://drive.google.com/drive/folders/1IbYgPtwaWpWC5pXLRqEaTaSoky37iK16

only source without executables (something going wrong with GitHub)
https://github.com/M2000Interpreter/Environment

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             