M2000 Interpreter and Environment

Version 12 Revision 37 active-X

Now reading values from file using Input #FileHandler when we use no character 34 for strings, handle null string.
So we have this ANSI csv file as test.csv:
header1;Header2;header3
label1;;
label2 bla bla;3;
label3 bla;0;
label4;8;

And we can read it using this (see after label1 we have null value, and for header3 we have for each row a null value:

Input With ";",,,true
string A, B, C
open "test.csv" for input as #F
while not eof(#F) 
	input #F, A, B, C		
	Print A, B, C
	print 
end while
close #F

// we can save the test.csv as UTF16LE without BOM, as "testUTF16LE.csv" and we change the code like this (we place Wide on OPEN statement):
open "test.csv" for wide input as #F

//Also you can check with Unix line separator:
// This Encode ANSI (based on Locale - change using Locale 1033 or other number)
// new line lf and no BOM

document a$
load.doc a$, "test.csv"
a=-13
save.doc a$, "testUnix.csv", a

Print " Encode "+if$(abs(a mod 10) mod 4+1->"UTF-16LE", "UTF-16BE", "UTF-8", "ANSI");
Print " | newline:"+if$(abs(a div 10) +1->"crlf", "lf", "cr");
Print " | BOM: "+if$(a<0->"No ", "Yes")

// using a=10 we have Encode UTF16LE, lf, BOM
// For Reading, we have to skip BOM using:
Seek #F, 2   

// For M2000 we can use Seek statement and Seek() function, to read/set the file cursor for specific file handler (we may have more than one file handler for any file). Also files may have more than 2Gbytes (value type for Seek is Currency).


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

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

M2000language.exe (Chrome can't scan, say it is a virus - heuristic choise)
All exe/dll files are signed
https://github.com/M2000Interpreter/Environment/releases

M2000 paper (305 pages). Included in M2000language.exe
https://drive.google.com/file/d/1pHBjLVeaGkyMhyyfvXyvh42cJ3njY7wa

M2000 Greek Small Manual (488 pages). Included in M2000language.exe
https://drive.google.com/file/d/1_2E-4_Eg10yvdGAhEaS3IxW2nI0jDtrH
                                                             