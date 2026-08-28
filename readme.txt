M2000 Interpreter and Environment
Version 15 Revision 35

Athens, August 29, 2026

1. Fixing the ActiveSheet=>Range("A1:B6")=a() which worked partially. The thing is that a() passed by reference and at the returned array not passed properly, and that make Interpreter to abandon next statements; Now fixed.

This example open a xlsx new file in temporary directory and pass two columns using an array to feed the columns. To read the array back we use this trick: a=ActiveSheet=>Range("A1:B6"): a()=array(&a)
1.The variable a get the variant array from Range() and place it in a UDT designed for this reason.
2. Array(&a) and Copy.Arr(&a) when a is a variant array, convert to mArray, by movig the array (without copy)
We need that because we get the copy first time when we got the variant array. So we didn't want another copy.
 

XlsxFilename=temporary$+"DeleteMe"+(int(Timecount))+".xlsx"
//	Declare ExcelSheet "Excel.Sheet"
' this is same as this
ExcelSheet=GetObject("","Excel.Sheet")

ActiveSheet=ExcelSheet=>Application=>ActiveSheet
Dim ole base 1, a(6,2)
a(1,1)="Hello",TODAY,244.123,334.123,55.23,55.34
a(1,2)="Second Column",TODAY,244.123,334.123,55.23,55.34

ActiveSheet=>name="SheetOne"
ActiveSheet=>Range("A1:B6")=a()

' uncomment these using ctrl+/ in M2000 editor
//	variant a=ActiveSheet=>Range("A1:B6"): a()=array(&a)
//	print a()#str$(", ")

ExcelSheet=>SaveAs XlsxFilename, 51

Declare ActiveSheet nothing
Declare ExcelSheet nothing
Win "excel",quote$(XlsxFilename)
Dir temporary$
Files "xlsx"
Wait 5000 ' 5 SECONDS WAIT
f=0
' 4hz watch...
Every 1000/4 {
	Print "Wait..."	
	Files "xlsx"
	Try { ' test for exclusive open
		Open XlsxFilename for input exclusive as #f
		Close #f
		Dos "del "+quote$(XlsxFilename);
		getout=true
	}
	Refresh
	If valid(getout) then exit
}
Print "done"
Files "xlsx"
Dir user ' return to user directory

2. Update the help file (I forgot to put the GetObject() function)


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
https://github.com/M2000Interpreter/Environment/releases/download/version15revision10/GreekManualVersion15_preview.pdf

Greek Book About OOP in M2000
https://github.com/M2000Interpreter/Environment/releases/download/version14revision51/OOP_M2000_2026.pdf

http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 