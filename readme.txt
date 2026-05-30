M2000 Interpreter and Environment
Version 14 Revision 40


1. Fix Modules ? "sreach"  ' was a mistake from some revisions before - although Modules ? "","alfa|beta" was ok
2. Some minor fixes (undocument GetObject())
3. Some major additions: We can use => to address objects (like dot). Works for properties, methods but bot for functions. For functions we can use the same way we use for indexes for properties or we have to use Method which can use buy reference arguments and named arguments. Also method use As for writing the result and can use WITHEVENTS AS name, which make event object for the object which we get as result.

See the example below: Cells is variable which hold two things: The App (Excel.Application) and then number of Cells property. We can use the value of Cells(1,1) or the range (object) as Cells(1,1)^
Also we can use Cells=>item(1,1)=>value 

so Print Cells(1,1) = Cells=>item(1,1)=>value
print True

//	Declare WithEvents App "Excel.Application" 
App = GetObject("", "Excel.Application")
Print type$(App)
WorkBooks=App=>WorkBooks
Print "Books:";WorkBooks=>count
If WorkBooks=>count=0 Then WorkBooks=>Add  ' same as Method WorkBooks, "Add"
Print "Books:";WorkBooks=>count
ActiveWorkBook=App=>ActiveWorkBook
ActiveSheet=App=>ActiveSheet
' App=>Cells(1,1)=12312.4   ' same as the next statement
App=>Cells(1,1)=>value=12312.4
' we get the Cells object which is a range
Cells=App=>Cells
Print Cells(1,1) = Cells=>item(1,1)=>value
Print Cells(1,1) = Cells=>item(1,1)
Print Type$(Cells(1,1)^)="Range"
Cells(1,2)="=A1*100"
Print Cells(1,1)
' there is no SET in M2000
' so we get the object (to not get the default property using ^
range1=Cells(1,1)^
Print range1=>Address="$A$1", Cells(1,1)=>Address="$A$1" 
Print Type$(Cells(1,2))="Double", Type$(Cells(1,2)^)="Range" 
Print Cells(1,2)=>Address
Print Cells(1,2)=>formula
Print Cells(1,2)=>value
ActiveWorkBook=>Saved=True
wait 400
App=>quit


Second Example:
if exist("mytest.xlsx") then
	WorkBook=GetObject("mytest.xlsx")
	App=WorkBook=>Parent : Print type$(App)
	WorkBook=>ActiveSheet=>Cells(1,1)=>value=12312.4
	Cells=WorkBook=>ActiveSheet=>Cells
	print Cells(1,1)=12312.4, Cells(1,1)=>Address="$A$1"
	Range=Cells(1,1)^   ' this is like Set Range=Cells(1,1) - M2000 has no a SET for objects
	print Range=>value=12312.4, Range=>Address="$A$1"
	WorkBook=>Saved=True
	wait 400
	App=>quit
end if


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