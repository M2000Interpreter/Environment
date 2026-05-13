M2000 Interpreter and Environment
Version 14 Revision 30

The Best.
1) Now simple functions (made using function/end function) not need the @ before the name.
2) Fix a problem which came with the new evaluator. The problem was inside parenthesis and the unary minus.
b=-1 : ?(-b+abs(b))  ' 2  - the fault return 0
3) Fix some problems not realy known.
4) Now we can write BSTR type of string in structures/Buffers.
The real string is written in a special hash table, and at the field we place the address of the first character or 0 (for empty)

Structure Alfa {
	a as double
	b as double
	Caption as string ' this is a BSTR string
}
Structure Beta {
	First as Alfa
	Second As Alfa*10
}
Beta Fields[20]
Fields[0]|First|Caption="This is a string - which take only 4 bytes in structure"
Fields[0]|First|a=1.23456e5
Fields[0]|First|b=0.45455
Print Fields[0]|First|Caption
PrintData(Fields[0]|First[]) ' pass a copy 
Fields[10]|Second[4]=Fields[0]|First[]
PrintData(Fields[10]|Second[4])  ' pass a copy 
Fields[10]|Second[4]=ChangeCaption(Fields[10]|Second[4], "New value for caption")
PrintData(Fields[10]|Second[4])  ' pass a copy 
Fields[10]|Second[4]|Caption="ok"
PrintData(Fields[10]|Second[4])  ' pass a copy 
Alfa JustAlfa
JustAlfa=Fields[10]|Second[4]
PrintData(JustAlfa)  ' pass a copy of a pointer
Print JustAlfa|Caption="ok..."  ' changed
PrintData(JustAlfa)  ' caption changed
JustAlfa|Caption="Final"
PrintData(JustAlfa[0]) ' pass a copy of buffer
Print JustAlfa|Caption="Final"  ' not changed
Print JustAlfa(0) ' absolute address of memory buffer

End
Sub PrintData(d as Alfa) ' pass by pointer
	Print "a:", d|a
	Print "b:", d|b
	Print "caption:",d|Caption
	d|Caption="ok..."	
End Sub
Sub ChangeCaptionByreference(&d as alfa)

End Sub
Function ChangeCaption(d as Alfa, newValue as string)
	d|Caption=newValue
	=d
End Function
George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and then open it again.

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