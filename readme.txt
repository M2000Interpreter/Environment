M2000 Interpreter and Environment
Not Finish: Version 15 Revision 28

August 23

1. Fixed a mistake for showing long long values from an array of variants through auto iterator in Print statement
Dim A(2)
A(0)=100, 200&&
Print A(2)=200 ' was ok
Print A()  ' now print 100 200 - A() replaced with an iterator and prints all values.

Clear ' clear variables for the current module.
Dim A(2) as long long
A(0)=100, 200&&  ' value 100 converted to long long.
Print A(2)=200 ' was ok
Print A()  ' was ok .


2. Fixed passing object to a com object property through choosing only Set (previous I use Let + Set in one call, which works but not for any object)
declare dict "scripting.dictionary"
dict(100)=12312213u ' the BigInteger is an object
Print Type(dict(100))="BigInteger"
Print dict(100)*12345=151994269485u
' M2000 by design can't handle Nothing as object.
' So to nullify an object we have this statement:
declare dict nothing

3. Adding Passing Objects for calling functions by address (by design this was locked)

declare global CLSIDFromString Lib "ole32.CLSIDFromString" {lpszCLSID$, Long Clsid}
global interface IUnknown, "{00000000-0000-0000-C000-000000000046}" {
	QueryInterface : {long riid,long &myptr},
	AddRef,
	Release
}
global peekLong=lambda -> {
	buffer ptr as long 'used only for peekInt32
	=lambda ptr ->ptr=>peekInt32(number)
}() 'execute now
global const S_OK=0& 
global  clsid  : buffer clsid as byte*16 ' clsid(0) is address of first byte
call CLSIDFromString("{00020400-0000-0000-c000-000000000046}", clsid(0))
global dictionary=getobject("","scripting.dictionary")
Module Inner {
	long ObjPtr=PeekLong(varptr(dictionary))
	' Part I: Dereference to find the address of QueryInterface 
	Long Vtable=PeekLong(ObjPtr)
	Queryinterface=PeekLong(Vtable)
	' now we make the function, as c decl call
	' declare code deigned for assembly probjects
	declare dictionary.Queryinterface code c Queryinterface {
		long ObjPtr, ' need a valid pointer
		long riid,   ' need address of first byte of 16 bytes 
		long &myptr  ' return object if we get S_OK 
	} as long
	object ret
	' passing object dictionary 
	If dictionary.queryinterface(dictionary, clsid(0), &ret)=S_OK then
		Print type$(ret)="Dictionary"
	end if
	declare ret nothing
	' passing address (objptr)
	If dictionary.queryinterface(ObjPtr, clsid(0), &ret)=S_OK then
		Print type$(ret)="Dictionary"
	end if
	declare ret nothing
	' Part II: Using interface call
	'passing object dictionary
	If IUnknown.QueryInterface(dictionary, clsid(0), &ret)=S_OK then
		Print type$(ret)="Dictionary"
	end if
	declare ret nothing
	' passing address (objptr)
	If IUnknown.QueryInterface(ObjPtr, clsid(0), &ret)=S_OK then
		Print type$(ret)="Dictionary"
	end if
}
Inner

4. Select an interface 
Class various {
	Interface IUnknown, "{00000000-0000-0000-C000-000000000046}" {
		QueryInterface : {long riid, long myptr}
		AddRef :{}
		Release :{}
	}
	Function get_unk(t as *IUnknown) {
		=t
	}
}
b=list
a->various()
M2000_list = a=>get_unk(b)
' b has a mHandler object which have a FastCollection
' the mHandler object has t1 property 1 : is a List or Queue (a list which allow same keys)
' Fumction get_unk find that we pass an mHandler object, and get the object under it
' then perform a QueryInterface using the ClsId "{00000000-0000-0000-C000-000000000046}"
' the Clsid stay as binary (not string) isnide Interface (is a type of Enumaretiion Object)
' Enumeration objects also came with mHandler with a t1=4.
Print type(b) = "Inventory"
Print type(M2000_list) = "FastCollection"
Print b is M2000_list 'true

5. Some extra work for the Shift F1 open the line of error.


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