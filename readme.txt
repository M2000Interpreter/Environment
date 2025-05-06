M2000 Interpreter and Environment
Version 13 revision 47 active-X

1. Fix a small problem with the later revision when we use fonts with break char different from space. The last problem was when we print at last line from specific statements which use a specific function (but not the Report statement).
This small program can run on revision 46 and 47. We will see that we see some blank lines inserted on revision 46 but not on revision 47 (which is the right one)
font "Segoe Ui"
for i=1 to 50
	Fkey
	? i
next
print "ok"

2. This is something from VB6. Some objects may have different interfaces. How we can get another interface? We have to use the QueryInterface function, using a specific code for the specific interface, and we get a new pointer to the new interface.

This example get a second pointer from an object, using the same idispatch interface, but we make it using the IUnknown.QueryInterface  function, which se declare as Interface. So we can use this function passing as first parameter the object which we want to run the specific function. When we pass the object the function first check if the object has the specific interface, and then call the specific function using the specific interface.

idispatch$="{00020400-0000-0000-C000-000000000046}"
enum IUnknown {
	QueryInterface=0
	AddRef
	Release
}
enum Idispatch {
	IUnknown
	GetTypeInfoCount
	GetTypeInfo
	GetIDsOfNames
	Invoke
}
DECLARE IUnknown.QueryInterface INTERFACE idispatch$, QueryInterface {long riid, long myptr}
K=List:=1,2,3,4
object Ptr
buffer clear riid as Long*4
return riid, 0:=0x0002_0400, 1:=0,  2:=0xc0, 3:=0x4600_0000
buffer clear myobj as long
? "0x"+hex$(uint(IUnknown.QueryInterface(K, riid(0), varptr(ptr))), 4)
? type$(ptr)  ' inentory type
? len(ptr)=4 
? ptr is K  = true


This is my first try. So I have to make some other examples. But for now this example works fine. See how I made the GUID (or RIID), the buffer with 16 bytes, which write the Interface ID for IDispatch type . Also see how we make enum with autonumbering, and using enum on enum.

The IUnKnown interface is this:
return riid, 0:=0, 1:=0,  2:=0xc0, 3:=0x4600_0000
Some objects has no IDispatch interface, and that object can't work with Method and With for properties. We have to declare functions (with specific offset) and signature of parameters (although M2000 need pointers of type Long for the signature). See also the use of Varptr for object Ptr. We place Ptr by reference when we place the Varptr(Ptr).




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

https://rosettacode.org/wiki/Category:M2000_Interpreter (544 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 