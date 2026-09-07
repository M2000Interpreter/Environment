M2000 Interpreter and Environment
Version 15 Revision 41

Athens, September 7, 2026

1. New read only variable ANY read strings or numbers or objects from stack. 

2. Upgrade Assembler (now we can pass unicode literals (not only with a variable) and we may have lables in any almost any language):

print "chapter 1"
print "strings returned as BSTR, as is or in a VARIANT"
print "We use SysAllocStringLen from oleaut32.dll"
Declare SysAllocStringLen Lib "oleaut32.SysAllocStringLen" {
	Long OleStr, Long BLen
} As Long
Print "  - using Variant - Calling with variant as return value M2000 automatic pass an empty variant"
mycode=assembly({
		push dword 14 			; Length of "Hello World"
		lea eax, [data1]        
		push eax
		call @SysAllocStringLen 	 ; @ read address of SysAllocStringLen()
		mov edx, [esp + 4]      		 ; get the address of hidden empty variant 
		mov word [edx], 8       		 ; vt type = 8 (string)
		mov dword [edx + 2], 0  	 ; Clear reserved fields (Offset 2)
		mov dword [edx + 6], 0  	 ; Clear reserved fields (Offset 6)
		mov [edx + 8], eax      		 ; Place the BSTR pointer into the Variant data (Offset 8)
		mov dword [edx + 12], 0 	 ; Clear reserved fields (Offset 12)
		; so now 16bytes returned via edx 
		ret 4
align 4
data1:	dw "Hello World ??" ; no need 0
})
declare HelloWorld code mycode(0) as variant
Print HelloWorld()
Print "  - using String - just return pointer to BSTR in eax"
mycode2=assembly({
		push dword 14  ; Length of "Hello World"
		lea eax, [data1]
		push eax
		call @SysAllocStringLen
		; string BSTR pointer is in EAX
		ret
align 4
data1:	dw "Hello World ??"
})

mycode2=assembly({
		push dword 14  ; Length of "Hello World"
		lea eax, [data1]
		push eax
		call @SysAllocStringLen
		; string BSTR pointer is in EAX
		ret
align 4
data1:	dw "Hello World ??"
})

declare HelloWorld2 code mycode2(0) as string
Print HelloWorld2()

print "chapter 2 - no need for SysAllocStringLen"
print "strings returned as pointer which have length depend of position of zero"
print "M2000 automatic produce BSTR from pointers"
print "1 - unicode string returned"
mycode3=assembly({
	lea eax, [data1]
	ret
align 4
data1: dw "?????? Hello World ??", 0
})
print "declared only by name of function HelloWorld3$"
declare HelloWorld3$ code mycode3(0)
Print HelloWorld3$()
	
print "declared as string pointer"
declare HelloWorld4 code mycode3(0) as string pointer
Print HelloWorld4()

print "2 - ansi string returned - name of function HelloWorld3"
mycode4=assembly({
	lea eax, [??????] ; we can use arabic also...
	ret
align 4
??????:	db "Hello World", 0  ; we use db not dw for ansi
})
print "  delcared as string pointer ansi"
declare HelloWorld5 code mycode4(0) as string pointer ansi
Print HelloWorld5()
print "  delcared as ansi"
declare HelloWorld6 code mycode4(0) as ansi
Print HelloWorld6()






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