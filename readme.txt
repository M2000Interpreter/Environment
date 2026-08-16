M2000 Interpreter and Environment
Version 15 Revision 25

August 16, 2026,
1. Fix a small mistake from Revision 24:

' 10 1000  - here was the error (returned fault: 1000 1000)
' Double Constant
TWOVALUE(Y:=1000) ' Y IS CONSTANT VALUE

' 50 1001
' Double Double
TWOVALUE(%Y=1000, %X=50) ' Y IS NORMAL VALUE

' 200 301
' Double Double
TWOVALUE(200, 300)

' 10 301
' Double Double
PUSH 1001
TWOVALUE(?, 300)
PRINT NUMBER=1001

SUB TWOVALUE(X=10, Y)
	TRY {Y++}
	PRINT X, Y
	PRINT TYPE(X), TYPE(Y)
END SUB

2. Upgrade Assembler
Use of @name to get values that are not variables, like:
HWND (the current window hanlder) and addresses of exrternal functions:
Also I made a variant in Declare starement to make functions to call at address using a signature for parameters
The variant is like Declare .... Lib which accept only string, but Declare Code accept an address
See Asm4 new advanced exampled (introduce local variables)

Declare MessageBox Lib "user32.MessageBoxW" {long alfa, lptext$, lpcaption$, long type}
ASM_TEST = {    
start_code3:
    push dword 2 | push dword mCaption | push dword mText |  push dword @HWND
	Call @MessageBox
    ret    
    mText:          dw "HELLO THERE", 0
    mCaption:       dw "GEORGE", 0
start_code4:  ; C call then StdCall
	push dword [esp+16] | push dword [esp+16]
	push dword [esp+16] | push dword [esp+16] ; copy arguments
	Call @MessageBox
	ret
start_code5:  ; StdCall -> StdCall
	push dword [esp+16]	| push dword [esp+16]
	push dword [esp+16] | push dword [esp+16] ; copy arguments
	Call @MessageBox
	ret 16
}

Assembler=getobject("","m2000.x86")
function x86 (b as string, &outbuffer, useprep as boolean=false) {
	if Assembler=>assemble(b, true) then
		local OutPutSize=Assembler=>OutputSize
		buffer code outbuffer as byte*OutputSize
		Assembler=>BaseAddress = outbuffer(0)
		if Assembler=>assemble(b) then
			outbuffer=>FillDataFromMem Assembler=>GetOutPtr
		else
			error "x86 fault 2"
		end if
	else
		error "x86 fault 1"
	end if
}
var example1
call local x86(ASM_TEST, &example1)
Declare CallCode code c example1(0) As Long
Print CallCode()
' c call - by default ret value is long but here we place the type.
Declare MsgBox code c assembler=>labelptr("start_code4") {long alfa, lptext$, lpcaption$, long type} as long
Print MsgBox(hwnd, "This is the text", "This is the Caption", 2&)
wait 300
' stdcall
Declare MsgBox2 code assembler=>labelptr("start_code5") {long alfa, lptext$, lpcaption$, long type}
Print MsgBox2(hwnd, "This is the text", "This is the Caption", 2&)


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