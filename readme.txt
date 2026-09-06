M2000 Interpreter and Environment
Version 15 Revision 40

Athens, September 6, 2026

Upgrade Assembler:
1: Align 4 add noop to move address to modulo 4.
2: Call_ext is an immediate call (not relative). Used from (3)
3: Make a dll (so now we can make Exe, Dll, or internal code for using by M2000)

' example
' no use of Assembly() functio  because we have to prepare Assembler first
' these are the steps to make a dll using 3 exported functions and one inported
' also this use relocarion table

name$="TESTASM7.DLL"
Static Function MachineCode
Enum PESubsystem {
    Subsystem_GUI = 2
    Subsystem_CUI = 3
}
Assembler =getobject("","m2000.x86")
Assembler=>Subsystem=Subsystem_GUI
Assembler=>PEDll=true 
Assembler=>DLLName = name$
Assembler=>AddExport "AddNumbers", "MyFunction"
Assembler=>AddExport "Time", "MyTime"
' we use a space so this sortrd to first place - ordinal 1
Assembler=>AddExport " 01", "MyTick"

Buffer2export=MachineCode({

extern "winmm", timeGetTime

; =========================================================================
; STANDARD WIN32 DLL ENTRY POINT (DLLMAIN)
; =========================================================================
DllMain:
    push ebp
    mov ebp, esp
    mov eax, 1
    leave
    ret 0xC

; =========================================================================
; EXPORTED FUNCTION (MyFunction to Addnumbers)
; =========================================================================
align 4
MyFunction:
    push ebp
    mov ebp, esp
    
    mov eax, dword [ebp + 8]   ; First parameter (a)
    add eax, dword [ebp + 0xC]  ; Second parameter (b)
    leave
;    mov esp, ebp ; same as leave
;    pop ebp
    ret 8                ; Clean up 2 arguments (2 * 4 bytes = 8) and return
    	
; =========================================================================
; EXPORTED FUNCTION (MyTime to Time)
; =========================================================================
	align 4
MyTime:
	call_ext timeGetTime   ; new directive for absolute call
	; so the code can be moved and the timeGetTime works fine
	ret

; =========================================================================
; EXPORTED FUNCTION (MyTick to #1 - no name)
; =========================================================================
	align 4 ; new directive for alignment
MyTick:
	call SomethingElse
	ret
align 16
SomethingElse:	
	push ebp
	mov ebp, esp
	mov eax, [data1]
	inc eax
	mov [data1], eax
	leave
	ret
data1:
	dd 0x20304050
})
Print "File name: ";name$
Print "Assembler Output Size: "; assembler=>outputsize
open name$ for output as #f
	put #f, Buffer2export
close #f
Print  "Saved ok. length:"; filelen(name$)
check$=file.name.only$(name$)
declare add2 lib dir$+check$+".AddNumbers" {long a, long b} as long
declare mytime lib dir$+check$+".Time" as long
' using ordinal number:
declare mytick lib dir$+check$+".#1" as long

try ok {
	? add2(1022322,22340) = 1022322+22340,  mytime(), mytick()
	? add2(1022322,-22340) = 1022322-22340,  mytime(), mytick()
	? add2(-1022322,22340) = -1022322+22340,  mytime(), mytick()
	? add2(10222,-2240) = 10222-2240, mytime(), mytick()
}
if ok then remove dir$+check$ else print error$

' this is the function for preparing the two pass assembler.

Function MachineCode(assembly as string)
	if Assembler=>assemble(assembly, true) then
		local OutPutSize=Assembler=>OutputSize
		local mc
		buffer code mc as byte*OutputSize
		' feed the base address to Assembler
		Assembler=>BaseAddress=&h10000000 		
		if Assembler=>assemble(assembly) then
			' get a copy of final machine code
			mc=>FillDataFromMem Assembler=>GetOutPtr
			=mc
			exit function
		End if
	End if
	error Assembler=>LastErrorMessage 
End Function






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