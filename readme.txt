M2000 Interpreter and Environment
Version 15 Revision 19

July 31, 2026,

1. Octal numbers: using 0o or 0q or &o
' by default are signed long but we can use % for 16bit integer (no 64bit yet)
? 0o77&=63, 0q77&=63, &o77&=63
? 0o177766=65526, 0q177766=65526, &o177766=65526
? 0o177766%=-10, 0q177766%=-10, &o177766%=-10

This is not so simple because I had to fix the TextViewer.cls and the GuiEditBox.cls (which they have the colouring algorithm) to display the octal values with the same color as the signed hex values for long and integers.

2. Assembler x86 updated.

2.1 Preprocessor now are very fast (I put a Fastcollection to use hash function for founding the EQU replacements). Also I put a document object (as a string of strings) to fast collect lines.
Also now lines with EQU may have leading spaces (original code do not import a symbol and the equ value if the symbol not starting from position 1)

2.2 Now we can use variables outside from Assembly code and Assembler use them as symbols for geting values (like Equ but the labels are variables of M2000. here the dd find alfa and alfa return value 123454321. An EQU get all characters and just replace it in every same name (no case sensitive. Fastcollection bydefault is case sensitive but we can change it if it is empty). An external varianle has value not characters as the EQU case.

alfa=123454321
E_NOINTERFACE=0x80004002&
bb={
Mystery_one EQU [Data]
 	mov eax, E_NOINTERFACE
    ret 16
ASM_TEST_RAWDATA:
		mov eax, Mystery_one
		ret 16
Data:
		dd alfa
}

2.3 Octal numbers for the Assembler plus 0x among &H for hexadecimal numbers.
     Now 0x works as &H (signed long), plus 0q, 0o and &o for Octal numbers.
     Also underscore can be used as phantom characters for numbers so 0o77_7 is 0o777.

2.4 Reading relative and absolute values for labels after the compiling of x86 code.
	assembler=>labeloffset("start_code") return the offset for code buffer
	assembler=>labelptr("start_code") for getting the absolute address.
	*** the star_code is a label defined in assembly source text.

2.5 lock a code buffer for using it for calling from other code buffer (or a callback),
	so the first buffer may have functions for objects. 
 
 
Assembler =getobject("","m2000.x86")
function MachineCode(assembly, &mc as buffer) {
		if Assembler=>assemble(assembly, true) then
		' get the output size
		local OutPutSize=Assembler=>OutputSize
		buffer code mc as byte*OutputSize
		' feed the base address to Assembler
		Assembler=>BaseAddress=mc(0)
		if Assembler=>assemble(assembly) then
				' get a copy of final machine code
				mc=>FillDataFromMem Assembler=>GetOutPtr
				exit
		end if
	end if
	error Assembler=>LastErrorMessage 
}

S_OK =0x00000000&
E_NOINTERFACE = 0x80004002&
bb={	
start_code:	
	mov eax, E_NOINTERFACE
    ret 16
start_code2:
	mov eax, E_NOINTERFACE
	ret
}
cc={	
start_code2:	
	mov eax, 0
    Call start_code_ptr  
    ret 16
}
buffer code mycode as byte*0x1000
buffer code mycode2 as byte*0x1000

' we use call local so the function called:
' like a module and using this module namespace.
call local MachineCode(bb, &mycode)

start_code_offset=assembler=>labeloffset("start_code")
' the offset used for calling by Execute Code
' the call (using Execute Call) pass 4 parameters
' (if we do not provide the any or all of them intepreter pass zero 32bit for each missing)
' so we need a Ret 16 (drop 16 bytes, 4x 32bit numbers/pointers)

' we can't use this function using a call from another point
' - we have to push eax four times
' but we may have proper Ret from another label, like start_code2
start_code_ptr=assembler=>labelptr("start_code2")
call local MachineCode(cc, &mycode2) 
Hex "start_code_ptr=";start_code_ptr
' M2000 when execute machine code from a buffer
' set the buffer for execution 
' also by default we can't write on the block, only read & execute
' and at return of execution interpreter reset execution flag for buffer.
' so the specific buffer can't execute code. Not from another another block.
' When we call with Execution Code! (notice the !) we can read/write  the block
' so this is good for using space for variables. But we can use another buffer
' for reading and writting (a buffer only for data)

' So the idea now is to set flag for execution and read/write before execution
' magic begin here:
mycode=>Callback=true

Execute Code! mycode, start_code_offset;
Print Eax
' now we call from mycode2 a routine in mycode
Execute Code! mycode2, 0;
Print Eax
' magic end here
mycode=>Callback=false
' If we forgot to turn to false the property callback then
' the block didn't deleted at the exit of this module
' We can delete it by using Clear from manual mode
' Interpreter mark the buffer for callback ...
' and keep it alive using a secont pointer to buffer.




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