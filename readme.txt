M2000 Interpreter and Environment
Version 15 Revision 6

1. Fix a mistake from revision 28 version 14 for structures using single/currency
2. Calling function using call FunctionName() now delete any returned object.
3. Added an x86 Assembler (a work of Arne Elster). This is the first try. Works
This is the module ASM86 (added in INFO file). We use the m2000.x86 object which is the assembler. First pass find the size and the second pass relocate to  as specific base (we provide it). We get a pointer from actual byte array and fill a M2000 buffer which if for code execution.

Now we can execute the code with the advanced Execute which take: the buffer object, the offset to call (0 for the first byte) plus (this is the new) 1 to 4 parameters of type Long (by value). Also there are two symbols: Execute Code ! ... (symbol ! make the buffer Execute_ReadWrite - by default is Execute_Read only - you have to write in another buffer which isn't for code execution). Another symbol is the last ; which bypass the default error checking - if Eax is no zero raise error with the number of Eax (as signed 32bit). So bypassing the error system we can get the value of eax from a new EAX read only value.

The examples are as the same from the Arne Elster work, only here are assembled all together. I put a Fastcollection inside x86 class, as a hash table for searching the labels. Also labels are compared upper case, so "Data" and "data" are the same. 



' x86.cls as  32 Bit X86 Assembler
' from Arne Elster 2007 / 2008
' I add hash table for find labels for now
' this is the first work

Assembler =getobject("","m2000.x86")

MachineCode= lambda Assembler (assembly)-> {
		if Assembler=>assemble(assembly, true) then
		' get the output size
		OutPutSize=Assembler=>OutputSize
		buffer code mc as byte*OutputSize
		' feed the base address to Assembler
		Assembler=>BaseAddress=mc(0)
		if Assembler=>assemble(assembly) then
				' get a copy of final machine code
				mc=>FillDataFromMem Assembler=>GetOutPtr
				=mc
				exit
		end if
	end if
	error Assembler=>LastErrorMessage 
}
Example=MachineCode({
ASM_TEST_RAWDATA:
		mov eax, [Data]
		ret 16
Data:
		dd 123454321
ASM_TEST_BUBBLESORT: 			; this is another program
		pushad
		mov esi, [ebp+16]			; Arraylength
outer_loop:
		mov ebx, [ebp+12]		; ArrPtr
		mov edx, [ebp+16]		; Arraylength
		xor edi, edi
inner_loop:
		mov eax, [ebx+0]			; arr(j)
		mov ecx, [ebx+4]			; arr(j+1)
		cmp eax, ecx
		jle byte next_loop			; swap if eax > ecx
		mov [ebx+0], ecx			; swap arr(j), arr(j+1)
		mov [ebx+4], eax
		mov edi, 1				; swapped
next_loop:
		add ebx, 4
		dec edx
		jnz byte inner_loop		; i > 0 => still in inner
		test edi, edi				; swapped?
		jz  byte return				; no => sorted
		dec esi
		jnz byte outer_loop
return:
		popad
		ret &H10
ASM_TEST_CPUID:	
		pushad
		mov edi, [ebp+12]
		xor eax, eax
		cpuid
		mov [edi+0], ebx
		mov [edi+4], edx
		mov [edi+8], ecx
		popad
		ret 16
ASM_TEST_FDIV:
		mov   eax, [ebp+20]		; Ptr to output float
		fild  dword [ebp+12]		; st0 = numerator
		fild  dword [ebp+16]		; st0 = divisor, st1 = numerator
		fdivp					; st1 = st1 / st0, pop st0
		fstp  float [eax]			; pop st0 to output float
		ret    16 
})
ASM_TEST_RAWDATA=assembler=>labeloffset("ASM_TEST_RAWDATA")
ALTERDATA=assembler=>labeloffset("data")
ASM_TEST_BUBBLESORT=assembler=>labeloffset("ASM_TEST_BUBBLESORT")
ASM_TEST_CPUID=assembler=>labeloffset("ASM_TEST_CPUID")
ASM_TEST_FDIV=assembler=>labeloffset("ASM_TEST_FDIV")
' a1: Numerator
' a2: Divisor
' a3: Ptr to result (float = single)

' **NEW** call using Code! used for ExecuteReadWrite
' call using Code used for ExecuteRead (no write bytes to buffer Example)
' **NEW** use ; to bypass error from eax<>0 - and use EAX to read the value
Execute Code! Example, ASM_TEST_RAWDATA;
Print Eax
Return Example, ALTERDATA:=uint(-11112222) as long '(unsigned)
Execute Code! Example, ASM_TEST_RAWDATA; 
' eax has value of eax from last execut code when we use the ; symbol
' eax is signed value
Print EAX=-11112222

buffer Bytes12 as byte*12
'**NEW** we can pass max 4 byvalue long values.
Execute Code! Example, ASM_TEST_CPUID, Bytes12(0);
HEX "CPUID:";chr$(Bytes12[0,12])
long lngDividend=2, lngDivisor=5
single sngQuotient
' here we pass by reference sngQuotient passing by value the Variable Pointer
Execute Code! Example, ASM_TEST_FDIV,lngDividend, lngDivisor, VarPtr(sngQuotient);
Print sngQuotient=0.4~
N=100 ' try 1000
buffer MyData as Long*n
for i=0 to N-1
	MyData[i]=Uint(random(1, 100000)-49999)
	print Sint(MyData[i]),
next
print
Execute Code! Example, ASM_TEST_BUBBLESORT, MyData(0), N-1;
for i=0 to N-1
	print sint(MyData[i]),
next
print	




  
George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
THEN press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 