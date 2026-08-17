M2000 Interpreter and Environment
Version 15 Revision 26

August 18, 2026,
1. Fix a small mistake from previous 2-3 revisions
a=(10,3i)
print a|r ' was ok
print a|r/3 ' had problem now fixed (revision 19 was ok)

2. New function Assembly()
This is from Help Assembly()

' using assembly(code$) we get the machince code for running from the first offset 
example1=Assembly({
    fild  dword [esp+4]    ; st0 = numerator
    fild  dword [esp+8]    ; st0 = divisor, st1 = numerator
    fdivp                  ; st1 = st1 / st0, pop st0
    mov eax, 0    
    ret 8
})
Declare Division Code example1(0) {long a, long b} as single
Print Division(34, 10)=3.4

' using assembly(code$, true) we get tuple, machinecode in a buffer and the object to get pointers from labels
(example1, Assembler)=Assembly({
    fild  dword [esp+4]    ; st0 = numerator
    fild  dword [esp+8]    ; st0 = divisor, st1 = numerator
    fdivp                  ; st1 = st1 / st0, pop st0
    mov eax, 0    
    ret 8
ASM_TEST_CPUID:
    ;mov eax, [esp+4]
    pushad ; 32 bytes
    xor eax, eax
    mov edi, [esp+36]
    xor eax, eax
    cpuid
    mov [edi+0], ebx
    mov [edi+4], edx
    mov [edi+8], ecx
    popad
    xor eax, eax
    ret 4
}, true)
Declare Division Code example1(0) {long a, long b} as single
Print Division(34, 10)=3.4
Dim Ret(12) as byte
addrPtr=Assembler=>LabelPtr("ASM_TEST_CPUID")
Hex "Call address of ASM_TEST_CPUID = ";addrPtr
Declare CPUID Code addrPtr {long ptrArrayItem}
call CPUID(VarPtr(Ret(0)))
' chr(number) return ansi string
for i=0 to len(Ret())-1
    Print chr(Ret(i));
next
Print
buffer clear retstring as byte*12
call CPUID(retstring(0))
' chr$(string_value) convert ANSI to UTF16LE
Print chr(retstring[0, 12])


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