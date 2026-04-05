M2000 Interpreter and Environment
Version 14 revision 14 active-X

1) BUG removed: In Trace code with Test form, you place the STOP statement in the Bottom textbox, and the execution move to console like cli, but in the running code, and you close the Test Form, at the exit the form not exist and by a mistake the system return to trace without the form, and this return the error "Object variable or With Block  variable not set" which is an error from VB6. This error not hang the M2000 Interpreter. Fixed.
2) BUG removed: In print if we have some wrong encoding text (from binary data) happen sometime to find a wrong character which cannot pring (no data to handle the print) and stop advance to next character. Fixed
3) Using
	str$(wide_string_encoding) return narrow_string_encoding based on locale
		and 
	chr$(narrow_string_encoding) return wide_string_encoding based on locale
   
sometime think that these can be used for binary date, to get the ASCII equivalent, and letter to cut the string as a part and then return that part to binary data from wide to narrow bytes. But this isn't true. The chr$() get the binary data, for example 10 bytes and produce 20 bytes with 10 characters mapping the locale. So for some numbers the mapping return the same number. The problem is with the other numbers which turn to something else. If we make the data using the STR$() with same locale we have no problem. although the actual converted to wide data would be not the same with the pure data which have the same number as wide (2 bytes) and as narrow (1 byte) encoding.
	So I write two encoders, one WIDE2BYTES and another BYTES2WIDE for using with STRING$() function which have many decoders/encoders.

This is an example of using BYTES2WIDE in STRING$() instead of CHR$() an prove why we get wrong data using the later:

buffer b as byte*256
for i=0 to 255:b[i]=i:next
locale 1032
z=chr$(eval$(b))
faults=0
for i=0 to 255
	if b[i]<>chrcode(mid$(z,i+1,1)) then
	 	// Print "Error "+i, b[i], chrcode(mid$(z,i+1,1)), mid$(z,i+1,1)
	 	faults++
	 end if
next i
Print "faults = ";faults  // 93 for locale 1032

locale 1033
z=chr$(eval$(b))
faults=0
for i=0 to 255
	if b[i]<>chrcode(mid$(z,i+1,1)) then
	 	// Print "Error "+i, b[i], chrcode(mid$(z,i+1,1)), mid$(z,i+1,1)
	 	faults++
	 end if
next i
Print "faults = ";faults  // 27 for locale 1033

z=String$(eval$(b) as BYTES2WIDE)   // and reverse WIDE2BYTES
faults=0
for i=0 to 255
	if b[i]<>chrcode(mid$(z,i+1,1)) then
	 	faults++
	 end if
next i
Print "faults = ";faults   /// 0 faults
GoodWideData=Z
// This is the real problem:
// This return True, because each function of str$() and chr$() use the same locale.
Print Eval$(b)=str$(chr$(Eval$(b)))
// But the intermediate result, the WideData are not the same as the GoodWideData:
Print GoodWideData = chr$(Eval$(b))   ' false
// so if we get binary data and want to expnad them to Wide char for farher processing
// we have to use the new decoder/encoder of String$() statement




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

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 