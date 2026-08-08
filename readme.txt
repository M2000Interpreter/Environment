M2000 Interpreter and Environment
Version 15 Revision 20

August 8, 2026,

1. Upgrade Assembly: Asm2 example now create exe with a form with a Unicode title. The BW directive place the Unicode string. Also now the external M2000 variables which previous used for numbers now works for string values too (see the Asm2 in new Info file).

2. Advance work for COM objects.
	2.1 we can pass by value or by reference arrays with parenthesis (mArray objects).
		this works by sending by reference the underline array, if we choose the reference pass (using & before name)
	2.2 we can get arrays like a long() as return value, and we can use it as parameter.
		we can convert to mArray object using Array() function:
		2.3 using array(variant_holding_array) we get a copy of array
		2.4 using array(&variant_holding_array) we get the array leaving empty array in variant_holding_array, without copy data.
	2.3 Now a biginteger object treated as non object (the same as Evaluator does, except when we have a name and place the => operator - for accessing properties and methos. We cant use this for biginteger literals)
	2.4. Now enumerations do not passed as objects (which are), but as values. Some revisions/versions they have default value the Value property but the last revisions do not have default property. So Older examples - using interpreter with Enumerations with default value works like the new one. 
	2.5 functions/properties which we call using object=>function( ) or  object=>property( ) may use by reference and call using name of parameter. For example if object has a function function1 with Argame1 and Argname2 arguments then we can do this:
	retvalue= object=>function1( Argname1:=ThsValueByVal, Argname2:=&ThisValueByRef)
	And for altering property1 (lets say that have the same arguments signature)
	object=>property1(Argname1:=ThsValueByVal, Argname2:=&ThisValueByRef)=NewValue
	
This works using the old way:
	Method object, "function1",  Argname1:=ThsValueByVal, Argname2:=&ThisValueByRef as RetValue.
	for properties there was no old way to use named arguments and by reference arguments for indexes. So the new way solve this.
	
	2.6 There is a lot of work for error control, so now we get feedback. Shapeex before the new error system, use a non exist property...(doing nothing but without a message...). So maybe older programs may have errors which the new interpreter may found it. For forms where an instance may delete from the user and some thread attempt to use it now we get error, so a Try { } block can recover from the error, until you master it and understand what cause it.

3. Upgrade sructures. We can use operators ++, --, +=, -=, *=, /= for numeric types, and += for BSTR type of strings.

4. Modules FUNCTOR and VBCOL2 now works as expected. Asm2 for Unicode form. Excel2 now get array from Excel, change items using the array and write back with one statement. See also struct3, struct4, jscript and mEditor (updated).

5. I did a little work for UDT with arrays.


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