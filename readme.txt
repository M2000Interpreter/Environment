M2000 Interpreter and Environment
Version 15 Revision 0

Change version to 15. The 14th version completed. Version 15 start with some new things, like the keys for lists and some string functions which have another name without the $ as prefix. See (2) 

1) Now we can use a=>property for list objects (we can use the With statement for accessing properties too). Adding a property to change the type of a key. Keys for lists may have type. See the example:
  a=list:="George":="M2000",2:="ok",3,4:=500
  Print a(2)="ok"
  if exist(a, 2) then Print "index:";eval(a!) ' should be 1
  Print eval(a, 1)=2, ' the key as number
  Print eval$(a)="ok", eval(a)="ok"  ' remeber last index
  m=each(a, 2, 2)
  while m
  	Print eval(m)="ok"
  	Print eval(m!)=1, m^=1
  	Print eval$(m)="ok"
  	Print eval$(m!)="2" ' always key as string
  	Print eval(a, eval(m!))=2
  	Print eval(a, m^)=2	
  end while
  m=each(a, -1, -1) ' last key only
  while m
  	Print eval(m)=500
  	Print eval(m!)=3, m^=3
  	Print eval$(m)="500"
  	Print eval$(m!)="4" ' always key as string
  	Print eval(a, eval(m!))=4
  	Print eval(a, m^)=4
  end while
  ' we have only key (no value)
  ' so value is key also
  m=each(a, 3, 3)
  while m
  	Print eval(m)=3 ' the key which was number
  	Print eval(m!)=2, m^=2
  	Print eval$(m)="3"
  	Print eval$(m!)="3" ' always key as string
  	Print eval(a, eval(m!))=3
  	Print eval(a, m^)=3
  end while
  ' first key is string and has a string value
  m=each(a, 1, 1)
  while m
  	Print eval(m)="M2000"
  	Print eval(m!)=0, m^=0
  	Print eval$(m)="M2000"
  	Print eval$(m!)="George" ' always key as string
  	Print eval(a, eval(m!)) = 0 ' no numeric value for key
  	Print eval(a, m^)	= 0 ' no numeric value for key
  end while
  ' copy as tuple for values
  Print Array(a)#Str$()="M2000 ok 3 500"
  ' copy as tuple for keys
  Print Array(a!)#Str$()="George 2 3 4"
  
  Delete a, "2"
  ' now place for 2 used by 4
  ' Deleting break the order of keys
  Print Array(a!)#Str$()="George 4 3"
  Sort ascending a as number
  Print Array(a!)#Str$()="3 4 George"
  if a=>ChangeKey("George", "35") then
  	if a=>ChangeKey("3", 1000) else exit
    if a=>ChangeKey(4, "200") else exit
  	Print Array(a!)#Str$(", ")="1000, 200, 35"
  	Sort ascending a as number
  	Print Array(a!)#Str$(", ")="35, 200, 1000"
  	Print Array(a)#Str$()="M2000 500 1000"	
  	a(1000)="Bingo"
  	Print Array(a)#Str$()="M2000 500 Bingo"
  	b=each(a)
  	while b
  		Print b^, eval(a, b^) ' print keys as  numeric values
  	end while
  	d=array(a!)
  	b=each(d)
  	while b
  		print array(b), type$((array(b))) ' first item is string not number
  	end while
  	if exist(a, 1000) then
  		? a=>keytypevalue=5 ' same as vbDouble = 5 
  	end if
  	if exist(a, 35) then
  		? a=>keytypevalue=8 ' same as vbString = 8
  		with a,"keytypevalue", 5	
  	end if	
  	d=array(a!)
  	b=each(d)
  	while b
  		print array(b), type$((array(b))) ' now all are double
  	end while
  end if
2) MID(), RIGHT(), LEFT() are valid functions (like MID$(), RIGHT$(), LEFT$()), and some more. They are not like in VB6 which the $ postfixed functions return strings and the non postfixed return variant. The value which came taking account as pure value and not variant. If we have a variant type variable we get the value and we have the option to store anoter type, but as a value is the same like a variable which have one type only if the type is the same. The only way the variant has signifant part, is when we pass it by reference in a method of a com object.

In 15th version I have to do more things, but for now I cut he expansion for the 14th and enter to 15th version. !4th version has no errors as I know.

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