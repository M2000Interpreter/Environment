M2000 Interpreter and Environment
Version 15 Revision 12

July 24, 2026,

1. Fix Sort for lists for keys with multiple parts (was a mistake from the day tuple split from mArray as different object)

2. A #fold() for arrays (now for lists too) can stop returning one value from the first search using =value in the lambda function. We use lambda function passing from stack of values parameters and get response to stack of values except when we want to abandon the search we return the value like normal function - the code behind watch the type of the return value, if it is not Empty the change plan and get this as the value to return. See also the example in (3)
	a=((1,2,3), (1,2), (1,2,3,4), (1,2))
	foldSum7=lambda (a)->{if a#sum()>6 then =a}
	Print a#fold(foldSum7,(,))#sum()=10

3. A list now get the tuple/mArray second interface as iboxarray, which we can do this:
	a=list:=1,2:=1000, 3, 5, 44:=155, 12
	? a#sort()#str$(", ")
	? a#sort(-1)#str$(", ")
	? a#str$(", ")
	? a#slice(2, 4)#str$(", ")
	? a#slice(2, 4)#rev()#str$(", ")
	? a#mat("+=", 100)#str$(", ")
	? a#start((1,2,3,4))#str$(", ")
	b=(1,2,3)
	z=a#expanse(10,b)
	 ' if b=list:=1,2,3 we get true but now we get false
	? z#val(9) is z#val(8)

	clear
	a=list:="d","b":="zoro","f","all","c"
	firstZ=lambda (a)->{if lcase(left(a,1))="z" then =a}
	' search keys 
	print array(a!)#fold(firstZ,"not found")="not found"
	' search values old made a copy of values in a tuple 
	print array(a)#fold(firstZ,"not found")="zoro"
	print "["+array(a!)#str$()+"]"="[d b f all c]"
	' search values new search without copy list to tuple
	print a#fold(firstZ,"not found")="zoro"
	print "["+a#str$()+"]"="[d zoro f all c]"

	As you see the "old code" was to use array(a) to get the tuple which copy the values of list (place key if no value exist for specific key). Now we use list as it is a tuple (only when we combine to it the serial function #something()), so this happen without copies of values. For keys we have to use array(a!)
	
4. Array(a, 0) is the same as a(0!) when a is a list and return the value from that index

5. #Expanse() now has optional then starting value:
	? (1,2,3)#Expanse(5, 1)#Str$(", ")="1, 2, 3, 1, 1"
	This make the tuple 5 items wide and place value 1 for the new items

	
  
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


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 