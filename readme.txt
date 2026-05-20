M2000 Interpreter and Environment
Version 14 Revision 34

Advanced handling of buffer arrays. Now we can use object arrays to hold buffer arrays and address individual fields of buffer.
1) Example Simple
structure simple {
	x as double
	y as double
	{ StrPtr as long} ' we make a union (4 bytes as long and string)
	name1 as string
}
simple one[5], zero
zero|name1="This is a BSTR string"
Print len(Simple)=20 ' 20 bytes, 2x8 bytes double, 4 bytes pointer to string
one[3]=zero ' copy to one at index 3 (4th item)
Print len(zero)=20, len(one)=20*5  ' 20 bytes and 100bytes
print Len(one[3]|name1)=21 ' characters = 21 words = 21*2 bytes
print one[3]|StrPtr  ' this is the address of the string
print one[0]|StrPtr=0  ' this is null string
print Len(one[0]|name1)=0
' now start the real example
Object alfa[3], beta[3][2]  ' max two index holder - but as you see we can put anything inside
alfa[1]=one
print alfa[1][3]|name1="This is a BSTR string"
print alfa[1] is one, " : True"  ' we have the same object
alfa[2]=buffer(one)  ' we get a copy of all items
alfa[3]=one[3] ' using one[3] we get a copy of item 3
print alfa[2] is one, " : False"  ' we have different objects
print alfa[2][3]|name1
print alfa[1][3]|StrPtr=one[3]|StrPtr, " : True"
print alfa[2][3]|StrPtr=one[3]|StrPtr, " : False"
print len(alfa[3])=20  ' one only item
print alfa[3]|name1="This is a BSTR string"
print alfa[3]|StrPtr=one[3]|StrPtr, " : False"
beta[3][0]=one
print beta[3][0] is alfa[1], " : True"
beta[3][0][3]|name1="New string"
print alfa[1][3]|name1="New string", " : True"
print one[3]|name1="New string", " : True"
' we can put an object in an array of objects too
beta[3][2]=alfa
print beta[3][2][1][3]|name1="New string", " : True"
zero=beta[3][2][1][3]  'we get the index 3 of object at beta[3][2][1]
' zero is a copy 
print zero|name1="New string", " : True"
print zero|StrPtr<>beta[3][2][1][3]|StrPtr, " : True"
zero=beta[3][2][1]
zero[3]|name1="last string"
' zero now point to beta[3][2][1]
print zero is one, " : True"
print zero is alfa[1], " : True"
print zero is beta[3][2][1], " : True"
print zero is beta[3][0], " : True"
print beta[3][0][3]|name1="last string", " : True"



2) Example Hard
structure simple {
	x as double
	y as double
	{StrPtr as long} ' we make a union
	name1 as string
}
structure bigone {
	{simple};   ' import simple (using ; we bypass the union mechanism)
	z as double * 100
}
structure bigtwo {
	s as simple * 4
	z as double * 10
}
bigtwo two[10]
bigone one[5]
simple zero
zero|name1="This is a BSTR string"
print len(Simple)=20 ' 20 bytes, 2x8 bytes double, 4 bytes pointer to string
one[3]=zero 
print len(zero)=20, len(one)=(20+100*8)*5  ' 20 bytes and 100bytes
print one[3]|name1=zero|name1, one[3]|StrPtr<>zero|StrPtr 
print Len(two)=(4*20+10*8)*10
two[2]|s[1]=zero
print two[2]|s[1]|name1=zero|name1, two[2]|s[1]|StrPtr<>zero|StrPtr 
two[2]|s[1]|name1="new string"
print two[2]|s[1]|name1=zero|name1, " : False"
zero=two[2]|s[1]  ' get a copy
print two[2]|s[1]|name1=zero|name1, two[2]|s[1]|StrPtr<>zero|StrPtr 
' now the real example
Object alfa[4], beta[3][2]
alfa[2]=two
print alfa[2][2]|s[1]|name1=zero|name1
beta[3][1]=one
print beta[3][1][3]|name1="This is a BSTR string"
beta[3][1][3]=zero  ' copy because ...[3] is index to buffer at beta[3][1]
print type$(beta[3][1])="Buffer"
print beta[3][1][3]|name1=zero|name1
print_name1(beta[3][1][3])
print_name2(zero) ' zero is object pass pointer
print zero|name1, zero|StrPtr ' we get new values.
print alfa[2][2]|s[1]|name1, alfa[2][2]|s[1]|StrPtr,  " values before call"
print_name2(alfa[2][2]|s[1])  ' pass copy
print alfa[2][2]|s[1]|name1, alfa[2][2]|s[1]|StrPtr,  " same values as before call"
print_name3(alfa[2],2,1) ' alfa[2] is object - pass pointer
print alfa[2][2]|s[1]|name1, "change name"
sub print_name1(a as bigone)
	print a|name1, a|StrPtr, a|z[30], " inside print_name1()"
end sub
sub print_name2(a as simple)
	print a|name1, a|StrPtr, " inside print_name2()"
	a|name1=a|name1+"...ok...."
end sub
sub print_name3(a as bigtwo, index, n)
	print a[index]|s[n]|name1, a[index]|s[n]|StrPtr, " inside print_name3()"
	a[index]|s[n]|name1=a[index]|s[n]|name1+"...ok"
end sub




George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and then open it again.

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