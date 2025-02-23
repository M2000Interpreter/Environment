M2000 Interpreter and Environment

Version 13 revision 14 active-X

1. Input statement now work as expected (I am working for the new manual)
1.1 Now we can read not only a value but a range of value using a group. The error make Input to erase the old input and wait for new one. The =0ub set the input to BYTE type (this is an old way for input to set the type for the input)

group a {
variant z
	set (x){
		if x<4 or x>10 then error "out of range"
		.z<=x
	}
	value {
		=.z
	}
}

input "3<x<11 = ", a=0ub



1.2 Also a new separator introduced using ;"/"; we press / and change field to next one.

integer M, D, Y
do
	try {
		input "(M/D/Y)=", M;"/";D;"/";Y
		date a=Y+"-"+M+"-"+D
	}
until valid(a) and a>1
Print a, type$(a)

2. type$() now work for expressions
group alpha {
	value {
		=100&&   ' this is Long Long (64bit)
	}
}
print type$(12*34@)="Decimal"
a=list:=1:="ok", 500:=10
' old append to list:
Append a, 200:=Group(alpha)
Return a, 1:=a(1)+".."
' Version 13 add keys like this (if "hello" not exist)
a("hello")=100&
print type$((list:=1,2,3,4,5))="Inventory"
print type$(a)="Inventory"
print type$(a, 1)="String"  ' old
print type$(a(1))="String" ' new
print type$(a("hello"))="Long" ' new
print type$(a, 200)="Group"
print type$(eval(a(200)))="Long Long"
print type$(12*34@)="Decimal"
print Len(a(1))=4, a(1)="ok.."
a(500)+=1000
print a(500)=1010, type$(a(500))="Double"
print a(200)=400, type$(eval(a(200)))="Long Long"
print type$((a(200)))="Long Long"


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

https://rosettacode.org/wiki/Category:M2000_Interpreter (534 tasks)
                 