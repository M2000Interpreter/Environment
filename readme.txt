M2000 Interpreter and Environment

Version 13 revision 10 active-X
This is the real 13 version.
1. The array (1,2,3,4) now has type tuple, not mArray.The reason for the change: mArray hold arrays for dimensions, so tuple is an one dimension (or none) array. 

Tuple has the same functionality for operators: a=(1,2,3): a++ ' add one to each item.
Also Tule use #functions. The Sort function works slighty different for tuple: The bigInteger is traited as string value (mArray show it as an object, and abandon the #sort() command). Tuple has no columns sort (only mArray at 2 dimensions).
The Car(), Cdr() and Cons() return tuple
A tuple class is lighter than the mArray class.

a=(1,2,3,4)
Print type$(a)="tuple"
link a to a()  ' a and a() point to same data site
Print type$(a())="tuple"
dim b()
b()=a  ' this is a copy from tuple to mArray
Print type$(b())="mArray"
b(2)+=100
' the a() interface works for copy | the a=b() just alter the pointer.
a()=b()  ' this is a copy from mArray to tuple
Print type$(a())="tuple"
Print a(2)=103
Print a#val(2)=103   ' a is a pointer to tuple
b=b()
Print type$(b)="mArray"
a=b  ' this can be done
print a is b  ' true
print type$(a)="mArray"
Print type$(a())="mArray"
k=car(b())
print type$(k)="tuple"
k1=cdr(b())
print type$(k1)="tuple"
k2=cons(b(), (1,2,3,4), (5,6,7))
print type$(k2)="tuple"
k2=cons((1,2,3,4),b(), (5,6,7))
print type$(k2)="tuple"
k3=b()#slice(0, 1)
print type$(k3)="mArray"
k4=b()#rev()
print type$(k4)="mArray"


2. Two more functions for tuple and mArray: The #mat() perform operators on array items (for numeric values, including biginteger). For the example, tuple a has five items, we add 100 for each and we place the values in a string. The original's tuple values not changed.

a=(1,2,3,4,5)
? a#mat("+=", 100)#str$()="101 102 103 104 105"

The #expanse() redim the array (not the original, but the product). So in the last line the tuple "a" has zero items, then we get a copy and we give an expanse of 100 items, and then we assign a value 1 for each, and last we get the sum of items.
a=(,)
? a#expanse(100)#mat("=", 1)#sum()=100


3. Lists. Before this revision used the statement Return obj, key:=value [, key1:=value1]
Now added a new one (so it is like we use sparse arrays):
a=list
// append to list using this form:
a(10)=30
a(3)=5
a("hello")="Hello"
a(3)++
a("hello")+=" World"
print  a(3)=6
print a("hello")="Hello World"
if exist(a, "hello") then print eval$(a)="Hello World"
Print eval(a, 0)=10  ' first index is 0, we get the numeric key
Print eval$(a, 2)="hello" ' we get the string key
// export keys
Print array(a!)#str$(", ")="10, 3, hello"
// export values
Print array(a)#str$(", ")="30, 6, Hello World"
// sort on tuple
Print array(a)#sort()#str$(", ")="6, 30, Hello World"
// sort on list object
sort a as number  ' by default is: as text
Print array(a)#str$(", ")="6, 30, Hello World"


4. Structures fixed - now works as expected.
structure alfa {
           structure beta {
                  low as integer
                  high as integer
            }  
            structure delta {
                  low as byte
                  middle1 as byte
                  middle2 as byte
                  high as byte
            }   
            value as long
}
Print Len(alfa)=4  ' we have a union of three parts: beta, delta, value
buffer clear alfa1 as alfa*10, alfa2 as alfa
Return alfa1,0!delta.high:=0xFF  ' old way
Print alfa1[0]|delta.high=0xFF ' True
Print alfa1[0]|beta.high=0xFF00 ' True
Print alfa1[0]|value=0xFF000000 ' True
alfa1[0]|delta.high=alfa1[0]|delta.high-1
Print alfa1[0]|value=0xFE000000 ' True
alfa2[0]|value=0xFFFFAAAA
alfa2[0]|delta.low=alfa1[0]|delta.high
Print alfa2[0]|value=0xFFFFAAFE ' True




George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement Settings to change font/language/colors and size of console letters.

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (384 tasks)

ExportM2000 all files with executables (you can get the ca.crt):
https://drive.google.com/drive/folders/1IbYgPtwaWpWC5pXLRqEaTaSoky37iK16

only source, with old revisions and a wiki, for executables see releases
https://github.com/M2000Interpreter/Environment

M2000language.exe (Chrome can't scan, say it is a virus - heuristic choice)
All exe/dll files are signed
https://github.com/M2000Interpreter/Environment/releases

M2000 paper (305 pages). Included in M2000language.exe
M2000 Greek Small Manual (488 pages). Included in M2000language.exe

                                                             