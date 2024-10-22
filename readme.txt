M2000 Interpreter and Environment

Version 12 Revision 39 active-X

1) a=(1,2,3): ? LEN(a#slice(1,0))=0 ' NOT RAISE ERROR, RETURN EMPTY ARRAY
? LEN(a#slice(2,1))=0 ' RAISE ERROR - AS PREVIOUS
? LEN(a#slice(3,))=0 ' NOT RAISE ERROR, RETURN EMPTY ARRAY
2) slice from fold function. Now fold function may return objects, and arrays too.
a=(1,2,"aaa","bbb",5,6,7)
fold1=lambda (skip, many) -> {
	=lambda t=(,), m=0, skip, many -> {
		if m=0 then shift 2: drop: m=1: data t
		read q, k as array
		if skip<=0 then
			if many>0 then append k, (q,)
			data k
			many--
		else
			data t
		end if
		skip--
	}
}
? a#fold(fold1(2, 5))#str$(", ")
? a#fold(fold1(1, 3))#str$(", ")
3)Map function exit when we want just flush the stack of values.
List1=(1,2,3,4,5,6,7,8,9)
List2=(10,11)' ,12,13,14,15,16,17,18)
List3=(19,20,21,22,23,24,25,26,27)
map1=lambda (list0 as array) ->{
	dim a()
	a()=list0  // we get a copy
	=lambda a(), n=-1 ->{
		if n<len(a())-1 then
			n++
			if islet then		
				push letter$+(a(n))
			else
				push number+""+(a(n))
			end if
		else
			flush  ' if not return value the map function exit ' addition from 39
		end if
	}
}
? "["+list1#map(map1(list2), map1(list3))#str$(", ")+"]"
map2=map1(list2)
map3=map1(list3)
list2=()
list3=()
// map2 and map3 has a copy of list2, list3 before we erase it
? "["+list1#map(map2, map3)#str$(", ")+"]"

4)Map allow now to make 1 to many maps.
dim a(3) as long
a(0):=1,2,3
map1=lambda ->{
	push random(10)+number, 100 ' two values for one
}
? a()#map(map1)#str$(", ")
// 100, 8, 100, 3, 100, 4

5)Fix Sum/Min/Max/Min$/Max$ for tuple using Enum values. Sort not work with enum values (they are objects, which know the name of value, the value (numeric as  number and sign or a string value), also know the order, so a variable of some enum type can advance to next using ++ operator or go back using -- operator. This example use Sort after a map function, which map the enum values only to simple values (without the object). Also a Map function can get more than one lambda function to make more maps, one after the other.

enum m {a=50, b=200, c=30}
k=(a, b, c)
? k#max(),  k#min(), k#val(0), k#Sum()
z=each(k)
strip = lambda->{
		read t as long:data t
}
strip1 = lambda->{
		read t as long:data t+100
}
k=k#map(strip,strip1)#sort()
? k#str$(", ")


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

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

                                                             