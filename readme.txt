M2000 Interpreter and Environment

Version 12 Revision 63 active-X
1. We can define user objects from classes using the compact format:
nameofclass var1, var2
or the standard format:
var var1=nameofclass(), var2=nameofclass()

if we use constructor then this is the compact:
nameofclass var1(1,2,3), var2(1,2,3)
... and this is the standard format:
var var1=nameofclass(1,2,3), var2=nameofclass(1,2,3)

Classes inside classes are local to class. Classes outside class definition is global.

Look this example. We have a global class Alfa and a private class Alfa inside class Something. We can make a class Alfa from the global one using the proper "global" name from modules (members of class).
If we change the .a<=.alfa(20, 30) with .a<=alfa(20, 30, 40)  we get the global class. So Beta has a with a.x, a.y and a.z members.

So our class can use the private alfa without problems from global classes which some previous module make it. The same hold for public alfa in class Something.

So the M2000 Interpreter allow objects to have same "type" but not same members. The scope distinguish each other.

Previous revisions use the compact form inside classes only.

if true then {
	class alfa {
		x, y, z
	class:
		module alfa (.x,.y, .z) {
		}
	}
	alfa a(20, 30, 50)
	Print a.x, a.y, a.z
}

Class Something {
private:
	class alfa {
		x, y
	class:
		module alfa (.x,.y) {
		}
	}
	alfa d(1, 2)  ' is the local one
public:
	group a
	module dothat {
		// use the global alfa
		try ok {
		alfa a(20, 30, 50)
		Print a.x, a.y, a.z
		? "ok........................."
		}
		if ok and not error then list else print error$
	}	
	module doit {
		// use the private alfa
		.alfa a(20, 30)
		Print a.x, a.y, valid(a.z)=false
		// list
	}
class:
	module Something {
		.a<=.alfa(20, 30)
	}
}
Something Beta
Beta.doit
Beta.dothat
Print beta.a.x, beta.a.y, valid(beta.a.z)=false
delta=beta  ' delta get a copy of beta
Print delta.a.x, delta.a.y, valid(delta.a.z)=false
list

2. Structures (work on progress)

I would like to include the notation a|x for the equivalent eval(a,0!x) although the later cast to other types.
Structures used for making buffers and using pointers for passing them to external functions. So a(4, y) is the real memory address of the 5th alfa item at offset y (which is 8). We can pass a number using Return a, 4!y:=-23.34e32, 3!k!50:=255 ' max value for byte, if we place a bigger no error raised, only one byte used, the lower one. 

structure alfa {
	x as double
	y as double*5
	k as byte*100
	z as long	
}
// buffer clear a as alfa*10  same as alfa a[10]
alfa a[10], b[100]
Print a(4, y)-a(4,x) = 8 ' 8 bytes for a double value
for i=0 to 99 
	return a, 4!k!i:=i+1
next
print a[4]|k[1], " address="; a(4, k, 1)  ' value, memory address
print a[4]|k[2], " address="; a(4, k, 2)  ' value, memory address
print a[4]|k[15]
byte a4k2= a[4]|k[2], a4k15= a[4]|k[15]
return a, 4!k!2:=a4k15, 4!k!15:=a4k2
print a[4]|k[2]
print a[4]|k[15]
k=a[4]  // get a copy of a[4] to k
return a, 0:=eval$(k)
print len(k)
print k|k[2]
print a[0]|k[2]
list

3. Enum now can be local/global and local shadowing local enum (in a sub).  Also Enum alfa here has a part beta, a previous defined enum. So we can define global, local (without using Local) and Local new (using Local statement).


global enum beta {
	kkk=2313
	aaa=100123123
	bbb=123213
}	
global enum alfa {
	beta  ' add members of beta to alfa
	epsilon=123123
}
module Inner {
	enum beta {
		kkk
		aaa=100
		bbb
	}	
	enum alfa {
		beta  ' add members of beta to alfa
		epsilon
	}
	kappa()
	kappa(kkk)
	other()
	list
	sub kappa(m as alfa=aaa)
		Print m
	end sub
	sub other()
	local enum beta {
		kkk=33
		aaa=100232
		bbb=21
	}	
	local enum alfa {
		beta  ' add members of beta to alfa
		epsilon
	}	
	kappa()
	kappa(kkk)
	end sub	
}
Inner



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

                                                             