M2000 Interpreter and Environment
Version 13 revision 51 active-X

We can make arrays of objects (of type RefArray) using for object of type aa: aa k[2] or aa k[2][20] or we can make more dimensions making an object array which get for each element another array. 
Also we can do that inside another object, type beta, which have private the class aa, and private k (as pointer to refarray) or at the second example the array k() (see parenthesis, is an mArray type), here we have from 0 to 10 dimensions.

1. Using RefArray type of array (using [ ])
class beta {
private:
	class aa {
		x=10&
	}
	aa k[2][20]
public:	
	module that {
		.k[2][3].x++
		? .k[2][3].x,  "k[2][3].x"
		.k[2][7].x*=100
		? .k[2][7].x
		for i=0 to len(.k)-1
			print i, .k[2][i].x
		next
		list
	}
}
beta z[4]   ' 5 times same object
' so we change the last 4 with new one
' so we get new k[][] array (is a pointer type array)
for i=1 to 4:z[i]=beta():next
z[1].that
Z[2].that
Z[2].that

2. Using Array of objects (of type mArray)
global counter=0
class beta {
private:
	class aa {
		a=0&
		x=10&
	class:
		module aa {
			counter++
			.a<=counter
		}
	}
	dim k() as object
public:	
	module that {
		.k(2,3).x++
		? .k(2,3).x , "k(2,3).x", .k(2,3).a
		.k(2,7).x*=100
		? .k(2,7).x
		for i=0 to dimension(.k(),1)-1
			print i, .k(2,i).x
		next
		list
	}
class:
	module beta {
		dim .k(0 to 2, 0 to 20)<<.aa()
	}
}
rem {   ' uncomment this block, press enter after rem 
	beta z[4]   ' 5 times same object
	rem : 
	for i=1 to 4:z[i]=beta():next   ' comment this line to see the difference.
	? counter
	z[1].that
	Z[2].that
	Z[2].that
	break
}

dim z(0 to 4)<<beta()
? counter
z(1).that
Z(2).that
Z(2).that


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

https://rosettacode.org/wiki/Category:M2000_Interpreter (544 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 