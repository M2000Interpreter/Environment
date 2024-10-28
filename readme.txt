M2000 Interpreter and Environment

Version 12 Revision 41 active-X

1) Fix the Point() internal function.
Cls #005580, 0: move 0, 6000
Pen 14 // yellow
Cursor !  // get the pos, row for characters from graphic cursor
Scroll Split Row 
a$=" "
Move 0,0: Copy 1000, 1000 to a$: Print Point(a$, 0, 0)=#005580
// Point return color from graphic cursor
Print Point(a$, 0, 0)=Point,  Point=#005580
for i=0 to Image.x.pixels(a$)-1: old=Point(a$, i, i, #FF5580):next
Print old=Point, old=#005580,  Point(a$, 0, 0)=#ff5580
Move 0,0 : image a$, 6000  // make the image X6
move 3000, 1000: image a$  // original size

2) Fix a rare situation, when the class has same name as an internal function. also the class has the module with the same name as constructor, and we call them from an inner module (if we change the name point() to point1() no error happened). A class definition is global (we can use it anywhere, as long as the module where the definition executed not exit yet. Objects defined from class definition no need the definition, so we can return an object from  a module/function which we have the object definition.

class point {
	X as integer, Y as integer
class:
	module point (a as integer, b as integer) {.X<=a:.Y<=b}
}
module inner {
	z1=point(10 ,20)
	print z1.X, z1.Y
	try {
		z2=point(30, 50)
		print z2.X, z2.Y
	}
}
Call inner 

3)The #fold() special function for tuple/arrays, can get an object as starting value. This not apply to #fold$() which used for returning string (and we can pass a string as starting value). The example use tuple as zero length as (,), the stack, the list and the queue. 

map=lambda (k, m as array)-> {
	append m, (k^2,): push m
}
? (1,2,3,4,5,6)#fold(map, (,))
? (1,2,3,4,5,6)#fold(map, (,))#fold(map, (,))
? (4,5,6)#fold(map, (1,2,3)#fold(map, (,)))#fold(map, (,))
map2stack=lambda -> {
	read k, m as stack   // we can put the read statement too
	stack m {data k^2}: push m
}
? (1,2,3,4,5,6)#fold(map2stack, stack)   // no #function for stack
map2list=lambda (k, m as list) ->{
	//using (k, m as list) is the same as Read k, m as list
	append m, k^2: push m
}
? (1,2,3,4,5,6)#fold(map2list, list)  // no #function for list
map2queue=lambda -> {
	read k, m as queue // here we place the read statement
	append m, k^2: push m
}
? (1,1,3,4,5,6)#fold(map2queue, queue)  // no #function for queue


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

                                                             