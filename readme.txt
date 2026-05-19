M2000 Interpreter and Environment
Version 14 Revision 33

1. Fix a problem as shown:

m=alfa() ' create a dummy alfa() in functions inner catalogue with a -1 index (which not exist, indexes are positive numbers)
// we can't make a reference for a static function
Push &alfa()  ' now raise error before hang the program

function alfa()
	=100
end function

2. Declare static function (only the name) to exclude a global function to bypass the static function. A static (or simple as we use to say) function for M2000 is a function like a sub, a part of code where we can use local variables and we see anything from the caller (a normal function can see only own identifiers - which are local by default- and global identifiers). Now we can call a static function using @ before the name (this is the fastet way), or without @. In any case we can't use a name for static function the same as a system identifier. So the function Abs() can't used for name for a simple function. The @Abs() is the original Abs(). We can make a normal function ABS() to bypass the original Abs() but we can use @Abs() to use the original always. So in a module lets use a new name for a simple or static function Hello(). 
So now we can make a global Hello() which can use the static function (but not the variables of module alfa). Using statement Static Function Hello or Static Function Hello() we insert a dummy local function which hide any global with same name, before any use of the static function. So using UseStaticFunction=true we get 0 from the Write variable, and using UseStaticFunction=false we get 4 from the Write variable.
As you see we can ude the static function because the search function always performed for the first call (then the function regostered in a local subs/function list). Modules and normal Functions (defined same as modules) are on a list of modules/fumctions and are more heavy definitions from subroutines and simple functions. Actually a module and a normall function have a basetask object which run on that task. Subs and static functions use the caller basetask (and each basetask has an owner the output object plus a stack of values for dynamic use of memory and passing (and returning for modules) values - like a stack for machine code, but we place values like strings or other types, like objects and arrays).

Also the statement Static Function check if there is an array with same name and raise error. We can use A() as function and A() as array, we have to use A(* ) as function call (the * say i am function). Always arrays have priority over functions with same name. Calling with @ interpreter know that is a call to a function (internal or a static one).

Global UseStaticFunction=true
Global Write as long
FUNCTION Global Hello(X) {
		=@Hello(X)
//			BEEP
//			WAIT 200
		Write++
}
if UseStaticFunction then Static Function Hello
Module alfa {		
	if UseStaticFunction then Static Function Hello()	
	Print Hello(10)	
	InternalUse()
	Gosub InternalUse()
	' this is a simple function
	Function Hello(x)
		=String$("Hello", x)
	End Function
	' this is a subroutine
	Sub InternalUse()
		Print Hello(5)
	End Sub		
}
Call alfa
' resolved automatic to Function Hello(x)/End Function
' it is in the same original code
Print Hello(3)
Print Write


3. I put a pool for objects mStiva (the stack of values), so we reuse the objects as we go.



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