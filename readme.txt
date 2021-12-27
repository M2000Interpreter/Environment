M2000 Interpreter and Environment

Version 10 revision 49 active-X
Upgrade INterpreter for Subroutines and simple Functions (static). Minimum code for a sub and a simple function.
alfa()
Print @Beta()
Sub alfa()
End Sub
Function Beta()
End Function
So now we can use any sub/simple function in the same module/function, let say the main, or in the same top module/function, the alternative space. A top module/function is one which defined at load time (or defined in immediate mode in m2000 console). Also the Test statement, which opens the Control Form, now can show code, from all subs, in the main or alternative space. Interpreter always search from main and if can't find search in alternative space. After first search  Interpreter keeps code in a hash table.
This is an example. In Inner module we define a lambda function. This lambda function use a sub in the "alterantive space" (not in the body of lambda function). 

Module Inner {
	alfa=lambda (x)-> {
		delta(&x)
		=x
	}
	push alfa
}
Inner
Read ret
Print ret(100)
Dim a(10)
a(3)=ret
Print a(3)(300)
sub delta(&a)
	a++
end sub

// So we can move the delta sub to the inner module (this is also alternative space for lambda alfa). So we can put this code in a top module B, and we can use Test B to execute the B using the Control Form, watching the code as executed.
Module Inner {
	alfa=lambda (x)-> {
		delta(&x)
		=x
	}
	push alfa
	sub delta(&a)
		a++
	end sub
}
Inner
Read ret
Print ret(100)
Dim a(10)
a(3)=ret
Print a(3)(300)



George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The fist time you run the interpreter do this in M2000 console:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)


http://georgekarras.blogspot.gr/

ExportM2000 all files with executables:
https://drive.google.com/drive/folders/1IbYgPtwaWpWC5pXLRqEaTaSoky37iK16

only source without executables (something going wrong with GitHub)
https://github.com/M2000Interpreter/Environment

https://drive.google.com/open?id=0BwSrrDW66vvvdER4bzd0OENvWlU

                                                             