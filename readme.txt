M2000 Interpreter and Environment

Version 10 revision 35 active-X
1. Fix const lambda to be used in a call by pass by reference
	const b=lambda k=1 -> {
		=k
		k++
	}
	module dosomething (&k) {
		Print k()
	}
	for k=1 to 10
		dosomething &b
	next
2. Final lambda in groups now can be used in a call when the group passed by reference.
	group M {
		final b=lambda k=1 -> {
			=k
			k++
		}
	}
	module dosomething (&k, &z) {
		Print k.b()  ' k.b() is the same as z()
		Print z()
	}
	for k=1 to 5
		dosomething &M, &M.b
	next

3.Update Demo1 in Info

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

                                                             