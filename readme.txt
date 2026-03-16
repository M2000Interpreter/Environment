M2000 Interpreter and Environment
Version 14 revision 8 active-X

1) Fix IsNum and Number read only variables.
push 12312312u   ' we push to stack a BigInteger (is an object)
Print isnum  ' fixed - now return true
long long z
z=number*100  ' now number pop the BigInteger
list ' see the variables
2) Fixed the Binary block (which have BASE64 bytes). The problem was the tabs. So if we place the Binary block and move it by placing tabs at the same lines as the data old code raising error at converting to bytes. Now fixed, character 9 is white space now for this statement
So now this work (there are tabs before the code
	BINARY {
		kwOVA5kDkQMgAKcDkQOhA5ED
	}  AS A$
	PRINT A$
3) The Matrix type of variable (exist on Math2 object), now has a library Matrix in info file. Math library now has upgraded Determinant function and a new function Inverse Matrix.

4) Update/Insert new Modules on Info file. Matrix, Server1 (we make an Http Server), FFT (fast fourier transform), Simpson (Simpson Rule, a numerical integration technique), Teacher (A speech program for English Language).

 
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

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 