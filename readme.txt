M2000 Interpreter and Environment

Version 12 Revision 66 active-X
1. A small bug prevent to read from stack of values the cxComplex number from a tuple. Now thid code run and return: cxComplex cxComplex
==example==
declare global m math2
method m,"cxone" as one
a=(one, one)
what(!a)

sub what(a, b)
	? type$(a), type$(b)
end sub
2. Added modules on info.gsb file:
ROOTS (using complex numbers to find Roots of Quadratic Equations)
TAXICAB (find TaxiCab numbers and display 1th to 25th and 2000th to 20006th. This module make and sort more than 719000 items on a List object, need less than 6 minutes on an old Intel(R) Core(TM) i5-3470 CPU @ 3.20GHz)
This is the code of TAXICAB (we call Taxicab inside Taxicab whithout a recursive call. Modules have to use Call NameOfModule to call itself. So the call using only the NAME call the local or global module (excluding this module if it is global). So we have the local TaxiCab so we call this one.

module TaxiCab (f as long){
	cls,0
	Print Part "Taxicab numbers"
	Print Under
	profiler
	var Cubes=list, 	Sums=list, 	Ret=list
	var st=0, en=1200
	st=@Proc(0, en)
	sort ret as number
	Display(1, 25)
	Display(2000, 2006)
	Print timecount
	print "done"
	end
	sub Display(from, to)
		local k=each(ret, from,to), s=""
		while k
			s= format$("{0:-6} {1:-12}",K^+1, val(eval$(k!)+"&&"))+eval$(k)
		 	Print s : if f>-1 then print #f, s
		end while
	end sub
	function Proc(ia, ib)
		local i, cube as long long, s as long long
		for i=ia to ib
			if i mod 10=1 then print over $("#0.00"), "Working..";(i-ia)/ib*100;"%"
			cube=i^3
			Append Cubes, cube
			k=each(cubes)
			while k
				s=cube+eval(k)
				if not exist(Sums, s) then
					append Sums,s:=(i)+"^3 + "+(k^)+"^3"
				else.if not exist(Ret, s) then
					append Ret, s:=" = "+(i)+"^3 + "+(k^)+"^3 = "+eval$(Sums)
				end if
			end while
		next
		print over $("#0.00"), "Working..";100;"%"
		print
		=i
	end function
}
file2export="TaxiCabNumbers.txt"
open file2export for wide output as #f
TaxiCab f
close #f
win dir$+"TaxiCabNumbers.txt"

A taxiCab number the sum of at least two different set of cubes of numbers: So for 1279 we have the set 12 - 1 and the set 10 - 9. Here ^ used to display the power operator (M2000 use the ** and ^ as power operators).

     1         1729 = 12^3 + 1^3 = 10^3 + 9^3
     2         4104 = 16^3 + 2^3 = 15^3 + 9^3
     3        13832 = 24^3 + 2^3 = 20^3 + 18^3
     4        20683 = 27^3 + 10^3 = 24^3 + 19^3
     5        32832 = 32^3 + 4^3 = 30^3 + 18^3
     6        39312 = 34^3 + 2^3 = 33^3 + 15^3
     7        40033 = 34^3 + 9^3 = 33^3 + 16^3
     8        46683 = 36^3 + 3^3 = 30^3 + 27^3
     9        64232 = 39^3 + 17^3 = 36^3 + 26^3
    10        65728 = 40^3 + 12^3 = 33^3 + 31^3
    11       110656 = 48^3 + 4^3 = 40^3 + 36^3
    12       110808 = 48^3 + 6^3 = 45^3 + 27^3
    13       134379 = 51^3 + 12^3 = 43^3 + 38^3
    14       149389 = 53^3 + 8^3 = 50^3 + 29^3
    15       165464 = 54^3 + 20^3 = 48^3 + 38^3
    16       171288 = 55^3 + 17^3 = 54^3 + 24^3
    17       195841 = 58^3 + 9^3 = 57^3 + 22^3
    18       216027 = 60^3 + 3^3 = 59^3 + 22^3
    19       216125 = 60^3 + 5^3 = 50^3 + 45^3
    20       262656 = 64^3 + 8^3 = 60^3 + 36^3
    21       314496 = 68^3 + 4^3 = 66^3 + 30^3
    22       320264 = 68^3 + 18^3 = 66^3 + 32^3
    23       327763 = 67^3 + 30^3 = 58^3 + 51^3
    24       373464 = 72^3 + 6^3 = 60^3 + 54^3
    25       402597 = 69^3 + 42^3 = 61^3 + 56^3
  2000   1671816384 = 1168^3 + 428^3 = 944^3 + 940^3
  2001   1672470592 = 1187^3 + 29^3 = 1124^3 + 632^3
  2002   1673170856 = 1164^3 + 458^3 = 1034^3 + 828^3
  2003   1675045225 = 1153^3 + 522^3 = 1081^3 + 744^3
  2004   1675958167 = 1159^3 + 492^3 = 1096^3 + 711^3
  2005   1676926719 = 1188^3 + 63^3 = 1095^3 + 714^3
  2006   1677646971 = 1188^3 + 99^3 = 990^3 + 891^3


You can measure a run using Profiler and Print TimeCount:
Profiler:ModuleName:Print TimeCount
The TaxiCab module use Profiler/Timecount so no need to use it external.


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

                                                             