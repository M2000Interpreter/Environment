M2000 Interpreter and Environment

Version 13 revision 21 active-X
1. Fix how names print for list of modules and variables.
You can check List statement for variables and modules ? for modules/functions.
2. Fix for list/queue/stack which use := for subs
(break from Version 13 Revision 17).
2.1 
alfa(list:=1,2,3)
sub alfa(a as list)
	? a
end sub

2.2
a=list := 1,2,3,4,5
beta(a,"string", 1)
sub beta(o as list, v1 as variant, v2 as double)
	? o
end sub

3. Added AS BIGINTEGER for CONST. Added AS CONST for parameters:
CONST P AS BIGINTEGER=-1232312323423423423334234U
ALFA1(P)
ALFA2(P)
SUB ALFA1(K AS BIGINTEGER)
	? K
	? TYPE$(K)="BigInteger"	
END SUB
SUB ALFA2(K AS CONST)
	? K
	? TYPE$(K)="Constant"
END SUB
4. Now we can make also CONST tuples
CONST P=(1,2,3,4)
This is read only.
CONST P=(1,2,3,4)
CONST A=P#MAT("+=", 500)
PRINT A#SUM()=2010
PRINT TYPE$(P)="Constant"
PRINT TYPE$(A)="Constant"
Z=P
PRINT TYPE$(Z)="tuple"
PRINT Z IS P=FALSE
Z++
PRINT Z#SUM()
PRINT P#VAL(3)=4
PRINT A#VAL(3)=504
PRINT Z#VAL(3)=5
5. EVALUATOR module in info.gsb now run as expected.

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

https://rosettacode.org/wiki/Category:M2000_Interpreter (534 tasks)
                 