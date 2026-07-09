M2000 Interpreter and Environment
Version 15 Revision 2

Added functions for bit manipulations for 64bit unsigned numbers, plus usnigned subtraction.  Also the new set for TEST, SET and RESET bits has two parts one for 64bit and one for 32bit. The 64bit start with BIT64. and for 32bit BINARY. 

b=BIT64.XOR(0xFFFF_FFFF_FFFF_FFFF,0xAAAA_AAAA_AAAA_AAAA)
hex b
b=BIT64.AND(b, 0xFFFF_0000_0000_FFFF)
hex b
b=BIT64.OR(b, 0x1234_5678_0000)
hex b
b=BIT64.ADD(b, 0X8000_0000_0000_0001)
hex b
b=BIT64.ADD(b, 0X8000_0000_0000_0001)
hex b
b=BIT64.SHIFT(b, 8)
hex b
b=BIT64.SHIFT(b, -32)
hex b
b=BIT64.ROTATE(b, -16)
hex b
b=BIT64.NOT(b)
hex b
b=BIT64.AND(b,0Xffff_ffff_ffff)
hex b
b=BIT64.NEG(b)
hex b
b=BIT64.NEG(0x7FF_FFFF_FFFF_FFFF)
hex b
b=BIT64.NEG(b)
hex b
b=HILOWLONG(0XAAAA_1234, 0xBBBB_5555)
hex b
b=HILOWWORD(0XAAAA, 0x5555)
hex b
' BINARY.ADD(a,b,c, ....)
b1=BINARY.ADD(0X8000_0001, 0xFFFF_1234, 1)
hex b1, HIWORD(b1), LOWORD(b1)

b=BIT64.ADD(b, 0xFFFF_1234_0000_0000)
hex b, HILONG(b), LOLONG(b)
hex HILOWLONG(LOLONG(b), HILONG(b))
' BINARY.SUB(a,b,c, ....)
b=BINARY.SUB(0XFFFF_0001, 0XFFFF_0005, 1)
hex b, HIWORD(b), LOWORD(b)
b=BIT64.SUB(0X0000_FFFF_FFFF_0001, 0XFFFF_FFFF_FFFF_0005, 1)
hex b
b=BIT64.SET(b,62)
hex b
b=BIT64.RESET(b,47)
hex b
print BIT64.TEST(b, 29, 27)=true
b=BIT64.RESET(b,29, 15)
hex b
B=BIT64.SET(0,60,58,4,1) ' make a 64bit setting 4 bits
hex b
print BIT64.TEST(b, 57)=false
B=BIT64.SET(b,57)
hex b
print BIT64.TEST(b, 57)=true
b2=BINARY.SET(0,9,8,0)
hex b2
print b2=0x301 ' bit 9,8,0
hex val(2^9+2^8+2^0->currency), b2
b2=BINARY.RESET(b2,12, 8)
hex b2,
print b2=0x201
print uint(-1%)=0×FFFF
print uint(-1)=0×FFFF_FFFF
print uint64(-1)=0×FFFF_FFFF_FFFF_FFFF
print type(uint64(-1))="Decimal"
print sint(0×FFFF_FFFF_FFFF_FFFF, 2)=0XFFFF%
print type(sint(0×FFFF_FFFF, 4))="Currency"
print sint(0×FFFF_FFFF_FFFF_FFFF, 4)=0XFFFF_FFFF&
print type(sint(0×FFFF_FFFF_FFFF_FFFF))="Long Long"
print sint(0×FFFF_FFFF_FFFF_FFFF)=0XFFFF_FFFF_FFFF_FFFF&&
print sint(0×FFFF_FFAA, 8)=0×FFFF_FFAA&&
print type(sint(0×FFFF_FFAA, 8))="Long Long"


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
THEN press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (560 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 