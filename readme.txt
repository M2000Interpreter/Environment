M2000 Interpreter and Environment
Version 13 revision 46 active-X

1. Now M2000 environment can use fonts with break char different from space, like Segoe UI which have char 13 and not space (char 32). Now the Report statement works nice with Segoe and left and right justification (the problem was on both). Also tab characters can be inserted and used with the full justification.


2. A new class for creating Various Barcodes from https://github.com/wqweto/ClipBar
Also I make Image to utilize the Picture object. see details in the following example:

CLS,0
declare bar "m2000.cBarcode"
print type$(bar)
Enum UcsBarCodeTypeEnum {
	ucsBctAuto = 0
	ucsBctEan13
	ucsBctEan8
	ucsBctEan128
	ucsBctUpcA
	ucsBctUpcE
}
Long ink_bleed_in_percents=10
boolean true=1=1, false=1=0
method bar, "init", ink_bleed_in_percents,  bHangSeparators:= TRUE, bShowDigits:=TRUE as ok
method bar, "GetBarCode", "123456789012:12345",  1 as Pic
With Pic, "Width" as pic.width, "Height" as pic.height
move  5000, 6000
' Pic is a Picture object
mem$=""
' Now Image can render Pic to string
Image Pic, 7500  to mem$
? image.x(mem$)=7500
rem	? len(mem$)
' so now we place a bitmap
image mem$			 ', 7500
' We can render Picture as emf directly
Step 0, -4000
Image Pic, 7500*1.2
' We can make a file in memory using Drawing.
' We can make another emf where we draw the picture
drawing {
	image Pic, 3000
} as Mem
? len(Mem)
step -3000, 7000
'  we can rotate the image now
image Mem, 10000, ,-90

3. Remove a Bug in Select Case
The MG module in info now run as expected.
The problem begins some revisions before this one. When a case match and another case wait after a comma and a block was under the case, then this block never run.
Now work fine.

select case 2
case 2, 3
	{
	print "ok"
	}
end select

4. DD8 module updated (load INFO from appdir$ after the installation of this revision).




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