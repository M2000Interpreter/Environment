M2000 Interpreter and Environment
Version 15 Revision 16

July 27, 2026,

1. Some minor fixes
2. Update Exif class (now we can use it from code too)

Enum IFD0 {	
    ImageDescription = &H10E&
    Make = &H10F&
    Model = &H110&
    Orientation = &H112&
    XResolution = &H11A&
    YResolution = &H11B&
    ResolutionUnit = &H128&
    Software = &H131&
    DateTime = &H132&
    WhitePoint = &H13E&
    PrimaryChromaticities = &H13F&
    YCbCrCoefficients = &H211&
    YCbCrPositioning = &H213&
    ReferenceBlackWhite = &H214&
    Copyright = &H8298&
    ExifOffset = &H8769&
    
    ImageWidth = &H100&
    ImageHeight = &H101&
 }
menu
' for constants see Help Path$()
Const CSIDL_DESKTOP           = 0x0
Const CSIDL_MY_DOCUMENTS      = 0x5
Const CSIDL_DESKTOPDIRECTORY  = 0x10
Const CSIDL_MYPICTURES        = 0x27
'dir path$(CSIDL_MY_DOCUMENTS)
'dir path$("%USERPROFILE%\Downloads\")
dir user
files + "jpg"
if menuitems>0 then
for i=1 to menuitems
	? "file:";menu$(i)+".jpg"
	a=getobject("","m2000.exifread")
	a=>load dir$+menu$(i)+".jpg"
	? a=>Tag(ImageDescription)
	? a=>Tag(Make)
	? a=>Tag(Model)
	Orientation(a=>Tag(Orientation))
	? a=>Tag(ImageWidth)+" X "+a=>Tag(ImageHeight)+" PIXELS"
	? a=>Tag(XResolution)+", "+a=>Tag(YResolution)
	ResolutionUnit(a=>Tag(ResolutionUnit))
	? a=>Tag(DateTime)
	? a=>Tag(Software)
	? a=>Comment
	?
	push key$: drop
next
end if
dir user

sub ResolutionUnit(a) 
	select case a
	case 1
		? "No absolute unit of measurement"
	case 2
		? "Inches"
	case 3
		? "Centimeters"
	case else
		? "Unknown unit (no 1,2 or 3):";a
	end select
end sub
sub Orientation(a)
	select case a
	case 1
	? "Normal (Top-Left) — The image is already upright, no rotation needed."
	Case 2
	? "Flipped horizontally (Top-Right) — Mirrored left-to-right."
	Case 3
	? "Rotated 180 degrees (Bottom-Right) — Upside down."
	Case 4
	? "Flipped vertically (Bottom-Left) — Mirrored top-to-bottom."
	Case 5
	? "Flipped horizontally and rotated 270 degrees clockwise (Left-Top)."
	Case 6
	? "Rotated 90 degrees clockwise (Right-Top)."
	Case 7
	? "Flipped horizontally and rotated 90 degrees clockwise (Right-Bottom)."
	Case 8
	? "Rotated 270 degrees clockwise / 90 counter-clockwise (Left-Bottom)."
	case Else
	? ""
	end select
end sub





George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and THEN open it again.

To get the INFO file, from M2000 console do this:

dir appdir$
load info

then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at GitHub to compile M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).

English old paper for M2000
https://github.com/M2000Interpreter/Environment/releases/download/ver13rev44/M2000paper.pdf

Greek Book for learning programming
https://github.com/M2000Interpreter/Environment/releases/download/version15revision10/GreekBookM2000.pdf

Greek Manual (a work in progress)
https://github.com/M2000Interpreter/Environment/releases/download/version15revision10/GreekManualVersion15_preview.pdf

Greek Book About OOP in M2000
https://github.com/M2000Interpreter/Environment/releases/download/version14revision51/OOP_M2000_2026.pdf

http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (578 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 