M2000 Interpreter and Environment

Version 13 revision 20 active-X

a FIX for EMF files.
On a Windows 11 laptop I found that the reference device - as stored in emf- has a bug from the OS, so I do my own calculation and now the emf has the size I want. The reference device cx/cy values used for adjusting the size, so my image before get a 1.75 bigger size, so my drawing drawing smaller at the top left corner. This happen when I choose to declare a bounding rectangle. The bounding rectangle saved ok on the original header but not on the emf file header (although the original header also included in the emf file).
So the fix in vb6 was:
' copy of the header to mheader
CopyMemory ByVal VarPtr(mHeader.iType), ByVal aPic.GetBytePtr(0), 88
If boundrect.Bottom > 0 Then
' then if we have a boundrect we get the pixels (which are ok)
' and calculate the milliimeters which aren't ok (fort Windows 10 are slight different, but for Windows 11 on a laptop with 1920X1080 screen was very bad)
mHeader.szlMillimeters.cx = mHeader.szlDevice.cx * 15 / 1440 * 25.4
mHeader.szlMillimeters.cy = mHeader.szlDevice.cy * 15 / 1440 * 25.4
' restore the bytes...on memory buffer
CopyMemory ByVal aPic.GetBytePtr(0), ByVal VarPtr(mHeader.iType), 88
End If

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
                 