M2000 Interpreter and Environment

Version 10 revision 27 active-X
1. A Fix for search string on files using statement file "gsb", "thatstring|otherstring"

2. Fix Uint() function to work with Integers (16bit -1%, 32bit -1&). I=-15065%: Print Uint(I), Uint(I*1&) ' 50471   4294952231

3. chrcode() now return Integer (16bit) or Currency.
Print chrcode("성")  ' -16079
Print chrcode$(-16079) ' 성
i=chrcode("성")
Print uint(i), uint(-16079%) ' 49457 49457
Print chrcode$(49457) ' 성
' surrogate 𐐷 (4 bytes in UTF-16)
C=chrcode(chrcode$(0x10437))
Print type$(C), c ' Currency 66615
Print chrcode$(66615)  ' 𐐷

4. Updated clsOsInfo class (Information object from m2000) to v1.13, by Dragokas, member of vbforums.com, https://www.vbforums.com/showthread.php?846709-OS-Version-information-class&p=5540187#post5540187
New for clsOsInfo 1.13:
Added detection of:
 - Windows 11
 - Windows Server 2019
 - Windows Server 2022
New properties added:
 - IsWin32
 - DisplayVersion 'e.g. 21H1
 - IsWindowsXP_SP3OrGreater
 - UserName
 - ComputerName
 - IsAdminGroup
 - IsSystemCaseSensitive
 - IsEmbedded
 - LCID_UserDefault
 - IsWow64 -> made public.

Example (also there are some readonly variables which alread use the clsOsInfo class)
Module TestOsInfo {
	declare osinfo information
	with osinfo, "IsWin32" as Win32
	method osInfo, "IsWow64" as Wow64
	print "IsWin32:";Win32
	print "IsWow64:";Wow64, " OsBit", osbit
	with osinfo, "DisplayVersion" as dv$
	print "DisplayVersion:";dv$
	with osinfo, "ComputerName" as cn$
	print "ComputerName:";cn$
	print "M2000 read only variable Computer$:";Computer$
	declare osinfo nothing
}
TestOsInfo

6. Update info.gsb
New program DownLoadAny:
Download using URLDownloadToFile from lib urlmon


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

                                                             