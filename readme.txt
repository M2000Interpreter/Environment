M2000 Interpreter and Environment
Version 15 Revision 14

July 26, 2026,

1. Links in M2000 Editor can selected and activated (double click)
2. Links in EditBox can be selected to work or not using EnableLink (boolean) property
3. The Method TextViewOnly of EditBox now show the caret, we can copy any part, we can activate a link, but we can't alter the text.4. 
4. Modules TextOnly and mEditor in INFO file show the changes. In mEditor a new menu item added in Edit menu, to enable or disable the link selection/event_raise. Also the event is prepared for either using keyboard Ctrl+Enter (when the marked text is a URL, we can select it using F4 automatic), or just by clicking and choosing from a messagebox to open the link.\

This is the event function for Pad (the EditBox). The link opened with the default application for html files. Statement Win (Windows) call the application passing the parameter (the link).
Function Pad.WWWlink(New link$) {
	if keypress(0x11) then ' this is the Control key (left or right)
		win file.app$("html"), link$		
	else.if ask(pad=>WWWLinkMes(link$), title$)=1 then
		win file.app$("html"), link$
	end if
}

  
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