M2000 Interpreter and Environment
Version 14 revision 15 active-X
Many Improvements

I am working on the Greek Manual and then I have to rewrote the English one. I think AI can help for this.

- ctrl 1 turn to upper case the caps lock, (not tongle, hust set upper case, if we mark text then text turn to upper case, with undo).
- ctrl 2 turn to lower case the caps lock, (not tongle, hust set lower case, if we mark text then text turn to lower case, with undo).

- ALt + Hex_digits now handle big numbers for codes above 5535 (for surrogates). Also work nice the input line at M2000 console.

- ToolTip for all M2000 Controls (Unicode with title), and for Nine Patch buttons (CTXNINEBUTTON) and ShapeEx control. There are four without ToolTip, three of them have internal system for tooltips on set of data (UCPIECHART, UCCHARTAREA, UCCHARTBAR) plus one UCRADIALPROGRESS which has label for showing the type. 

- PLAYVALUE(), PLAYNOTE(), PLAYVOLUME(), PLAYNOW() for the 16 voices of sound card, which return the current value of note (1 for 1/1 of base duration, to 6, for 1/32 of the hole note duration - (1-1/1, 2-1/2, 3-1/4, 4-1/8, 5-1/16, 6-1/32). PLAYNOTE() return -1 for pause or 0 to 119 for 120 notes, starting from C (10 octaves). PLAYVOLUME() return current volume for specific voice, from 0 to 127. PLAYNOW() return true if not finished yet. For the music score I found a bug for pauses and time delays, so now works perfect. Also I add the flat notes using the symbol ♭ which now we can insert using ctrl + 3 (same for # but Shift + 3).

- IMAGE statement improved for quality
- New LIST ERRORS (see HELP ERROR) for watching the errors (and those you never see like some which you hide using the Try { } block. Errors now stamped with time from the running period of current interpreter.

- PSET now works with Width N { } using either circle as point or square using PSET ! or PSET!! (with one ! we print the square using center the current x and y, but the other one !! use the defined square as part of the screen, so the current x and y point to some point in that square



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