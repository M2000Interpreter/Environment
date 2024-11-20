M2000 Interpreter and Environment

Version 12 Revision 47 active-X
1. Fixed a bug in WIN statement, and now the Win "file.txt" open file.txt from current folder to default app for txt files.
2. Fixed some bugs in MovieModule for video player (see also Vplayer on info file)
3. Movie.Height and Movie.Width now return the height and width of movie in twips.
4. A new event for use form, called MoveTo which return X and Y the left and top values before the form get the new one, when we moving the window. This is usefull if we have a video on user form and we want to move the video together with the form.

boolean mymovie
// here we have the code for making the window (see examples in info)
Declare form1 form

// before the form close we have to erase the movie using this:
function form1.unload {
  movie
}
// we can play from the start by clicking the form
function form1.click {
  if mymovie else exit            
  movie to 0
  movie restart
}
// this is the new one to move the movie when we move the window.
function form1.moveto {
  read new x, y
  movie x+7000, y+1000
}
// before we turn visible the form:
layer form1 {
  cls #333333, 0
  movie motion.x+7000, motion.y+1000, 2000
  // we can load the movie without playing using movie load statement.
  movie load "Second"  
  mymovie=true  ' so now we mark tha the video Second.avi is loaded
  // this code run after 300 msec (when the form is open)
  after 300 {
    movie show
    movie restart
  }
}
// so now we can show the form
      Method form1, "show", 1
      Declare form1 nothing

5. Updated Info.gsb



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

                                                             