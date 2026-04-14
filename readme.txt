M2000 Interpreter and Environment
Version 14 revision 17 active-X
Many Improvements

1) Now we can split Write to two or more statements using the comma as last character. (Before this revision the comma produce error and you have to include everything in one line, although M2000 has no line length limit - editor has a limit for line length, and then wrap the line without breaking the paragraph, so we can edit it)
The Input statement was ready for splitting from a lot of revisions/versions before;


// You can erase wide for using the current LOCALE for converting strings to ANSI (strings are in UTF16LE format if they are made it from literals, but actual they are data buffers)
// Write With and Input With are statements to change how a csv written/read
open "tsts.csv" for wide output as #f
	Write #f, "alfa", "beta", 100,
	Write #f, 200, "ok"
close #f
open "tsts.csv" for wide input as #f
	Input #f, a$, b$, num1, 
	Input #f, num2, c$
	Print a$, b$, num1, num2, c$
close #f

2) I use ChatGpt to learn about music - about dots, ties and tuplets. So from the help of ΑΙ: no code, only advice, making test programs using M2000 code which learn in a minute after I told it how it works and then checking data of songs playings note serving against time. Also AI make the Ode of Joy version for M2000. 

So now we can program the internal midi organ to play with staccato n percent note plus (100-n) percent silence/legato 100% no silence. 

This is the Joy module in info file:


CONST  Acoustic_Bass = 33
TEMPO=960*1.2
VOLUME 80
PLAY 1, 0

' Notes: 12, 7 physical and 5 semitones
' You can use C# or D♭ (is the same). The ♭ character from CTRL + 3

' Octave 4 (5th)
' we can change octave once for next notes too, until a new change happen.
'Octaves 0 to 9

' VALUE PART
' @1 TEMPO  - IF YOU OMIT THEN IT IS @1
' @2 TEMPO/2
' @3 TEMPO/4
' @4 TEMPO/8
' @5 TEMPO/16
' @6 TEMPO/32
' dotted note @2+  (@2 plus @3)
' double dotted @2++ (@2 plus @3 plus @4)
' we can add Velocity (Volume), V0 to V127, but this change until next V excuted

' You can make TIES adding values (max 20), so @4@2 is a TEMPO/8+TEMPO/2 duration for one note.
' A ! means legato, but a !95 is a staccato 95% - You can assign it per note if you like (or at SCORE, 4th parameter for global in score)

' You can make Tuplets: [CD]3@3 two notes played for 3 x quarters. We compute scale=(3/4)/(2)=3/8, so C played for 1x3/8 and D played for 1x3/8
' Tuplets can't have tuplets inside (for this version).  Tuplets may have staccato/velocity for all individul notes
' Value for tuplet may have Multiplier before or not. A tuplet without defined value and or multiplier has value 1 Tempo by default.
' At the end of tuplet Staccato, Volume, Octave restored.


' USE SPACE FOR SILENCE - ALSO SILENCE HAVE VALUES
' SO A SPACE IS A SILENCE FOR A TEMPO TIME.
' A SPACE@4 IS A SILENCE FOR TEMPO/8

' Volume/Velocity V0 TO V127   - you can place it before note or after value (value may missing, is 1/1 by default) for the note to take account the change.
' Inside Tuplets not use this at start - only after value (value may missing, is 1/1 by default)
' For defining Volume for Tuplets you can do that aftet the value of Tuplet and before Staccato. so [C@3 @3E@3]2V60!75 place V60 for Tuplet and staccato 75%. See there is only multiplier, so this Tuplet has duration 2xTEMPO. Space inside is the silence of a value 1/4*scale. Scale = 2/(3/4)=8/3, so the 1/4*8/3=8/12=2/3. C@3 has finally Tempo*2/3*0.75 duration time
partitura="V80"
partitura+="E4@3E@3F@3G@3"
partitura+="G@3F@3E@3D@3"
partitura+="C@3C@3D@3E@3"
partitura+="E@4D@4D@2"

partitura+="E@3E@3F@3G@3"
partitura+="G@3F@3E@3D@3"
partitura+="C@3C@3D@3E@3"
partitura+="D@4C@4C@2"

partitura+="D@3D@3E@3C@3"
partitura+="D@3E@4F@4E@3C@3"
partitura+="D@3E@4F@4E@3D@3"
partitura+="C@3D@3G3@2"

partitura+="E4@3E@3F@3G@3"
partitura+="G@3F@3E@3D@3"
partitura+="C@3C@3D@3E@3"
partitura+="D@4C@4C@2!"  ' LAST NOTE LEGATO - 100%

SCORE 1, TEMPO, partitura, 75  ' 75% STACCATO  -  100% IS LEGATO (OMIT THE PARAMETER)- NO SILENCE BETWEEN NOTES.
PLAY 1, Acoustic_Bass
PRINT "ODE TO JOY"

3) new functions to monitoring the midi channels.
where n is the channel 1 to 16. (10 for drums always)
PLAYNOW(n) return true if n channel is running
PLAYNOTE(n) return -2 - not playing, -1 - play silence, 0 - 119 play midi note.
PLAYVALUE(n) return the nominal value, so a 8 means full_note_duration/8  (In score statement, tempo is the full note duration in msec)
PLAYVOLUME(n) return from 0 to 127 the current output volume/velocity for midi
PLAYDOTS(n) return 0 -note, 1 -dotted note, 2 -double dotted note (dotted note has value+value/2, doubled dotted note has value+value/2+value/4)
PLAYGATE(n) return 100 for legato, <100 for staccato.
PLAYTUPLET(n) return 1 for out of tuplet play, or the scale of tuplet, can be any value including 1. The value computed from the system f=(defined duration of tuplet)/(duration of all members of tuplet as computed). So each member's duration multiply by f for the final duration (without apply the gate, the staccato percent)

In Test 7 we have this:[E@3G@3 @3C@3]@2+!80
The set [ ] is like one note, say X so we have X@2+!80, a value of 1/2 plus 1/4 (dotted note) and a staccato 80%. But because we have multiple notes the staccato affect every one, except members of silence. We can put V50 before staccato !80 and we pass the volume only to members. This tuplet has a 0.5 scale, because inside we have 4*1/4=1 and we ask for @2 or 1/2, so we get scale=(1/2)/1=1/2=0.5

This is a test based form a test from ChatCpt (spaces are for silence and they have value too). The new one include a Tuplet at TEST 7, plus two more PLAY_functions, the PLAYGATE() and the PLAYTUPLET()

TEMPO=960

' @1=960 msec
' @2=480
' @3=240
' @4=120
' @5=60
' @6=30

' TEST 1: basic values
a$="C4@3D@3E@3F@3G@3A@3B@3C5@3"

' TEST 2: dotted notes with +
b$="C4@3+D@3+E@3+F@3+"
b$+="G@2+A@2+"

' TEST 3: double-dotted notes with ++
c$="C4@3++D@3++"
c$+="E@2++F@2++"

' TEST 4: rests with + and ++
d$="C4@3 @3+D@3 @3++E@3"
d$+=" @2+F@3 @4+G@2"

' TEST 5: velocity changes
e$="C4@3V40D@3E@3F@3"
e$+="G@3V90A@3B@3C5@3"
e$+="B4@3V127A@3G@3F@3"

' TEST 6: same note repeated versus long held note
f$="C4@3C@3C@3 @3C@1"
g$="C4@2+C@3 @2"

' TEST 7: mixed example
h$="E4@3E@3E@2+ @4"
h$+="E@3E@3E@2++ @3"
h$+="[E@3G@3 @3C@3]@2+!80D@3E@1"

SCORE 1, TEMPO, a$, 100
SCORE 2, TEMPO, b$, 95
SCORE 3, TEMPO, c$, 95
SCORE 4, TEMPO, d$, 95
SCORE 5, TEMPO, e$, 95
SCORE 6, TEMPO, f$, 95
SCORE 7, TEMPO, g$, 95
SCORE 8, TEMPO, h$, 95

LOCALE 1033
OPEN "DURATION_TEST.TXT" FOR OUTPUT AS #F
PROFILER
PLAY 1,1,2,1,3,1,4,1,5,1,6,1,7,1,8,1
EVERY TEMPO/128 {
    WRITE #F, PLAYNOTE(1),PLAYNOTE(2),PLAYNOTE(3),PLAYNOTE(4),
    WRITE #F,PLAYNOTE(5),PLAYGATE(5) ,PLAYNOTE(6),PLAYNOTE(7),PLAYNOTE(8), PLAYTUPLET(8),PLAYGATE(8),
    WRITE #F,TIMECOUNT
    IF NOT PLAYSCORE THEN EXIT
}
CLOSE #F
WIN "NOTEPAD", DIR$+"DURATION_TEST.TXT"




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