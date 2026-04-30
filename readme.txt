M2000 Interpreter and Environment
Version 14 revision 23 active-X
1. I split the taskmanager to two objects, the second one now has the music score player tasks. And now we can use these in any degree of execution speed (SLOW, FAST and FAST !) without using wait or other similar task oriented statemends like Every {} and Main.Task {}. There is a new tstjoy2 module in INFO file.
(Versions before 11 can play midi messages for SLOW and FAST but not for FAST!, from 11 to play that kind of music need something which give time to run the music tasks. Now because I made faster the M2000 Interpreter I think it again and I find a fast way to play the tasks and have a nice time slice between these tasks and the execution of statements.
So on an old computer (~14 year old, using Intel(R) Core(TM) i5-3470 CPU @ 3.20GHz), with a lot of other programs running:
times milliseconds
Slow: insert sleep time, fast normal with screen refresh based on intervals, and fast ! without screen refresh.
We have to use this from a module or without it from M2000 console.
SET SLOW
                 slow   fast   fast!  Play song
For Loop 100000: 3949   2366   2366   no
For Loop 100000: 4998   4568   4030   yes
While 100000   : 5523   3852   3874   no
While 100000   : 7110   6720   6753   yes
Do 100000      : 5522   3816   3899   no
Do 100000      : 7632   6460   6942   yes
Loop 100000    : 8261   5729   5644   no
Loop 100000    :11980   9897   9539   yes
As we see playing a song reduce the speed but for the song we notice nothing, and that is the goal, to get time when we need for what we need. The solution based from a previous solution of how we make the Ctrl C to break the fast ! which not perform a refresh or a "Doevents" in VB6 context of programming (which the M2000 Interpreter written). The idea was to use a way to fast compare two long values stripping these from some not wanted bits. So we do not perform division or modulus but just a binary AND (very fast for CPU).

These are the test modules the TESTD has three instructions inside a block. The Loop raise a flag for block which make the block to run again. We use the IF statement and a boolean expression to check to exit. This is more time consuming from the other types of loops which has only one statement I-- to run.
MODULE TESTA {
	PROFILER:Z=100000
	FOR I=1 TO 100000:Z++:NEXT
	PRINT PART $(0),TIMECOUNT
	PRINT
}

MODULE TESTB {
	PROFILER:I=100000
	WHILE I>0:I--:END WHILE
	PRINT PART $(0),TIMECOUNT
	PRINT
}
MODULE TESTC {
	PROFILER:I=100000
	DO:I--:UNTIL I<1
	PRINT PART $(0),TIMECOUNT
	PRINT
}
MODULE TESTD {
	PROFILER:I=100000
	{LOOP:I--:IF I<1 THEN EXIT
	}:PRINT PART $(0),TIMECOUNT
	PRINT
}

2 Fix the PSET !! and ! options (was a mistake and in any way we get the !! option for both symbols)
3 Italic and Italics (was Italic and now has two words)


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