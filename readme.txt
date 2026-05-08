M2000 Interpreter and Environment
Version 14 Revision 27

Some last big changes (not for new version), which made Version 14 the best
1. Select Case with no restrictions (like those from all other revisions/versions)
When First time select case introduced to M2000 has many restrictions: 
1.1. Between each CASE only one statement or a block of statements (using curly brackets).
After some versions we can use multiple Case statements without something but need one of them to have something so if one of them has a solution (is in case) then the statement executed:
Select case a+b
case 1
case 2
Print "1 - 2"
end select

So this was not something that the previous can do, but the previous implementation want all the cases in one case like this:
 
Select case a+b
case 1, 2
Print "1 - 2"
end select

What was the problem and have those restrictions? Because there was 2 vocabularies, in Greek one, was a word which used as CASE and as WITH <object>... statement. So to not have problem with this I put the restriction. Now the are two versions for the Greek statement for the With <object> statement, the old and the new one.
Another reason was the speed. If we have restriction to what we have in front, we can choose faster. A select case has a skip function to skip cases so by using the restriction founding open block (a curly bracket) skip the block and expect the next Case or the End Select.

Also the ELSE CASE was just ELSE, so what happen If i put multiline code without using a block { }?
So I skip all the work to find out a solution. And I make some small changes, to have more lines but for each line execution can be done for statement which executed in one line only.
See this code. The while has a block but the system read intentional "while i<10 {" so this break
select case 1
case 1 ' not run on revision 26 at lower
i=0:while i<10 {?i:i++}
end select  

The same for this code:
select case 1
case 1 ' not run on revision 26 or lower
i=0:while i<10 :?i:i++:end while
end select

Because the interpreter get the "while i<10 :" and try to run it (the other part missing for that time).
 
These two scripts runs on revision 27.
Look something better (we do a jump which break the inner select case, the while, the multiline if and the for/next loop:

limit=30
if rnd>.5 then limit*=100
? "Limit=";limit
P=(100,0,1,5)#val(random(0,3))
? "P=";p
for j=1 to 10
	if limit>20 then
		select case P
		case 0
		// zero has something so not join with 1 and 5
		? j
		case 1 ' not run on revision 26 at lower
		case 5 ' so this 
			i=0
			print "Starting i=";i
			while i<50
				select case i
				case <5			
					i++
				case >limit
					? "i="+i, "limit="+limit
					goto earlyexit
				case else
					i+=2
				end select
				print "current i=";i;", j=";j
			end while
		case else
		end select	
	end if
next j
print "normal exit"
end 
earlyexit:
print "exit"


So I found a solution, using two things: A mark on the current task return stack (this holds only for return from subs, plus a mark for multiline if). So now we mark the "run" of the select case. Using this we can determine an ELSE if belong to a multiline IF or a Select Case. The second was a way to skip the case. Because the return stack hold values only for the running code, the skip "case" or "case else" has to do something to keep the multiline if and then inner select case as fold structures in line, so we work recursive and use counters. If a counter finish as not zero this means we have something odd. It is more complicated because the select case do something peculiar:
This example find case 2 and print 2 but send a break event which break only this block and continue all the cases until case 5 where in another block a continue event  break the block and  continue the "select case" job as normal, so skip case 6 and case else. This exist from the first version of Select Case (from 2002). So here we get 2,3,4,5 only if we start from 2. In any other case we get one number only.
Statements break and continue out of the block in a Case works for the block which the Select belong. Two ore more folded Select cases do not consist multiple logical block, but all are belong to the same. So an exit, or a goto "break" the select case...breaking the top block which the select case belong. M2000 do not use AST (yet), instead use object, which hold the "then" state pf execution, and all the functions passed the current task (so the task can be a thread, or an event routine). Variables are stored in a flat array/hash table, so every return clear the last frame. Using this we can use variables as references for older variables (at the clearing stage the system knows about references and just clear the pointer not the value).

select case 2
case 1
	//	nothing
case 2
	{? 2 : break}
case 3
	? 3
case 4
	? 4
case 5
	{?5:continue}
case 6
	? 6
case else
	?"do something else"
end select




2. Enum variables can get the Error Value if we supply wrong value instead raising an error (this is the default action, the error).
Also operator ^ now return for Enum variable the index on the set of Enum values.
Module Example1 {
	Enum UnaryOperatorSymbol {
		UO_plus="+"
		UO_minus="-"
	error:
		UO_no=""
	}
	
	Var z as UnaryOperatorSymbol
	' excluding the fault value..(is at the end)
	k=each(UnaryOperatorSymbol)
	while k
		print k^, eval$(k) 
	end while
	z="+"
	Print eval$(z)="UO_plus", z=UO_plus
	z="??"  ' not exist
	Print eval$(z)="UO_no", z=UO_no
}
Module Example2 {
	Enum UnaryOperatorSymbol {
		UO_plus="+"
		UO_minus="-"
		UO_no=""
	}
	
	Var z as UnaryOperatorSymbol
	' excluding the fault value..(is at the end)
	k=each(UnaryOperatorSymbol)
	while k
		print eval$(k) 
	end while
	z="+"
	Print eval$(z)="UO_plus", z=UO_plus
	try ok {
		z="??"  ' not exist
	}
	if not ok then z=UO_no
	Print eval$(z)="UO_no", z=UO_no

}
example1
example2

3. I change the way which the M2000 Editor change the color set for the source. It is better. The pen color determine the final comment color (for entire line). All other colors are standard (depends of set). The program make a paperwhite background so we get the "black color" our identifiers.

4. Before this revision only users but not supervisor has a private definition for colors. Although using Settings we can set colors, we can't set any color (Settings is limited to 16 basic colors). So now I made m2000.exe to send the message to start using the Desktop.inf file if exist for the Supervisor also. The M2000.exe still run ok with old M2000.dll.

This is a desktop.inf file (use Edit "desktop.inf" to paste this and press esc to save to file). You can hide/unhide settings selecting the lines you want and press ctrl /
Also to make the changes to have to give a Break (Start or Start "" or Start "", "")

font "verdana" ' white strings
Bold 1
Cls #332211, 0
Pen #AACCDD     ' Cyan all internal identifiers

//	font "verdana" ' magenta strings
//	Bold 1
//	Cls #F9FBFF, 0  ' paperwhite
//	Pen 0  ' blue known identifiers

//	font "verdana"  ' white strings
//	Bold 1
//	Cls 5, 0    ' magenta background
//	Pen 14     ' Cyan all internal identifiers

//	font "verdana"  ' white strings
//	Bold 1
//	Cls 1, 0  ' Blue background
//	Pen 11   ' Cyan all internal identifiers

//	font "verdana"  ' magenta strings
//	Bold 1
//	Cls #FFAACC, 0    ' light red bacground
//	Pen #224455     ' Blue all internal identifiers



Or you can use this Desktop.Inf (set English keyboard/messages, Courier New for full HD screen and set tab width to 2 characters - not 4 the default:
English
font "Courier new"
form 136, 36 ' for 1920X1080 - 42 with zero linespace.
back {cls #002200,0}
form;   ' this extend back to full screen
// form  ' this make back same size as the console
bold 0
Cls #003300, 0 '#226611,0
Pen #99ffbb 
edit !2  ' 2 places for TAB


Using: Start Desktop ""  you do a warm restart (this not erase the loaded modules)
Using: Start you do a cold restart (but this ask before perform the restart, if you have not save the loaded/edited modules/functions).
The Cold restart bypass the Desktop.inf if the user is the Supervisor
Use USER GEORGE to set the name of user GEORGE
Use LIST USERS to see the list of users
Use DIR MASTER to return to Supervisor but the folder (except if you give Dir User which change to the user directory, so master has different user directory).
? USER.NAME$ return the name of the user. Master's name is the Window's user name.
A user other than the supervisor cannot do: Use of Win and Dos statements. Create files in a space other then the space of user (user folder and sub folders).

Users have no passwords. The user folder change to specific user, and each user may have a desktop.inf
use Files "inf" to see if there are files *.inf
use Edit "desktop.inf" to edit it as text file. The file if not exist created if we write something.
Shift F12 close editor without save changes (there is a same option in the dropdown menu of editor).
 

George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows did some work behind the scenes so the M2000 console slowed down. You can type END to close the program and then open it again.

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