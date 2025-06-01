M2000 Interpreter and Environment
Version 14 revision 1 (updated - info file) active-X

A lot of updates.

1. Comp module now work as expected (the a(index)=a, b, c if b and c was objects not assigment perform for that b and c. The a(index):=a, b, c was ok, now the a(index)=a, b, c works for objects too.
2. Form X, Y statement works very nice (I found that sometimes a Linespac with -15 value happen; so I found what to do to correct it).
3. For groups/classes we can include some statements inside a block inside the group/class body. Although we can use a constructor, this constuctor executed after the execution of the class body.
module InAClass {
	class alfa {
		{ //  x and a are local to function alfa()
		 // but not part of object
		 // we can take numbers from stack
			read ? x=10
			a=today+x
		}
		date lastdate=a
	}
	alfa a
	? a.lastdate
	alfa b(100)
	? b.lastdate, type$(b.lastdate)
	? val(b.lastdate-a.lastdate -> double)=90
	? a is type alfa, b is type alfa	
	list
	push a  ' we push the objct a to stack of values
}
InAClass
// we pop value from stack of values using Read
read export1
// class alfa erased but the object can be used
print export1.lastdate

module InAGroup {
	function alfa {
	 	read ? x=10   // we can put these two lines
		a=today+x     // in a block inside group alfa, before date definition
		Group alfa {
			type: alfa
			date lastdate=a
		}
		=alfa
	}
	' we can't use "alfa a" because function alfa isn't "class function"
	' also function alfa is local but class alfa is global until eraded (at exit from InAClass)
	a=alfa()  
	? a.lastdate
	b=alfa(100)
	? b.lastdate, type$(b.lastdate)
	? val(b.lastdate-a.lastdate -> double)=90
	? a is type alfa, b is type alfa	
	list
	push a
}
InAGroup ' same result as module InAClass
read export2
print export2.lastdate

4. Private external functions:
class StartTime {
// check os before make functions
// but all declaretions found function at the first call
// so because we have constructor (we need it to call the function QueryUnbiasedInterruptTime)
// this block can be erased
{
	declare info information
	with info, "IsWindows7OrGreater" as win7
	if not win7 then Error "Minimum os: Windows 7"
}
private:
	declare QueryUnbiasedInterruptTime lib "Kernel32.QueryUnbiasedInterruptTime" { long long &a} as integer
	timebase=1000000&&
public:
	lastvalue=0&&
	value {
		long long ret
		if .QueryUnbiasedInterruptTime(&ret) then
			.lastvalue<=ret/.timebase
		end if		
		=uint(.lastvalue)
	}
class:
	module StartTime {
		// so we can do the test at the constructor.
		declare info information
		with info, "IsWindows7OrGreater" as win7
		if not win7 then Error "Minimum os: Windows 7"
		long long tbase, ret, ret1
		decimal a
		profiler
		a = timecount+1000@
		if .QueryUnbiasedInterruptTime(&ret) then {
			while uint(timecount)<a {}
			if .QueryUnbiasedInterruptTime(&ret1) then {
				.timebase<= 10^int(log(uint(ret1)-uint(ret)))
			}
		}
	}
}
// the StartTime return value
StartTime=StartTime()
print StartTime;"sec"
print int(StartTime/60/60);"hours", now
print round(StartTime/60/60/24, 1);"Days"
// we can get a copy
Another=Group(StartTime)
Print Another;"sec"
// using a fake pointer - is a reference to StartTime as weak pointer
pointer2StartTime->StartTime
Print Eval(pointer2StartTime);"sec"
// now we get a true pointer to a copy of StartTime
pointer2StartTime->group(StartTime)
Print Eval(pointer2StartTime);"sec"
wait 1000
Print StartTime-pointer2StartTime=>lastvalue=1 ' true

 
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

https://rosettacode.miraheze.org/wiki/M2000_Interpreter (544 tasks)
Old (not working rosettacode.org)
https://rosettacode.org/wiki/Category:M2000_Interpreter (544 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 