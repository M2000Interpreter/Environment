M2000 Interpreter and Environment
Version 14 revision 18 active-X

1) Fix a bug from version 14 revision 15, which prevent to call subs inside lambda functions. I found the crypto module with problem and I found my mistake.
Test:
a=lambda ->{
	malakas()
	=100
	sub malakas()
		beep
	end sub
}
z=a()
? z=100 ' now is ok...
2) List show objects better (for those which are childs of an ExtControl)

3) #Reduce() like #fold() but simple one. Read Help
? (1,2,3,4,5,6,7,8,9,10)#reduce(lambda (a,b)->a*b)=3628800

4) Ctrl+4 | Ctrl+/  used for mark one or more lines as comment or uncomment if they are. This works for EditBox too.

5) I found and fix the mysterious mistake when we search to up direction (to previous paragraphs in a document) some time not move to the top one. Now works as expected. The same function used for searching subs so some fixes not worked for the subs too.

6) I make Default property for External controls (not my controls). This property register the control to form, so when we move the form, we get the focus on the header and when we leave the from the focus return to control. Because each external controls  is in a dedicated ExtControl object, which knows the receiver of the events (the GUIM2000 type form) we have to handle the property Default to ExtControl. So we place an asterisk before the name "*Default" and now M2000 knows that this is for the ExtControl. Look at the OPENGL module, and at the NINEBUTTONS module.

7) New class ASF_RegexEngine.cls (I change the internal name as Regex), from W. García, https://ecp-solutions.github.io/
 
The class used as is. Export to vb collection which we can handle them easy.
Although we can use the RegExp of VBScript using:
declare global ObjRegEx "VBscript.RegExp"

I Prefer the nice solution of ASF_RegexEngine.cls. 

The program is in Info.gsb as regexNew

This is the output and have the first two categories of tests (transfer to M2000 from VBA):

' -----------------------------------------------------------------------
' Test functions (each returns Boolean). For brevity they follow same
' logic used in earlier conversion: initialize, run Exec/Replace, assert.
' -----------------------------------------------------------------------

' Category 1

T_1_01_basic_literal_exact
ok: match for 'abc'
Len = 1
abc

T_1_02_basic_literal_mismatch
ok: expected no match for 'abx'

T_1_03_dot_matches_any
ok: '.' to match 'a'

T_1_04_dot_requires_at_least_one
ok: expected '.' not to match empty

T_1_05_dot_in_sequence
ok: 'a.c' to match 'abc'
Len = 1
abc

T_1_06_dot_dotall_true_newline
Use of dotAll: True
ok: dotAll True expected match across newline
Len = 1
a
c

T_1_07_dot_dotall_false_newline
ok: expected dotAll False expected no match across newline

T_1_08_anchor_dot_single
ok: ^.$ to match 'x'
Len = 1
x

T_1_09_anchor_dot_multi_fail
ok: expected ^.$ to fail on 'xy'


' Category 2 (Escapes)

T_2_01_escape_digit_true
ok: \d to match '5'
Len = 1
5

T_2_02_escape_digit_false
ok: expected \d not to match 'a'

T_2_03_escape_word_true
ok: \w+ to match 'hello_123'
Len = 1
hello_123

T_2_04_escape_word_false
ok: expected \w+ not to match '!@#'

T_2_05_escape_space_true
ok: \s+ to match whitespace
Len = 1
 	


T_2_06_escape_space_false
ok: expected \s+ not to match 'abc'

T_2_07_escape_lf
ok: \n to match vbLf
Len = 1



T_2_08_escape_cr
ok: \r to match vbCr
Len = 1



T_2_09_escape_tab
ok: \t to match vbTab
Len = 1
	

T_2_10_escape_escaped_metachar
ok: \. to match '.'
Len = 1
.

T_2_11_escape_d2_exec
ok: \d{2} to match '12'
12 match 12

T_2_12_escape_d2_partial_fail
ok: expected \d{2} to fail on '1a'




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