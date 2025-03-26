M2000 Interpreter and Environment
Version 13 revision 33 active-X
Fix an old program which uses properties of com objects with indexes saved to arrays and inventories.

This is an example which was ok for version 8.9 (2017) but not for later versions until now.
// Using Inventories (lists)
declare form1 form
declare list1(3) combobox form form1
inventory controls, Enabled
For i=0 to 2 {
      Method list1(i), "move", 2000,1200+i*800, 5000,600
      with list1(i),"MenuStyle", True, "MenuWidth", 3000,  "MenuEnabled" as new list1_enabled()
      with list1(i),"label", "Menu"+str$(i), "list" as new list$()
      Method list1(i), "MenuItem","Command 1",True
      Method list1(i), "MenuItem","Command 2",false
      Method list1(i), "MenuItem","Command 3",True
      Append controls, i:=list$()
      Append Enabled, i:=list1_enabled()
      controls$(i)(0)="ok"+str$(i)
      Print type$(Enabled(i)()), type$(controls$(i)())
      if i=1 then Enabled(i)(1)=true
      Print list$(0), "ok", controls$(i)(0), Enabled(i)(1)
}
method form1, "show",1
declare list1() nothing
declare form1 nothing

This is an example which work until version 11:
// using tuple
declare form1 form
declare list1(3) combobox form form1
controls=(,)
Enabled=(,)
link controls, Enabled to controls$(), Enabled()
For i=0 to 2 {
      Method list1(i), "move", 2000,1200+i*800, 5000,600
      with list1(i),"MenuStyle", True, "MenuWidth", 3000,  "MenuEnabled" as new list1_enabled()
      with list1(i),"label", "Menu"+str$(i), "list" as new list$(), "MenuGroup","group_a"
      Method list1(i), "MenuItem","Command 1",True
      Method list1(i), "MenuItem","Command 2",false
      Method list1(i), "MenuItem","Command 3",True
      Append controls, (list$(),)
      Append Enabled, (list1_enabled(),)
      Print Type$(list1_enabled())
      controls$(i)(0)="ok"+str$(i)
      Print type$(Enabled(i)()), type$(controls$(i)())
      if i=1 then Enabled(i)(1)=true
      Print list$(0), "ok", controls$(i)(0), Enabled(i)(1)
}
method form1, "show",1
declare list1() nothing
declare form1 nothing







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
                 