M2000 Interpreter and Environment
Version 13 revision 48 active-X

This is a work in progress (for the example, the code is perfect).

Example of the new Interface construction.

Using:
INTERFACE name, "{xx..xx}" {
}
we get a special enum object, with name, but the members are hidden, we can use this name|member to get the offset (start from 0, then 1...and internal we get the product times 4 then length of 32bit pointer).

So this is the example:

Structure GUID {
	p as long
	i as integer * 2
	b as byte * 8
}
Structure TYPEATTR {
	p as long
	i as integer * 2
	b as byte * 8
	tLCID as Long
	dwReserved as Long
	memidConstructor as Long
	memidDestructor as Long
	pstrSchema as Long
	cbSizeInstance  as Long
	typekind as Long
	cFuncs as Integer
	CVars  as Integer
	cImplTypes  as Integer
	cbSizeVft  as Integer
	cbAlignment as Integer
	wTypeFlags  as Integer
	wMajorVerNum  as Integer
	wMinorVerNum  as Integer
	tdescAlias as Long
	idldescType as Long
}

INTERFACE IUnknown, "{00000000-0000-0000-C000-000000000046}" {
	QueryInterface : {long riid, long myptr}
	AddRef
	Release
}
INTERFACE Idispatch, "{00020400-0000-0000-C000-000000000046}" {
	IUnknown
	GetTypeInfoCount
	GetTypeInfo : {long iTInfo, long lcid, Long ITypeInfo }
	GetIDsOfNames
	Invoke
}
INTERFACE ITypeInfo, "{00020401-0000-0000-C000-000000000046}" {
	IUnknown
	GetTypeAttr : {long ppTypeAttr}
	GetTypeComp
	GetFuncDesc
	GetVarDesc
	GetNames : {long memid, long PTRrgBstrNames, long cMaxNames, long &pcNames}
	GetRefTypeOfImplType
	GetImplTypeFlags
	GetIDsOfNames
	Invoke
	GetDocumentation
	GetDllEntry
	GetRefTypeInfo
	AddressOfMember
	CreateInstance
	GetMops
	GetContainingTypeLib
	ReleaseTypeAttr : {long pTypeAttr}  ' void
	ReleaseFuncDesc
	ReleaseVarDesc
}
declare math math  ' math is an internal object, always with same pointer
object obj=@iDISPATCH_OBJ(math)
Declare LCID_def1 Lib "kernel32.GetSystemDefaultLCID" { }
long LCID_DEF=LCID_def1() mod 0x10000
pRINT  LCID_DEF
object oTypeInfo

print Idispatch.GetTypeInfo(obj, 0&, LCID_DEF, VarPtr(oTypeInfo))

buffer clear riid as GUID
IDLfromString(riid, "{00020401-0000-0000-C000-000000000046}")
object ret1
HEX  IUnknown.QueryInterface(oTypeInfo, riid(0), varptr(ret1))
Print  ret1 is oTypeInfo
list 
long ppTypeAttr
ret11= iTypeInfo.GetTypeAttr(oTypeInfo, VarPtr(ppTypeAttr))
hex ret11
buffer clear  copyAttr as TYPEATTR
if ret11=0 then
	? "ok", ppTypeAttr
	// copy from ppTypeAttr to copyAttr(0)
	method copyAttr, "FillDataFromMem", ppTypeAttr
	print "ok"
	call void  iTypeInfo.ReleaseTypeAttr(oTypeInfo, ppTypeAttr)
	print "released ok"
end if
print  @getClassId(copyAttr)
END
// return for sure the iDispatch interface - if exist.
FUNCTION iDISPATCH_OBJ(K)
	local riid
	buffer clear riid as LONG*4
	return riid, 0:=0x0002_0400, 1:=0,  2:=0xc0, 3:=0x4600_0000
	local object ret
	local long Hresult
	=ret
	try {
		Hresult=Idispatch.QueryInterface(K, riid(0), varptr(ret))
		=ret
	}
END FUNCTION
FUNCTION getClassId(riid as buffer)
	local i, last$=hex$(riid|b[0],1)+hex$(riid|b[1],1)+"-"
	for i=2 to 7:last$+=hex$(riid|b[i],1):next
	="{"+hex$(riid|p)+"-"+hex$(riid|i[0],2)+"-"+hex$(riid|i[1],2)+"-"+last$+"}"	
END FUNCTION
SUB IDLfromString(riid as buffer, st$)
	st$=filter$(st$,"{}")
	local i, p$=leftpart$(st$, "-")
	st$=rightpart$(st$,"-")
	riid|p=val("0x"+p$)
	p$=leftpart$(st$, "-"): st$=rightpart$(st$,"-")
	riid|i[0]=val("0x"+p$)
	p$=leftpart$(st$, "-"): st$=rightpart$(st$,"-")
	riid|i[1]=val("0x"+p$)
	p$=left$(st$,2): st$=mid$(st$,3)
	riid|b[0]=val("0x"+p$)
	p$=left$(st$,2): st$=mid$(st$,4)
	riid|b[1]=val("0x"+p$)
	for i=2 to 7
		p$=left$(st$,2): st$=mid$(st$,3)
		riid|b[i]=val("0x"+p$)
	next
END SUB




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

https://rosettacode.org/wiki/Category:M2000_Interpreter (544 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 