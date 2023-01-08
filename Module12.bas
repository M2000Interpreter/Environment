Attribute VB_Name = "Module12"
Option Explicit
 
Private Const OPEN_EXISTING             As Long = &H3
Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const GENERIC_READ              As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Private Const FILE_BEGIN                As Long = &H0
Private Const RT_ICON                   As Long = &H3
Private Const RT_GROUP_ICON             As Long = &HE
Private Const RT_RCDATA                 As Long = 10&
Private Type ICONDIRENTRY
    bWidth          As Byte
    bHeight         As Byte
    bColorCount     As Byte
    bReserved       As Byte
    wPlanes         As Integer
    wBitCount       As Integer
    dwBytesInRes    As Long
    dwImageOffset   As Long
End Type
 
Private Type ICONDIR
    idReserved      As Integer
    idType          As Integer
    idCount         As Integer
End Type
 
Private Type GRPICONDIRENTRY
    bWidth          As Byte
    bHeight         As Byte
    bColorCount     As Byte
    bReserved       As Byte
    wPlanes         As Integer
    wBitCount       As Integer
    dwBytesInRes    As Long
    nID             As Integer
End Type
 
Private Type GRPICONDIR
    idReserved      As Integer
    idType          As Integer
    idCount         As Integer
    idEntries()     As GRPICONDIRENTRY
End Type
 
'Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal lFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal lFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceW" (ByVal pFileName As Long, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceW" (ByVal lUpdate As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceW" (ByVal lUpdate As Long, ByVal fDiscard As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 
Public Function CopyM2000Exe(NewName$, Optional copyHelp) As Boolean
' copy the app.EXEName from appdir$
On Error GoTo 1234
Dim ap$, bp$, cp1 As Boolean, R4, there$, D As New Document, mycoder As New coder
there$ = ExtractPath(NewName$, , True)
If there$ = "" Then
    If ExtractPath(NewName$) <> "" Then
        MyEr "Path not exist", "Δεν υπάρχει ο φάκελος"
        Exit Function
    End If
    there$ = UserPath
Else
    NewName$ = ExtractNameOnly(NewName$)
End If

Dim R$
R$ = GetLongName(App.Path)
If Right(R$, 1) <> "\" Then R$ = R$ + "\"
ap$ = "\\?\" + R$ + App.EXENAME + ".exe"
bp$ = "\\?\" + there$ + NewName$ + ".exe"
If CFname(there$ + NewName$ + ".exe") <> "" Then
    KillFile there$ + NewName$ + ".exe"
End If
If 0 <> CopyFile(StrPtr(ap$), StrPtr(bp$), 1) Then
    If CFname(R$ + "M2000.dll") <> "" Then
        ap$ = "\\?\" + R$ + "M2000.dll"
    ElseIf CFname(R$ + "lib.bin") <> "" Then
        ap$ = "\\?\" + R$ + "lib.bin"
    End If
    bp$ = "\\?\" + there$ + "lib.bin"
    If CFname(there$ + "lib.bin") <> "" Then
        KillFile there$ + "lib.bin"
    End If
    If 0 <> CopyFile(StrPtr(ap$), StrPtr(bp$), 1) Then
        If Not IsMissing(copyHelp) Then
            If CFname(R$ + "help2000utf8.dat") <> "" Then
                If CFname(there$ + "help2000utf8.dat") <> "" Then
                    KillFile there$ + "help2000utf8.dat"
                End If
                ap$ = "\\?\" + R$ + "help2000utf8.dat"
                bp$ = "\\?\" + there$ + "help2000utf8.dat"
                If 0 = CopyFile(StrPtr(ap$), StrPtr(bp$), 1) Then
                    MyEr "Can't copy help2000utf8.dat", "δεν μπορώ να αντιγράψω το help2000utf8.dat"
                End If
            End If
        End If

        If CFname(there$ + NewName$ + ".ico") <> "" Then
            ap$ = there$ + NewName$ + ".exe"
            bp$ = there$ + NewName$ + ".ico"
            CopyM2000Exe = ChangeIcon(ap$, bp$)
            
        Else
            MyEr "Can't find  " + ExtractNameOnly(ap$) + ".ico", "δεν βρήκα το " + ExtractNameOnly(ap$) + ".ico"
        End If
        If CFname(there$ + NewName$ + ".gsb") = "" Then
             
            D.SaveUnicodeOrAnsi there$ + NewName$ + ".gsb", 2
        Else
           D.ReadUnicodeOrANSI there$ + NewName$ + ".gsb"
           '' bb() = D.textDoc 'mycoder.must1()
           ap$ = D.textDoc
           Dim LCID, i As Long
           Dim ResHdl As Long: ResHdl = BeginUpdateResource(StrPtr("\\?\" + there$ + NewName$ + ".exe"), 0)
           If ResHdl Then
            On Error GoTo 1
           Debug.Print Len(ap$)
           For i = 1 To Len(ap$) \ 8192
                bp$ = space$(8192)
                LSet bp$ = Mid$(ap$, (i - 1) * 8192 + 1, 8192)
                UpdateResource ResHdl, RT_RCDATA, i + 100, LCID, StrPtr(bp$), LenB(bp$)
           Next i
           If Len(ap$) Mod 8192 > 0 Then
                bp$ = space$(8192)
                LSet bp$ = Mid$(ap$, (i - 1) * 8192 + 1)
                UpdateResource ResHdl, RT_RCDATA, i + 100, LCID, StrPtr(bp$), LenB(bp$)
           End If
1          Debug.Print EndUpdateResource(ResHdl, IIf(Err, 1, 0)), i
           End If
        End If
    Else
        MyEr "Can't copy " + ExtractName(ap$), "δεν μπορώ να αντιγράψω το " + ExtractName(ap$)
    End If
Else
    MyEr "Can't copy " + App.EXENAME + ".exe", "δεν μπορώ να αντιγράψω το " + App.EXENAME + ".exe"
End If
1234 Err.Clear
End Function
 
Public Function ChangeIcon(ByVal strExePath As String, ByVal strIcoPath As String) As Boolean
    Dim lFile               As Long
    Dim lUpdate             As Long
    Dim lRet                As Long
    Dim i                   As Integer
    Dim tICONDIR            As ICONDIR
    Dim tGRPICONDIR         As GRPICONDIR
    Dim tICONDIRENTRY()     As ICONDIRENTRY
   
    Dim bIconData()         As Byte
    Dim bGroupIconData()    As Byte
   
    lFile = CreateFile(StrPtr("\\?\" + strIcoPath), GENERIC_READ, 0, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
   
    If lFile = INVALID_HANDLE_VALUE Then
        ChangeIcon = False
        CloseHandle (lFile)
        Exit Function
    End If
   
    Call ReadFile(lFile, tICONDIR, Len(tICONDIR), lRet, ByVal 0&)
   
    ReDim tICONDIRENTRY(tICONDIR.idCount - 1)
   
    For i = 0 To tICONDIR.idCount - 1
        Call ReadFile(lFile, tICONDIRENTRY(i), Len(tICONDIRENTRY(i)), lRet, ByVal 0&)
    Next i
   
    ReDim tGRPICONDIR.idEntries(tICONDIR.idCount - 1)
   
    tGRPICONDIR.idReserved = tICONDIR.idReserved
    tGRPICONDIR.idType = tICONDIR.idType
    tGRPICONDIR.idCount = tICONDIR.idCount
   
    For i = 0 To tGRPICONDIR.idCount - 1
        tGRPICONDIR.idEntries(i).bWidth = tICONDIRENTRY(i).bWidth
        tGRPICONDIR.idEntries(i).bHeight = tICONDIRENTRY(i).bHeight
        tGRPICONDIR.idEntries(i).bColorCount = tICONDIRENTRY(i).bColorCount
        tGRPICONDIR.idEntries(i).bReserved = tICONDIRENTRY(i).bReserved
        tGRPICONDIR.idEntries(i).wPlanes = tICONDIRENTRY(i).wPlanes
        tGRPICONDIR.idEntries(i).wBitCount = tICONDIRENTRY(i).wBitCount
        tGRPICONDIR.idEntries(i).dwBytesInRes = tICONDIRENTRY(i).dwBytesInRes
        tGRPICONDIR.idEntries(i).nID = i + 1
    Next i
   
    lUpdate = BeginUpdateResource(StrPtr("\\?\" + strExePath), False)
    For i = 0 To tICONDIR.idCount - 1
        ReDim bIconData(tICONDIRENTRY(i).dwBytesInRes)
        SetFilePointer lFile, tICONDIRENTRY(i).dwImageOffset, ByVal 0&, FILE_BEGIN
        Call ReadFile(lFile, bIconData(0), tICONDIRENTRY(i).dwBytesInRes, lRet, ByVal 0&)
   
        If UpdateResource(lUpdate, RT_ICON, tGRPICONDIR.idEntries(i).nID, 0, bIconData(0), tICONDIRENTRY(i).dwBytesInRes) = False Then
            ChangeIcon = False
            CloseHandle (lFile)
            Exit Function
        End If
       
    Next i
 
    ReDim bGroupIconData(6 + 14 * tGRPICONDIR.idCount)
    CopyMemory ByVal VarPtr(bGroupIconData(0)), ByVal VarPtr(tICONDIR), 6
 
    For i = 0 To tGRPICONDIR.idCount - 1
        CopyMemory ByVal VarPtr(bGroupIconData(6 + 14 * i)), ByVal VarPtr(tGRPICONDIR.idEntries(i).bWidth), 14&
    Next
               
    If UpdateResource(lUpdate, RT_GROUP_ICON, 1, 0, ByVal VarPtr(bGroupIconData(0)), UBound(bGroupIconData)) = False Then
        ChangeIcon = False
        CloseHandle (lFile)
        Exit Function
    End If
   
    If EndUpdateResource(lUpdate, False) = False Then
        ChangeIcon = False
        CloseHandle (lFile)
    End If
 
    Call CloseHandle(lFile)
    ChangeIcon = True
End Function

Public Function WriteResData(ByVal ResSubID As Long, ResTypeOrID, BytesOrString, FileNameToExeOrDll As String, Optional ByVal LCID As Long) As Boolean
  
  Dim ResTyp As Long: ResTyp = IIf(VarType(ResTypeOrID) = vbString, StrPtr(ResTypeOrID), ResTypeOrID)
  'Dim Data() As Byte: Data = IIf(IsArray(BytesOrString), BytesOrString, StrConv(BytesOrString, vbFromUnicode))
  'Dim LenDat As Long: LenDat = UBound(Data) - LBound(Data) + 1
  'Dim lpData As Long: If LenDat Then lpData = VarPtr(Data(LBound(Data)))
  Dim ResHdl As Long: ResHdl = BeginUpdateResource(StrPtr("\\?\" + FileNameToExeOrDll), 0)
  
  If ResHdl Then
     On Error GoTo 1
     
     Debug.Print UpdateResource(ResHdl, ResTyp, ResSubID, LCID, StrPtr(BytesOrString), LenB(BytesOrString))
     
1    WriteResData = EndUpdateResource(ResHdl, IIf(Err, 1, 0))
  End If
End Function
Public Function UpdateManifestData(sManifestContent As String, FileNameToExeOrDll As String) As Boolean
  Dim IsExe As Boolean: IsExe = LCase$(Right$(FileNameToExeOrDll, 4)) = ".exe"
  Const RT_MANIFEST As Long = 24
  
  UpdateManifestData = WriteResData(IIf(IsExe, 1, 2), RT_MANIFEST, sManifestContent, FileNameToExeOrDll, 1033)
End Function

Public Function RemoveManifestData(FileNameToExeOrDll As String) As Boolean
  RemoveManifestData = UpdateManifestData("", FileNameToExeOrDll)
End Function
