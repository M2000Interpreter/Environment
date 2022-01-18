VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "recDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private LastPath As String
Private Const MAX_PATH As Long = 260
Private Const MAX_PATH_UNICODE As Long = 260 * 2 - 1
Private onetime As Boolean
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type
Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName(MAX_PATH * 2 - 1) As Byte
cAlternate(14 * 2 - 1) As Byte
End Type
Private Declare Function CompareString Lib "kernel32" Alias "CompareStringW" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long

Dim myfile$(), myLevel() As Integer
Public Event DirFinished()
Public Event feedback(Filename$) ' dir or filenames...
Dim Filetypes2find As String  ' input  "jpg|bmp"  or ""
Dim mrec As Boolean
Public abort As Boolean
Public SortType As Long ' 0 is the default, 1 by name, 2 by type
Public IncludedFolders As Boolean
Public Nofiles As Boolean
Public Pattern As String
Private mTopFolder As String
Dim InUseItems As Long
Public LevelStop As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32" _
 Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
 lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
 lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) _
 As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
  Alias "GetLogicalDriveStringsW" _
  (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long

Private Declare Function GetLongPathName Lib "kernel32" _
   Alias "GetLongPathNameW" _
  (ByVal lpszShortPath As Long, _
   ByVal lpszLongPath As Long, _
   ByVal cchBuffer As Long) As Long
Public lastTimeStamp As Date
Public lastTimeStamp2 As Date
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public UseUTC As Boolean
Private Function FileTimeToDate(File_Time As FILETIME) As Date
Dim System_Time As SYSTEMTIME
Dim local_Time As FILETIME
    ' post from JasonC in vbforums
    ' Convert the FILETIME structure into a Date.
    ' changed to convert from UTC to Local Time
    If UseUTC Then
    Call FileTimeToLocalFileTime(File_Time, local_Time)
    
    Call FileTimeToSystemTime(local_Time, System_Time)
    Else
    Call FileTimeToSystemTime(File_Time, System_Time)
    End If
    FileTimeToDate = SystemTimeToDate(System_Time)
    
End Function
 
Private Function SystemTimeToDate(System_Time As SYSTEMTIME) As Date
    
     With System_Time
     SystemTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + _
                   TimeSerial(.wHour, .wMinute, .wSecond)
    End With
   
    
End Function

Public Function GetLongName(strTest As String) As String
   Dim sLongPath As String
   Dim buff As String
   Dim cbbuff As Long
   Dim Result As Long
 
   buff = space$(MAX_PATH_UNICODE)
   cbbuff = Len(buff)
 
   Result = GetLongPathName(StrPtr(strTest), StrPtr(buff), cbbuff)
 
   If Result > 0 Then
      sLongPath = Left$(buff, Result)
   End If
 
   GetLongName = sLongPath
 
End Function
  Function TotalDiskSpace(DriveLetter As String) As Currency
' Returns the total storage capacity for a drive
On Error GoTo 100
Dim Dletter As String, x As Long

    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberofFreeClusters As Long
    Dim TotalClusters As Long

    Dletter = Left(DriveLetter, 1) & ":\"
    x = GetDiskFreeSpace(Dletter, SectorsPerCluster, _
      BytesPerSector, NumberofFreeClusters, TotalClusters)
    
    If x = 0 Then 'Error occurred
        TotalDiskSpace = -99 'Assign an arbitrary error value
        Exit Function
    End If
    TotalDiskSpace = _
    Int(CCur(uintnew(SectorsPerCluster) * uintnew(BytesPerSector)) * uintnew(TotalClusters) / 1024@ / 1024@)
    Exit Function
100
    MsgBox "error at totaldiskspace"
End Function
Function NumberofDrives() As Integer

    Dim i As Integer
 
    
    Dim Buffer As String
    Buffer = String$(260, 0)
    Dim BuffLen As Long
    Dim TheDrive As String
    Dim DriveCount As Integer
    BuffLen = GetLogicalDriveStrings(MAX_FILENAME_LEN, StrPtr(Buffer))
    DriveCount = 0
' Search for a null -- which separates the drives
    For i = 1 To BuffLen
        If Asc(Mid$(Buffer, i, 1)) = 0 Then _
          DriveCount = DriveCount + 1
    Next i
    NumberofDrives = DriveCount
    Exit Function
End Function

Function DriveName(index As Integer) As String
    Dim i As Integer
    Dim Buffer As String
    Buffer = String$(260, 0)
    Dim BuffLen As Long
    Dim TheDrive As String
    Dim DriveCount As Integer
    BuffLen = GetLogicalDriveStrings(MAX_FILENAME_LEN, StrPtr(Buffer))

    TheDrive = vbNullString
    DriveCount = 0
    For i = 1 To IIf(BuffLen > 260, 260, BuffLen)
        If Asc(Mid$(Buffer, i, 1)) <> 0 Then _
          TheDrive = TheDrive & Mid$(Buffer, i, 1)
        If Asc(Mid$(Buffer, i, 1)) = 0 Then 'null separates drives
            DriveCount = DriveCount + 1
            If DriveCount = index Then
                DriveName = UCase(Left(TheDrive, 1))
                Exit Function
            End If
            TheDrive = vbNullString
        End If
    Next i
End Function
    

Function FreeDiskSpace(DriveLetter As String) As Currency
' Returns the number of free bytes for a drive
On Error GoTo 400
Dim Dletter As String, x As Currency
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberofFreeClusters As Long
    Dim TotalClusters As Long

    Dletter = Left(DriveLetter, 1) & ":\"
    x = uintnew(GetDiskFreeSpace(Dletter, SectorsPerCluster, _
      BytesPerSector, NumberofFreeClusters, TotalClusters))
    
    If x = 0 Then 'Error occurred
        FreeDiskSpace = -99@ 'Assign an arbitrary error value
        Exit Function
    End If
    FreeDiskSpace = uintnew(SectorsPerCluster) * uintnew(BytesPerSector) * uintnew(NumberofFreeClusters)
      Exit Function
400
      MsgBox "error at freediskspace"
      
End Function




Public Property Get Path() As String
If Len(LastPath) = 0 Then
Else
If Right$(LastPath, 1) <> "\" Then
Path = LastPath & "\"
Else
Path = LastPath
End If
End If
End Property

Public Property Get IsEmpty() As Boolean

IsEmpty = InUseItems = 0

End Property
Public Property Get listcount() As Long
If IsEmpty Then
listcount = 0
Else
listcount = InUseItems  ' UBound(myfile$())
End If
End Property
Public Sub MarkLevel(item As Long)
myLevel(item) = -Abs(myLevel(item))
End Sub
Public Sub unMarkLevel(item As Long)  ' this mark can be used for multiple choosing or for hide/show.
myLevel(item) = Abs(myLevel(item))    'prefered the first one because glist cannot hide item...yet
End Sub
Public Function ReadLevel(item As Long) As Integer
ReadLevel = Abs(myLevel(item))
End Function
Public Function ReadMark(item As Long) As Boolean
ReadMark = Sgn(myLevel(item)) = -1
End Function
Public Function FindFolder(item As Long) As String
Dim i As Long, lev As Integer
lev = Abs(myLevel(item))
If lev = 1 Then
FindFolder = Path
Else
For i = item - 1 To 0 Step -1
If Abs(myLevel(i)) < lev Then Exit For
Next i
 FindFolder = Mid$(myfile$(i), 2) & "\"
End If
End Function

Public Sub AddAFile(File$, Level As Integer)
If Not onetime Then onetime = True: once
If InUseItems >= UBound(myfile$()) Then
ReDim Preserve myfile$(UBound(myfile$()) * 2)
ReDim Preserve myLevel(UBound(myLevel()) * 2) As Integer
End If
 myfile$(InUseItems) = File
myLevel(InUseItems) = Level  ' level must be not 0, mark = false
InUseItems = InUseItems + 1


End Sub
Public Sub ClearList()
ReDim myfile$(100)
ReDim myLevel(100) As Integer
InUseItems = 0
End Sub
Public Property Get list(index As Long) As String
If index < 0 Then
list = vbNullString
Else
If Not onetime Then onetime = True: once
list = myfile$(index)
End If
End Property
Public Property Let list(index As Long, RHS As String)
myfile$(index) = RHS
End Property

Private Function CheckFile(ByVal expan$) As Boolean
Dim ee As Integer
If Filetypes2find = "||" Then
CheckFile = True
Else
CheckFile = InStr(Filetypes2find, "|" & UCase(Mid$(expan$, InStrRev(expan$, ".") + 1)) & "|") > 0
End If
End Function

Public Function Dir2(Optional ByVal ff As String = vbNullString, Optional FileType$ = vbNullString, Optional Recursive As Boolean = True, Optional Append As Boolean = False) As String
Dim i As Integer
If LastPath <> "" And ff = vbNullString Then

Else
LastPath = ff

Static ndx As Long
If ff = vbNullString Then
ClearList

ndx = 0
For i = 1 To NumberofDrives
 'Me.additem DriveName(i) & ":\" & CStr(TotalDiskSpace(DriveName(i)))
 AddAFile Chr$(2) + DriveName(i) & ":", 1 + Abs(TotalDiskSpace(DriveName(i)) < 0) 'use level as not enabled
Next i
If listcount > 0 Then Dir2 = list(0)
RaiseEvent DirFinished
Exit Function
End If
'If Right$(ff, 1) <> "\" Then ff = ff & "\"
'' filetype$ input  "jpg|bmp"  or ""

End If
Filetypes2find = "|" & UCase(FileType$) & "|"
mrec = Recursive
Dim s As String


If ff <> "" Then
If Append Then
ndx = listcount
Else
            ClearList ' CHANGE THIS TO APPEND...
           ndx = 0
    End If
                abort = False
' ff must be a path
' so strip \
            If Right$(ff, 1) = "\" Then ff = Left$(ff, Len(ff) - 1)
        
                If look(1, ff, (ndx), IncludedFolders, Nofiles) > 0 Then
               
                    Dir2 = list(ndx)
                    ndx = ndx + 1
     
     
                    End If
           
Else
            If listcount > ndx Then
            Dir2 = list(ndx)
            ndx = ndx + 1   ' read with auto increment
            Else
            ndx = 0  ' reset
                Dir2 = vbNullString ' this is the end
            End If
End If
RaiseEvent DirFinished
End Function
Public Function FileAttr(a$, attr As Long) As Boolean
Dim wfd As WIN32_FIND_DATA
On Error GoTo there1
Dim fhandle As Long

fhandle = FindFirstFile(StrPtr(a$), wfd)

If (fhandle > 0) Then
attr = wfd.dwFileAttributes
lastTimeStamp = FileTimeToDate(wfd.ftLastWriteTime)
FileAttr = True
FindClose fhandle
Else
Dim b$
If Right$(a$, 1) = "\" Then
b$ = a$ & "."
Else
b$ = a$ & "\."
End If
fhandle = FindFirstFile(StrPtr(b$), wfd)

If (fhandle > 0) Then
attr = wfd.dwFileAttributes
lastTimeStamp = FileTimeToDate(wfd.ftLastWriteTime)
FileAttr = True
FindClose fhandle
Else
If GetLongName(a$) <> "" Then
attr = CLng(GetAttr(a$))
lastTimeStamp = 0
FileAttr = True
End If
End If
End If
Exit Function
there1:
End Function

Public Function ExistFile(a$) As Boolean
Dim wfd As WIN32_FIND_DATA
On Error GoTo there2
Dim fhandle As Long
fhandle = FindFirstFile(StrPtr(a$), wfd)
ExistFile = (fhandle > 0)
If ExistFile Then FindClose fhandle: lastTimeStamp = FileTimeToDate(wfd.ftLastWriteTime): lastTimeStamp2 = FileTimeToDate(wfd.ftCreationTime)
Exit Function
there2:
End Function
Public Function ExistFolder(a$) As Boolean
Dim wfd As WIN32_FIND_DATA
On Error GoTo there3
Dim fhandle As Long
fhandle = FindFirstFile(StrPtr(a$ & "\*.*"), wfd)
ExistFolder = (fhandle > 0)
If ExistFolder Then FindClose fhandle: lastTimeStamp = FileTimeToDate(wfd.ftLastWriteTime): Exit Function
ExistFolder = ExistFolder And (wfd.dwFileAttributes And 16) > 0

Exit Function
there3:
End Function

Public Function Dir1(ByVal ThisFile As String, curdir As String) As String

    If curdir = vbNullString Or ExtractPath(ThisFile) = ThisFile Then
    
      If ExistFolder(ThisFile) Then
      Dir1 = ThisFile
      End If
      
      Exit Function
    End If

    If ExtractPath(ThisFile) = vbNullString Then ThisFile = curdir & ThisFile
    If ExistFile(ThisFile) Then Dir1 = ThisFile
End Function
Private Function look(ByVal Level As Integer, ByVal folder As String, Optional Offset As Long = 0, Optional foldersincluded As Boolean = False, Optional foldersonly As Boolean = False) As Long
Dim lookpath As String, guard As String, mf$(), maxf As Long, k As Long
If LevelStop > 0 Then If Level > LevelStop Then Exit Function
Dim fold$(), fff As Long, addthis As Long, mlen As Long
Dim Filter As Long
  Dim wfd As WIN32_FIND_DATA, fhandle As Long
On Error GoTo 1000:
Dim thatmax As Long, thatmaxtoo As Long
thatmax = 20
thatmaxtoo = 20
ReDim mf$(1 To thatmax)
ReDim fold$(1 To thatmaxtoo)
If (Right$("\" & folder, 2) = "\." Or Right$("\" & folder, 3) = "\..") Then Exit Function
'  The paths need to have the correct unicode lookpath prefix "\\?\" for local HDs or "\\?\UNC" for network shares.
If Pattern <> "" Then
If Left$(folder, 2) = "\\" Then
fhandle = FindFirstFile(StrPtr(folder & "\" & Pattern), wfd)
Else
fhandle = FindFirstFile(StrPtr("\\?\" & folder & "\" & Pattern), wfd)
End If

Else
If Left$(folder, 2) = "\\" Then
fhandle = FindFirstFile(StrPtr(folder & "\*.*"), wfd)
Else
fhandle = FindFirstFile(StrPtr("\\?\" & folder & "\*.*"), wfd)
End If
End If
' if fhandle = -1 then we have a failure but this isn't the end of search
' for now we exit from such situation.
' maybe the better is to wait for a period of 100ms before we quit...reading again.
If fhandle > 0 Then
lookpath = StripNulls(wfd.cFileName)
End If
If Level = 3 Then mywait Basestack1, 5

While lookpath <> ""
If Pattern <> "" Then
If StripNulls(wfd.cAlternate) <> "" Then
If Not LCase(lookpath) Like LCase(Pattern) Then
      
    GoTo conthere
    End If
End If
End If

If fff Mod 30 = 29 Then mywait Basestack1, 1
    If isdirP(wfd) Or lookpath = ".." Then
         If lookpath = "." Or lookpath = ".." Then
        Else
        fff = fff + 1
        If fff > thatmaxtoo Then thatmaxtoo = thatmaxtoo * 2: ReDim Preserve fold$(1 To thatmaxtoo)
        If foldersonly Then RaiseEvent feedback("\" & lookpath)
        Select Case SortType
        Case 1, 2
        fold$(fff) = lookpath  ' no type for folder
        Case Else
        fold$(fff) = myFILEDATeTIME(wfd.ftLastWriteTime) + Format$(fff, "00000") + lookpath
        End Select
        End If
    ' do nothing
    ElseIf Not foldersonly Then
    If CheckFile(lookpath) Then
        maxf = maxf + 1
         If maxf > thatmax Then thatmax = thatmax * 2: ReDim Preserve mf$(1 To thatmax)
         Select Case SortType  ' case sensitive
         Case 1
         mf$(maxf) = lookpath
         Case 2
         mf$(maxf) = LCase(ExtractType(lookpath)) & "#" & lookpath
         Case Else
         ' maybe here filedatatime can be converted to a unicode reader
         ' I am sure that this isn't ok...for unicode
         mf$(maxf) = myFILEDATeTIME(wfd.ftLastWriteTime) + Format$(maxf, "00000") + lookpath
        End Select
        End If
    End If
conthere:
  If FindNextFile(fhandle, wfd) = 0& Then
  lookpath = vbNullString
  Else
    lookpath = StripNulls(wfd.cFileName)
  
    End If
Wend
FindClose fhandle
' #1
' maybe we can hide that with a switch
 If Level = 1 And folder <> "" And ((LCase(folder & "\") <> LCase(TopFolder)) Or (TopFolder = "\")) Then AddAFile "...", Level: addthis = addthis + 1  '  may we needaswitch here to exclude that
If foldersonly Or foldersincluded Then
     ' #2
    If fff > 1 Then
    If SortType <> 0 Then
    QuickSortItemsFileNames fold$(), 1, fff      ' we can sort the folders here
    Else
    QuickSortExtended fold$(), 1, fff
    End If
    End If
End If
    ' 3#  for each folder we  want to put  files  first and folders + files later...(this is my way)
If maxf > 0 Then

    If SortType <> 0 Then
    QuickSortItemsFileNames mf$(), 1, maxf
    Else
    QuickSortExtended mf$(), 1, maxf
    End If
    Select Case SortType
    Case 1
    For k = 1 To maxf
        AddAFile mf$(k), Level
    Next k
    Case 2
    For k = 1 To maxf
        AddAFile Mid$(mf$(k), InStr(mf$(k), "#") + 1), Level
    Next k
    Case Else
    For k = 1 To maxf
        AddAFile Mid$(mf$(k), 20), Level
    Next k
    End Select
  
  If Level < 3 Then RaiseEvent feedback(folder & " (" & CStr(maxf) & ") files")
End If
If Not mrec Then
If foldersonly Or foldersincluded Then
    If fff > 0 Then
    Select Case SortType
    Case 1, 2
        For k = 1 To fff
            AddAFile Chr$(2) + folder & "\" & fold$(k), Level
    Next k
    Case Else
        For k = 1 To fff
            AddAFile Chr$(2) + folder & "\" & Mid$(fold$(k), 20), Level
        Next k
   End Select
        maxf = maxf + fff  'IF FOLDERONLY THEN MAXF=0
   
    End If
End If
Else
If fff > 0 Then
Dim i As Long
For i = 1 To fff
If foldersincluded Or foldersonly Then
Select Case SortType
Case 1, 2
AddAFile Chr$(2) + folder & "\" & fold$(i), Level: maxf = maxf + 1
Case Else
AddAFile Chr$(2) + folder & "\" & Mid$(fold$(i), 20), Level: maxf = maxf + 1
End Select
End If
Select Case SortType
Case 1, 2
maxf = maxf + look(Level + 1, folder & "\" & fold$(i), maxf + Offset, IncludedFolders, Nofiles)

Case Else
maxf = maxf + look(Level + 1, folder & "\" & Mid$(fold$(i), 20), maxf + Offset, IncludedFolders, Nofiles)
End Select
If abort Then Exit For
Next i
End If
End If


look = maxf + addthis
Exit Function
1000:
look = 0
End Function


Public Sub QuickSortExtended(Arr() As String, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
Dim Piv As String, tmp As String '<- adjust types here, when switching to something different than Long
     If UB - LB = 1 Then
     M1 = LB
      If Arr(M1) > Arr(UB) Then tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = tmp
      Exit Sub
     Else
       M1 = (LB + UB) \ 2
             If Arr(M1) = Arr(LB) Then
                M2 = UB - 1
                M1 = LB
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If Arr(UB) < Arr(LB) Then tmp = Arr(LB): Arr(LB) = Arr(UB): Arr(UB) = tmp
                        Exit Sub
                    End If
                Loop Until Arr(M1) <> Arr(LB)
                Piv = Arr(M1)
                If M1 > LB Then If Arr(LB) > Piv Then Arr(M1) = Arr(LB): Arr(LB) = Piv: Piv = Arr(M1)
            Else
                Piv = Arr(M1)
                M1 = LB
                Do While (Arr(M1) < Piv): M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While (Arr(M2) > Piv): M2 = M2 - 1: Loop
      If M1 <= M2 Then
        tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = tmp 'swap
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While (Arr(M1) < Piv): M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortExtended Arr, LB, M2
    If M1 < UB Then QuickSortExtended Arr, M1, UB
End Sub
Private Sub QuickSortItemsFileNames(Arr() As String, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
Dim Piv As String, tmp As String '<- adjust types here, when switching to something different than Long
     If UB - LB = 1 Then
     M1 = LB
       If compareStr4(Arr(M1), Arr(UB)) = 1 Then tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = tmp
      Exit Sub
     Else
       M1 = (LB + UB) \ 2 '+ 1
       Piv = Arr(M1)
             If compareStr4(Arr(LB), Piv) = 0 Then
                M2 = UB - 1
                M1 = LB
                Piv = Arr(LB)
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If compareStr4(Arr(UB), Piv) = -1 Then tmp = Arr(LB): Arr(LB) = Arr(UB): Arr(UB) = tmp
                        Exit Sub
                    End If
                Loop Until compareStr4(Arr(M1), Piv)
                Piv = Arr(M1)
                If M1 > LB Then If compareStr4(Arr(LB), Piv) = 1 Then tmp = Arr(LB): Arr(LB) = Arr(M1): Arr(M1) = tmp: Piv = Arr(M1)
            Else
                Piv = Arr(M1)
                M1 = LB
                Do While compareStr4(Arr(M1), Piv) = -1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While compareStr4(Arr(M2), Piv) = 1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
       If M1 <> M2 Then tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = tmp
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While compareStr4(Arr(M1), Piv) = -1: M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortItemsFileNames Arr, LB, M2
    If M1 < UB Then QuickSortItemsFileNames Arr, M1, UB
    Exit Sub
abc1:
    

End Sub


Private Function IsSymbol3(a$, c$) As Boolean
Dim b$
b$ = NLtrim$(a$)
  If b$ <> "" Then
    If Left$(b$, Len(c$)) = c$ Then
        a$ = Mid$(b$, Len(c$) + 1)
        IsSymbol3 = True
End If
End If
End Function
Private Function PointPos(f$) As Long
Dim er As Long, er2 As Long
While IsSymbol3(f$, """")
f$ = GetStrUntil(Chr(34), f$)
Wend
Dim i As Long, j As Long, oj As Long
If f$ = vbNullString Then
PointPos = 1
Else
er = 3
er2 = 3
For i = 1 To Len(f$)
er = er + 1
er2 = er2 + 1
Select Case Mid$(f$, i, 1)
Case "."
oj = j: j = i
Case "\", "/", ":", Is = """"
If er = 2 Then oj = 0: j = i - 2: Exit For
er2 = 1
oj = j: j = 0
If oj = 0 Then oj = i - 1: If oj < 0 Then oj = 0
Case " "
If j > 0 Then Exit For
If er2 = 2 Then oj = 0: j = i - 1: Exit For
er = 1
Case "|", "'"
j = i - 1
Exit For
Case Is > " "
If j > 0 Then oj = j Else oj = 0
Case Else
If oj <> 0 Then j = oj Else j = i
Exit For
End Select
Next i
If j = 0 Then
If oj = 0 Then
j = Len(f$) + 1
Else
j = oj
End If
End If
While Mid$(f$, j, i) = " "
j = j - 1
Wend
PointPos = j
End If
End Function
Public Function ExtractType(f$, Optional JJ As Long = 0) As String
Dim i As Long, j As Long, D$
If IsSymbol3(f$, """") Then f$ = GetStrUntil(Chr(34), f$)
If f$ = vbNullString Then ExtractType = vbNullString: Exit Function
If JJ > 0 Then
j = JJ
Else
j = PointPos(f$)
End If
D$ = f$ & " "
If j < Len(D$) Then
For i = j To Len(D$)
Select Case Mid$(D$, i, 1)
Case "/", "|", "'", " ", Is = """"
i = i + 1
Exit For
End Select
Next i
If (i - j - 2) < 1 Then
ExtractType = vbNullString
Else
ExtractType = Mid$(D$, j + 1, i - j - 2)
End If
Else
ExtractType = vbNullString
End If
End Function
Public Function ExtractPath(f$, Optional IsAfile As Boolean = True) As String
If f$ = ".." Then ExtractPath = f$: Exit Function
Dim i As Long, j As Long
j = PointPos(f$)
If Mid$(f$, j, 1) = "." Then
'
Else
j = Len(f$)
End If
For i = j To 1 Step -1
Select Case Mid$(f$, i, 1)
Case Is < " ", "\", "/", ":"
Exit For
End Select
Next i
If Not IsAfile Then
    For i = i - 1 To 1 Step -1
    Select Case Mid$(f$, i, 1)
    Case Is < " ", "\", "/", ":"
    Exit For
    End Select
    Next i
End If
If i = 0 Then
ExtractPath = vbNullString
Else
ExtractPath = Mid$(f$, 1, i)
End If
End Function
Public Function ExtractName(f$) As String
Dim i As Long, j As Long, k$
If f$ = vbNullString Then Exit Function
'j = PointPos(F$)

'If Mid$(F$, j, 1) = "." Then
'k$ = ExtractType(F$, j)
'Else
j = Len(f$)
'End If
For i = j To 1 Step -1
Select Case Mid$(f$, i, 1)
Case Is < " ", "\", "/", ":"
Exit For
End Select
Next i
If k$ = vbNullString Then
If Mid$(f$, i + j - i, 1) = "." Then
ExtractName = Mid$(f$, i + 1, j - i - 1)
Else
ExtractName = Mid$(f$, i + 1, j - i)

End If
Else
ExtractName = Mid$(f$, i + 1, j - i) + k$
End If

End Function
Public Function isReadOnly(ByVal f$) As Boolean
Dim v&

Err.Clear
On Error Resume Next
If Right$(f$, 1) = "\" Then f$ = Mid$(f$, 1, Len(f$) - 1)
If f$ <> "" Then
If FileAttr(f$, v&) Then isReadOnly = (v& And 1) > 0 Else isReadOnly = True

Else
isReadOnly = True

End If
End Function
Private Function isdirP(wfd As WIN32_FIND_DATA) As Boolean
isdirP = (wfd.dwFileAttributes And 16) > 0
End Function
Public Function isdir(f$) As Boolean
Dim v&
Err.Clear
If f$ = vbNullString Then Exit Function
On Error Resume Next
If FileAttr(f$, v&) Then
isdir = (v& And 16) > 0
End If
End Function
Private Sub Class_Initialize()
''
End Sub
Sub once()
ReDim myfile$(100)
ReDim myLevel(100) As Integer
End Sub
Public Function GetStrUntil(ByVal sStr As String, fromstr As String, Optional RemoveSstr As Boolean = True) As String
Dim i As Long
If fromstr = vbNullString Then GetStrUntil = vbNullString: Exit Function
i = InStr(fromstr, sStr)
If (i < 2) And Not ((i > 0) And RemoveSstr) Then
GetStrUntil = vbNullString
fromstr = vbNullString
Else
GetStrUntil = Left$(fromstr, i - 1)
If RemoveSstr Then
fromstr = Mid$(fromstr, Len(sStr) + i)
Else
fromstr = Mid$(fromstr, i)
End If
End If
End Function
Public Sub DropLeft(ByVal uStr As String, fromstr As String)
Dim i As Long
i = InStr(fromstr, uStr)
If i > 0 Then
fromstr = Mid$(fromstr, i + Len(uStr))
Else
fromstr = vbNullString
End If
End Sub
Private Function StripNulls(ByVal OriginalStr As String) As String
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   StripNulls = OriginalStr
End Function


Public Property Get TopFolder() As String
TopFolder = mTopFolder
End Property

Public Property Let TopFolder(ByVal RHS As String)
If Right$(RHS, 1) <> "\" Then
mTopFolder = RHS & "\"
Else
mTopFolder = RHS
End If
End Property
Function CleanName(ByVal Name$) As String
Name$ = Replace(Name$, """", "-")
Name$ = Replace(Name$, "\", "-")
Name$ = Replace(Name$, "<", "-")
Name$ = Replace(Name$, ">", "-")
Name$ = Replace(Name$, ":", "-")
Name$ = Replace(Name$, "/", "-")
CleanName = Replace(Name$, "|", "-")
End Function

Private Sub Class_Terminate()
abort = True
End Sub
Private Function myFILEDATeTIME(aa As FILETIME) As String
' i need 14 bytes,
myFILEDATeTIME = Right$("00000000000000" & Hex$(aa.dwHighDateTime) + Left$("0000" + Hex$(HighWord((aa.dwLowDateTime))), 4), 14)
End Function
Function compareStr4(ByVal a$, ByVal b$) As Long
Const useclid = 0&
Dim a0$, b0$
a$ = LCase(a$)
b$ = LCase(b$)
If CompareString(useclid, 0, StrPtr(a$), Len(a$), StrPtr(b$), Len(a$)) = 2 Then Exit Function
a0$ = Mid$(a$, InStrRev(a$, ".") + 1)
If Len(a0$) = Len(a$) Then a0$ = vbNullString Else a$ = Left$(a$, Len(a$) - Len(a0$) - 1)

b0$ = Mid$(b$, InStrRev(b$, ".") + 1)
If Len(b0$) = Len(b$) Then b0$ = vbNullString Else b$ = Left$(b$, Len(b$) - Len(b0$) - 1)
'GoTo 1111
Select Case CompareString(useclid, 0, StrPtr(a0$), Len(a0$), StrPtr(b0$), Len(a0$))
    Case 1
    compareStr4 = -1
    Exit Function
    Case 3
    compareStr4 = 1
    
End Select

Dim i As Long, j As Long, a1$, b1$, p1 As Double, p2 As Double, k As Long, k1 As Long
k = Sgn(Len(a$) - Len(b$))
k1 = k
Dim once As Boolean
again:
j = IIf(Len(a$) >= Len(b$), Len(b$), Len(a$))
For i = 1 To j

    If Mid$(a$, i, 1) Like "[0-9]" Then
    a1$ = Mid$(a$, i)
    b1$ = Mid$(b$, i)
        IsNumberInFilename a1$, p1
        If IsNumberInFilename(b1$, p2) Then
            compareStr4 = Sgn(MyRound(p1, 8) - MyRound(p2, 8))
            If compareStr4 = 0 Then
                If Len(a1$) * Len(b1$) <> 0 Then
                    k = Sgn(Len(a1$) - Len(b1$))
                    a$ = a1$
                    b$ = b1$
                   
                    GoTo again
                End If
                If Len(a1$) + Len(b1$) = 0 Then
                    compareStr4 = k
                Else
                    compareStr4 = Sgn(Len(a1$) - Len(b1$))
                End If
                Exit Function
            End If
            Exit Function
        Else
        a$ = Mid$(a$, i)
    b$ = Mid$(b$, i)
    IsNumberInFilename a$, p1
       
          If Len(a$) > 0 Then
       
            Select Case CompareString(useclid, &H1000, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))
                Case 1, 3
                compareStr4 = -1
                Case 2
                If k > 0 Then
                compareStr4 = -1
                Else
                compareStr4 = 1
                End If
            End Select
        Else
        compareStr4 = -1
        End If
        End If
        Exit Function
    ElseIf Mid$(b$, i, 1) Like "[0-9]" Then
    a$ = Mid$(a$, i)
    b$ = Mid$(b$, i)
    IsNumberInFilename b$, p1
    If Len(b$) > 0 Then
       
            Select Case CompareString(useclid, &H1000, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))
                Case 3
                compareStr4 = 1
                Case 1, 2
                If k < 0 Then
                compareStr4 = 1
                Else
                compareStr4 = -1
                End If
                
                
            End Select
        Else
        compareStr4 = 1
        End If
        Exit Function
    End If
    compareStr4 = -2 + CompareString(useclid, 0, StrPtr(a$) + i * 2 - 2, 1, StrPtr(b$) + i * 2 - 2, 1)
    If compareStr4 <> 0 Then
    Exit Function
    End If
Next i
compareStr4 = Sgn(Len(a$) - Len(b$))
If compareStr4 = 0 Then compareStr4 = k1
End Function
Function compareStr444(ByVal a$, ByVal b$) As Long
Const useclid = 0&
Dim a0$, b0$
a$ = LCase(a$)
b$ = LCase(b$)
a0$ = Mid$(a$, InStrRev(a$, ".") + 1)
If Len(a0$) = Len(a$) Then a0$ = vbNullString Else a$ = Left$(a$, Len(a$) - Len(a0$) - 1)

b0$ = Mid$(b$, InStrRev(b$, ".") + 1)
If Len(b0$) = Len(b$) Then b0$ = vbNullString Else b$ = Left$(b$, Len(b$) - Len(b0$) - 1)
'GoTo 1111
Select Case CompareString(useclid, 0, StrPtr(a0$), Len(a0$), StrPtr(b0$), Len(a0$))
    Case 1
    compareStr444 = -1
    Exit Function
    Case 3
    compareStr444 = 1
    
End Select

1111
If CompareString(useclid, 0, StrPtr(a$), Len(a$), StrPtr(b$), Len(a$)) = 2 Then Exit Function
Dim i As Long, j As Long, a1$, b1$, p1 As Double, p2 As Double, k As Long, k1 As Long
k = Sgn(Len(a$) - Len(b$))
k1 = k
Dim once As Boolean
again:
j = IIf(Len(a$) >= Len(b$), Len(b$), Len(a$))

For i = 1 To j
    If Mid$(a$, i, 1) Like "[0-9]" Then
        a1$ = Mid$(a$, i)
        b1$ = Mid$(b$, i)
        IsNumberInFilename a1$, p1
        If IsNumberInFilename(b1$, p2) Then
            compareStr444 = Sgn(MyRound(p1 - p2))
            If compareStr444 = 0 Then
                If Len(a1$) * Len(b1$) <> 0 Then
                    k = Sgn(Len(a1$) - Len(b1$))
                    a$ = a1$
                    b$ = b1$
                    GoTo again
                End If
                If Len(a1$) + Len(b1$) = 0 Then
                    compareStr444 = k
                Else
                    compareStr444 = Sgn(Len(a1$) - Len(b1$))
                Exit Function
                End If
            End If
            Exit Function
        Else
        a$ = Mid$(a$, i)
        b$ = Mid$(b$, i)
        IsNumberInFilename a$, p1
        If once Then
        compareStr444 = -1
        ElseIf Len(a$) > 0 Then
            Select Case CompareString(useclid, &H1000, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))
            Case 1
            compareStr444 = -1
            Case 3
            compareStr444 = 1
            Case 2
        
                compareStr444 = 1
                
            End Select
        Else
        compareStr444 = -1
        End If
        End If
        Exit Function
    ElseIf Mid$(b$, i, 1) Like "[0-9]" Then
    a$ = Mid$(a$, i)
    b1$ = Mid$(b$, i)
    IsNumberInFilename b1$, p1
    
     If once Then
        compareStr444 = 1
        ElseIf Len(b$) > 0 Then
            Select Case CompareString(useclid, &H1000, StrPtr(a$), Len(a$), StrPtr(b1$), Len(b1$))
                Case 3
                compareStr444 = 1
                Case 1
                compareStr444 = -1
                Case 2
                compareStr444 = -1
            End Select
        Else
        compareStr444 = 1
        End If
        Exit Function
    End If
    once = False
    compareStr444 = -2 + CompareString(useclid, 0, StrPtr(a$) + i * 2 - 2, 1, StrPtr(b$) + i * 2 - 2, 1)
    If compareStr444 <> 0 Then Exit Function
Next i
compareStr444 = Sgn(Len(a$) - Len(b$))
If compareStr444 = 0 Then compareStr444 = k1

End Function

Private Function IsNumberInFilename(a$, D As Double) As Boolean
Dim a1 As Long
If a$ <> "" Then
a1 = Len(a$)

If IsNumberOnly(a$, 1, D, a1) Then
a$ = Mid$(a$, a1)
IsNumberInFilename = True
Else
IsNumberInFilename = False
End If
Else
IsNumberInFilename = False
End If
End Function
Private Function IsNumberOnly(a$, fr As Long, R As Double, lr As Long) As Boolean
Dim SG As Long, i As Long, ig$
If a$ = vbNullString Then IsNumberOnly = False: Exit Function
i = fr


If val("0" & Mid$(a$, i, 1)) = 0 And Left(Mid$(a$, i, 1), i) <> "0" And Left(Mid$(a$, i, 1), i) <> "." Then
IsNumberOnly = False

Else
    Do While i <= Len(a$)
        
        Select Case Mid$(a$, i, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, i, 1)
        Case Else
        Exit Do
        End Select
       i = i + 1
    Loop
    End If
    
    If ig$ = vbNullString Then
    IsNumberOnly = False
    lr = 1
    Else
    R = val(ig$)
    lr = i - fr + 1
    IsNumberOnly = True
End If

End Function


Public Function FindItemStartWidth(ByVal Key As String, NoCase As Boolean, ByVal Offset) As Long
Dim i As Long, j As Long, cp$

j = Len(Key)
i = -1
FindItemStartWidth = -1
If j = 0 Then Exit Function
If NoCase Then
For i = Offset To listcount - 1
cp$ = myfile(i)
If Len(cp$) > 0 Then If AscW(cp$) = 2 Then cp$ = Mid$(cp$, rinstr(cp$, "\") + 1)
If StrComp(Left$(cp$, j), Key, vbTextCompare) = 0 Then Exit For
Next i
Else
For i = Offset To listcount - 1
cp$ = myfile(i)
If Len(cp$) > 0 Then If AscW(cp$) = 2 Then cp$ = Mid$(cp$, rinstr(cp$, "\") + 1)
If StrComp(Left$(cp$, j), Key, vbBinaryCompare) = 0 Then Exit For
Next i
End If
If i < listcount Then
FindItemStartWidth = i
End If
End Function
