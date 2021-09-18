VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents Zip As cZipArchive
Attribute Zip.VB_VarHelpID = -1
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Event Cancel(NoCancel As Boolean)
Event Progress(ByVal FileIdx As Long, ByVal Current As Long, ByVal Total As Long, mCancel As Boolean)
Dim mTotal As Double, mSum As Double, LastTotal As Double, oIndex As Long
Dim mLastCancel As Boolean
Dim mLastCurent
Dim LockControl As Boolean
Private Sub Class_Initialize()
Set Zip = New cZipArchive
End Sub
Public Function OpenZip(ByVal FileName$) As Boolean
    If LockControl Then
        MessErr
        Exit Function
    End If
    OpenZip = Zip.OpenArchive(FileName$)
End Function
Public Function OpenZipBuf(FileBuf() As Byte) As Boolean
    If LockControl Then
        MessErr
        Exit Function
    End If
    OpenZipBuf = Zip.OpenArchive(CVar(FileBuf))
End Function
Public Sub CreateZipFile(ByVal FileName$)
    If LockControl Then
        MessErr
        Exit Sub
    End If
    LockControl = True
    FileName$ = mylcasefILE(FileName$)
    Dim ret As Boolean
    ret = Zip.CompressArchive(FileName$, , , vbTrue)
    LockControl = False
    mLastCancel = False
End Sub
Public Function CreateZipBuffer() As mHandler
    If LockControl Then
        MessErr
        Exit Function
    End If
    LockControl = True
   
    Dim ret As Boolean, bufmem() As Byte
    
    ret = Zip.CompressArchive(bufmem, , , vbTrue)
    LockControl = False
    mLastCancel = False
End Function
Public Function AddFromMemory(FileBuf() As Byte, Optional Name As String, Optional Comment As String, Optional Password As String, Optional EncrStrength As Long, Optional Level As Long = -1) As Boolean
    AddFromMemory = Zip.AddFile(CVar(FileBuf), Name, Comment, Password, EncrStrength, Level)
End Function
Public Function AddFile(FileStr As String, Optional Name As String, Optional Comment As String, Optional Password As String, Optional EncrStrength As Long, Optional Level As Long = -1) As Boolean
    AddFile = Zip.AddFile(CVar(FileStr), Name, Comment, Password, EncrStrength, Level)
End Function
Function AddFileFromFolder(sFolderAndMask As String, _
            Optional Recursive As Boolean, _
            Optional TargetFolder As String, _
            Optional IncludeEmptyFolders As Boolean, _
            Optional Password As String, _
            Optional EncrStrength As Long, _
            Optional Level As Long = -1) As Boolean
    'Zip.AddFile
    AddFileFromFolder = Zip.AddFromFolder(sFolderAndMask, Recursive, TargetFolder, IncludeEmptyFolders, Password, EncrStrength, Level)
End Function
Public Property Get FileInfo(FileNum) As Object
Dim Handler As New mHandler, touple As New mArray
Dim pp
pp = Zip.FileInfo(FileNum)
touple.LoadTuple pp
Set Handler.objref = touple
Handler.t1 = 3
Set FileInfo = Handler
End Property


Public Sub ExtractToPath(ByVal Path$, Optional ByVal FileMask$, Optional ByVal Pass$)
    If LockControl Then
        MessErr
        Exit Sub
    End If
    LockControl = True
    Dim p, ret As Boolean
    p = Path$
    mLastCancel = False
    If FileMask$ = "" Then
        ret = Zip.Extract(Path$, , Pass$)
    Else
        ClearStr FileMask$
        ret = Zip.Extract(Path$, FileMask$, Pass$)
    End If
    LockControl = False
    mLastCancel = False
End Sub
Public Function ExtractOneToBuffer(ByVal FileName$, Optional ByVal Pass$) As mHandler
    If LockControl Then
        MessErr
        Exit Function
    End If
    ClearStr FileName$
    FileName$ = CleanStr(FileName$, "?*")
    LockControl = True
    Dim p() As Byte, ret As Boolean, Handler As New mHandler, mem As New MemBlock
    mLastCancel = False
    
    ret = Zip.Extract(p(), FileName$, Pass$)
    If ret Then
    mem.Construct 1, UBound(p) + 1
    CopyMemory ByVal mem.GetPtr(0), p(0), UBound(p) + 1
    Handler.t1 = 2
    Set Handler.objref = mem
    Set ExtractOneToBuffer = Handler
    End If
    LockControl = False
    mLastCancel = False
End Function
Public Property Get Busy() As Boolean
    Busy = LockControl
End Property
Public Property Get FileCount()
    FileCount = Zip.FileCount
End Property
Private Sub MessErr()
MyEr "ZipTool Busy", "Το εργαλείο συμπίεσης είναι σε εργασία"
End Sub
Private Sub Class_Terminate()
Set Zip = Nothing
End Sub

Private Sub Zip_BeforeCompress(ByVal FileIdx As Long, Level As Long, SkipFile As Boolean, mCancel As Boolean)
    RaiseEvent Cancel(mCancel)
End Sub

Private Sub Zip_BeforeExtract(ByVal FileIdx As Long, File As Variant, SkipFile As Boolean, mCancel As Boolean)
    RaiseEvent Cancel(mCancel)
End Sub

Private Sub Zip_Progress(ByVal FileIdx As Long, ByVal Current As Long, ByVal Total As Long, mCancel As Boolean)
    RaiseEvent Progress(FileIdx, Current, Total, mCancel)
    
End Sub

