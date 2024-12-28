VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mPtr As Long
Private mitems As Long
Private itemlen As Long
Private mBytes As Long
Private ErrorMes As Long
Public structref As FastCollection
Public UseStruct As Boolean
Private mExecBloc As Boolean
Private Strings As LongHash
Private BasicItem As Integer
Private Type METAFILEPICT
    mm As Long
    xExt As Long
    yExt As Long
    hmf As Long
End Type
Private Type PWMFRect16
    Left   As Integer
    top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hmf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    Reserved    As Long
    CheckSum    As Integer
End Type
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    onebound(0 To 0) As SAFEARRAYBOUND
End Type
Private Width As Long, Height As Long, mSubtype As Long
' mSubtype = 0 undefined
' mSubtype = 1 Bitmap/Png/jpg/gif/ico
' mSubtype = 2 Emf
' mSubtype = 3 Wmf
Private Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As Any) As Long
Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hmf As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Byte) As Long

Private Declare Function GetEnhMetaFileHeader Lib "gdi32" (ByVal hmf As Long, ByVal cbBuffer As Long, lpemh As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal addr As Long, RetVal As Byte)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetEnhMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpData As Any) As Long
Private Declare Function GetEnhMetaFileBits Lib "gdi32" (ByVal hmf As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function CopyEnhMetaFile Lib "gdi32.dll" Alias "CopyEnhMetaFileW" (ByVal hemfSrc As Long, lpszFile As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As Long) As Long


Const CF_ENHMETAFILE = 14
Sub Construct(ByteLen As Long, Optional Times As Long = 1, Optional clearme As Long = 0, Optional executable As Boolean = False, Optional thatbasicitem As Integer)
If mitems <> 0 Then
If mitems <> Times Or mBytes <> ByteLen * Times Then
    ErrorMes = 1: Exit Sub
    End If
End If
If Times < 1 Then ErrorMes = 2: Exit Sub
If ByteLen < 0 Then
If thatbasicitem <> 0 Then
BasicItem = thatbasicitem
Else
BasicItem = vbString
End If
ElseIf Not IsMissing(thatbasicitem) Then
BasicItem = thatbasicitem
End If
ByteLen = Abs(ByteLen)

If ByteLen < 1 Then ErrorMes = ErrorMes + 4
If ErrorMes <> 0 Then Exit Sub
itemlen = ByteLen
mitems = Times
mExecBloc = Not executable
If clearme Then
    If mPtr <> 0 Then
        If NoRun Then
            If BlockSize(mPtr) > 0 Then
                BlockFree mPtr
            End If
        Else
                BlockFreeVirtual mPtr, mBytes
        End If
        mPtr = 0
    End If
End If
If mPtr = 0 Then
    mPtr = Blockmalloc(ByteLen * Times, clearme, mExecBloc)
End If
'If mExecBloc Then
If mPtr <> 0 Then mBytes = ByteLen * Times
End Sub
Public Sub RemovePtr(hMem As Long)
' don't use this sub if you didn't get the number.
    hMem = mPtr
    mPtr = 0
End Sub
Property Get ItemSize()
ItemSize = itemlen
End Property
Public Sub ClearError()
    ErrorMes = 0
End Sub
Public Property Get status()
    status = ErrorMes
End Property
Public Sub ResizeItems(NewSize As Long, Optional clearme As Long = 0)

If mBytes = 0 Then Exit Sub
Dim oPtr As Long


If NewSize > 0 Then
    If mExecBloc Then
        oPtr = BlockResizemalloc(mPtr, itemlen * NewSize, clearme)
        If oPtr <> 0 Then
            mitems = NewSize
            mPtr = oPtr
        Else
            ErrorMes = 3
        End If
    Else
    
    If itemlen * NewSize <= mBytes Then
            mPtr = BlockResizeVirtual(mPtr, itemlen * NewSize)
            mitems = NewSize
            mBytes = itemlen * NewSize
    Else
       oPtr = Blockmalloc(itemlen * NewSize, clearme, mExecBloc)
       If oPtr <> 0 Then
            VirtualBlockCopy oPtr, mPtr, itemlen * NewSize, mBytes
           
            BlockFreeVirtual mPtr, mBytes
            mitems = NewSize
            mBytes = itemlen * NewSize
            mPtr = oPtr
        Else
            ErrorMes = 3
        End If
    End If
    End If
Else
    ErrorMes = 2
End If
End Sub
Public Property Get ItemsAreStrings() As Boolean
ItemsAreStrings = BasicItem = vbString
End Property
Public Property Get WhatIsBasicItem() As Integer
WhatIsBasicItem = BasicItem
End Property

Public Property Get GetBytePtr(Offset As Variant) As Long
If ErrorMes = 0 And mitems > 0 Then
    If Offset < 0 Or Offset >= mitems * itemlen Then
        ErrorMes = 4
    Else
    GetBytePtr = signlong(uintnew(mPtr) + Offset)
    End If
Else
    GetBytePtr = 0  ' null
End If
End Property
Public Property Get GetPtr(Offset As Variant) As Long
If ErrorMes = 0 And mitems > 0 Then
If Offset < 0 Or Offset >= mitems Then
    ' zero
    ErrorMes = 4
Else
    GetPtr = cUlng(uintnew(mPtr) + itemlen * Offset)
    End If
Else
    GetPtr = 0  ' null
End If
End Property
Public Property Get SizeByte() As Long
If ErrorMes = 0 And mitems > 0 Then
If mExecBloc Then
    SizeByte = BlockSize(mPtr)
    Else
    SizeByte = mBytes
    End If
End If
End Property
Public Property Get items() As Long
If ErrorMes = 0 And mitems > 0 And itemlen > 0 Then
If mExecBloc Then
    items = BlockSize(mPtr) \ itemlen
    Else
    items = mBytes \ itemlen
    End If
    
End If
End Property
Public Function ValidArea(ByVal ThisStart, ByVal ThisWidth) As Boolean
Dim A As Currency
' ThisStart is a sign long which hold same bits with Unsigned address
If ErrorMes = 0 And mitems > 0 Then
A = uintnew(CCur(mPtr))   ' make mptr an Unsigned
If ThisStart = 0 Then
    ThisStart = A
Else
    ThisStart = uintnew(ThisStart)
End If
ThisWidth = CCur(ThisWidth)
' we have to make both Unsigned and then check them
If ThisStart - A >= 0 Then
If mExecBloc Then
If ThisStart + ThisWidth <= A + BlockSize(mPtr) Then ValidArea = True
Else
If ThisStart + ThisWidth <= A + mBytes Then ValidArea = True
End If
End If
End If

End Function
Public Function ValidArea2(ByVal ThisOffsetU As Double, ByVal ThisWidth As Double) As Boolean
' ThisOffsetU is Unsigned
If ErrorMes = 0 And mitems > 0 Then
If ThisOffsetU >= 0 Then
If mExecBloc Then
If ThisOffsetU * itemlen + ThisWidth <= uintnew2(CDbl(BlockSize(mPtr))) Then ValidArea2 = True
Else
If ThisOffsetU * itemlen + ThisWidth <= mBytes Then ValidArea2 = True
End If
End If
End If

End Function

Private Sub Class_Initialize()
Width = -1&
Height = -1&
End Sub

Private Sub Class_Terminate()
Set structref = Nothing
If mPtr <> 0 Then
    If NoRun Then
        If BlockSize(mPtr) > 0 Then
            BlockFree mPtr
        End If
    Else
            BlockFreeVirtual mPtr, mBytes
            mBytes = 0
    End If
End If
mPtr = 0
End Sub

Public Property Get NoRun() As Boolean
NoRun = mExecBloc
End Property

Public Function PutStringAtOffset(addr As Variant, AnyString As String) As Variant
Dim Offset
Offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mitems > 0 Then
    If Offset < 0 Or Offset + 3 >= mitems * itemlen Then
        ErrorMes = 4
    Else
        If Strings Is Nothing Then Set Strings = New LongHash
        With Strings
        If .ExistKey(cUlng(Offset + 1)) Then
           .ValueStr = AnyString
           PutStringAtOffset = .StrPointer
        Else
        .ItemCreator2 cUlng(Offset + 1), AnyString
        PutStringAtOffset = .StrPointer
        End If
        .Done = False
        End With
    End If
Else
    PutStringAtOffset = 0  ' null
End If
End Function
Public Function GetStringFromOffset(addr As Variant, AnyString As String) As Variant
Dim Offset
Offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mitems > 0 Then
    If Offset < 0 Or Offset + 3 >= mitems * itemlen Then
        ErrorMes = 4
        AnyString = "Crash"
    Else
        If Strings Is Nothing Then Set Strings = New LongHash
        With Strings
        If .ExistKey(cUlng(Offset + 1)) Then
           AnyString = .Value
           GetStringFromOffset = .StrPointer
        Else
            AnyString = "Crash"
            GetStringFromOffset = addr
        End If
        .Done = False
        End With
    End If
    
Else
    AnyString = "Crash"
    GetStringFromOffset = 0  ' null
End If

End Function
Public Function GetData1(f As Long, addr, lengthInBytes As Long) As Boolean
Dim Offset, place
Offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mitems > 0 Then
    If Offset < 0 Or Offset + lengthInBytes - 1 >= mitems * itemlen Then
        ErrorMes = 4
        GetData1 = False
    Else
    On Error GoTo there
    Dim a1 As SAFEARRAY1D, A() As Byte
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = lengthInBytes
        .pvData = addr
    End With
        CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4

        place = 1
        Get #f, , A()
        CopyMemory ByVal VarPtrArray(A), 0&, 4
        GetData1 = True
  End If
    
Else
there:
        If place = 1 Then
                CopyMemory ByVal VarPtrArray(A), 0&, 4
        End If
        GetData1 = False
End If
        
End Function
Public Function GetData(f As Long, ByVal addr, lengthInBytes As Long) As Boolean
Dim Offset, range
range = uintnew(addr) - uintnew1(GetPtr(0))
If range < 0 Or range >= mitems * itemlen Then
        ErrorMes = 4
        GetData = False
        Exit Function
End If
Offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mitems > 0 Then
    If Offset < 0 Or Offset + lengthInBytes - 1 >= (mitems * itemlen - range) Then
        ErrorMes = 4
        GetData = False
    Else
        On Error GoTo there
        Dim FH As Long
        FH = Module10.ReadFileHandler(CLng(f))
        If FH = 0 Then Exit Function
        Module10.API_ReadBLOCK FH, lengthInBytes, cUlng(addr)
        GetData = True
        Exit Function
  End If
Else
there:
        GetData = False
End If
        
End Function
Public Function ExportToByte() As Byte()
Dim place
If ErrorMes = 0 And mitems > 0 Then

    On Error GoTo there
    Dim a1 As SAFEARRAY1D, A() As Byte
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = SizeByte
        .pvData = GetPtr(0)
    End With
        CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4

        place = 1
        On Error Resume Next
        ExportToByte = A()
        If Err.Number > 0 Then Exit Function
        CopyMemory ByVal VarPtrArray(A), 0&, 4
       
  End If
    
Exit Function
there:
        If place = 1 Then
                CopyMemory ByVal VarPtrArray(A), 0&, 4
        End If
        

End Function



Public Function putData(f As Long, addr, lengthInBytes As Long) As Boolean
Dim Offset, place, range
range = uintnew(addr) - uintnew1(GetPtr(0))
If range < 0 Or range >= mitems * itemlen Then
        ErrorMes = 4
        putData = False
        Exit Function
End If
Offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mitems > 0 Then
    If Offset < 0 Or Offset + lengthInBytes - 1 >= mitems * itemlen Then
        ErrorMes = 4
        putData = False
    Else
    On Error GoTo there
    Dim a1 As SAFEARRAY1D, A() As Byte
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = lengthInBytes
        .pvData = addr
    End With
        CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
        place = 1
        On Error Resume Next
        Dim FH As Long
        FH = Module10.ReadFileHandler(CLng(f))
        Module10.FileWriteBytes FH, A()
        If Err.Number > 0 Then Exit Function
        CopyMemory ByVal VarPtrArray(A), 0&, 4
        putData = True
  End If
    
Else
there:
        If place = 1 Then
                CopyMemory ByVal VarPtrArray(A), 0&, 4
        End If
        putData = False
End If
        
End Function
Friend Function GetStdPicture1(Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional RetainRatio As Boolean = False, Optional Bitmap As Boolean = False, Optional magia1 As Boolean = False) As IPicture
   Dim a1 As SAFEARRAY1D, A() As Byte, token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
If mSubtype = 0 Then
If IsWmf Then
mSubtype = 2
ElseIf IsEmf Then
SubType = 2
End If
ElseIf mSubtype = 30 Then
      '  CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4
      '  Set GetStdPicture1 = LoadImageFromBuffer2(a(), Width, Height, backcolor)
      '   CopyMemory ByVal VarPtrArray(a), 0&, 4
ElseIf mSubtype = 300 Then
GoTo there1
End If
    If mSubtype <> 2 Or Bitmap Then
        CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
        Set GetStdPicture1 = LoadImageFromBuffer3(A(), Width, Height, BackColor, magia1)
         CopyMemory ByVal VarPtrArray(A), 0&, 4
    Else
       Dim hEmf As Long
    On Error Resume Next
    Dim bb As MemBlock, hMem As Long
there1:
    Set bb = Copy()
    bb.RemovePtr hMem
    Set bb = Nothing
    Set GetStdPicture1 = HandleToPictureFromBits(hMem, Me.SizeByte)
    If BlockSize(hMem) > 0 Then
    BlockFree hMem
    End If
    End If

   
 
End Function
Public Function GetStdPicture(Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal BackColor As Long = vbWhite, Optional FlipOrRotate As Long = 0, Optional Bitmap As Boolean = False) As IPicture
   Dim a1 As SAFEARRAY1D, A() As Byte, token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    If mSubtype = 0 Then
    If IsWmf Then
    SubType = 2
    ElseIf IsEmf Then
    SubType = 2
    End If
    ElseIf mSubtype = 300 Then
    GoTo there1
    End If
    If mSubtype <> 2 Or Bitmap Then
        CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
        Set GetStdPicture = LoadImageFromBuffer2(A(), Width, Height, BackColor, FlipOrRotate)
         CopyMemory ByVal VarPtrArray(A), 0&, 4
    Else
       Dim hEmf As Long
    
    On Error Resume Next
    
    hEmf = SetEnhMetaFileBits(SizeByte, ByVal GetBytePtr(0))
    If hEmf = 0 Then
there1:
    Set GetStdPicture = Me.GetStdPicture1()

    Else
    Set GetStdPicture = gCreatePicture(hEmf, 4)
    End If
   
    End If

   
 
End Function
Public Function DrawEmfToHdc(Scr As Object, X As Long, Y As Long, Angle!, Width As Long, Height As Long) As Boolean
   Dim a1 As SAFEARRAY1D, A() As Byte, token As Long
   If mSubtype = 0 Then
        If IsEmf Then
        mSubtype = 2
        ElseIf Not IsWmf Then
        mSubtype = 1
        End If
   End If
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
    DrawEmfToHdc = DrawEmfFromBuffer(Scr, A(), X, Y, Angle!, Width, Height)
    
    CopyMemory ByVal VarPtrArray(A), 0&, 4
 
End Function
Public Function DrawSpriteToHdc(bstack As basetask, sprt As Boolean, Angle!, zoomfactor!, blend!, Optional BackColor As Long = -1) As Boolean
   Dim a1 As SAFEARRAY1D, A() As Byte, token As Long
   If mSubtype = 0 Then
        If IsEmf Then
        mSubtype = 2
        ElseIf Not IsWmf Then
        mSubtype = 1
        End If
   End If
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
    DrawSpriteToHdc = DrawSpriteFromBuffer(bstack, A(), Not sprt, Angle!, zoomfactor!, blend!, BackColor, mSubtype = 2)
    
    CopyMemory ByVal VarPtrArray(A), 0&, 4
 
End Function
Public Function DrawImageToHdc(where As Object, Optional X As Long = 0&, Optional Y As Long = 0&, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
   Dim a1 As SAFEARRAY1D, A() As Byte, token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
   CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
   DrawImageToHdc = DrawImageFromBuffer(A(), where.hDC, X, Y, (Width), (Height))
   CopyMemory ByVal VarPtrArray(A), 0&, 4
 
End Function
Public Function ReadImageSizeX(SizeX As Variant) As Boolean
If Width <> -1 Then SizeX = Width: ReadImageSizeX = True: Exit Function
   Dim a1 As SAFEARRAY1D, A() As Byte, token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
    If ReadSizeImageFromBuffer(A(), Width, Height) Then
    SizeX = Width
    ReadImageSizeX = True
        Else
        Width = -1&
        Height = -1&
    End If
    
    CopyMemory ByVal VarPtrArray(A), 0&, 4


End Function
Public Function ReadImageSizeY(SizeY As Variant) As Boolean
If Height <> -1 Then SizeY = Height: ReadImageSizeY = True: Exit Function
   Dim a1 As SAFEARRAY1D, A() As Byte, token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
    If ReadSizeImageFromBuffer(A(), Width, Height) Then
    SizeY = Height
    ReadImageSizeY = True
    Else
        Width = -1&
        Height = -1&
    End If
    
    CopyMemory ByVal VarPtrArray(A), 0&, 4
End Function
Public Sub ResetImageSize()
Dim a1 As SAFEARRAY1D, A() As Byte, token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLbound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(A()), VarPtr(a1), 4
    If Not ReadSizeImageFromBuffer(A(), Width, Height) Then
     Width = -1&
        Height = -1&
    End If
    
    CopyMemory ByVal VarPtrArray(A), 0&, 4
End Sub
Function Copy(Optional R As Long = 0, Optional P As Long = 0) As Object
Dim aCopy As New MemBlock, BLen As Long, hMem As Long
If mitems > 0 Then
    If mExecBloc Then
        BLen = BlockSize(mPtr)
    Else
        BLen = mBytes
    End If
    mitems = BLen \ itemlen
    If R <> 0 Then
        If R >= 0 And R < mitems Then
            If P >= 1 Then
                aCopy.Construct itemlen, CLng(P), , R = 0 And Not mExecBloc
                If mitems > P Then
                    P = mitems
                End If
                CopyMemory ByVal aCopy.GetPtr(0), ByVal Me.GetPtr(R), itemlen * P
            Else
           
                aCopy.Construct itemlen, 1&, , Not mExecBloc
                
                CopyMemory ByVal aCopy.GetPtr(0), ByVal Me.GetPtr(R), itemlen
            End If
        Else
            aCopy.Construct itemlen, 1&, , Not mExecBloc
            CopyMemory ByVal aCopy.GetPtr(0), ByVal Me.GetPtr(R), itemlen
        End If
    Else
    aCopy.Construct itemlen, mitems, , Not mExecBloc
    hMem = GlobalLock(aCopy.GetPtr(0))
    
    CopyMemory ByVal hMem, ByVal mPtr, BLen
    GlobalUnlock hMem
    End If
    Set Copy = aCopy
End If
End Function
Public Function Search(from As Long, what) As Long
Dim limit As Long, b As Byte, I As Long, W As Byte
b = Abs(what) And &HFF
limit = SizeByte - 1

For I = from To limit
    GetMem1 GetBytePtr(I), W
    If b = W Then
    Search = I + 1: Exit Function
    End If
Next
Search = I + 1
End Function

Public Property Get SubType() As Variant
SubType = mSubtype
End Property

Public Property Let SubType(ByVal RHS As Variant)
mSubtype = RHS
End Property
Public Sub SentToClipBoard(Optional pWidth As Long = -1, Optional ByVal pHeight As Long = -1, Optional BackColor As Long = vbWhite, Optional lockratio As Boolean = False, Optional toBitmap As Boolean = False)
    Dim hBitmap As Long
    Dim mm As StdPicture, hEmf As Long
    On Error Resume Next
    If pWidth <> -1 Then pWidth = pWidth / DXP
    If pHeight <> -1 Then pHeight = pHeight / DYP
    If pWidth = -1 Then
        pWidth = pHeight
        lockratio = True
    ElseIf pHeight = -1 Then
        lockratio = True
    End If

    Set mm = GetStdPicture1(pWidth, pHeight, BackColor, lockratio, toBitmap)

    If Err.Number > 0 Or (mm Is Nothing) Then
        MyEr "Not a proper picture to send it to clipboard", "Δεν είναι μια σωστή εικόνα για να τη στείλω στο πρόχειρο"
        Exit Sub
    End If
    If mm.Handle <> 0 Then
        If (OpenClipboard(Form1.hWnd) <> 0) Then
           If EmptyClipboard Then
                If mm.Type = 2 Then
                    hEmf = getEmf(mm)
                    If SetClipboardData(14, hEmf) = 0 Then
                        DeleteEnhMetaFile hEmf
                    End If
                ElseIf mm.Type = 4 Then
                    SetClipboardData 14, mm.Handle
                Else
                    SetClipboardData 2, mm.Handle
                End If
            End If
            CloseClipboard
        End If

    End If

End Sub
Public Function IsPng() As Boolean
    Dim h As Long
       CopyMemory h, ByVal GetBytePtr(0), 4
        If h = &H474E5089 Then
            mSubtype = 1
            IsPng = True
        End If
End Function
Public Function IsIco() As Boolean
    Dim h As Long
       CopyMemory h, ByVal GetBytePtr(0), 4
        If h = &H10000 Then
            mSubtype = 1
            IsIco = True
        End If
End Function
Public Function IsBmp() As Boolean
    Dim h As Long
       CopyMemory h, ByVal GetBytePtr(0), 4
        If (h And &HFFFF&) = &H4D42 Then
            mSubtype = 1
            IsBmp = True
        End If
End Function
Public Function IsJpg() As Boolean
    Dim h As Long
       CopyMemory h, ByVal GetBytePtr(0), 4
        If (h And &HFFFF&) = &HD8FF Then
            mSubtype = 1
            IsJpg = True
        End If
End Function

Public Function IsWmf() As Boolean
    Dim hEmf As Long
       CopyMemory hEmf, ByVal GetBytePtr(0), 4
        If hEmf = &H9AC6CDD7 Then
            mSubtype = 2
            IsWmf = True
        End If
End Function
Public Function IsGif() As Boolean
    Dim h As Long
       CopyMemory h, ByVal GetBytePtr(0), 4
        If h = &H38464947 Then
            mSubtype = 1
            IsGif = True
        End If
End Function
Public Function IsEmf() As Boolean
    Dim hEmf As Long
    CopyMemory hEmf, ByVal GetBytePtr(0), 4
    
    If hEmf = 1 Then mSubtype = 2: IsEmf = True: Exit Function
End Function
Public Sub SentEmfToClipBoard()
    Dim hEmf As Long
    
    On Error Resume Next
    hEmf = SetEnhMetaFileBits(SizeByte, ByVal GetBytePtr(0))
  
    
    If Err.Number > 0 Or hEmf = 0 Then
        MyEr "Not a proper picture to send it to clipboard", "Δεν είναι μια σωστή εικόνα για να τη στείλω στο πρόχειρο"
        Exit Sub
    End If
    If hEmf <> 0 Then
        If (OpenClipboard(Form1.hWnd) <> 0) Then
            EmptyClipboard
            If SetClipboardData(CF_ENHMETAFILE, hEmf) = 0 Then
                DeleteEnhMetaFile hEmf
            End If
            CloseClipboard
        End If
    End If

End Sub
Private Function getEmf(Wmf As StdPicture) As Long
 Dim zz As METAFILEPICT
 Dim bb As wmfPlaceableFileHeader
 CopyMemory bb, ByVal GetBytePtr(0), LenB(bb)
 zz.xExt = bb.BoundingBox.Right - bb.BoundingBox.Left
 zz.yExt = bb.BoundingBox.Bottom - bb.BoundingBox.Bottom
 
 Dim bytes As Long
Dim Buffer() As Byte
   bytes = GetMetaFileBitsEx(Wmf.Handle, bytes, ByVal 0&)
 If bytes Then
ReDim Buffer(1 To bytes) As Byte
Call GetMetaFileBitsEx(Wmf.Handle, bytes, Buffer(1))
Else

Exit Function
End If
    On Error Resume Next
    zz.mm = 1
    getEmf = SetWinMetaFileBits(bytes, Buffer(1), Form1.hDC, zz)
End Function

Public Sub SentWmfToClipBoard()
 Dim hEmf As Long
    hEmf = getEmf(GetStdPicture1())
    If Err.Number > 0 Or hEmf = 0 Then
        MyEr "Not a proper picture to send it to clipboard", "Δεν είναι μια σωστή εικόνα για να τη στείλω στο πρόχειρο"
        Exit Sub
    End If
    If hEmf <> 0 Then
        If (OpenClipboard(Form1.hWnd) <> 0) Then
            EmptyClipboard
            If SetClipboardData(CF_ENHMETAFILE, hEmf) = 0 Then
                DeleteEnhMetaFile hEmf
            End If
            CloseClipboard
        End If
    End If

End Sub
Friend Function CreateFromPicture(BackColor As Long, w1 As Long, w2 As Long, rot As Long) As mHandler
Dim aa As New cDIBSection
If w1 <> -1 Then w1 = w1 / dv15
If w2 <> -1 Then w2 = w2 / dv15
   
    
    Dim sp As StdPicture
    Set sp = GetStdPicture(w1, w2, mycolor(BackColor), rot)
    If sp.Type = 4 Then
    aa.emfSizeFactor = 1
    aa.CreateFromPicture sp, mycolor(BackColor), w1, w2
    Else
    aa.CreateFromPicture sp
    End If
    Dim k As Object
    Set k = New MemBlock
    aa.SaveDibToMeMBlock k
    If rot <> 0 Then
        aa.ClearUp
        aa.CreateFromPicture k.GetStdPicture(-1, -1, , rot)
        Set k = New MemBlock
        aa.SaveDibToMeMBlock k
        
    End If
    Dim Handler As New mHandler
    Handler.T1 = 2
    Set Handler.objref = k
    Set CreateFromPicture = Handler
    
End Function

