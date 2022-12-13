Attribute VB_Name = "Module5"
Option Explicit
Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal ByteLen As Long, ByVal Destination As Long, ByVal Source As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Const Pi = 3.14159265359
Private Type SAFEARRAYBOUND
    cElements As Long
    lLBound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Sub SaveBmp(sFile As String, ByVal Scr As Object)
' as a .DIB with .bmp extension
       Dim photo As New cDIBSection
            photo.CopyPicture Scr
            photo.SaveDib sFile
       
End Sub
Public Function Decode64toMemBloc(ByVal a$, ok As Boolean, Optional forcode As Boolean = False) As Object
    Dim mem As New MemBlock, BLen As Long
    a$ = Decode64(a$, ok)
    If ok Then
        BLen = LenB(a$)
        mem.Construct 1, BLen, , forcode
        CopyBytes BLen, mem.GetPtr(0), StrPtr(a$)
        Set Decode64toMemBloc = mem
    End If
    
End Function
Public Function File2newMemblock(FileName As String, R, p) As Object
    Dim mem As New MemBlock, BLen As Long, i As Long
    R = -1#
    FileName = CFname(FileName)
    If FileName <> "" Then
     
     
    BLen = FileLen(GetDosPath(FileName))
    If BLen Then
    mem.Construct 1, BLen, , CBool(p)
    i = FreeFile
    On Error Resume Next

    Open GetDosPath(FileName) For Binary Access Read As i

    If Err.Number > 0 Then MyEr Err.Description, Err.Description: Close i: Exit Function
    If mem.GetData1(i, mem.GetPtr(0), BLen) Then
    R = 0#
    Set File2newMemblock = mem
    If Not mem.IsWmf Then
    If Not mem.IsEmf Then
    If Not mem.IsBmp Then
    If Not mem.IsIco Then
    If Not mem.IsGif Then
    If Not mem.IsJpg Then
    If Not mem.IsPng Then
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    Close i
    End If
    End If
End Function

Public Function SaveStr2MemBlock(a, R) As Object
Dim aa As New cDIBSection
R = -1#
If cDib(a, aa) Then
    aa.GetDpi 96, 96
    If Not aa.SaveDibToMeMBlock(SaveStr2MemBlock) Then
    MyEr "Can't save to buffer", "��� ����� �� ���� ��� ���������"
    Else
    R = 0#
    End If
End If
End Function

Public Function SaveJPG( _
      ByRef cDib As cDIBSection, _
      ByVal sFile As String, Optional ByVal lQuality As Long = 90, _
      Optional UserComment As String) As Boolean
   Dim j As New cJpeg
j.Quality = lQuality
If UserComment = vbNullString Then j.Comment = "M2000 User" Else j.Comment = Left$(UserComment, 64)
If lQuality <= 50 Then
j.SetSamplingFrequencies 2, 2, 1, 1, 1, 1  ' for screen
Else
j.SetSamplingFrequencies 1, 1, 1, 1, 1, 1  ' as camera
End If
With cDib
.needHDC
j.SampleHDC .HDC1, .Width, .Height
.FreeHDC
j.SaveFile sFile
End With
   End Function

Public Sub CheckOrientation(a As cDIBSection, f As String)

          If LCase(ExtractType(f, (0))) = "jpg" Then
           Dim qw As New ExifRead
             qw.Load f
 
            Select Case qw.Tag(Orientation)
                  Case 3
              RotateDib180 a
              Case 8
              RotateDib90 a
              Case 6
              RotateDib270 a
              End Select
                         
                         End If
End Sub
Public Function RotateDib90(cDibbuffer0 As cDIBSection, Optional MEDOEV As Boolean = False)
Dim piw As Long, pih As Long
If cDibbuffer0.hDib = 0 Then Exit Function
   piw = cDibbuffer0.Width
   pih = cDibbuffer0.Height
      If piw = 0 Then Exit Function

 Dim cDIBbuffer1 As New cDIBSection
If cDIBbuffer1.create(pih, piw) Then

cDIBbuffer1.GetDpiDIB cDibbuffer0
Dim bDib() As Byte, bDib1() As Byte
Dim X As Long, Y As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
On Error Resume Next
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLBound = 0
        .Bounds(0).cElements = cDibbuffer0.Height
        .Bounds(1).lLBound = 0
        .Bounds(1).cElements = cDibbuffer0.BytesPerScanLine()
        .pvData = cDibbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLBound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLBound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4
'dDib1 �� ������ pih, piw
Dim myx As Long, oldx As Long
Dim ttt As Long, TTC As Long

ttt = 1 + 100000 / piw

myx = 0
If MEDOEV Then
  For Y = pih - 1 To 0 Step -1
    oldx = 0
    TTC = TTC - 1
    If TTC = 0 Then DoEvents: TTC = ttt
        For X = 0 To piw - 1
                        bDib1(myx, X) = bDib(oldx, Y)
                        bDib1(myx + 1, X) = bDib(oldx + 1, Y)
                        bDib1(myx + 2, X) = bDib(oldx + 2, Y)
                        oldx = oldx + 3
       Next X
       myx = myx + 3
    Next Y
Else
  For Y = pih - 1 To 0 Step -1
    oldx = 0
        For X = 0 To piw - 1
                        bDib1(myx, X) = bDib(oldx, Y)
                        bDib1(myx + 1, X) = bDib(oldx + 1, Y)
                        bDib1(myx + 2, X) = bDib(oldx + 2, Y)
                        oldx = oldx + 3
       Next X
       myx = myx + 3
    Next Y
    End If

    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
 cDibbuffer0.CreateFromPicture cDIBbuffer1.Picture
cDibbuffer0.GetDpiDIB cDIBbuffer1
    End If



End Function
Public Function RotateDib270(cDibbuffer0 As cDIBSection, Optional MEDOEV As Boolean = False)
Dim piw As Long, pih As Long
   piw = cDibbuffer0.Width
   pih = cDibbuffer0.Height
    If piw = 0 Then Exit Function
' Dim cDIBbuffer1 As Object
 Dim cDIBbuffer1 As New cDIBSection
If cDIBbuffer1.create(pih, piw) Then
'cDIBbuffer1.WhiteBits
cDIBbuffer1.GetDpiDIB cDibbuffer0
cDIBbuffer1.Cls
Dim bDib() As Byte, bDib1() As Byte
Dim X As Long, Y As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
On Error Resume Next
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLBound = 0
        .Bounds(0).cElements = cDibbuffer0.Height
        .Bounds(1).lLBound = 0
        .Bounds(1).cElements = cDibbuffer0.BytesPerScanLine()
        .pvData = cDibbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLBound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLBound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4
'dDib1 �� ������ pih, piw
Dim myx As Long, oldx As Long
myx = pih * 3 - 3
Dim ttt As Long, TTC As Long

ttt = 1 + 100000 / piw
If MEDOEV Then
    For Y = pih - 1 To 0 Step -1
    oldx = piw * 3 - 3
      TTC = TTC - 1
    If TTC = 0 Then DoEvents: TTC = ttt
        For X = 0 To piw - 1
                        bDib1(myx, X) = bDib(oldx, Y)
                        bDib1(myx + 1, X) = bDib(oldx + 1, Y)
                        bDib1(myx + 2, X) = bDib(oldx + 2, Y)
                        oldx = oldx - 3
       Next X
       myx = myx - 3
    Next Y
       Else
        For Y = pih - 1 To 0 Step -1
    oldx = piw * 3 - 3
        For X = 0 To piw - 1
                        bDib1(myx, X) = bDib(oldx, Y)
                        bDib1(myx + 1, X) = bDib(oldx + 1, Y)
                        bDib1(myx + 2, X) = bDib(oldx + 2, Y)
                        oldx = oldx - 3
       Next X
       myx = myx - 3
    Next Y
    End If
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
 cDibbuffer0.CreateFromPicture cDIBbuffer1.Picture
cDibbuffer0.GetDpiDIB cDIBbuffer1
    End If



End Function
Public Function RotateDib180(cDibbuffer0 As cDIBSection, Optional MEDOEV As Boolean = False)
' MEDOEV if true then with doevents
Dim piw As Long, pih As Long

If cDibbuffer0.hDib = 0 Then Exit Function
   piw = cDibbuffer0.Width
   pih = cDibbuffer0.Height
      If piw = 0 Then Exit Function
 Dim cDIBbuffer1 As New cDIBSection
 If cDIBbuffer1.create(piw, pih) Then
 cDIBbuffer1.GetDpiDIB cDibbuffer0

Dim bDib() As Byte, bDib1() As Byte
Dim X As Long, Y As Long
Dim tSA As SAFEARRAY2D
Dim tSA1 As SAFEARRAY2D
On Error Resume Next
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLBound = 0
        .Bounds(0).cElements = cDibbuffer0.Height
        .Bounds(1).lLBound = 0
        .Bounds(1).cElements = cDibbuffer0.BytesPerScanLine()
        .pvData = cDibbuffer0.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    With tSA1
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLBound = 0
        .Bounds(0).cElements = cDIBbuffer1.Height
        .Bounds(1).lLBound = 0
        .Bounds(1).cElements = cDIBbuffer1.BytesPerScanLine()
        .pvData = cDIBbuffer1.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib1()), VarPtr(tSA1), 4
'dDib1 �� ������ pih, piw
Dim myx As Long, oldx As Long, oldy As Long
oldx = piw * 3
oldy = pih - 1
Dim ttt As Long, TTC As Long

ttt = 1 + 100000 / piw
If MEDOEV Then
    For Y = 0 To pih - 1
  myx = oldx - 3
  oldx = 0
      TTC = TTC - 1
    If TTC = 0 Then DoEvents: TTC = ttt
        For X = 0 To piw - 1
                        bDib1(myx, oldy) = bDib(oldx, Y)
                        bDib1(myx + 1, oldy) = bDib(oldx + 1, Y)
                        bDib1(myx + 2, oldy) = bDib(oldx + 2, Y)
                        myx = myx - 3
                        oldx = oldx + 3
       Next X
       oldy = oldy - 1
       myx = myx + 3
    Next Y
Else
    For Y = 0 To pih - 1
  myx = oldx - 3
  oldx = 0
        For X = 0 To piw - 1
                        bDib1(myx, oldy) = bDib(oldx, Y)
                        bDib1(myx + 1, oldy) = bDib(oldx + 1, Y)
                        bDib1(myx + 2, oldy) = bDib(oldx + 2, Y)
                        myx = myx - 3
                        oldx = oldx + 3
       Next X
       oldy = oldy - 1
       myx = myx + 3
    Next Y
    End If
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    CopyMemory ByVal VarPtrArray(bDib1), 0&, 4
cDibbuffer0.CreateFromPicture cDIBbuffer1.Picture
cDibbuffer0.GetDpiDIB cDIBbuffer1
    End If
End Function


