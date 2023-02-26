VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' M2000 ver 1.
Option Explicit
' chr(34) is """" maybe somewhere in a module as public
Const mway As Long = DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP
Private Col() As Long, lastdata As String, thispara As Long
Private TextPixelOffset As Long, mark1 As Long, mark2 As Long
Public ColorSet As Long, sp As String, sp1 As String, sp2 As String
Public PreserveRightSpaces As Boolean
Public NoCenterLineEdit As Boolean
Public NumberOnly As Boolean
Public NumberIntOnly As Boolean
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Dim profundo As New clsProfiler
Private blockKeyboard As Boolean
Public WithEvents mDoc As Document    ' document
Attribute mDoc.VB_VarHelpID = -1
Public Form1mn1Enabled As Boolean
Public Form1mn2Enabled As Boolean
Public Form1mn3Enabled As Boolean
Public Form1sdnEnabled As Boolean
Public Form1supEnabled As Boolean
Public Form1mscatEnabled As Boolean
Public Form1rthisEnabled As Boolean
Dim WithEvents mDoc1 As Document  ' header
Attribute mDoc1.VB_VarHelpID = -1
Dim UndoDoc As New Document  ' one level of undo no events ' no break lines
Dim undopara As Long, undopos As Long, undotext As String, undomode As Long, undotexthtml As String
Dim Redopara As Long, Redopos As Long, Redotext As String, Redomode As Long
Dim undoprev As Boolean, Redoprev As Boolean
Dim undostate As Boolean
Dim Redostate As Boolean
Dim NoHeadwrap As Boolean
Private Type UndoChain
    undopara As Long
    undopos As Long
    undomode As Long
    undotext As String
    undoprev As Boolean
End Type
Dim NeoUndo As Long, UndoMax As Long
Dim NeoRedo As Long, RedoMax As Long

Dim MultiUndo() As UndoChain
Dim MultiRedo() As UndoChain

Dim GetUndoNext As Boolean
Dim myfilename As String
Dim all As Long, onelineheight As Long
Event inform(tLine As Long, tPos As Long)
Event ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Event CtrlPlusF1()
Public LastSelStart As Long
Public NoMark As Boolean
Public NoTitle As Boolean
Public NoColor As Boolean
Public Title As String
Public UsedAsTextBox As Boolean, HighLightColor As Long  '' only in m2000. Enter exit - use maxchar - No wrap
Private mEditDoc As Boolean
Public AutoNumber As Boolean
Public Charpos As Long, expandCharPos As Long
Public showparagraph As Boolean
Private showparagraphWork As Boolean
Public SelectionColor As Long
Public SelectionTextColor As Long
Private showparagraphonce As Boolean
Private mNoWrap As Boolean
Private BlockStartPara As Long, BlockStartPos As Long, BlockStartLine As Long, BlockTopline As Long
Private BlockStartCursor As Long, BlockEndCursor As Long
Private BlockEndPara As Long, BlockEndPos As Long, BlockEndline As Long, BlockBottomline As Long
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000&
Private Const GMEM_DISCARDABLE = &H100&
Private Const GMEM_DISCARDED = &H4000&
Private Const GMEM_FIXED = &H0&
Private Const GMEM_INVALID_HANDLE = &H8000&
Private Const GMEM_LOCKCOUNT = &HFF&
Private Const GMEM_MODIFY = &H80&
Private Const GMEM_MOVEABLE = &H2&
Private Const GMEM_NOCOMPACT = &H10&
Private Const GMEM_NODISCARD = &H20&
Private Const GMEM_NOT_BANKED = &H1000&
Private Const GMEM_NOTIFY = &H4000&
Private Const GMEM_SHARE = &H2000&
Private Const GMEM_VALID_FLAGS = &H7F72&
Private Const GMEM_ZEROINIT = &H40&
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "user32" _
    (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" _
    () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&) ' not NT?
Private Const DWL_ANYTHREAD& = 0&
Const LOCALE_ILANGUAGE = 1&
Private OldPar As Long, dragCharpos As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Integer)
Private Type DRAWTEXTPARAMS
     cbSize As Long
     iTabLength As Long
     iLeftMargin As Long
     iRightMargin As Long
     uiLengthDrawn As Long
End Type
Dim tParam As DRAWTEXTPARAMS
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (ByRef lpRect As RECT) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hDC As Long, ByVal lpsz As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long, ByVal lpDrawTextParams As Long) As Long
Private Declare Function GetStringTypeExW Lib "kernel32.dll" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByRef lpCharType As Integer) As Long

'nTabPositions  here we use 1
' lpnTabStopPositions here we usen just a long variable, passed by reference.
'
'Private Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutW" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lplong As Long, ByVal nCount As Long, ByVal nTabPositions As Long, ByRef lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long

Private tmpshow As Boolean
Private LastListIndex As Long
Public Property Set Container(glistNN As gList)
Set glistN = glistNN
glistN.DropEnabled = True
glistN.DragEnabled = True
glistN.PreserveNpixelsHeaderRightTwips = 8
End Property
Sub Show()
If glistN Is Nothing Then Exit Sub
TextPixelOffset = Int(glistN.TextHeightOffset)
If tmpshow Then tmpshow = False: Exit Sub
glistN.PrepareToShow 5
End Sub
Public Sub PasteDoc(thisdoc As String)

 Const CF_UNICODETEXT = 13
 ClearRedo
If thisdoc = vbNullString Then
thisdoc = GetTextData(CF_UNICODETEXT)
Else
PushUndo
End If

undotext = thisdoc
undopara = mDoc.MarkParagraphID
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + CharPosStart - glistN.SelStart
Else
undopos = CharPosStart
End If
undomode = 0
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, (thisdoc)
Render
End Sub
Public Sub PasteText(thisdoc As String)
PushUndo
undotext = thisdoc + vbCrLf
undopara = mDoc.MarkParagraphID
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + CharPosStart - glistN.SelStart
Else
undopos = CharPosStart
End If
undomode = 0
InsertText = thisdoc + vbCrLf
End Sub

Private Function GetTextData(ByVal lFormatId As Long) As String
Dim bData() As Byte, sr As String, sr1 As String
sr1 = Clipboard.GetText(1)
If (OpenClipboard(0) <> 0) Then

        
    If (GetBinaryData(lFormatId, bData())) Then
        sr = bData
        
        If IsWine Then
            sr1 = Myleft$(sr, Len(sr1))
            GetTextData = Myleft$(sr1, Len(sr1))
        Else
            GetTextData = Myleft$(sr, Len(sr1))
        End If
    End If

End If
CloseClipboard
End Function
Private Function GetClipboardMemoryHandle( _
        ByVal lFormatId As Long _
    ) As Long

    

    If (IsClipboardFormatAvailable(lFormatId) <> 0) Then
         
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
        
    End If
End Function
Private Function GetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean

Dim hMem As Long, lSize As Long, lPtr As Long
    

    Erase bData
    
    hMem = GetClipboardMemoryHandle(lFormatId)

    If (hMem <> 0) Then

        lSize = GlobalSize(hMem)

        lPtr = GlobalLock(hMem)
        If (lSize > 0) Then

            ReDim bData(0 To lSize - 2) As Byte

            CopyMemory bData(0), ByVal lPtr, lSize - 1
        End If

        GlobalUnlock hMem

        GetBinaryData = (lSize > 0)

    End If
End Function


Public Sub HorSrcoll(widthtwips As Long)
glistN.PanPos = widthtwips
End Sub
Public Property Get FileName() As String
FileName = myfilename
End Property
Sub NewTitle(RHS As String, addpixelstoo As Long, Optional factor As Single = 1)
Title = RHS


Set mDoc1 = New Document
If Title = vbNullString Then
Else
NoHeadwrap = False
mDoc1.textDoc = Title
If glistN Is Nothing Then Exit Sub
glistN.VerticalCenterText = True
If mDoc1.DocLines >= glistN.lines + 1 Then
mDoc1.EmptyDoc
NoHeadwrap = True
mDoc1.textDoc = Title
End If
glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight + addpixelstoo
glistN.addpixels = 4 * factor ' 'line spacing
End If
all = mDoc.DocLines
glistN.MultiLineEditBox = True
glistN.ShowBar = True
glistN.ShowBar = False  ' now is in autohide mode
''glistN.ShowMe
glistN.StickBar = True
mDoc.WrapAgain
mDoc1.WrapAgain
glistN.PrepareToShow 10
all = mDoc.DocLines
glistN.CalcAndShowBar
End Sub
Public Property Let FileName(ByVal RHS As String)
myfilename = RHS
onelineheight = 0
Dim thisdoc As New Document, Header As New Document
Set mDoc = thisdoc
Set mDoc1 = Header
Dim i As Long
glistN.NoCaretShow = True
ClearRedo
ClearUndo
If RHS = vbNullString Then
mDoc.textDoc = vbNullString
Else
mDoc.textDoc = ReadUnicodeOrANSI(RHS)
End If

Dim what$, ThatPara As Long, pastchars
    pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
NoTitle = Title = vbNullString
nowrap = True


If Not NoTitle Then
mDoc1.textDoc = Title

glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight
End If
''glistN.AddPixels = 4  ' use this before  ---- changed for m2000 version 2
all = mDoc.DocLines
glistN.MultiLineEditBox = True
glistN.ShowBar = True
glistN.ShowBar = False  ' now is in autohide mode
glistN.ShowMe
glistN.StickBar = True
mDoc.WrapAgain
mDoc1.WrapAgain
glistN.PrepareToShow 10
'glistN.ShowMe2
all = mDoc.DocLines
glistN.CalcAndShowBar
'glistN.FloatList = True
End Property
Public Sub ResetSelColors()
SelectionTextColor = &HFFFFFF
SelectionColor = &H666666
End Sub
Property Let TabWidth(RHS As Long)
    tParam.iTabLength = Abs(RHS)
    glistN.TabWidthChar = Abs(RHS)
End Property


Private Sub Class_Initialize()
tParam.cbSize = LenB(tParam)
tParam.iTabLength = 4
sp1 = ",+-*/=<>^!.&[]{}()^;""" + "'!"
sp = " " + ChrW(9) + ChrW(160)
sp2 = sp1 + sp
Set mDoc = New Document
Set mDoc1 = New Document
ResetSelColors
mDoc.LCID = Clid
mDoc1.LCID = Clid
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
undopara = -1
Redopara = -1
HighLightColor = rgb(66, 66, 66)
End Sub

Private Sub glistN_addone(that As String)  '
' push undo addone
' from delete or backspace..If Not enabled Then Exit Sub
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 2
ElseIf that = vbCr Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = vbCrLf
undopos = SelStart
undomode = 1
ElseIf undomode = 1 And undotext <> "" Then
If undopos = SelStart Then
' we add after
undotext = undotext + that
ElseIf undopos = SelStart + 1 Then
    undotext = that + undotext
    undopos = SelStart
Else
GoTo there
End If
Else
there:
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart ' always the left side of char - so 0 is inside
undomode = 1

End If
End Sub
Public Sub AddUndo(that As String)
If Not enabled Then Exit Sub
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 2
Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart ' always the left side of char - so 0 is inside
undomode = 1
End If
End Sub
Public Sub GroupUndo()
undoprev = True
End Sub
Public Sub RemoveUndo(that As String)
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 3
ElseIf that = vbCr Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 3
Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
undopos = SelStart
End If
undomode = 0
End If
End Sub
Public Sub dothis()
glistN.MarkWord
End Sub
Public Sub insertbrackets()
Dim i As Long, j As Long, ss$, idnt$, tc As Long
Dim ww As Long
If SelLength > 0 Then
ClearRedo
PushUndo
glistN_MarkDelete False
ss$ = undotext
undoprev = True
End If
Dim A$
If glistN.UseTab Then
tc = 0
Else
tc = EditTabWidth
End If
glistN.enabled = False
    A$ = Myleft(CurrentParagraph, Charpos)
    
    If MyTrimL2NoTab(A$) <> 1 And Len(A$) > 0 And tc = 0 Then
      tc = Int(glistN.UserControlTextWidthPixels(Chr(9)) / glistN.UserControlTextWidthPixels("m") * 2)
    End If
    If tc = 0 Then

    j = SelStart
    i = Len(A$) - Len(NLtrim$(A$))
    idnt$ = Myleft$(A$, i)
   
    If ss$ <> "" Then
        ss$ = Replace(ss$, vbCrLf, vbCrLf + vbTab)
        RemoveUndo "{" + vbCrLf + idnt$ + vbTab + ss$ + vbCrLf + idnt$ + "}"
        InsertText = "{" + vbCrLf + idnt$ + vbTab + ss$ + vbCrLf + idnt$ + "}"
    Else
        RemoveUndo "{" + vbCrLf + idnt$ + vbTab + vbCrLf + idnt$ + "}"
        InsertText = "{" + vbCrLf + idnt$ + vbTab + vbCrLf + idnt$ + "}"
    End If
    SelStartSilent = j + Len(idnt$) + 4
    
    
    Else
    j = SelStart
    
    ww = tc
    i = Len(A$) - Len(NLtrim$(A$))
    idnt$ = Myleft$(A$, i)
   
    If ss$ <> "" Then
        ss$ = Replace(ss$, vbCrLf, vbCrLf + space$(ww))
        RemoveUndo "{" + vbCrLf + idnt$ + space$(ww) + ss$ + vbCrLf + idnt$ + "}"
        InsertText = "{" + vbCrLf + idnt$ + space$(ww) + ss$ + vbCrLf + idnt$ + "}"
    Else
        RemoveUndo "{" + vbCrLf + idnt$ + space$(ww) + vbCrLf + idnt$ + "}"
        InsertText = "{" + vbCrLf + idnt$ + space$(ww) + vbCrLf + idnt$ + "}"
    End If
    SelStartSilent = j + Len(idnt$) + 3 + ww
    End If
glistN.enabled = True
Render
End Sub

Public Sub ManualInform()

If Not glistN.enabled Then Exit Sub
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)

Else

RaiseEvent inform(glistN.ListIndex + 1, glistN.SelStart)
End If
If NoCenterLineEdit Then

Else
glistN.ShowMe
End If
End Sub

Private Sub glistN_CorrectCursorAfterDrag()
    If Not mDoc.InvalidPara(OldPar) Then
        SelLengthSilent = 0
        mDoc.MarkParagraphID = OldPar
        glistN.enabled = False
        ParaSelStart = dragCharpos
        glistN.enabled = True
        ManualInform
        
    End If
End Sub

Private Sub glistN_CtrlPlusF1()
RaiseEvent CtrlPlusF1
End Sub

Private Sub glistN_DragOverCursor(ok As Boolean)
ok = True
End Sub

Private Sub glistN_ExpandSelStart(val As Long)
If val < 1 Then val = 1
glistN.SelStartNoEvents = expandCharPos
val = expandCharPos
FixCharPos
expandCharPos = val
End Sub
Private Sub glistN_SetExpandSS(val As Long)
expandCharPos = val
End Sub
Private Sub glistN_DelExpandSS()
expandCharPos = 0
End Sub
Private Sub FixCharPos()
Dim what$, ThatPara As Long, pastchars As Long
pastchars = mDoc.TextParagraphLine2(glistN.ListIndex + 1, what$, ThatPara, False)
If Len(what$) + 1 < glistN.SelStart Then
    Charpos = pastchars + Len(what$) + 1
    glistN.SelStartNoEvents = Len(what$) + 1
Else
    Charpos = pastchars + glistN.SelStart
End If
End Sub

Private Sub glistN_Fkey(A As Integer)
Dim s$
If Me.UsedAsTextBox Then
If FK$(A) <> "" Then
    s$ = FK$(A)
    interpret Basestack1, s$
Fkey = 0
End If
A = 0
End If
End Sub

Private Sub glistN_getpair(A As String, b As String)
If Not UsedAsTextBox Then
Select Case A
Case Chr$(34)
b = Chr$(34)
Case "("
b = ")"
Case "{"
b = "}"
Case "["
b = "]"
End Select
End If
End Sub

Private Sub glistN_GroupUndo()
If undopara <> -1 Then
undoprev = True
End If
End Sub

Private Sub glistN_HaveMark(Yes As Boolean)
Yes = HaveMarkedText
End Sub

Private Sub glistN_MarkCut(ThatData As String)
expandCharPos = 0
ThatData = undotext
End Sub

Private Sub glistN_Maybelanguage()
If Not glistN.enabled Then Exit Sub

If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, glistN.SelStart)
End If
Render
End Sub

Private Sub glistN_PureListOff()
showparagraphWork = True

End Sub

Private Sub glistN_PureListOn()
showparagraphWork = False

End Sub

Private Sub glistN_PushMark2Undo(that As String)
ClearRedo
undopara = BlockEndPara
undotext = that
undomode = 1
undopos = SelStart
Debug.Print "PushMark2Undo"
End Sub

Private Sub glistN_PushUndoIfMarked()
If HaveMarkedText Then PushUndo
End Sub

Private Sub glistN_RealCurReplace(A As String)
If showparagraph Then
A = Replace(A, vbTab, ChrW(&H21E5))
End If

End Sub



Private Sub glistN_RemoveOne(that As String)
' push undo removeone
If Not enabled Then Exit Sub
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 3
ElseIf undomode = 0 And undotext <> "" Then
If undopos + Len(undotext) = SelStart Then
' we add after
undotext = undotext + that

Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
If glistN.ListIndex = all Then
undopos = SelStart - 1
Else
undopos = SelStart
End If
End If
 ' always the left side of char - so 0 is inside
undomode = 0
End If

Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
undopos = SelStart
End If
undomode = 0
End If
End Sub


Private Sub glistN_MayRefresh(ok As Boolean)
ok = Not HaveMarkedText
End Sub


Private Sub glistN_ChangeListItem(item As Long, content As String)
Dim selstart2 As Long, checklistindex As Long
'' only for M2000, use NumberOnly
If NumberOnly Then
If Not ValidNum(content, False, NumberIntOnly) Then
undotext = vbNullString
content = mDoc.TextLine(item + 1)
Exit Sub
End If
End If
If UsedAsTextBox Then
If Len(content) > glistN.maxchar Then
'IsRTLrun
Charpos = glistN.maxchar
glistN.SelStartEventAlways = glistN.SelStart - 1
Exit Sub
End If
End If
mDoc.TextLine(item + 1) = content
'IsRTLrun
WrapMarkedPara
If all <> mDoc.DocLines Then
    Render
Else
    checklistindex = mDoc.FirstParagraphLine(mDoc.MarkParagraphID) + mDoc.TextParagraph2Line(mDoc.MarkParagraphID, Charpos, selstart2)
    If (checklistindex - glistN.ScrollFrom) >= glistN.lines + 1 Then
     glistN.ListindexPrivateUse = checklistindex
     glistN.ScrollToTextEdit checklistindex - glistN.lines
     glistN.SelStartEventAlways = selstart2
     

    
    ElseIf checklistindex < glistN.ScrollFrom Then
     glistN.ListindexPrivateUse = checklistindex
     glistN.ScrollToTextEdit checklistindex
          glistN.SelStartEventAlways = selstart2
    Else
    glistN.ListindexPrivateUse = checklistindex
    glistN.SelStart = selstart2
        End If
End If

End Sub

Private Sub glistN_ChangeSelStart(thisselstart As Long)
If glistN.ListIndex < 0 Then Exit Sub

Dim what$, ThatPara As Long, pastchars As Long, oldpanpos As Long
pastchars = mDoc.TextParagraphLine2(glistN.ListIndex + 1, what$, ThatPara, False)
Charpos = pastchars + glistN.SelStart

'mDoc.MarkParagraphID = ThatPara

If Not glistN.enabled Then Exit Sub
oldpanpos = glistN.PanPos
tmpshow = True
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, glistN.SelStart)
End If
On Error Resume Next

'If NoCenterLineEdit Then Exit Sub

'If LastListIndex <> glistN.ListIndex Then

'LastListIndex = glistN.ListIndex: glistN.ShowMe

'End If

If LastListIndex <> glistN.ListIndex Then

LastListIndex = glistN.ListIndex

    If NoCenterLineEdit Then
       If Not tmpshow Then glistN.ShowMe2
    Else
    tmpshow = False
    If oldpanpos <> glistN.PanPos Then Show Else glistN.ShowMe
    End If
Else
If NoCenterLineEdit Then
 'glistN.ShowPan
Else
Show
End If
    
End If
tmpshow = False
Exit Sub
If NoCenterLineEdit Then
    If LastListIndex <> glistN.ListIndex Then
        LastListIndex = glistN.ListIndex:  glistN.ShowMe: If Not tmpshow Then glistN.ShowMe2
    Else
        If Not tmpshow Then glistN.ShowMe2
    End If

Else
    glistN.ShowMe
    If Not tmpshow Then glistN.ShowMe2
End If

End Sub

Private Sub glistN_DragData(ThatData As String)
FillBlock
ThatData = UndoDoc.textDoc
 OldPar = mDoc.MarkParagraphID
 dragCharpos = Charpos
End Sub

Private Sub glistN_DragPasteData(ThatData As String)
Dim A() As String, k As Long, a77$
a77$ = space$(EditTabWidth)
If Not glistN.UseTab Then

A() = Split(ThatData, ChrW(9))
PasteDoc Join(A(), a77$)
Else
PasteDoc ThatData
End If
RelocateMarkedText

End Sub

Private Sub glistN_DropFront(ok As Boolean)
ok = False
' we need to know if we delete mark before paste data
' when we perform dragdrop to the same document
' by defult ok=false so we drop behind so we can delete marked text and then we paste
If HaveMarkedText Then
If glistN.ListIndex < BlockStartLine Then
' we are behind
ElseIf glistN.ListIndex = BlockStartLine Then
If glistN.SelStart < BlockStartCursor Then
Else
ok = True
End If
Else
ok = True
End If
End If
End Sub

Private Sub glistN_DropOk(ok As Boolean)
If HaveMarkedText Then
If ClickInsideMarkedText Then
Else
ok = True
End If
Else
ok = True
End If
End Sub
Private Sub glistN_DragOverDone(A As Boolean)

If Not mDoc.InvalidPara(OldPar) Then
        A = True
        SelLengthSilent = 0
        mDoc.MarkParagraphID = OldPar
        glistN.enabled = False
        ParaSelStart = dragCharpos
        glistN.enabled = True
        ManualInform
        OldPar = 0
    End If
End Sub
Private Sub glistN_DragOverNow(A As Boolean)
If OldPar = 0 Then
A = True
 SelLengthSilent = 0
 OldPar = mDoc.MarkParagraphID
 dragCharpos = Charpos
End If
End Sub
Private Sub glistN_ExposeListcount(cListCount As Long)
cListCount = all
End Sub
Private Sub gListN_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim A As RECT, b As RECT, i As Long, tmp As Long
Dim Final$, fbypass As Boolean
Dim bb As RECT, cc As RECT
Dim OldColor As Long
Dim what$, ThatPara As Long
CopyFromLParamToRect A, thisrect
b = A
bb = A
cc = A
cc.Left = 0
If UsedAsTextBox Then
b.Bottom = A.Bottom - 2
End If
Dim skipcolor As Boolean
OldColor = glistN.ForeColor
If item < 0 Then  ' this is a headline
    FillBack thisHDC, b, glistN.CapColor
    mDoc1.TextParagraphLine2 1, what$, ThatPara, False
    showparagraphonce = False
    glistN.HeadlineHeight = mDoc1.DocLines * onelineheight
    For i = 1 To mDoc1.DocLines
        PrintItem2 thisHDC, mDoc1.TextLine2(i, False, True), A
        A.top = A.top + onelineheight
    Next i
    RaiseEvent ExposeRect(item, thisrect, thisHDC, skip)
    skip = True
    Exit Sub
Else
    showparagraphWork = showparagraph
    A.top = A.top + glistN.addpixels \ 2
    mDoc.TextParagraphLine2 item + 1, what$, ThatPara, showparagraph
    If MarkedTextAny Then
        showparagraphWork = True
        If item > BlockTopline And item < BlockBottomline Then
            FillBack thisHDC, b, SelectionColor
            SetTextColor thisHDC, SelectionTextColor
            skipcolor = True
        ElseIf item = BlockTopline And item <> BlockBottomline Then
            If BlockEndline < BlockStartLine Then
                If BlockEndCursor > 1 Then  ' if not leave it as is
                    b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1))
                    PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, mDoc.lastcolordata()
                    Final$ = Mid$(glistN.list(item), BlockEndCursor)
                    fbypass = True
                End If
            Else
                If BlockStartCursor > 1 Then  ' if not leave it as is
                    b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1))
                    PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, mDoc.lastcolordata()
                    Final$ = Mid$(glistN.list(item), BlockStartCursor)
                    fbypass = True
                 End If
            End If
            FillBack thisHDC, b, SelectionColor
            SetTextColor thisHDC, SelectionTextColor
            skipcolor = True
        ElseIf item = BlockBottomline And item <> BlockTopline Then
            If BlockBottomline = BlockEndline Then
                If BlockEndCursor <= Len(glistN.list(item)) Then
                    If BlockEndCursor > 1 Then ' if not leave it as is
                        b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                        PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, mDoc.lastcolordata()
                        Final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
                        fbypass = False
                    Else
                        If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor2 RealX1(thisHDC, glistN.SelStart, mDoc.TextLine2(item + 1, showparagraph, NoColor)), OldColor
                        GoTo conthere
                    End If
                End If
            Else
                If BlockStartCursor <= Len(glistN.list(item)) Then  ' if not leave it as is
                    If BlockStartCursor > 1 Then
                        b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                        PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, mDoc.lastcolordata()
                        Final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
                        fbypass = False
                    Else
                        If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor2 RealX1(thisHDC, glistN.SelStart, mDoc.TextLine2(item + 1, showparagraph, NoColor)), OldColor
                        GoTo conthere
                    End If
                End If
            End If
            FillBack thisHDC, b, SelectionColor
            A.Right = b.Right + 1
            SetTextColor thisHDC, SelectionTextColor
            skipcolor = True
        ElseIf BlockBottomline = item Then ' in one line only
            skipcolor = False
            If BlockEndCursor > BlockStartCursor Then
                If BlockStartCursor > 0 Then b.Left = b.Left + RealX1(thisHDC, BlockStartCursor, mDoc.TextLine2(item + 1, showparagraph, NoColor))
                If BlockEndCursor > 0 Then b.Right = RealX1(thisHDC, BlockEndCursor, mDoc.TextLine2(item + 1, showparagraph, NoColor)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                If b.Left > b.Right Then
                    tmp = b.Right
                    b.Right = b.Left
                    b.Left = tmp
                End If
                If BlockStartCursor = 0 Then
                    Final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
                Else
                    skipcolor = False
                    Final$ = Mid$(glistN.list(item), BlockStartCursor, BlockEndCursor - BlockStartCursor)
                End If
                PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, mDoc.lastcolordata()
                FillBack thisHDC, b, SelectionColor
                A.Right = b.Right + 1
                SetTextColor thisHDC, SelectionTextColor
            ElseIf BlockEndCursor < BlockStartCursor Then
                If BlockEndCursor > 0 Then b.Left = b.Left + RealX1(thisHDC, BlockEndCursor, mDoc.TextLine2(item + 1, showparagraph, NoColor))
                If BlockStartCursor > 0 Then b.Right = RealX1(thisHDC, BlockStartCursor, mDoc.TextLine2(item + 1, showparagraph, NoColor)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                If b.Left > b.Right Then
                    tmp = b.Right
                    b.Right = b.Left
                    b.Left = tmp
                End If
                If BlockEndCursor > 0 Then
                    Final$ = Mid$(glistN.list(item), BlockEndCursor, BlockStartCursor - BlockEndCursor)
                Else
                    Final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
                End If
                PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, mDoc.lastcolordata()
                FillBack thisHDC, b, SelectionColor
                A.Right = b.Right + 1
                SetTextColor thisHDC, SelectionTextColor
            End If
            fbypass = False
        End If
    ElseIf mEditDoc And Not UsedAsTextBox And Not NoColor Then
       If ThatPara = mDoc.MarkParagraphID Then
        bb = cc
        bb.top = b.top
        bb.Bottom = b.Bottom
        FillBack thisHDC, bb, HighLightColor
        End If
    End If
    If ThatPara = mDoc.MarkParagraphID Then
        If EditDoc Then
            If glistN.ListIndex = item Then
            ' why???
                If mNoWrap Or (Not skipcolor) Then glistN.ExternalCursor2 RealX1(thisHDC, glistN.SelStart, mDoc.TextLine2(item + 1, showparagraph, NoColor)), OldColor
            End If
        ElseIf Not NoMark Then
            FillBack thisHDC, b, 0
            SetTextColor thisHDC, SelectionTextColor
        End If
    ElseIf Not skipcolor Then
        SetTextColor thisHDC, glistN.ForeColor
    End If
    If Final$ <> "" Or fbypass Then
        If Not fbypass Then

            PrintItem2X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, b, mDoc.lastcolordata(), False
        Else

             PrintItem2X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, b, mDoc.lastcolordata(), True
        End If
        showparagraphWork = False
    Else
conthere:
        If skipcolor Or NoColor Then
            PrintItem3X thisHDC, mDoc.TextLine2(item + 1, showparagraph, True), A, mDoc.lastcolordata()
        Else
            PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), A, mDoc.lastcolordata()
        End If
    End If
End If
SetTextColor thisHDC, OldColor
showparagraphonce = False
skip = True
End Sub

Public Sub Render()
If mDoc.Busy Then Exit Sub
Dim selstart2 As Long
mDoc.Busy = True
all = mDoc.DocLines
glistN.ListIndex = mDoc.FirstParagraphLine(mDoc.MarkParagraphID) + mDoc.TextParagraph2Line(mDoc.MarkParagraphID, Charpos, selstart2, NoColor)
glistN.SelStart = selstart2

If glistN.ListIndex < glistN.ScrollFrom Then
    If glistN.ListIndex > glistN.lines Then
        glistN.ScrollToSilent glistN.ListIndex - glistN.lines / 2
    Else
        glistN.ScrollToSilent 0
    End If
'ElseIf glistN.ListIndex > glistN.lines And glistN.listcount - glistN.ListIndex <= glistN.lines Then
 '   glistN.ScrollToSilent glistN.listcount - glistN.lines - 1
ElseIf glistN.ListIndex > glistN.ScrollFrom + glistN.lines Then
    If glistN.listcount - glistN.ListIndex < glistN.lines Then
        glistN.ScrollToSilent glistN.listcount - glistN.lines
    Else
        glistN.ScrollToSilent glistN.ListIndex - glistN.lines / 2
    End If
ElseIf glistN.ListIndex > glistN.lines And glistN.listcount - glistN.ListIndex < glistN.lines Then
If glistN.enabled Then glistN.ScrollToSilent glistN.listcount - glistN.lines
End If

glistN.ScrollToSilent glistN.ScrollFrom

If Not NoCenterLineEdit Then
If glistN.enabled Then glistN.ShowPan
Else
If glistN.enabled Then glistN.ShowMe2
End If
'If glistN.enabled Then glistN.ShowPan
mDoc.Busy = False
End Sub

Public Sub WrapMarkedPara()
If glistN.ListIndex >= 0 Then
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID)
End If
End Sub




Private Sub glistN_KeyDown(keycode As Integer, shift As Integer)
If keycode = vbKeyTab Then
If shift = 2 Then choosenext
        
    keycode = 0
    Exit Sub
    

End If

Dim selstart2 As Long, b As Boolean, Pad$, cc As Long
Dim l1 As Long, c1 As Long
If keycode = 18 Then
Exit Sub
End If
If keycode = 16 And shift <> 0 Then
keycode = 0
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, glistN.SelStart)
End If
Exit Sub
End If
If NoMark Then keycode = 0: Exit Sub

If keycode = vbKeyLeft Or keycode = vbKeyUp Or keycode = vbKeyDown Or keycode = vbKeyRight Or keycode = vbKeyEnd Or keycode = vbKeyHome Or keycode = vbKeyPageUp Or keycode = vbKeyPageDown Then
ElseIf keycode = vbKeyX And (shift And 3) = 2 Then
MarkCut
Show
keycode = 0
ElseIf keycode = vbKeyC And (shift And 3) = 2 Then
MarkCopy
keycode = 0
ElseIf keycode = vbKeyV And (shift And 3) = 2 Then
Pad$ = GetTextData(13)
If Pad$ = vbNullString Then Pad$ = Clipboard.GetText(1)
If ParaSelStart = 2 And glistN.list(glistN.ListIndex) = vbNullString Then
SelStart = SelStart - 1
End If
AddUndo ""
If UsedAsTextBox Then
Pad$ = GetNextLine((Pad$))
If Len(Text) + Len(Pad$) - SelLength > glistN.maxchar Then
Pad$ = Left$(Pad$, glistN.maxchar - Len(Text))
End If
End If

SelText = Pad$
RemoveUndo SelText

keycode = 0
'mDoc.WrapAgainColorBlock mDoc.MarkParagraphID, mDoc.MarkParagraphID
ElseIf keycode = vbKeyA And (shift And 3) = 2 Then
glistN.MarkNext = 0
glistN.MarkALL
keycode = 0
ElseIf keycode = vbKeyY And (shift And 3) = 2 Then
' Redo
profundo.MARKONE
againY:
If profundo.MARKTWO > 1000 Then ProcTask2 Basestack1: profundo.MARKONE
If Redopara <> -1 Then

PushUndo

undopara = Redopara
undopos = Redopos
undotext = Redotext
undomode = Redomode
If undomode = 2 Then undopos = Redopos - 2
undoprev = Redoprev
b = enabled
enabled = False
SelLength = 0
If Redomode = 3 Then
SelStart = Redopos - 2
Else

SelStart = Redopos
End If
If Redomode = 0 Or Redomode = 5 Then ' now we do the opposite

If Redomode = 5 Then


PopRedo

If glistN.SelStart = 0 Then
SelLength = Len(undotext) + 1
Else
SelLength = Len(undotext)
End If
    glistN_MarkDelete True
    glistN.MarkNext = 0
PushUndo
InsertText = Redotext
undopara = Redopara
undopos = Redopos
undotext = Redotext
undomode = Redomode
undoprev = Redoprev
Else
InsertText = Redotext

End If
ElseIf Redomode = 1 Then
If glistN.SelStart = 0 Then
SelLength = Len(Redotext) + 1
Else
SelLength = Len(Redotext)
End If
SelText = vbNullString
ElseIf Redomode = 2 Then
   keyzLineUp
ElseIf Redomode = 3 Then
glistN_SplitLine

End If

enabled = b
If Not undoprev Then
glistN.ShowPan
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, glistN.SelStart)
End If
End If
If Not PopRedo Then
Redopara = -1
Redotext = vbNullString
Redomode = 0
End If
If undoprev Then GoTo againY
End If

keycode = 0
ElseIf keycode = vbKeyZ And (shift And 3) = 2 Then
' 2nd version of undo system
' undopos now is selstart
profundo.MARKONE
againZ:
If profundo.MARKTWO > 1000 Then ProcTask2 Basestack1: profundo.MARKONE
If undopara <> -1 Then
' seve to redo...
PushRedo
Redopara = undopara
Redopos = undopos
Redotext = undotext
If undomode = 2 Then Redopos = undopos + 2
Redomode = undomode
Redoprev = undoprev

' MOVE TO PARAGRAPH...AND CHARPOS
b = enabled
enabled = False

   SelLength = 0
   SelStart = undopos

If undomode = 0 Or undomode = 4 Then

    If glistN.SelStart = 0 Then
        SelLength = Len(undotext) + 1
    Else
        SelLength = Len(undotext)
    End If
    
    If undomode = 4 Then
    glistN_MarkDelete True
    glistN.MarkNext = 0
    PopUndo
    PushRedo
        Redopara = undopara
        Redopos = undopos
        Redotext = undotext
        Redomode = undomode
        Redoprev = undoprev
    ' now we have undomode 5
    InsertText = undotext
    Else
    SelText = vbNullString
    End If
ElseIf undomode = 3 Then
    keyzLineUp
ElseIf undomode = 2 Then
    glistN_SplitLine
Else
        InsertText = undotext
End If
enabled = b
If Not undoprev Then
glistN.ShowPan
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, glistN.SelStart)
End If

End If

End If
 If Not PopUndo Then
 ClearUndo
undopara = -1
undoprev = False
undotext = vbNullString
undomode = 0
End If
If undoprev Then GoTo againZ

keycode = 0
ElseIf keycode = vbKeyA And (shift And 3) = 3 Then
glistN.MarkNext = 0
BlockEndPara = mDoc.MarkParagraphID
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.ShowPan
keycode = 0

ElseIf keycode = vbKeyDelete Then
If HaveMarkedText Then
ClearRedo
PushUndo
glistN_MarkDelete False

keycode = 0
End If
ElseIf keycode = vbKeyBack Then
If HaveMarkedText Then
ClearRedo
PushUndo
glistN_MarkDelete False

keycode = 0
End If
End If

End Sub

Private Sub glistN_KeyDownAfter(keycode As Integer, shift As Integer)
If NoMark Then Exit Sub
If HaveMarkedText Then
If shift = 0 Then
glistN.MarkNext = 0
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.RefreshNow
ElseIf (shift And 1) = 1 Then

If keycode = vbKeyLeft Or keycode = vbKeyUp Or keycode = vbKeyDown Or keycode = vbKeyRight Or keycode = vbKeyEnd Or keycode = vbKeyHome Or keycode = vbKeyPageUp Or keycode = vbKeyPageDown Then
glistN_MarkOut
End If

End If
ElseIf (shift And 1) = 1 And glistN.MarkNext = 0 Then
If keycode = vbKeyLeft Or keycode = vbKeyUp Or keycode = vbKeyDown Or keycode = vbKeyRight Or keycode = vbKeyEnd Or keycode = vbKeyHome Or keycode = vbKeyPageUp Or keycode = vbKeyPageDown Then

glistN.MarkNext = 1
glistN_MarkIn

End If
ElseIf (shift And 1) = 1 And glistN.MarkNext > 0 Then
If keycode = vbKeyLeft Or keycode = vbKeyUp Or keycode = vbKeyDown Or keycode = vbKeyRight Or keycode = vbKeyEnd Or keycode = vbKeyHome Or keycode = vbKeyPageUp Or keycode = vbKeyPageDown Then
glistN.MarkNext = 2
glistN_MarkOut
End If
End If
End Sub

Private Sub glistN_LineDown()
Dim b$, nline As Long
mDoc.TextParagraphLine glistN.ListIndex + 1, b$, nline
mDoc.MarkParagraphID = nline

If mDoc.IsLast(mDoc.MarkParagraphID) Then Exit Sub

  
 
 If mDoc.TextParagraphLen(nline) = 0 Then
 Charpos = 1
  mDoc.RemoveEmptyParagraph nline
 mDoc.MarkParagraphID = nline
 Else
 Charpos = ParaSelStart
  mDoc.ReWritePara nline, mDoc.TextParagraph(nline) + mDoc.DeleteNextParagraph(nline)

 End If
 
  all = mDoc.DocLines
 
  ParaSelStart = Charpos
End Sub

Private Sub glistN_LineUp()
Dim olID As Long, b$ ', c$
olID = mDoc.MarkParagraphID
If ParaSelStart < 2 Then
If mDoc.ParagraphOrder(mDoc.MarkParagraphID) > 1 Then
glistN.PressKey vbKeyLeft, 0
If olID <> mDoc.MarkParagraphID Then
b$ = mDoc.DeleteNextParagraph(mDoc.MarkParagraphID)
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID) + b$
 all = mDoc.DocLines
ClearRedo
PushUndo
undotext = ""
undopara = mDoc.MarkParagraphID
undopos = SelStart
undomode = 2
'End If
Render
End If
End If
Else
glistN.PressKey vbKeyLeft, 0
End If
all = mDoc.DocLines

End Sub

Private Sub keyzLineUp()
Dim olID As Long, b$ ', c$
olID = mDoc.MarkParagraphID
If ParaSelStart < 2 Then
If mDoc.ParagraphOrder(mDoc.MarkParagraphID) > 1 Then
glistN.PressKey vbKeyLeft, 0
If olID <> mDoc.MarkParagraphID Then
b$ = mDoc.DeleteNextParagraph(mDoc.MarkParagraphID)
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID) + b$
Render
End If
End If
Else
glistN.PressKey vbKeyLeft, 0
End If
all = mDoc.DocLines
End Sub

Public Sub MarkCut()
MarkCopy
PushUndo
glistN_MarkDelete False

End Sub
Private Sub FillBlock()
Dim order1 As Long, order2 As Long, A$, b$, i As Long, nn As Long
' TAKE FIRST VBCRLF IF FIRST LINE =""
On Error GoTo fbl

 Set UndoDoc = New Document
 UndoDoc.CrTail = False
    If BlockStartPara = BlockEndPara Then
        b$ = mDoc.TextParagraph(BlockStartPara)
        
        UndoDoc.CrTail = False
        If b$ = vbNullString Then
        Else
        If BlockStartPos < BlockEndPos Then
           '' UndoDoc.textDoc = Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos)
            UndoDoc.textParagraphPlusColor Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos), mDoc.Paracolordata(BlockStartPara), BlockStartPos - 1
            If BlockEndPos > Len(b$) + 1 Then UndoDoc.CrTail = True
        Else
          '' UndoDoc.textDoc = Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos)
            UndoDoc.textParagraphPlusColor Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos), mDoc.Paracolordata(BlockStartPara), BlockEndPos - 1
            If BlockStartPos > Len(b$) + 1 Then UndoDoc.CrTail = True
        End If
        End If
    Else
        order1 = mDoc.ParagraphOrder(BlockStartPara)
        order2 = mDoc.ParagraphOrder(BlockEndPara)
        If order2 > order1 Then '  normal
            A$ = mDoc.TextParagraph(BlockStartPara)
            UndoDoc.LfLead = False
         If BlockStartPos < 1 Then
               UndoDoc.AppendParagraphPlusColor A$, mDoc.Paracolordata(BlockStartPara)
               '' UndoDoc.AppendParagraph a$
            Else
            If Len(A$) + 1 < BlockStartPos Then
                UndoDoc.LfLead = True
                UndoDoc.AppendParagraph ""  '' no color
               Else
            
                ''UndoDoc.AppendParagraph Mid$(a$, BlockStartPos)
                UndoDoc.AppendParagraphPlusColor Mid$(A$, BlockStartPos), mDoc.Paracolordata(BlockStartPara), BlockStartPos - 1
                End If
            End If
            UndoDoc.MarkParagraphID = 1
            nn = BlockStartPara
            order2 = order2 - 1
            If order2 > order1 Then
                For i = order1 + 1 To order2
               ' UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn))
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraphPlusColor(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn), mDoc.Paracolordata(nn))
                Next i
            End If
                b$ = mDoc.TextParagraph(BlockEndPara)
                UndoDoc.CrTail = False
                
               If BlockEndPos > 0 Then
               ''UndoDoc.AppendParagraph Left$(b$, BlockEndPos - 1) 'maybe is empty
               UndoDoc.AppendParagraphPlusColor Left$(b$, BlockEndPos - 1), mDoc.Paracolordata(BlockEndPara)
               Else
               UndoDoc.CrTail = True
               End If
        Else
            A$ = mDoc.TextParagraph(BlockEndPara)
        '   If a$ = VbNullString Then UndoDoc.AppendParagraph ""
         UndoDoc.LfLead = False
            If BlockEndPos < 1 Then
        
              UndoDoc.AppendParagraphPlusColor A$, mDoc.Paracolordata(BlockEndPara)
               ' UndoDoc.AppendParagraph a$
                
            Else
                 If Len(A$) + 1 < BlockEndPos Then
                UndoDoc.LfLead = True
                UndoDoc.AppendParagraph ""
            Else
                ' UndoDoc.AppendParagraph Mid$(a$, BlockEndPos)
                If BlockEndPos >= 1 Then
                  UndoDoc.AppendParagraphPlusColor Mid$(A$, BlockEndPos), mDoc.Paracolordata(BlockEndPara), BlockEndPos - 1
                  Else
                  UndoDoc.AppendParagraphPlusColor Mid$(A$, BlockEndPos), mDoc.Paracolordata(BlockEndPara)
                  End If
                End If
            End If
            UndoDoc.MarkParagraphID = 1
             nn = BlockEndPara
            order1 = order1 - 1
            If order1 > order2 Then
                For i = order2 + 1 To order1
                ''UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn))
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraphPlusColor(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn), mDoc.Paracolordata(nn))
                Next i
            End If
                b$ = mDoc.TextParagraph(BlockStartPara)
                If BlockStartPos < 1 Then BlockStartPos = 1
                  'UndoDoc.AppendParagraph Left$(b$, BlockStartPos - 1) 'maybe is empty
                  UndoDoc.AppendParagraphPlusColor Left$(b$, BlockStartPos - 1), mDoc.Paracolordata(BlockStartPara)
        
        End If
End If
fbl:
End Sub

Public Sub MarkCopy()
If Not NoColor Then ColorBlockOnly
If HaveMarkedText Then
FillBlock
Dim pp1 As String, pp2 As String
pp1 = UndoDoc.textDoc

If Not NoColor Then pp2 = SelHtmlText
Clipboard.Clear  ' always

blockKeyboard = True
MyDoEvents
blockKeyboard = False
 On Error Resume Next
 ProcTask2 Basestack1
''mywait basestack1, 10

If Not NoColor Then SimpleHtmlData pp2
 SetTextData 13, pp1

End If
End Sub
Private Function SetTextDataLong( _
        ByVal lFormatId As Long, _
         dLong As Long _
    ) As Boolean
    ' use strptr and lenb
    Dim hMem As Long, lPtr As Long
 ''   Dim checkme As Long
    Dim lSize As Long
        lSize = 4
    hMem = GlobalAlloc(0, lSize)
If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, dLong, lSize
       '' CopyMemory checkme, ByVal lPtr, lSize
     
        GlobalUnlock hMem
       If (OpenClipboard(0) <> 0) Then
     SetClipboardData lFormatId, hMem
      CloseClipboard
       End If
          
    End If
    

End Function
Private Function SetTextData( _
        ByVal lFormatId As Long, _
        ByVal sText As String _
    ) As Boolean
    If lFormatId = 0 Then Exit Function
    Dim hMem As Long, lPtr As Long
    Dim lSize As Long
        lSize = LenB(sText)
    hMem = GlobalAlloc(0, lSize + 2)
If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, ByVal StrPtr(sText), lSize + 1
        GlobalUnlock hMem
       If (OpenClipboard(0) <> 0) Then
     SetClipboardData lFormatId, hMem
      CloseClipboard
      Else
      GlobalFree hMem
       End If
    End If
End Function

Private Sub glistN_MarkDelete(preservecursor As Boolean)
Dim order1 As Long, order2 As Long, A$, b$, i As Long
Dim oldid As Long, OLDPOS As Long
oldid = mDoc.MarkParagraphID
OLDPOS = Charpos
 If HaveMarkedText Then
 Set UndoDoc = New Document
    If BlockStartPara = BlockEndPara Then
        b$ = mDoc.TextParagraph(BlockStartPara)
        If BlockStartPos < BlockEndPos Then
            UndoDoc.textDoc = Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos)
            b$ = Mid$(b$, 1, BlockStartPos - 1) + Mid$(b$, BlockEndPos)
                        If OLDPOS > BlockStartPos And oldid = BlockStartPara Then
            OLDPOS = OLDPOS - (BlockEndPos - BlockStartPos)
            End If
            Charpos = BlockStartPos
        Else
           UndoDoc.textDoc = Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos)
            b$ = Left$(b$, BlockEndPos - 1) + Mid$(b$, BlockStartPos)
            If OLDPOS > BlockEndPos And oldid = BlockStartPara Then
            OLDPOS = OLDPOS - (BlockStartPos - BlockEndPos)
            End If
            Charpos = BlockEndPos
        End If
        
        mDoc.ReWritePara BlockStartPara, b$
        
    Else
        order1 = mDoc.ParagraphOrder(BlockStartPara)
        order2 = mDoc.ParagraphOrder(BlockEndPara)
  
        If order2 > order1 Then '  normal
            UndoDoc.AppendParagraph "" ' start paragraph no 1
            UndoDoc.MarkParagraphID = 1
        
            b$ = mDoc.TextParagraph(BlockEndPara)
            mDoc.RemoveDocParaIndex BlockEndPara
            If BlockEndPos < 1 Then
            UndoDoc.AppendParagraph ""
            b$ = vbNullString
            Else
            UndoDoc.AppendParagraph Left$(b$, BlockEndPos - 1) 'maybe is empty
            b$ = Mid$(b$, BlockEndPos)
            End If
            
            
            order2 = order2 - 1
            If order2 > order1 Then
                For i = order1 + 1 To order2
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.DeleteNextParagraph(BlockStartPara))
                    
                Next i
            End If
            A$ = mDoc.TextParagraph(BlockStartPara)
            UndoDoc.ReWritePara 1, A$, True
            If BlockStartPos = 1 Then
                mDoc.ReWritePara BlockStartPara, b$
                UndoDoc.ReWritePara 1, A$, True
            Else
            UndoDoc.ReWritePara 1, Mid$(A$, BlockStartPos), True
                mDoc.ReWritePara BlockStartPara, Left$(A$, BlockStartPos - 1) + b$
            End If
            mDoc.MarkParagraphID = BlockStartPara
            Charpos = BlockStartPos
        Else
            UndoDoc.AppendParagraph "" ' start paragraph no 1
            UndoDoc.MarkParagraphID = 1
             b$ = mDoc.TextParagraph(BlockStartPara)
            mDoc.RemoveDocParaIndex BlockStartPara
             UndoDoc.AppendParagraph Left$(b$, BlockStartPos - 1) 'maybe is empty
            b$ = Mid$(b$, BlockStartPos)
            order1 = order1 - 1
            If order1 > order2 Then
                For i = order2 + 1 To order1
                    
                    UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.DeleteNextParagraph(BlockEndPara))
               
                Next i
            End If
            A$ = mDoc.TextParagraph(BlockEndPara)
            
            If BlockEndPos = 1 Then
                mDoc.ReWritePara BlockEndPara, b$
                UndoDoc.ReWritePara 1, A$
            Else
            UndoDoc.ReWritePara 1, Mid$(A$, BlockEndPos)
                mDoc.ReWritePara BlockEndPara, Left$(A$, BlockEndPos - 1) + b$
            End If
            mDoc.MarkParagraphID = BlockEndPara
            
            Charpos = BlockEndPos
        
        
        End If

End If
all = mDoc.DocLines
If Not preservecursor Then
undopara = mDoc.MarkParagraphID
undopos = SelStart
undotext = UndoDoc.textDoc
undomode = 1 ' insert text at position changed to 1
End If
Set UndoDoc = New Document


glistN.MarkNext = 0
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
If preservecursor Then
If mDoc.ParagraphOrder(oldid) > 0 Then

mDoc.MarkParagraphID = oldid
Charpos = OLDPOS
End If
End If
Render
End If

End Sub

'
Private Sub glistN_MarkDestroy()
' destroy or open to change
If Not EditDoc Then Exit Sub
If Not ClickInsideMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
Else
glistN.MarkNext = 1
' we have to raise event now
End If
End Sub

Private Sub glistN_MarkDestroyAny()
' destroy always
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End Sub
Friend Sub markin()
    glistN_MarkIn
End Sub
Friend Sub MarkOut()
    glistN_MarkOut
End Sub
Private Sub glistN_MarkIn()
If Not glistN.EditFlag Then glistN.MarkNext = 0: Exit Sub
expandCharPos = 0
BlockStartPara = mDoc.MarkParagraphID
BlockStartPos = ParaSelStart
BlockStartCursor = glistN.SelStart '- 1
BlockEndCursor = BlockStartCursor + 1

BlockStartLine = glistN.ListIndex
BlockEndline = BlockStartLine
BlockEndPara = BlockStartPara
BlockEndPos = BlockStartPos


End Sub
Function HaveMarkedText() As Boolean
If BlockEndPos < 1 Then BlockEndPos = 1
If glistN.MarkNext = 0 Then
Else
If BlockStartPara <> BlockEndPara Then
HaveMarkedText = True
ElseIf BlockStartPos <> BlockEndPos Then
HaveMarkedText = True
End If
End If
End Function
Function MarkedTextAny() As Boolean
If BlockStartPara <> BlockEndPara Then
MarkedTextAny = True
ElseIf BlockStartPos <> BlockEndPos Then
MarkedTextAny = True
End If
End Function


Function ClickInsideMarkedText() As Boolean
' we need that for drag and drop, copy or move
' it is simple
' we now the current line and selstart of glistN
If MarkedTextAny Then
If BlockTopline <> BlockBottomline Then
If glistN.ListIndex = BlockTopline Then
    If BlockEndline < BlockStartLine Then
    If glistN.SelStart >= BlockEndCursor Then ClickInsideMarkedText = True
    Else
    If glistN.SelStart >= BlockStartCursor Then ClickInsideMarkedText = True
    End If
ElseIf glistN.ListIndex = BlockBottomline Then
    If BlockEndline < BlockTopline Then
    If glistN.SelStart <= BlockStartCursor Then ClickInsideMarkedText = True
    Else
    If glistN.SelStart <= BlockEndCursor Then ClickInsideMarkedText = True
    End If
ElseIf glistN.ListIndex < BlockTopline Or glistN.ListIndex > BlockBottomline Then
' FALSE
Else
    ClickInsideMarkedText = True
End If
ElseIf BlockTopline = glistN.ListIndex Then
If BlockEndCursor < BlockStartCursor Then
If glistN.SelStart >= BlockEndCursor And glistN.SelStart <= BlockStartCursor Then ClickInsideMarkedText = True
Else
If glistN.SelStart >= BlockStartCursor And glistN.SelStart <= BlockEndCursor Then ClickInsideMarkedText = True
End If

End If
End If
End Function
Sub RelocateMarkedText()
' when we resize glist, need to calculate the position of marked text
' we use a second set of variables
'  BlockStartPara As Long, BlockStartPos As Long
' BlockEndPara As Long, BlockEndPos As Long
'
BlockStartLine = mDoc.FirstParagraphLine(BlockStartPara) + mDoc.TextParagraph2Line(BlockStartPara, BlockStartPos, BlockStartCursor)
BlockEndline = mDoc.FirstParagraphLine(BlockEndPara) + mDoc.TextParagraph2Line(BlockEndPara, BlockEndPos, BlockEndCursor)
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If
End Sub
Private Sub glistN_MarkOut()
'
If Not EditDoc Then Exit Sub
expandCharPos = 0
BlockEndPara = mDoc.MarkParagraphID
BlockEndPos = ParaSelStart
BlockEndline = glistN.ListIndex
BlockEndCursor = glistN.SelStart
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If

End Sub

Private Sub glistN_ReadListItem(item As Long, content As String)
If item >= 0 Then

content = mDoc.TextLine2(item + 1, showparagraphWork And showparagraph)

End If
End Sub




'End Sub

Private Sub glistN_ScrollSelected(item As Long, y As Long)
If NoMark Then Exit Sub
Dim what$, ThatPara As Long
   mDoc.TextParagraphLine item, what$, ThatPara
   mDoc.MarkParagraphID = ThatPara
End Sub

Private Sub glistN_Selected(item As Long)
       If undomode = 1 And undotext <> "" Then PushUndo
    
       
If NoMark Then Exit Sub

EditDoc = EditDoc
Dim what$, ThatPara As Long
   mDoc.TextParagraphLine2 item, what$, ThatPara, showparagraphWork And showparagraph
   mDoc.MarkParagraphID = ThatPara

If glistN.DragEnabled Then Exit Sub
If ClickInsideMarkedText Then
    Dim m As Long, z As Long
    If BlockStartLine = BlockEndline Then
        z = InStr(2, SelText, ".")
        If z > 1 Then
            SelStartSilent = SelStart + z
            SelLengthSilent = SelLength - z
            FillBlock
        ElseIf Right$(SelText, 1) = "(" Then
            SelStartSilent = SelStart
            SelLengthSilent = SelLength - 1
            FillBlock
        Else
            m = SelLength
            SelStartSilent = SelStart
            SelLengthSilent = SelLength + 1
            If m + 1 <> SelLength Then
                SelLengthSilent = m
                FillBlock
            ElseIf Right$(SelText, 1) = "(" Then
                SelStartSilent = LastSelStart
                FillBlock
            ElseIf Right$(SelText, 1) = "." Then
               If BlockStartPos < BlockEndPos Then
               m = BlockEndPos
               Else
               m = BlockStartPos
               End If
               z = mDoc.TextParagraphLen(ThatPara)
                While InStr(glistN.WordCharLeft, Mid$(mDoc.TextParagraph(ThatPara), m, 1)) < 2 And m < z
                    m = m + 1
                Wend
                If BlockStartPos < BlockEndPos Then
                SelLengthSilent = m - BlockStartPos
                Else
                SelLengthSilent = m - BlockEndPos
                End If
                
                FillBlock
            Else
                SelLengthSilent = SelLength - 1
            End If
        End If
    End If
End If
End Sub
Private Sub glistN_Selected2(item As Long)
If undomode = 1 And undotext <> "" Then PushUndo
If ClickInsideMarkedText Then
' start drag
If glistN.DragEnabled Then
glistN.MarkNext = 4

glistN.DragNow
End If
End If
End Sub
Private Sub glistN_ShowExternalCursor()
glistN.PrepareToShow 5
End Sub
Private Sub glistN_SplitLine()
'
Dim b$, dummy As Long, A As Long

If glistN.ListIndex >= 0 Then
    b$ = mDoc.TextParagraph(mDoc.MarkParagraphID)
If Charpos = 1 And b$ <> "" Then
    dummy = mDoc.InsertBeforeParagraph(mDoc.MarkParagraphID, "")
        mDoc.ColorThis (dummy)
     all = mDoc.DocLines
    mDoc.MarkParagraphID = dummy
    SetRowColumn glistN.ListIndex + 2, 0

    
  ''
   
ElseIf ParaSelStart = 1 Then
    mDoc.MarkParagraphID = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, "")
    all = mDoc.DocLines
    ParaSelStart = 1

ElseIf ParaSelStart > Len(b$) Then

    mDoc.MarkParagraphID = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, "")
    
      all = mDoc.DocLines
    ParaSelStart = Len(b$)

Else  ' ANYWHERE

dummy = ParaSelStart

    mDoc.ReWritePara mDoc.MarkParagraphID, Left$(b$, ParaSelStart - 1), True    ' write and format
    
    dummy = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, Mid$(b$, dummy))
'    mDoc.startfrom(mDoc.MarkParagraphID) = 0
If mDoc.startfrom(mDoc.MarkParagraphID) <> 0 Then
    mDoc.ColorThis (mDoc.MarkParagraphID)
    If mDoc.startfrom(mDoc.MarkParagraphID) <> mDoc.startfrom(dummy) Then
          mDoc.ColorThis (dummy)
    End If
Else
    mDoc.ColorThis (mDoc.MarkParagraphID), , False
    mDoc.ColorThis (dummy)

End If
 mDoc.MarkParagraphID = dummy
    all = mDoc.DocLines
    ParaSelStart = 1

End If
End If
End Sub


Private Sub glistN_SubSelStart(val As Long, shift As Integer)
expandCharPos = 0
Dim W As Long, A$, l As Long, b$
If HaveMarkedText And Not (shift And 1) = 1 Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
val = 0
Exit Sub
End If
If Charpos < 2 Or NumberOnly Or NumberIntOnly Then Exit Sub
'If Charpos > 0 And Not (NumberOnly Or NumberIntOnly) Then



    A$ = Mid$(mDoc.TextParagraph(mDoc.MarkParagraphID), Charpos - 1)
    If Len(A$) > 0 Then
        W = AscW(A$)
        If (shift And 2) = 2 And Charpos > 1 Then
            A$ = Mid$(mDoc.TextParagraph(mDoc.MarkParagraphID), 1, Charpos - 1)
            l = Len(A$)
            val = l
            If val = 1 Then Exit Sub
            b$ = Mid$(A$, l, 1)
            If InStr(sp1, b$) = 0 Then
                Do While l > 1 And InStr(sp, b$) > 0
                    l = l - 1
                    b$ = Mid$(A$, l, 1)
                Loop
                If InStr(sp1, b$) = 0 Then
                    If l = val Then
                        Do While l > 1 And InStr(sp2, b$) = 0
                            l = l - 1
                            b$ = Mid$(A$, l, 1)
                        Loop
                        If l = 1 And Not InStr(sp2, b$) > 0 Then l = l - 1
                        
                    Else
                     If InStr(sp2, b$) > 0 Then l = l - 1
                    End If
                End If
            Else
                l = l - 1
            End If
            val = val - l
        ElseIf W > -9217 And W < -8192 Then
            val = 2


        End If
       ElseIf Charpos > 0 Then
    W = Len(mDoc.TextParagraph(mDoc.MarkParagraphID))
    If W < Charpos Then val = Charpos - W
    End If
End Sub

Private Sub glistN_AddSelStart(val As Long, shift As Integer)
expandCharPos = 0
If NumberOnly Or NumberIntOnly Then Exit Sub
Dim W As Long, A$, l As Long, b$
If HaveMarkedText And Not (shift And 1) = 1 Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
val = 0
Exit Sub
End If
If Charpos > 0 Then
    A$ = Mid$(mDoc.TextParagraph(mDoc.MarkParagraphID), Charpos, 1)
    If Len(A$) > 0 Then
        W = AscW(A$)
        If (shift And 2) = 2 Then
            A$ = Mid$(mDoc.TextParagraph(mDoc.MarkParagraphID), Charpos)
            l = Len(A$)
            val = 1
            If l < 2 Then Exit Sub
                b$ = Mid$(A$, val, 1)
                If InStr(sp1, b$) = 0 Then
                Do While val < l And InStr(sp, b$) > 0
                    val = val + 1
                    b$ = Mid$(A$, val, 1)
                Loop
                If InStr(sp1, b$) = 0 Then
                    If val = 1 Then
                        Do While val < l And InStr(sp2, b$) = 0
                            val = val + 1
                            b$ = Mid$(A$, val, 1)
                        Loop
                        If InStr(sp2, b$) > 0 Then val = val - 1
                    Else
                        If val < l Then val = val - 1
                    End If
                Else
                    val = val - 1
                End If
            End If
        ElseIf W > -10241 And W < -9216 Then
            val = 2
        
        End If
    End If
End If



End Sub

Private Sub glistN_SyncKeyboard(item As Integer)
If NoMark Then item = 0
If blockKeyboard Then item = 0
If item > 32 And item <> 127 Then GetUndoNext = True
End Sub


Private Sub mDoc_BreakLine(Data As String, datanext As String)
 ' Nowrap = True mean no Wrap
 Dim data1 As String, datanext1 As String
 If nowrap Then
    datanext = Data
    Data = vbNullString
 Else
    If showparagraph Then
         data1 = Replace(Replace(Data, " ", ChrW(&HB7)), ChrW(&H2007), ChrW(&H2423))
        
        glistN.BreakLine data1, datanext1, , ChrW(&HB7)
        If data1 <> "" Then
        
        datanext = Left$(Data, Len(datanext1))
        Data = Right$(Data, Len(data1))
        Else
            datanext = Data
             Data = vbNullString
        End If
      
    Else
        glistN.BreakLine Data, datanext
    End If
End If

End Sub
Private Sub makebase(basecolor As Long)
ReDim Col(1) As Long
Col(0) = basecolor
Col(1) = 1

End Sub
Private Sub addone(ByVal thispos As Long, ByVal thiscolor As Long)
Dim k As Long
If thiscolor = 0 Then thiscolor = 14
k = UBound(Col())
If Col(1) = 0 Then
makebase 14
If thispos <> 1 Then
k = UBound(Col())
Else
Col(0) = thiscolor
Exit Sub
End If
End If
' LOOK THIS SECOND TIME
If k >= 1 Then
If k > 1 Then If Col(k - 1) = thiscolor Then Exit Sub
If Col(k) = thispos Then
Col(k - 1) = thiscolor: Exit Sub
End If
End If

''''If col(k - 1) = thiscolor And k > 1 Then Exit Sub

ReDim Preserve Col(k + 2) As Long
Col(k + 1) = thiscolor
Col(k + 2) = thispos


End Sub
Private Sub ColorizeInnerStatements(A$, k As Long, pos2 As Long, feedback As Long, Optional jump As Boolean = False)
Dim drop$, pos1 As Long, onetime As Boolean, kind As Long, lastkind As Long, feedthis As Long, counter As Long
Dim lastcol As Long
A$ = Replace(A$, vbTab, " ")
If jump Then
    kind = IsLabelOnly((A$), drop$)
    If kind = 0 Then
        addone pos2, 14
        lastcol = 14
        If FastSymbol(A$, "(") Then
            pos2 = k - Len(A$) + 1
        End If
        GoTo again2
    End If
    GoTo fromjump
End If
onetime = True
If k < 1 Then
    makebase 14
    mDoc.Paracolordata(thispara) = Col()
    If mDoc.startfrom(thispara) = thispara Then
        feedback = thispara
        ReDim Col(1)
        Do While mDoc.Advance(thispara)
            If mDoc.startfrom(thispara) = feedback Then
                mDoc.startfrom(thispara) = 0
                mDoc.Paracolordata(thispara) = Col()
            End If
        Loop
        feedback = feedback <> thispara
        Exit Sub
    End If
End If
If feedback > 0 Then
    A$ = Mid$(A$, feedback + 1)
     addone feedback + 1, lastcol
    feedback = 0
    GoTo checkthistoo
Else
    lastcol = 14
End If
again:
If RTrim$(A$) = vbNullString Then
    If pos2 > 0 Then
        If Col((UBound(Col()) \ 2) * 2) <> lastcol Then
            If pos2 < k Then
                addone pos2 + 1, lastcol
            End If
        ElseIf UBound(Col()) < 2 Then
            addone pos2, lastcol
        End If
    Else
        makebase 14
    End If
    mDoc.Paracolordata(thispara) = Col()
    Exit Sub
' FOR TESTS
    pos1 = 0
    For k = 1 To UBound(Col()) Step 2
        If VarType(Col(k)) = 0 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
        If Col(k) > pos1 Then pos1 = Col(k)
        If Col(k) < pos1 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
    Next k
    For k = 0 To UBound(Col()) Step 2
        If Col(k) > 15 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
    Next k
End If
If MaybeIsSymbol(A$, ":", "\") Then
    A$ = Mid$(A$, 2)
    If AscW(Left$(A$, 1) + " ") = 92 Then
        If pos2 > 1 Then
            If Col(UBound(Col())) = 1 Or InStr(" #@.+-~^%$,!({'?|/&)}" + Chr$(34), Mid$(lastdata, pos2 - 1, 1)) > 0 Then
                If pos2 > 0 Then
                    addone pos2, 14
                Else
                    addone 1, 14
                End If
                pos2 = k - Len(A$) + 1
                GoTo novar
            End If
        End If
        A$ = Mid$(A$, 2)
    End If
    If Len(A$) > 1 Then
        If AscW(A$) <> 32 Then
            lastkind = 1
            GoTo again3
        End If
    End If
Else
novar:
    While FastSymbol(A$, ":")
        If pos2 > 0 Then
            addone pos2, 14
        Else
            addone 1, 14
        End If
        pos2 = k - Len(A$) + 1
    Wend
End If
If onetime And MaybeIsSymbol(A$, "(") Then
    GoTo CheckThis
End If
drop$ = "@(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then
    If IsSymbol(A$, "@") Then
        If pos2 > 0 Then
            addone pos2, 14
        Else
            addone 1, 14
        End If
        pos2 = k - Len(A$) + 1
    End If
End If
again500:
kind = IsLabelDot("", A$, drop$)

If kind = 0 Then drop$ = "$(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "->": kind = Abs(FastSymbol(A$, drop$, , 2))
fromjump:
If kind = 0 Then drop$ = vbNullString
If kind > 0 Then
    If ismine(drop$) Then
        lastcol = 11
        If pos2 > 0 Then
            addone pos2, 11
        Else
            addone k - Len(A$) - Len(drop$) + 1, 11
        End If
        If ismine2(drop$) Then
            kind = -2
            If Trim$(A$) <> "" Then
                pos2 = k - Len(A$) + 1
                If ismine22(drop$) Then
                ElseIf MaybeIsSymbol(A$, "{") Then
                    addone pos2, 14
                    pos2 = k - Len(A$) + 1
                    onetime = True
                    GoTo again
                ElseIf ismine5(drop$) Then
                    GoTo again500
                ElseIf ismine3(drop$) Then
                    onetime = True
                    GoTo again
                End If
            End If
        ElseIf ismine1(drop$) Then
            kind = -1
            If Trim$(A$) <> "" Then
                pos2 = k - Len(A$) + 1
                If FastSymbol(A$, "{") Then
                    addone pos2, 14
                    pos2 = k - Len(A$) + 1
                    onetime = True
                    GoTo again
                Else
                    mDoc.Paracolordata(thispara) = Col()
                    Exit Sub
                End If
            End If
        ElseIf ismine33(drop$) Then
            ' fix for case >100, =50...
            pos2 = k - Len(A$) + 1
            Do
                lastcol = 14
                addone pos2, 14
                If FastSymbol(A$, "<") Then
                    If Not FastSymbol(A$, "=") Then
                        FastSymbol A$, ">"
                    End If
                ElseIf FastSymbol(A$, ">") Then
                    FastSymbol A$, "="
                ElseIf FastSymbol(A$, "=") Then
                    FastSymbol A$, "="
                ElseIf FastSymbol(A$, "~") Then
                End If
                pos1 = 1
                drop$ = aheadstatus(A$, False, pos1)
                If pos1 > 0 Then
                    drop$ = Mid$(A$, 1, pos1 - 1)
                    ColorizeExpressions drop$, k - Len(A$) + Len(drop$), pos2, feedthis
                    A$ = Mid$(A$, pos1)
                End If
            Loop Until Not FastSymbol(A$, ",")
                onetime = True
                GoTo here1
            End If
        Else
            If kind = 8 Then
                lastcol = 13
            Else
                lastcol = 10
            End If
            If Len(A$) > 0 Then
                If AscW(A$) = 92 Then
                    lastkind = 1
                End If
            End If
            If pos2 > 0 Then
                If kind = 8 Then
                   addone k - Len(A$) - Len(drop$) + 1, 10
                   addone k - Len(A$), 13
                Else
                    addone pos2, 10
                End If
            Else
                lastkind = 0
                If kind = 8 Then
                    addone k - Len(A$) - Len(drop$) + 1, 10
                    addone k - Len(A$), 13
                Else
                    addone k - Len(A$) - Len(drop$) + 1, 10
                End If
            End If
        End If
        pos2 = k - Len(A$) + 1
        If kind > 3 Then
            If kind = 8 Then
                lastcol = 13  '10
                If Left$(A$, 1) = "]" Then
                    GoTo checkthistoo
                End If
            End If
CheckThis:
            pos1 = 1
            drop$ = aheadstatus(A$, False, pos1)
marvel:
            If pos1 > 1 Then
marvel2:
                drop$ = Mid$(A$, 1, pos1 - 1)
                ColorizeExpressions drop$, k - Len(A$) + Len(drop$), pos2, feedthis
                If feedthis = 11 Or feedthis = 14 Or feedthis = 10 Then
                    feedthis = 0
                End If
newjump:
                feedback = feedthis
                If Col((UBound(Col()) \ 2) * 2) = 7 Then
                    A$ = vbNullString
                Else
                    A$ = Mid$(A$, pos1)
                End If
checkthistoo:
                If RTrim$(A$) = vbNullString Then
                    If pos2 > 0 Then
                        If feedthis <> 15 Then
                            If Col((UBound(Col()) \ 2) * 2) <> lastcol And pos2 < pos1 Then addone pos2, lastcol
                        End If
                    Else
                        makebase 14
                    End If
                    mDoc.Paracolordata(thispara) = Col()
                    If feedthis = 15 Then GoTo use15
                    Exit Sub
                End If
                    If kind = 8 And Left$(A$, 1) = "]" Then '****************************fix this
                        pos2 = k - Len(A$) + 1
                        addone pos2, 13
                        A$ = Mid$(A$, 2)
                        pos2 = pos2 + 1 'k - Len(a$) + 1
                        If Left$(A$, 1) = "[" Then
                            A$ = Mid$(A$, 2)
                            pos2 = k - Len(A$) + 1
                            GoTo CheckThis
                        End If
                        onetime = False
                    ElseIf kind > 4 And onetime Then
                        If FastSymbol(A$, ",") Then
                            addone pos2, 14
                            pos2 = k - Len(A$) + 1
                            pos1 = 1
                            GoTo CheckThis
                        End If
                    End If
                    If FastSymbol(A$, "->", , 2) Then
                        addone pos2, 14
                        pos2 = k - Len(A$) + 1
                        If FastSymbol(A$, "{") Then
                            pos2 = k - Len(A$) + 1
                            onetime = True
                            GoTo again
                        End If
                    ElseIf FastSymbol(A$, ",") Then
                        addone pos2, 14
                        pos2 = k - Len(A$) + 1
                        pos1 = 1
                    End If
                ElseIf kind > 4 And onetime Then
                    If FastSymbol(A$, ",") Then
                        addone pos2, 14
                        pos2 = k - Len(A$) + 1
                        pos1 = 1
                        GoTo CheckThis
                      End If
                End If
                If FastSymbol(A$, ")") Then
                    pos2 = k - Len(A$)
                    If lastcol = 0 Then
                    addone pos2, 14
                Else
                    addone pos2, lastcol
                End If
                pos2 = k - Len(A$) + 1
            End If
        End If
here3:
    If kind > 0 Then
        If MaybeIsSymbol(A$, ".") Then
            If FastSymbol(A$, ".(", , 2) Then
                addone pos2, 14
                A$ = vbNullString
                pos2 = k + 1
            ElseIf FastSymbol(A$, "..", , 2) Then
                lastcol = 10
                lastkind = 1
                addone pos2, lastcol
                pos2 = k - Len(A$) + 1
                GoTo here3
            Else
                pos1 = pos1 + 1
                lastcol = 14
                GoTo CheckThis
            End If
        End If
    End If
    If FastSymbol(A$, "<=", , 2) Then
        addone pos2, 15
        pos2 = k - Len(A$) + 1
    Else
again4:
        If MaybeIsSymbolNoSpace(A$, "~", True) Then
            addone pos2, 15
            pos2 = k - Len(A$) + 1
        ElseIf FastSymbol(A$, "=>", , 2) Then
            addone pos2, 14
            pos2 = k - Len(A$) + 1
        ElseIf FastSymbol(A$, "=") Then
            If onetime And lastcol = 11 Then
            Else
                addone pos2, 15
            End If
            pos2 = k - Len(A$) + 1
        ElseIf FastSymbol(A$, "//", , 2) Then
            GoTo here1
        ElseIf MaybeIsSymbol(A$, "+-*/", "=", True) Then
            addone pos2, 15
            pos2 = k - Len(A$) + 1
        ElseIf FastSymbol(A$, "/*-+=~^|<>") Then
            addone pos2, 14
            pos2 = k - Len(A$) + 1
            If FastSymbol(A$, "/*-+=~^&|<>") Then
                addone pos2, 14
                pos2 = k - Len(A$) + 1
            End If
        End If
    End If
again2:
    Do
again3:
        pos1 = 1
        drop$ = aheadstatus(A$, False, pos1)
        If pos1 > 1 Or drop$ = "S" Then
            If pos1 = 1 Then
                drop$ = A$
                pos1 = Len(A$) + 1
                counter = 1
            Else
                drop$ = Mid$(A$, 1, pos1 - 1)
                If drop$ = vbNullString Then drop$ = " "
            End If
            feedback = 0
            ColorizeExpressions drop$, k - Len(A$) + Len(drop$), pos2, feedback
            If feedback <> 15 Then
                If feedback = 0 And lastcol = 0 Then feedback = 14
                If drop$ = "(" Then
                    If feedback = 100 Then
                        lastcol = 14
                    Else
                        lastcol = feedback
                    End If
                Else 'If lastcol <> 11 And lastcol <> 10 Then
                    If feedback = 100 Then
                        lastcol = 14
                    Else
                        If feedback > 0 And kind > 0 Then
                            lastcol = feedback
                            feedback = 0
                            A$ = Mid$(A$, pos1)
                            ColorizeInnerStatements A$, k, pos2, feedback, True
                            If MaybeIsSymbol(A$, ")") Then
                                addone pos2, lastcol
                                FastSymbol A$, ")"
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                feedback = 0
            End If
            If feedback = 100 Then
                A$ = Mid$(A$, pos1)
                GoTo again112
            ElseIf feedback = 15 Then
use15:
                counter = 1
nextline:
                feedback = 0
                drop$ = mDoc.NextTextParagraph(thispara, feedback)
                If feedback = 0 Then
                    A$ = Mid$(A$, pos1)
                    mDoc.Paracolordata(thispara) = Col()
                    Exit Sub
                End If
                pos1 = 1
                countbrakcet drop$, pos1, counter
                mDoc.Paracolordata(thispara) = Col()
                A$ = drop$
                
                k = Len(A$)
                
                If mDoc.startfrom(thispara) = 0 Then
                    mDoc.startfrom(thispara) = thispara
                End If
                mDoc.startfrom(feedback) = mDoc.startfrom(thispara)
                thispara = feedback
                lastdata = A$
                pos2 = 1
                If Len(A$) > 0 Then
                    ReDim Col(3)
                    Col(0) = 15
                    Col(1) = 1
                    Col(2) = 15
                    Col(3) = k + 1
                Else
                    ReDim Col(1)
                    Col(0) = 15
                    Col(1) = 1
                End If
                If counter = 0 Then
                    If UBound(Col()) > 1 Then
                        Col(3) = pos1
                        If Col(3) = Col(1) Then
                            pos1 = pos1 + 1
                            Col(3) = pos1
                            Col(2) = 14
                        End If
                    Else
                        Col(1) = pos1
                    End If
                    If Left(A$, 1) <> "}" Then
                        A$ = Mid$(A$, pos1 + 1)
                        pos2 = k - Len(A$) + 1
                    End If
                    feedback = 0 'lastcol
                    feedthis = 0
                    If MaybeIsSymbol(A$, "<>=~") Then
                        onetime = False
                        pos2 = k - Len(A$) + 1
                        addone pos2, 14
                        If FastSymbol(A$, "<") Then
                            If FastSymbol(A$, ">") Then
                        ElseIf FastSymbol(A$, "=") Then
                        End If
                        'pos2 = k - Len(a$) + 1
                        'addone pos2, 14
                    ElseIf FastSymbol(A$, ">") Then
                        If FastSymbol(A$, "=") Then
                        End If
                    ElseIf FastSymbol(A$, "=") Then
                    ElseIf FastSymbol(A$, "~") Then
                    End If
                End If
            Else
                GoTo nextline
            End If
            pos1 = 1
            GoTo CheckThis
        ElseIf feedback = 0 And pos1 > 1 And kind > 0 Then
            GoTo newjump
        End If
        If pos1 > 0 Then
            If Mid$(A$, pos1 - 1, 1) = " " Then lastkind = 0
            A$ = Mid$(A$, pos1)
        End If
        If MaybeIsSymbol(A$, ":{") Then
            onetime = True
            GoTo again
        End If
        If InStr(drop$, " ") = 0 Or kind >= 0 Then
            GoTo CheckThis
        End If
    End If
again112:
    If MaybeIsSymbol(A$, ",;", , True) Then
        addone pos2, 14
        pos2 = k - Len(A$) + 1
        lastkind = 0
        onetime = False
    ElseIf FastSymbol(A$, "}") Then
        lastkind = 0
        addone pos2, 14
        pos2 = k - Len(A$) + 1
        If MaybeIsSymbol(A$, ".;+(<>=~") Then
            onetime = False
            If FastSymbol(A$, "<") Then
                If FastSymbol(A$, ">") Then
                ElseIf FastSymbol(A$, "=") Then
                End If
            ElseIf FastSymbol(A$, ">") Then
                If FastSymbol(A$, "=") Then
                End If
            ElseIf FastSymbol(A$, "=") Then
            ElseIf FastSymbol(A$, "~") Then
            End If
            GoTo CheckThis
        End If
        GoTo again112
        ElseIf MaybeIsSymbol(A$, "'") Then
            lastkind = 0
            pos2 = k - Len(A$) + 1
            FastSymbol A$, "'"
            If pos2 > 0 Then
                addone pos2, 7
            Else
                addone 1, 7
            End If
            A$ = vbNullString
            pos2 = k
            If pos2 > 0 Then
                addone pos2 + 1, 14
            End If
            mDoc.Paracolordata(thispara) = Col()
            Exit Sub
        Else
            onetime = True
            mDoc.Paracolordata(thispara) = Col()
            Exit Do
        End If
    Loop
    GoTo again
ElseIf FastSymbol(A$, "//", , 2) Then
    GoTo here1
ElseIf MaybeIsSymbol(A$, "\", , True) Then
    If lastkind = 1 Then
        pos2 = k - Len(A$) + 1
        If AscW(A$ + ".") <> 32 Then
            If MaybeIsSymbol(A$, ".", ".", True) Then
            End If
            If MaybeIsSymbolNoSpace(A$, "\") Then GoTo again
            pos2 = k - Len(A$) + 1
            lastkind = 0
            If MaybeIsSymbol(A$, "\", , True) Then GoTo here1
            GoTo again
        End If
    Else
here1:
        If pos2 > 0 Then
            addone pos2, 7
        Else
            addone 1, 7
        End If
        pos2 = k - Len(A$) + 1
        mDoc.Paracolordata(thispara) = Col()
        Exit Sub
    End If
ElseIf MaybeIsSymbol(A$, "'") Then
    If pos2 > 1 Then
        addone pos2, 7
    Else
        addone 1, 7
    End If
    pos2 = k - Len(A$) + 1
    mDoc.Paracolordata(thispara) = Col()
    Exit Sub
ElseIf Not onetime Then
    If FastSymbol(A$, ",") Then
        mDoc.Paracolordata(thispara) = Col()
        Exit Sub  ' stop
    ElseIf FastSymbol(A$, ":") Then
        onetime = True
    ElseIf FastSymbol(A$, ")") Then
        addone pos2, lastcol
        pos2 = k - Len(A$) + 1
        GoTo again4
    ElseIf FastSymbol(A$, ";") Then
        addone pos2, 14
        pos2 = k - Len(A$) + 1
        GoTo again2
    ElseIf FastSymbol(A$, "}") Then
        addone pos2, 14
        addone k - Len(A$), 11
        pos2 = k - Len(A$) + 1
        GoTo again3
    ElseIf FastSymbol(A$, "{") Then
        onetime = True
    ElseIf FastSymbol(A$, "?") Then
        addone pos2, 14
        addone k - Len(A$), 11
        pos2 = k - Len(A$) + 1
        GoTo again2
    Else
        addone pos2, lastcol
        mDoc.Paracolordata(thispara) = Col()
        Exit Sub
    End If
    addone pos2, 14
    pos2 = k - Len(A$) + 1
    GoTo again
ElseIf IsSymbol(A$, "?") Then
    addone 1, 11
    pos2 = k - Len(A$) + 1
    GoTo again2
ElseIf IsSymbol(A$, "=") Then
    If pos2 > 0 Then
        addone pos2, 11
    Else
        addone 1, 11
    End If
    pos2 = k - Len(A$) + 1
    onetime = False
    GoTo again2
ElseIf FastSymbol(A$, "{") Then
    onetime = True
    If pos2 > 0 Then
        If Col((UBound(Col()) \ 2) * 2) <> 14 Then addone pos2, 14
        pos2 = k - Len(A$) + 1
        If FastSymbol(A$, "&") Then
            addone pos2, 10
            pos2 = k - Len(A$) + 1
        End If
    Else
        addone 1, 14
        pos2 = k - Len(A$) + 1
    End If
    
    GoTo again
ElseIf FastSymbol(A$, "}") Then
    If pos2 > 0 Then
        addone pos2, 14
    Else
        addone 1, 14
        If MaybeIsSymbol(A$, "(") Then
            GoTo CheckThis
        ElseIf MaybeIsSymbol(A$, ",;") Then
            A$ = Mid$(A$, 2)
            GoTo CheckThis
        End If
    End If
    pos2 = k - Len(A$) + 1
    GoTo again
Else
    If IsNumberLabel(A$, drop$) Then
        If pos2 > 0 Then
            addone k - Len(A$) - Len(drop$) + 1, 14
        Else
            addone k - Len(A$) - Len(drop$) + 1, 15
        End If
        pos2 = k - Len(A$) + 1
        GoTo again
    End If
End If
If pos2 > 0 Then addone pos2, 14
If Col(1) = 0 Then
    makebase 0
End If
End Sub

Private Sub ColorizeExpressions(ByVal A$, k As Long, pos2 As Long, feedthis As Long)
Dim drop$, pos1 As Long, kind As Long, localfeedthis As Long, thispara1 As Long, lastpar As Long
Dim lastcol As Long, bracket As Long
Dim n$
again:
If A$ = vbNullString Then
If feedthis = 0 Then
If lastpar > 0 Then
feedthis = lastpar
End If
End If
Exit Sub
End If
drop$ = "@(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then
    If IsSymbol(A$, "@") Then
        If pos2 > 0 Then
            addone pos2, 14
        Else
            addone 1, 14
        End If
        pos2 = k - Len(A$) + 1
    
    End If
getlabelonly:
    n$ = A$
    kind = IsLabelOnly(n$, drop$)
    pos2 = k - Len(A$) + 1
    If kind = 0 And n$ <> A$ Then
    feedthis = 100
    A$ = n$: Exit Sub
    End If
    A$ = n$
    If kind = 100 Then
    If pos2 > 0 Then
            addone pos2, 14
        Else
            addone 1, 14
        End If
         pos2 = k - Len(A$) + 1
         
    kind = 0
    GoTo again
    
    End If
End If
If kind = 0 Then drop$ = "$(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then
  If FastSymbol(A$, "->", , 2) Then
    pos2 = k - Len(A$) - 1
    addone pos2, 11
        A$ = NLtrim$(A$)
        pos2 = k - Len(A$) + 1
        If MaybeIsSymbolNoSpace(A$, "{") Then
        pos2 = k - Len(A$) + 1
        addone pos2, 14
        A$ = Mid$(A$, 2)
         localfeedthis = 0
           pos2 = k - Len(A$) + 1
           If pos2 <= k Then
                drop$ = block(A$)
                If Len(drop$) > 0 Then
                
                    ColorizeInnerStatements drop$, k - Len(A$), pos2, localfeedthis

                    If Len(A$) > 0 Then ColorizeInnerStatements A$, k, pos2, localfeedthis
          
                    drop$ = vbNullString
                Else
                ColorizeInnerStatements A$, k, pos2, localfeedthis
               End If

           End If
            kind = 0
            drop$ = vbNullString
            If pos2 > k Then Exit Sub
        Else
           GoTo checklambda
            
        End If
    
    End If
End If
If kind = 0 Then drop$ = vbNullString
If kind > 0 Then
    If ismine(drop$) Then
    lastcol = 11
    If kind > 4 Then lastpar = 11
    If pos2 > 0 Then
       addone pos2, 11
    Else
       addone k - Len(A$) - Len(drop$) + 1, 11
       End If
        '' for the end
   Else
   If kind > 4 And kind < 9 Then lastpar = 10 - (kind = 8) * 3
   If kind = 10 Then
        lastcol = 13
   ElseIf kind = 20 Then
        lastcol = 12
   Else
        lastcol = 10 - (kind = 8) * 3
   End If
   
   If pos2 > 0 Then
        If kind = 8 Then
        addone pos2, 10
        addone pos2 + Len(drop$) - 1, lastcol
        
        Else
       addone pos2, lastcol
       End If
   Else
        If kind = 8 Then
        addone k - Len(A$) - Len(drop$), 10
        addone k - Len(A$) - Len(drop$) + 1, 11
        Else
        addone k - Len(A$) - Len(drop$) + 1, lastcol
        End If
       End If
  End If
  pos2 = k - Len(A$) + 1
 If kind > 3 Then
checklambda:
  Do
  pos1 = 1
  drop$ = aheadstatus(A$, False, pos1)
  If pos1 > 1 Then
  ColorizeExpressions Mid$(A$, 1, pos1 - 1), k - Len(A$) + pos1 - 1, pos2, feedthis
  If feedthis = 14 Or feedthis = 11 Then
  If lastcol <> 11 And lastcol <> 10 And lastcol <> 13 Then lastcol = feedthis
  feedthis = 0
  End If
  If feedthis > 0 And pos1 > Len(A$) Then Exit Sub
  feedthis = 0
   A$ = Mid$(A$, pos1)
   End If
   If FastSymbol(A$, ",") Then
     addone pos2, 14
    pos2 = k - Len(A$) + 1
   Else
   Exit Do
   End If
   Loop
If MaybeIsSymbol(A$, "(") Then
addone pos2, 14
lastpar = 14
FastSymbol A$, "("

pos2 = k - Len(A$) + 1
  pos1 = 1
  drop$ = aheadstatus(A$, False, pos1)
    thispara1 = thispara
  If pos1 > 1 Then

  ColorizeExpressions Mid$(A$, 1, pos1 - 1), k - Len(A$) + pos1 - 1, pos2, feedthis
    A$ = Mid$(A$, pos1)
    End If
    If thispara1 <> thispara Then
    Exit Sub
    End If
    If kind = 8 Then
    addone pos2, 11
    FastSymbol A$, "]"
    
    Else
    addone pos2, 14
    FastSymbol A$, ")"
    End If
    pos2 = k - Len(A$) + 1
ElseIf IsSymbol(A$, ",") Then
addone pos2, 14
pos2 = k - Len(A$) + 1
    GoTo again
ElseIf MaybeIsSymbol(A$, "]") Then

    pos2 = k - Len(A$) + 1
    addone pos2, lastcol
    If FastSymbol(A$, "][", , 2) Then
    GoTo checklambda
    Else
        pos2 = k - Len(A$) + 2
        addone pos2, 14
        FastSymbol A$, "]"
        pos2 = k - Len(A$) + 1
    End If

ElseIf MaybeIsSymbol(A$, ")") Then

    pos2 = k - Len(A$) + 1
    If kind >= 10 Then
    addone pos2, 14
    
    ElseIf kind > 0 Then
  addone pos2, lastcol
  Else
  addone pos2, 14
  End If
  pos2 = k - Len(A$) + 2
  addone pos2, 14
  FastSymbol A$, ")"
   
    pos2 = k - Len(A$) + 1
  End If
  End If

GoTo again
ElseIf MaybeIsSymbol(A$, "(", , True) Then
    addone pos2, 14
    pos2 = k - Len(A$) + 1
   GoTo again
ElseIf MaybeIsSymbol(A$, "{") Then
addone pos2, 15
pos1 = 1
bracket = 0
'' fix this
blockStringAhead2 A$, pos1, bracket
If bracket = 0 Then feedthis = 0 Else feedthis = 15
A$ = Mid$(A$, pos1 + 1)
drop$ = vbNullString
pos2 = k - Len(A$) + 1
If feedthis = 15 Then Exit Sub
    GoTo again

ElseIf MaybeIsSymbol(A$, Chr(34)) Then

If Not ISSTRINGA(A$, drop$) Then
A$ = vbNullString
addone pos2, 10
Else
addone pos2, 15
End If
drop$ = vbNullString
pos2 = k - Len(A$) + 1
    GoTo again
Else
If Col((UBound(Col()) \ 2) * 2) <> 14 Then addone pos2, 14
A$ = Mid$(A$, 2)
If Len(A$) > 0 Then GoTo again
End If

If pos2 > 0 Then addone pos2, 14


End Sub


Private Sub mDoc_ColorDataEvent(thispara1 As Long, Data As String, feedback As Long)

thispara = thispara1
Col() = mDoc.Paracolordata(thispara)
If NoColor Then
If UBound(Col()) > 1 Then
ReDim Col(1)
mDoc.Paracolordata(thispara) = Col()
End If
Exit Sub
End If
If feedback = 0 Then
ReDim Col(1)
End If
lastdata = Data
ColorizeInnerStatements (Data), Len(Data), (0), feedback
If feedback = -1 Then
If thispara <> 0 Then
mDoc.Paracolordata(thispara) = Col()
thispara1 = thispara
End If
Else
mDoc.Paracolordata(thispara) = Col()
thispara1 = thispara
End If
End Sub




Function IsLabelDot(where$, A$, R$) As Long 'ok
' for left side...no &

Dim rr&, one As Boolean, c$, firstdot$, gr As Boolean
R$ = vbNullString
If A$ = vbNullString Then IsLabelDot = 0: Exit Function

A$ = NLtrim$(A$)
    Do While Len(A$) > 0
    c$ = Left$(A$, 1)
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 64  '"@"
        
            If Mid$(A$, 2, 1) <> "(" And R$ <> "" Then

                R$ = R$ + "@"
                A$ = Mid$(A$, 2)
            Else
                IsLabelDot = 0: A$ = firstdot$ + A$: Exit Function
            End If
        Case 63 '"?"
        If R$ = vbNullString And firstdot$ = vbNullString Then
        R$ = "?"
        A$ = Mid$(A$, 2)
        IsLabelDot = 1
        Exit Function
    
        ElseIf firstdot$ = vbNullString Then
        IsLabelDot = 1
        Exit Function
        Else
        IsLabelDot = 0
        Exit Function
        End If
        Case 46 '"."
            If one Then
            Exit Do
            Exit Do
            ElseIf R$ <> "" And Len(A$) > 1 Then
            If Mid$(A$, 2, 2) = ". " Or Mid$(A$, 2, 1) = " " Then Exit Do
            R$ = R$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            rr& = 1
            Else
            firstdot$ = firstdot$ + "."
            A$ = Mid$(A$, 2)
            End If
       Case 92, 94, 123 To 126, 160 '"\","^", "{" To "~"
        Exit Do

        Case 48 To 57, 95 '"0" To "9", "_"
           If one Then
            If firstdot$ <> "" Then A$ = firstdot$ + A$
            Exit Do
            ElseIf R$ <> "" Then
            R$ = R$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit Do
            End If
        Case Is < 0, Is > 64 ' >=A and negative
            If one Then
            Exit Do
            Else
                If Left$(R$, 1) <> "[" And c$ = "]" Then
                    If InStr(R$, ".[") > 0 Then
                        R$ = R$ + "]"
                        A$ = Mid$(A$, 2)
                    Else
                        Exit Do
                    End If
                ElseIf Left$(R$, 1) = "[" And c$ = "]" And Mid$(A$, 2, 1) = "." Then
                    R$ = R$ & "]"
                    A$ = Mid$(A$, 2)
                    Exit Do
                ElseIf Right$(R$, 2) = "]$" And c$ = "]" Then
                    Exit Do
                ElseIf Right$(R$, 1) = "]" And c$ = "]" Then
                    Exit Do
                ElseIf c$ = "[" And Len(R$) > 0 Then
                    R$ = R$ & "["
                    A$ = Mid$(A$, 2)
                    If Right$(R$, 2) <> ".[" Then
                        If Left$(R$, 1) <> "[" Then
                            rr& = 8
                            Exit Do
                        End If
                    End If
                Else
                    R$ = R$ & Left$(A$, 1)
                    A$ = Mid$(A$, 2)
                End If
                rr& = 1 'is an identifier or floating point variable
            End If
        Case 36 ' "$"
        
            If one Then
            If InStr(R$, "$.") Then
            rr& = 3          ' is string variable with string pointer
            R$ = R$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            End If
            Exit Do
            End If
            If R$ <> "" Then
            one = True
            If Mid$(A$, 2, 1) <> "." Then
            rr& = 3
            R$ = R$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            Else
            rr& = 1       ' is string variable
            R$ = R$ & Left$(A$, 2)
            A$ = Mid$(A$, 3)
            
            End If
            
            Else
            Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If R$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            R$ = R$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            Else
            Exit Do
            End If
        Case 40 ' "("
            If R$ <> "" Then
                            If Mid$(A$, 2, 2) = ")@" Then
                                    R$ = R$ & "()."
                                  
                                 A$ = Mid$(A$, 4)
                               Else
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit Do
                                       End Select
                                       R$ = R$ & Left$(A$, 1)
                                       A$ = Mid$(A$, 2)
                                   Exit Do
                            
                          End If
               Else
                        Exit Do
            
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
                gr = True
                If Left$(R$, 1) <> "[" And c$ = "]" Then
                    If InStr(R$, ".[") > 0 Then
                        R$ = R$ + "]"
                        A$ = Mid$(A$, 2)
                    Else
                        Exit Do
                    End If
                ElseIf Left$(R$, 1) = "[" And c$ = "]" And Mid$(A$, 2, 1) = "." Then
                    R$ = R$ & "]"
                    A$ = Mid$(A$, 2)
                    Exit Do
                ElseIf Right$(R$, 2) = "]$" And c$ = "]" Then
                    Exit Do
                ElseIf Right$(R$, 1) = "]" And c$ = "]" Then
                    Exit Do
                ElseIf c$ = "[" And Len(R$) > 0 Then
                    R$ = R$ & "["
                    A$ = Mid$(A$, 2)
                    If Right$(R$, 2) <> ".[" Then
                        If Left$(R$, 1) <> "[" Then
                            rr& = 8
                            Exit Do
                        End If
                    End If

                Else
                    R$ = R$ & Left$(A$, 1)
                    A$ = Mid$(A$, 2)
                End If
                rr& = 1 'is an identifier or floating point variable
            End If
        End If

    Loop
       R$ = firstdot$ + myUcase(R$, gr)
    IsLabelDot = rr&
   'a$ = LTrim$(a$)

End Function
Function IsLabelOnly(A$, R$) As Long 'ok
Dim rr&, one As Boolean, c$, dot&, Extra&, num As Boolean
R$ = vbNullString
If A$ = vbNullString Then IsLabelOnly = 0: Exit Function
A$ = NLtrim$(A$)
Do While Len(A$) > 0
    c$ = Left$(A$, 1) 'ANYCHAR HERE
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 35 ' "#"
            If R$ <> "" Then
                Exit Do
            Else
                R$ = "#"
                A$ = Mid$(A$, 2)
                rr& = 10
           End If
        Case 64  '"@"
           If R$ = vbNullString Then
              A$ = Mid$(A$, 2)
              ElseIf Mid$(A$, 2, 1) <> "(" And R$ <> "" Then
            R$ = R$ & "."
              A$ = Mid$(A$, 2)
              Else
                 IsLabelOnly = 0: Exit Function
            End If
        Case 46 '"."
            If one Then
                Exit Do
            ElseIf R$ <> "" Then
                R$ = R$ & Left$(A$, 1)
                A$ = Mid$(A$, 2)
            ElseIf Not Mid$(A$, 2, 1) Like "[0-9]" Then
                       If R$ <> "" Then
                            R$ = R$ & Left$(A$, 1)
                            rr& = 1
                        Else
                            dot& = dot& + 1
                        End If
                        If Mid$(A$, 2, 1) Like "[eEÂ≈]" Then
                            R$ = R$ & Mid$(A$, 2, 1)
                            A$ = Mid$(A$, 3)
                        Else
                             A$ = Mid$(A$, 2)
                        End If
            Else
                If R$ = vbNullString And dot& > 0 Then
                    R$ = String$(dot& + 1, ".")
                    A$ = Mid$(A$, 2)
                    IsLabelOnly = 1
                Else
                If MaybeIsSymbolNoSpace(Mid$(A$, 2, 1), "1234567890") Then
                    FastSymbol A$, "."
                    While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                    Wend
                    If MaybeIsSymbolNoSpace(Mid$(A$, 3, 1), "Ee≈Â") Then
                        A$ = Mid$(A$, 4)
                        MaybeIsSymbolNoSpace A$, "+-", True
                        MaybeIsSymbolNoSpace A$, "123456789", True
                        While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                        Wend
                        If MaybeIsSymbolNoSpace(A$, ".") Then
                            A$ = vbNullString
                        End If
                    End If
                    If MaybeIsSymbolNoSpace(A$, ".") Then
                        IsLabelOnly = 1
                    Else
                        IsLabelOnly = 100
                    End If
                    Exit Function
                End If
                IsLabelOnly = 0
            End If
            Exit Function
        End If
    Case 92, 94, 123 To 126, 160 '"\","^", "{" To "~"
        Exit Do
    Case 48 To 57, 95 '"0" To "9", "_"
        If one Then
            Exit Do
        ElseIf R$ <> "" Then
            If c$ = "_" Then
            R$ = " " & R$
            Else
            R$ = R$ & Left$(A$, 1)
            End If
            A$ = Mid$(A$, 2)
            If Extra& > 0 Then
                If Left$(A$, 1) = "%" Then
                    A$ = Mid$(A$, 2)
                    If Len(LTrim$(R$)) > 6 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                    Exit Do
                ElseIf Left$(A$, 2) = "&&" Then
                    A$ = Mid$(A$, 3)
                    If Len(LTrim$(R$)) > 18 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                    Exit Do
                ElseIf Left$(A$, 1) = "&" Then
                    A$ = Mid$(A$, 2)
                    If Len(LTrim$(R$)) > 10 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                    Exit Do
                ElseIf Left$(A$, 1) = "@" Then
                    A$ = Mid$(A$, 2)
                    If Len(LTrim$(R$)) > 18 Then rr& = 0 Else rr& = 1
                    Exit Do
                ElseIf LCase$(Left$(A$, 2)) = "ub" Then
                    A$ = Mid$(A$, 3)
                    If Len(LTrim$(R$)) > 4 Then rr& = 0 Else rr& = 1
                    Exit Do
                End If
            End If
            If Len(LTrim$(R$)) > 18 And Extra Then rr& = 0 Else rr& = 1
        ElseIf Mid$(A$ + "__", 2, 2) Like "[◊˜Xx][0-9A-Fa-f]" Then
evryka:
            R$ = Left$(A$, 3)
            A$ = Mid$(A$, 4)
            Extra& = 9
            If Left$(A$, 1) = "%" Then
                A$ = Mid$(A$, 2)
                rr& = 1: Extra& = Extra& + 10
                Exit Do
            ElseIf Left$(A$, 2) = "&&" Then
                Extra& = Extra& + 10
                A$ = Mid$(A$, 3)
                Exit Do
            ElseIf Left$(A$, 1) = "&" Then
                rr& = 1: Extra& = Extra& + 10
                A$ = Mid$(A$, 2)
                Exit Do
            ElseIf Left$(A$, 1) = "@" Then
                A$ = Mid$(A$, 2)
                rr& = 1
                Exit Do
            ElseIf LCase$(Left$(A$, 2)) = "ub" Then
                A$ = Mid$(A$, 3)
                rr& = 1
                Exit Do
            Else
            rr& = 1
            End If
        ElseIf Mid$(A$, 2, 1) = "%" Then
            A$ = Mid$(A$, 3)
            Extra& = 0
            rr& = 100
            Exit Do
        ElseIf Mid$(A$, 2, 1) = "&" Then
            A$ = Mid$(A$, 3)
            Extra& = 0
            MaybeIsSymbolNoSpace A$, "&", True

            rr& = 100: Exit Do
            
            Exit Do
        Else
            If dot& > 0 Then
                A$ = "." + A$: dot& = 0
            ElseIf Mid$(A$, 2, 1) = "." Then
                A$ = Mid$(A$, 3)
checkexp:
                While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                Wend
                If MaybeIsSymbolNoSpace(A$, "Ee≈Â", True) Then
                    MaybeIsSymbolNoSpace A$, "+-", True
                    If MaybeIsSymbolNoSpace(A$, "+-", True) Then
                        IsLabelOnly = 1
                        Exit Function
                    End If
                    MaybeIsSymbolNoSpace A$, "123456789", True
                    While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                    Wend
                    If MaybeIsSymbolNoSpace(A$, ".") Then
                      IsLabelOnly = 1
                      A$ = vbNullString
                      Exit Function
                    End If
                Else
                    While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                    Wend
                    If MaybeIsSymbolNoSpace(A$, "uU", True) Then
                    If MaybeIsSymbolNoSpace(A$, "dD", True) Then
                    End If
                End If
                End If
                If MaybeIsSymbolNoSpace(A$, ".") Then
                    IsLabelOnly = 1
                Else
                    IsLabelOnly = 100
                End If
                Exit Function
            ElseIf MaybeIsSymbolNoSpace(A$, "1234567890", True) Then
                While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                Wend
                If MaybeIsSymbolNoSpace(A$, ".", True) Then
                    GoTo checkexp
                ElseIf MaybeIsSymbolNoSpace(A$, "Ee≈Â") Then
                    GoTo checkexp
                End If
                If MaybeIsSymbolNoSpace(A$, "&", True) Then
                MaybeIsSymbolNoSpace A$, "&", True
                ElseIf MaybeIsSymbolNoSpace(A$, "%", True) Then
                ElseIf MaybeIsSymbolNoSpace(A$, "uU", True) Then
                    If MaybeIsSymbolNoSpace(A$, "bB", True) Then
                    ElseIf MaybeIsSymbolNoSpace(A$, "dD", True) Then
                    End If
                End If
                IsLabelOnly = 100
                Exit Function
            End If
            Exit Do
        End If
    Case Is < 0, Is > 64 ' >=A and negative
        If one Then
            Exit Do
        Else
            If Extra& > 0 Then
                If Not c$ Like "[a-fA-F]" Then Exit Do
                R$ = R$ & Left$(A$, 1)
                A$ = Mid$(A$, 2)
                If Left$(A$, 1) = "%" Then
                    If Len(LTrim$(R$)) > 6 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                    A$ = Mid$(A$, 2)
                    Exit Do
                ElseIf Left$(A$, 2) = "&&" Then
                   A$ = Mid$(A$, 3)
                   If Len(LTrim$(R$)) > 18 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                   Exit Do
                ElseIf Left$(A$, 1) = "&" Then
                   A$ = Mid$(A$, 2)
                   If Len(LTrim$(R$)) > 10 Then rr& = 0 Else rr& = 1: Extra& = Extra& + 10
                   Exit Do
                ElseIf Left$(A$, 1) = "@" Then
                    A$ = Mid$(A$, 2)
                    If Len(LTrim$(R$)) > 18 Then rr& = 0 Else rr& = 1
                    Exit Do
                End If
                If Len(LTrim$(R$)) > 18 Then rr& = 0
            Else
                If Left$(R$, 1) <> "[" And c$ = "]" Then
                    If InStr(R$, ".[") > 0 Then
                        R$ = R$ + "]"
                        A$ = Mid$(A$, 2)
                    Else
                        Exit Do
                    End If
                ElseIf Left$(R$, 1) = "[" And c$ = "]" And Mid$(A$, 2, 1) = "." Then
                    R$ = R$ & "]"
                    A$ = Mid$(A$, 2)
                    Exit Do
                ElseIf Right$(R$, 2) = "]$" And c$ = "]" Then
                    Exit Do
                ElseIf Right$(R$, 1) = "]" And c$ = "]" Then
                    Exit Do
                ElseIf c$ = "[" And Len(R$) > 0 Then
                    R$ = R$ & "["
                    A$ = Mid$(A$, 2)
                    If Right$(R$, 2) <> ".[" Then
                        If Left$(R$, 1) <> "[" Then
                        If dot& > 0 Then
                            R$ = String$(dot&, ".") + R$
                        End If
                            rr& = 8
                            Exit Do
                        End If
                    End If

                Else
                    R$ = R$ & Left$(A$, 1)
                    A$ = Mid$(A$, 2)
                End If
                rr& = 1 'is an identifier or floating point variable
            End If
        End If
    Case 36 ' "$"
       If one Then Exit Do
            If R$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            R$ = R$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            Else
            Exit Do
            End If
     Case 37 ' "%"
      If one Then
            If InStr(R$, "$.") Then
            rr& = 3          ' is string variable with string pointer
            R$ = R$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            End If
            Exit Do
            End If
            If R$ <> "" Then
            one = True
            If Mid$(A$, 2, 1) <> "." Then
            rr& = 3
            R$ = R$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            Else
            rr& = 1       ' is string variable
            R$ = R$ & Left$(A$, 2)
            A$ = Mid$(A$, 3)
            
            End If
            
            Else
                Exit Do
            End If
        Case 38 ' "&"
        If one Then Exit Do
        If LCase$(Mid$(A$, 2, 1)) = "h" Then
        If Mid$(A$ + "__", 3, 1) Like "[0-9A-Fa-f]" Then
            GoTo evryka
        End If
        End If
        If R$ = vbNullString Then rr& = 10: A$ = Mid$(A$, 2): R$ = "&"
        Exit Do
        Case 40 ' "("
            If R$ <> "" Then
                If Mid$(A$, 2, 2) = ")@" Then
                    R$ = R$ & "()."
                    A$ = Mid$(A$, 4)
                Else
                Select Case rr&
                Case 1
                rr& = 5 ' float array or function
                Case 3
                rr& = 6 'string array or function
                Case 4
                rr& = 7 ' long array
                Case Else
                Exit Do
                End Select
                R$ = R$ & Left$(A$, 1)
                A$ = Mid$(A$, 2)
               Exit Do
            
            End If
           Else
           Exit Do
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
                If Left$(R$, 1) <> "[" And c$ = "]" Then
                    If InStr(R$, ".[") > 0 Then
                        R$ = R$ + "]"
                        A$ = Mid$(A$, 2)
                    Else
                        Exit Do
                    End If
                ElseIf Left$(R$, 1) = "[" And c$ = "]" And Mid$(A$, 2, 1) = "." Then
                    R$ = R$ & "]"
                    A$ = Mid$(A$, 2)
                    Exit Do
                ElseIf Right$(R$, 2) = "]$" And c$ = "]" Then
                    Exit Do
                ElseIf Right$(R$, 1) = "]" And c$ = "]" Then
                    Exit Do
                ElseIf c$ = "[" And Len(R$) > 0 Then
                    R$ = R$ & "["
                    A$ = Mid$(A$, 2)
                    If Right$(R$, 2) <> ".[" Then
                        If Left$(R$, 1) <> "[" Then
                        If dot& > 0 Then
                            R$ = String$(dot&, ".") + R$
                        End If
                            rr& = 8
                            Exit Do
                        End If
                    End If

                Else
                    R$ = R$ & Left$(A$, 1)
                    A$ = Mid$(A$, 2)
                End If
                rr& = 1 'is an identifier or floating point variable
            End If
        End If
    Loop
    If rr& > 0 Then
        If Left$(R$, 1) = "#" Then
          If Len(R$) > 1 Then
            If Mid$(R$, 2) Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                rr& = 10
            Else
                A$ = R$ + A$
                rr& = 0
            End If
         Else
        A$ = R$ + A$
        rr& = 0
        End If
        End If
    End If
    
    If R$ = "&" Then
        IsLabelOnly = 10
    ElseIf rr& > 3 And rr& < 9 Then
        IsLabelOnly = rr& + Extra&
    ElseIf CheckFree(Left$(A$, 1)) Then
        IsLabelOnly = rr& + Extra&
    ElseIf rr& > 0 Then
        If Left$(R$, 1) = "&" Then
            A$ = Mid$(R$, 2) + A$: R$ = "&"
            IsLabelOnly = 10
        Else
            IsLabelOnly = 1
        End If
    End If
   
End Function


Function MaybeIsSymbol(A$, c$, Optional tail$ = vbNullString, Optional Remove As Boolean = False) As Boolean
Dim i As Long, ok As Boolean
If A$ = vbNullString Then Exit Function
i = 1
While Mid$(A$, i, 1) = " "
    i = i + 1
Wend
If i > Len(A$) Then Exit Function
If i > 1 Then A$ = Mid$(A$, i): i = 1

ok = InStr(c$, Mid$(A$, i, 1)) > 0
If tail$ <> "" And ok Then
ok = Mid$(A$, i + 1, Len(tail$)) Like tail$
End If
If Remove Then If ok Then A$ = Mid$(A$, i + 1 + Len(tail$))
MaybeIsSymbol = ok
End Function
Function MaybeIsSymbolNoSpace(A$, c$, Optional Remove As Boolean = False) As Boolean
Dim ok As Boolean

If A$ = vbNullString Then Exit Function
ok = InStr(c$, Left$(A$, 1))
If Remove And ok Then A$ = Mid$(A$, 2)
MaybeIsSymbolNoSpace = ok
End Function
Private Sub mDoc_Lastline()
If showparagraph Then showparagraphonce = True
End Sub

Private Sub mDoc1_BreakLine(Data As String, datanext As String)
If NoHeadwrap Or (NoTitle) Or UsedAsTextBox Then
    datanext = Data
    Data = vbNullString
Else
If Not glistN Is Nothing Then
 glistN.BreakLine Data, datanext
End If
End If
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub PrintItem1(mHdc As Long, c As String, R As RECT, Optional way As Long = DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
Dim cc As String
Dim tabw
If showparagraphonce Then
    cc = c & "∂"
    
    'TabbedTextOut mHdc, r.Left, r.top, StrPtr(cc), Len(cc), 1, tParam.iTabLength, 0
    

    DrawTextEx mHdc, StrPtr(cc), -1, R, way, VarPtr(tParam)
Else
'tabw = tParam.iTabLength * glistN.AveCharWith
   ' TabbedTextOut mHdc, r.Left, r.top, StrPtr(c), Len(c), 1, tabw, r.Left
    DrawTextEx mHdc, StrPtr(c), -1, R, way, VarPtr(tParam)
End If
End Sub
Private Sub PrintItem(mHdc As Long, c As String, R As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
  
    DrawTextEx mHdc, StrPtr(c), -1, R, way, VarPtr(tParam)
 
End Sub
Private Sub PrintItem2(mHdc As Long, c As String, R As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
    Dim that As Long
    that = way
    R.Left = R.Left + 4
    If glistN.VerticalCenterText Then that = that Or DT_VCENTER Or DT_EXPANDTABS Or DT_TABSTOP
    DrawTextEx mHdc, StrPtr(c), -1, R, that, VarPtr(tParam)
End Sub
    
Public Property Get EditDoc() As Boolean
EditDoc = mEditDoc
End Property

Public Property Let EditDoc(ByVal RHS As Boolean)
mEditDoc = RHS
If Not RHS Then glistN.MarkNext = 0
glistN.EditFlag = RHS
If glistN.enabled Then
glistN.ShowMe2
End If
End Property

Public Property Get ParaSelStart() As Long
If glistN.ListIndex < 0 Then ParaSelStart = 1: Exit Property
Dim what$, ThatPara As Long, pastchars
    pastchars = mDoc.TextParagraphLine(glistN.ListIndex + 1, what$, ThatPara)
Charpos = pastchars + glistN.SelStart
If Charpos = 0 Then
ParaSelStart = 1
Else
ParaSelStart = Charpos
End If
End Property
Public Sub WrapAll()
mDoc.WrapAgain
End Sub
Public Property Let ParaSelStart(ByVal RHS As Long)
' we have a new ParaSelStart
' we want a new listindex...and a new selstart

Charpos = RHS
Render
End Property

Public Property Get nowrap() As Boolean
nowrap = mNoWrap
End Property

Public Property Let nowrap(ByVal RHS As Boolean)
If Not (mNoWrap = RHS) Then
mNoWrap = RHS
glistN.NoScroll = Not mNoWrap
    If Not mDoc.IsEmpty Then
    WrapAll
   If HaveMarkedText Then RelocateMarkedText
    
    HorSrcoll 0
    Render
    End If
End If
End Property

Public Property Get enabled() As Boolean
enabled = glistN.enabled
End Property

Public Property Let enabled(ByVal RHS As Boolean)
If Not RHS Then
glistN.Noflashingcaret = True
Show

Else
If IsWine Then
glistN.Noflashingcaret = Not glistN.NoCaretShow And Not glistN.BackColor = rgb(255, 255, 355)
Else
glistN.Noflashingcaret = Not glistN.NoCaretShow
End If
End If
glistN.enabled = RHS
End Property
Sub SetFocus()
glistN.SetFocus
End Sub

Public Property Get ScrollKeyUpDown() As Boolean
ScrollKeyUpDown = glistN.NoFreeMoveUpDown
End Property

Public Property Let ScrollKeyUpDown(ByVal RHS As Boolean)
glistN.NoFreeMoveUpDown = RHS
End Property
Public Sub SetRowColumn(Row As Long, ByVal Column As Long, Optional noShow As Boolean = False)
' base 0 for column as input
Column = Column + 1
' base 1
If Column < 0 Then Column = 0
If Row = -1 Then
glistN.ListindexPrivateUse = glistN.listcount - 1
Else
If Row > glistN.listcount + 1 Then
Row = glistN.listcount + 1
End If
glistN.ListindexPrivateUse = Row - 1  ' no events
End If
Dim A As String, b As Long
mDoc.TextParagraphLine glistN.ListIndex + 1, A, b
mDoc.MarkParagraphID = b
If Len(glistN.list(glistN.ListIndex)) < Column Then
If Row = glistN.listcount + 1 Then
glistN.SelStartEventAlways = Len(glistN.list(glistN.ListIndex))
Else
glistN.SelStartEventAlways = Column  'Len(glistN.List(glistN.listindex)) + 1
End If
Else
glistN.SelStartEventAlways = Column
End If
If glistN.enabled Then
If Not noShow Then glistN.ShowMe2
End If
End Sub
Public Function TextFormatBreak(lineBreaker As String) As String
Text = mDoc.textFormat(lineBreaker)
End Function

Public Property Get Text() As String
Text = mDoc.textDoc
End Property
Public Property Let Text(RHS As String)
Dim oldenabled As Boolean
LastListIndex = 0
If Me.HaveMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End If
oldenabled = glistN.enabled
glistN.enabled = False
Set mDoc = New Document
mDoc.ColorEvent = Not NoColor
mDoc.textDoc = RHS
all = mDoc.DocLines
expandCharPos = 0
Charpos = 0
BlockStartPara = 0
BlockStartPos = 1
BlockStartLine = 0
BlockTopline = 0
BlockStartCursor = 0
BlockEndCursor = 0
BlockEndPara = 0
BlockEndPos = 1
BlockEndline = 0
BlockBottomline = 0
ClearRedo
ClearUndo
glistN.enabled = oldenabled
Dim what$, ThatPara As Long, pastchars
   pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
End Property
Public Property Let Text2(RHS As String)
Dim oldenabled As Boolean
LastListIndex = 0
If Me.HaveMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End If
oldenabled = glistN.enabled
glistN.enabled = False
Set mDoc = New Document
mDoc.ColorEvent = Not NoColor
mDoc.textDoc2 = RHS
all = mDoc.DocLines
expandCharPos = 0
Charpos = 0
BlockStartPara = 0
BlockStartPos = 1
BlockStartLine = 0
BlockTopline = 0
BlockStartCursor = 0
BlockEndCursor = 0
BlockEndPara = 0
BlockEndPos = 1
BlockEndline = 0
BlockBottomline = 0
ClearRedo
ClearUndo
glistN.enabled = oldenabled
Dim what$, ThatPara As Long, pastchars
   pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
End Property
Public Property Let TextColor(RHS As String)
Dim oldenabled As Boolean
If Me.HaveMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End If
oldenabled = glistN.enabled
glistN.enabled = False
Set mDoc = New Document
mDoc.textDocColor = RHS
all = mDoc.DocLines
ClearRedo
ClearUndo
glistN.enabled = oldenabled
Dim what$, ThatPara As Long, pastchars
   pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
End Property
Public Property Get TextParagraphs() As Long
TextParagraphs = mDoc.DocParagraphs
End Property
Public Property Get TextLines() As Long
TextLines = mDoc.DocLines
End Property
Public Property Let LineSpacingSilent(ByVal RHS As Long)
glistN.addpixels = RHS
''Render
End Property
Public Property Let LineSpacing(ByVal RHS As Long)
glistN.addpixels = RHS
Render
End Property

Public Property Get SelLength() As Long
If HaveMarkedText Then
If BlockStartLine = BlockEndline Then
SelLength = Abs(BlockStartCursor - BlockEndCursor)
Else
Dim A$(), oldmydoc As Document, i As Long, sum As Long
' maybe we didn't need to preserve old UndoDoc...I need to clarify this..
Set oldmydoc = UndoDoc
Set UndoDoc = New Document
FillBlock
sum = UndoDoc.SizeCRLF + 2
If UndoDoc.LfLead Then sum = sum - 1
If UndoDoc.CrTail Then sum = sum + 1
SelLength = sum - 2 ' no the last vbcrlf
End If
Set UndoDoc = oldmydoc
End If
End Property
Public Property Let SelLengthSilent(ByVal RHS As Long)
If RHS <= 0 Then
    ' a copy of glistN_MarkDestroyAny
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
'If glistN.Enabled Then glistN.ShowPan
Else
glistN.MarkNext = 1
   BlockStartCursor = glistN.SelStart
    BlockStartPara = mDoc.MarkParagraphID
    BlockStartPos = ParaSelStart  ' selstart isn't from start of the doc...
 
    BlockStartLine = glistN.ListIndex
        BlockEndPara = BlockStartPara
        BlockEndPos = BlockStartPos
        BlockEndline = BlockStartLine
    BlockEndCursor = BlockStartCursor + 1
    
mDoc.FindPos BlockStartLine + 1, BlockStartCursor, RHS, BlockEndline, BlockEndCursor, BlockEndPara, BlockEndPos

BlockEndline = BlockEndline - 1 ' base 0


    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
If RHS <> 0 Then
glistN.MarkNext = 2
'If glistN.Enabled Then glistN.ShowPan
End If
End If
End Property
Public Property Let SelLength(ByVal RHS As Long)
If RHS <= 0 Then
    ' a copy of glistN_MarkDestroyAny
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
If glistN.enabled Then glistN.ShowPan
Else
glistN.MarkNext = 1
   BlockStartCursor = glistN.SelStart
    BlockStartPara = mDoc.MarkParagraphID
    BlockStartPos = ParaSelStart  ' selstart isn't from start of the doc...
 
    BlockStartLine = glistN.ListIndex
        BlockEndPara = BlockStartPara
        BlockEndPos = BlockStartPos
        BlockEndline = BlockStartLine
    BlockEndCursor = BlockStartCursor + 1
    
mDoc.FindPos BlockStartLine + 1, BlockStartCursor, RHS, BlockEndline, BlockEndCursor, BlockEndPara, BlockEndPos

BlockEndline = BlockEndline - 1 ' base 0

    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
If RHS <> 0 Then
glistN.MarkNext = 2
If glistN.enabled Then glistN.ShowPan
End If
End If
End Property

Public Property Get Visible() As Boolean
Visible = glistN.Visible
End Property

Public Property Let Visible(ByVal RHS As Boolean)
''If Not rhs Then LastSelStart = SelStart
glistN.Visible = RHS

End Property
Public Property Get FontSize() As Single

  FontSize = glistN.FontSize
 
End Property

Public Property Let FontSize(New_FontSize As Single)
glistN.FontSize = New_FontSize
End Property
Public Property Get Font() As Font
 Set Font = glistN.Font
End Property

Public Property Set Font(New_Font As Font)
    Set glistN.Font = New_Font
End Property
Public Property Let BackColor(ByVal RHS As OLE_COLOR)

    
glistN.BackColor = RHS
  
    
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = glistN.BackColor
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = glistN.ForeColor
End Property

Public Property Let ForeColor(ByVal RHS As OLE_COLOR)

    glistN.ForeColor = RHS

End Property

Public Property Get top() As Long
top = glistN.TopTwips
End Property

Public Property Let top(ByVal RHS As Long)
glistN.TopTwips = RHS
End Property
Public Property Get Height() As Long
Height = glistN.HeightTwips
End Property

Public Property Let Height(ByVal RHS As Long)
glistN.HeightTwips = RHS
End Property
Public Sub move(ByVal mleft As Long, ByVal mtop As Long, mWidth As Long, mHeight As Long)
glistN.MoveTwips mleft, mtop, mWidth, mHeight
End Sub
Public Sub Dereference()
Set glistN = Nothing
Set mDoc = Nothing
Set mDoc1 = Nothing

End Sub

Public Property Get SelText() As String
If HaveMarkedText Then
FillBlock

If UndoDoc.LfLead Then
SelText = Mid$(UndoDoc.textDoc, 2)
ElseIf UndoDoc.CrTail Then
SelText = UndoDoc.textDoc + vbCr
Else  ''If UndoDoc.LfLead = UndoDoc.CrTail Then
SelText = UndoDoc.textDoc
End If
End If
End Property
Public Property Get SelHtmlText() As String
If HaveMarkedText Then
''FillBlock
Dim curcolors(0 To 15) As String, i As Long
For i = 0 To 15
curcolors(i) = Right$("000000" + Hex$(htmlM2000(i)), 6)

Next i

If UndoDoc.LfLead Then
SelHtmlText = Replace(Mid$(UndoDoc.textHtmlDoc(curcolors()), 2), vbCrLf, "<br>")
Else
SelHtmlText = Replace(UndoDoc.textHtmlDoc(curcolors()), vbCrLf, "<br>")
End If


End If
End Property
Public Property Let SelText(ByVal RHS As String)
Dim kk As Long, Pad$, prev As Boolean


kk = SelStart
SelStart = kk
prev = HaveMarkedText
glistN_MarkDelete False
glistN.MarkNext = 1

BlockStartPara = mDoc.MarkParagraphID
BlockStartPos = ParaSelStart
BlockStartCursor = glistN.SelStart '- 1
BlockEndCursor = BlockStartCursor + 1
BlockStartLine = glistN.ListIndex
BlockEndline = BlockStartLine
BlockEndPara = BlockStartPara
BlockEndPos = BlockStartPos
Dim k As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then
If prev Then GroupUndo
If NumberOnly Then

Pad$ = space$(Len(RHS) + Len(glistN.list(0)))
Mid$(Pad$, 1, Charpos - 1) = Left$(glistN.list(0), Charpos - 1)
Mid$(Pad$, Charpos, Len(RHS)) = RHS
If Len(glistN.list(0)) > 0 Then
Mid$(Pad$, Charpos + Len(RHS) - 1) = Mid$(glistN.list(0), Charpos)
End If
RHS = glistN.list(0)
glistN.list(0) = Pad$
If Not glistN.list(0) = Pad$ Then glistN.list(0) = RHS
Else
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS, , , Not NoColor
End If
Render
End If
BlockEndPara = mDoc.MarkParagraphID
BlockEndPos = ParaSelStart
BlockEndline = glistN.ListIndex
BlockEndCursor = glistN.SelStart
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If


k = kk
glistN.MarkNext = 2
SelStart = k

End Property
Public Property Let InsertTextNoRender(ByVal RHS As String)
Dim b As Boolean
ClearRedo
PushUndo
glistN_MarkDelete False
glistN.MarkNext = 0
undomode = 5 ' like 1
RemoveUndo (RHS)
undomode = 4 ' like 0
b = enabled
enabled = False
Dim k As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then


mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
all = mDoc.DocLines

End If
enabled = b
End Property
Public Property Let InsertText2(ByVal RHS As String)
glistN_MarkDelete True
glistN.MarkNext = 0
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
mDoc.ResetColor mDoc.MarkParagraphID
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
End Property

Public Property Let InsertText(ByVal RHS As String)
glistN_MarkDelete False
glistN.MarkNext = 0
Dim k As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then
mDoc.ResetColor mDoc.MarkParagraphID
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
Render

End If
End Property
Public Property Get CharPosStart() As Long
Dim findit As Long
Dim A As Long, b As Long, c As Long, D As Long
findit = -1
mDoc.FindPos glistN.ListIndex + 1, glistN.SelStart, findit, A, b, c, D
CharPosStart = findit
End Property
Public Property Get SelStart() As Long
'
Dim findit As Long
Dim A As Long, b As Long, c As Long, D As Long
findit = -1
If glistN.MarkNext > 0 Then
If BlockTopline = BlockBottomline Then
If BlockStartCursor < BlockEndCursor Then
mDoc.FindPos BlockTopline + 1, BlockStartCursor, findit, A, b, c, D
Else
mDoc.FindPos BlockTopline + 1, BlockEndCursor, findit, A, b, c, D
End If
ElseIf BlockTopline = BlockStartLine Then
mDoc.FindPos BlockTopline + 1, BlockStartCursor, findit, A, b, c, D
Else
mDoc.FindPos BlockTopline + 1, BlockEndCursor, findit, A, b, c, D
End If
Else
mDoc.FindPos glistN.ListIndex + 1, glistN.SelStart, findit, A, b, c, D
End If
If findit <> -1 Then LastSelStart = findit
SelStart = findit
End Property

Public Property Let SelStart(ByVal RHS As Long)
Dim findit As Long, sel1 As Long
sel1 = SelLength
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
Dim A As Long, b As Long, c As Long, D As Long
findit = RHS
mDoc.FindPos 1, 0, findit, A, b, c, D
SetRowColumn A, b, sel1 <> 0
If sel1 <> 0 Then SelLength = sel1: LastSelStart = sel1
End Property
Public Property Let SelStartSilent2(ByVal RHS As Long)
Dim findit As Long, sel1 As Long
sel1 = SelLength
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
Dim A As Long, b As Long, c As Long, D As Long
findit = RHS
mDoc.FindPos 1, 0, findit, A, b, c, D
SetRowColumn A, b, True
If sel1 <> 0 Then SelLength = sel1: LastSelStart = sel1
End Property
Public Property Let SelStartSilent(ByVal RHS As Long)
Dim findit As Long
Dim A As Long, b As Long, c As Long, D As Long
LastSelStart = RHS
findit = RHS
mDoc.FindPos 1, 0, findit, A, b, c, D
SetRowColumn A, b, True
End Property

Sub ClearUndo()
undopara = -1
NeoUndo = 0
UndoMax = 20
ReDim MultiUndo(1 To UndoMax)
End Sub
Function FindNewUndo() As Long
If NeoUndo = 0 Then
UndoMax = 20
ReDim MultiUndo(1 To UndoMax)
End If
NeoUndo = NeoUndo + 1
If NeoUndo > UndoMax Then
UndoMax = UndoMax * 2
ReDim Preserve MultiUndo(1 To UndoMax)
ElseIf UndoMax > 20 And NeoUndo < UndoMax / 2 Then
UndoMax = UndoMax / 2
ReDim Preserve MultiUndo(1 To UndoMax)
End If
FindNewUndo = NeoUndo
End Function
Private Sub PushUndo()
If undopara = -1 Then Exit Sub
Dim b As Long
b = FindNewUndo
With MultiUndo(b)
.undotext = undotext  ' new  seltext
.undopara = undopara  ' not need for this 2nd version
.undomode = undomode  ' needed
.undopos = undopos  ' new not charpos but selstart
.undoprev = undoprev
undopos = ParaSelStart
undotext = vbNullString
undopara = mDoc.MarkParagraphID  ' nod needed maybe for trace errors later
undomode = -1
undopara = -1
undoprev = False
End With
End Sub
Private Function PopUndo() As Boolean
' false ...no undo saved
If NeoUndo < 1 Then Exit Function

With MultiUndo(NeoUndo)
undotext = .undotext
.undotext = vbNullString  ' flush
undopara = .undopara
undomode = .undomode
undopos = .undopos
undoprev = .undoprev
End With
NeoUndo = NeoUndo - 1
PopUndo = True
If undopara < 0 Then Exit Function

End Function
Sub ClearRedo()
Redopara = -1
NeoRedo = 0
RedoMax = 20
ReDim MultiRedo(1 To RedoMax)
End Sub
Function FindNewRedo() As Long
If NeoRedo = 0 Then
RedoMax = 20
ReDim MultiRedo(1 To RedoMax)
End If
NeoRedo = NeoRedo + 1
If NeoRedo > RedoMax Then
RedoMax = RedoMax * 2
ReDim Preserve MultiRedo(1 To RedoMax)
ElseIf RedoMax > 20 And NeoRedo < RedoMax / 2 Then
RedoMax = RedoMax / 2
ReDim Preserve MultiRedo(1 To RedoMax)
End If
FindNewRedo = NeoRedo
End Function
Private Sub PushRedo()

If Redopara = -1 Then Exit Sub
Dim b As Long
b = FindNewRedo
With MultiRedo(b)
.undotext = Redotext
.undopara = Redopara
.undomode = Redomode
.undopos = Redopos
.undoprev = Redoprev
Redopos = ParaSelStart
Redotext = vbNullString
Redopara = mDoc.MarkParagraphID
Redomode = -1
End With
End Sub
Private Function PopRedo() As Boolean
' false ...no Redo saved
If NeoRedo < 1 Then Exit Function
With MultiRedo(NeoRedo)
Redotext = .undotext
.undotext = vbNullString  ' flush
Redopara = .undopara
Redomode = .undomode
Redopos = .undopos
Redoprev = .undoprev
End With
NeoRedo = NeoRedo - 1
PopRedo = True
End Function

Public Sub ResetUndoRedo()
ClearUndo
ClearRedo
Redopara = -1
undopara = -1
End Sub
Public Function CurrentParagraph() As String
'If showparagraphWork And showparagraph Then
'CurrentParagraph = Replace(Replace(mdoc.TextParagraph(mdoc.MarkParagraphID), " ", ChrW$(&HB7)), ChrW$(&H2007), ChrW$(&H2423))
'Else
CurrentParagraph = mDoc.TextParagraph(mDoc.MarkParagraphID)
'End If
End Function
Public Function NextParagraph() As String
Dim dummy As Long
If showparagraphWork And showparagraph Then
NextParagraph = Replace(Replace(mDoc.NextTextParagraph(mDoc.MarkParagraphID, dummy), " ", ChrW$(&HB7)), ChrW$(&H2007), ChrW$(&H2423))
Else
NextParagraph = mDoc.NextTextParagraph(mDoc.MarkParagraphID, dummy)
End If
End Function
Public Function AutoIntNewLine() As Boolean
' if autonumber
Dim A$, newid As Long 'mDoc.MarkParagraphID
Dim b$, c$, D$, e$, skip As Boolean, tr$, maxw As Long, Ns As String, och As Long

showparagraphWork = False

b$ = CurrentParagraph
' FIND CHARPOSNOW
If Charpos > Len(b$) + 1 Then SelStartSilent = SelStart - (Charpos - Len(b$) - 1)
FixCharPos
On Error Resume Next
If AutoNumber And Not mDoc.IsPart(mDoc.MarkParagraphID) Then
If IsNumberLabel(b$, c$) Then
D$ = NextParagraph: If D$ <> "" Then If IsNumberLabel(D$, e$) Then skip = True
If Not skip Then c$ = Trim$(c$)
If e$ <> "" Then If val(e$) > val(c$) + 10 Then skip = False
If (val(c$) Mod 10 = 0) And Not skip Then
tr$ = String$(Len(NLTrim2(c$)), "0")
A$ = Format((val(c$) + 10), tr$) + Mid$(CurrentParagraph, 1, MyTrimL(CurrentParagraph) - 1) + Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph), MyTrimL(Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph))) - 1)
Else
If e$ <> "" And Trim$(str((val(c$) + 1))) = Trim$(e$) Then
maxw = glistN.UserControlTextWidth(c$)
A$ = space$(Len(c$))
While glistN.UserControlTextWidth(A$) < maxw
A$ = A$ + " "
Wend
c$ = vbNullString

Else
tr$ = String$(Len(NLTrim2(c$)), "0")
A$ = Format((val(c$) + 1), tr$) + Mid$(CurrentParagraph, 1, MyTrimL(CurrentParagraph) - 1) + Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph), MyTrimL(Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph))) - 1)

End If
End If
Else
  GoTo there
End If
Else
there:
        If HaveMarkedText Then
        ClearRedo
        PushUndo
        glistN_MarkDelete False
        undoprev = True
        End If
        Ns = NLtrim(CurrentParagraph)
        A$ = Myleft$(CurrentParagraph, Len(CurrentParagraph) - Len(Ns))
 If Len(A$) > Charpos - 1 Then
SelStartSilent = SelStart - Charpos + Len(A$) + 1
Else
If Charpos > 1 And Not PreserveRightSpaces Then
maxw = MyTrimRNoCr(Myleft$(CurrentParagraph, Charpos - 1))
If Charpos > maxw Then
och = Charpos - maxw - 1
If och > 0 Then
skip = enabled
enabled = False

SelStartSilent2 = SelStart - och
SelLengthSilent = och
InsertTextNoRender = ""
undoprev = True
enabled = skip
End If
End If
End If
 End If

End If

If A$ <> "" Then
RemoveUndo vbCrLf + A$
InsertText = vbCrLf + A$
AutoIntNewLine = True
End If
showparagraphWork = True

End Function

Public Property Get Length() As Long
Length = mDoc.SizeCRLF
End Property


Function GetLCIDFromKeyboard() As Long
    Dim Buffer As String, ret&, R&
    Buffer = String$(514, 0)
      R = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF&
      R = val("&H" & Right(Hex(R), 4))
        ret = GetLocaleInfo(R, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    GetLCIDFromKeyboard = CLng(val("&h" + Left$(Buffer, ret - 1)))
End Function
Public Property Let ReplaceTitle(A$)

Dim curlines As Long
NoTitle = False
NoHeadwrap = False
curlines = mDoc1.DocLines
mDoc1.EmptyDoc
mDoc1.textDoc = Title$ + A$
mDoc1.WrapAgain
If mDoc1.DocLines >= curlines + glistN.lines + 1 Then
mDoc1.EmptyDoc
NoHeadwrap = True
mDoc1.textDoc = Title$ + A$
End If
glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight
glistN.HeadlineHeight = mDoc1.DocLines * onelineheight

there12345:
glistN.ShowMe False, True
End Property
Private Sub PrintItem1X(mHdc As Long, ByVal c As String, R As RECT, Col() As Long)
On Error Resume Next
Dim cc As String, all As Long, Offset As Long
Dim b As RECT, Origin As Long, tabw As Long
tabw = tParam.iTabLength * glistN.AveCharWith
Origin = R.Left
Dim ww$(), items&, i, rleft
If UBound(Col()) < 1 Or NoColor Then
    PrintItem1 mHdc, c, R
    Exit Sub
End If
If UBound(Col()) = 2 Then
If Col(0) = 0 And Col(1) = 0 And Col(2) = 0 Then GoTo there
End If
If UBound(Col()) = 1 And (Col(0) = 14 Or Col(0) = 0 Or Col(0) = 7) Then
there:
    PrintItem1 mHdc, c, R
    Exit Sub
End If
rleft = R.Left

ww$() = SplitStringUsingArrayNoTab(c, Col(), items&)


Offset = rleft
For i = 0 To items&
    c = ww$(i)
       R.Left = Offset
       If R.Left > R.Right Then Exit For
    R.top = R.top + TextPixelOffset
    
    If Not NoColor Then SetTextColor mHdc, M2000(Col(i * 2))
    Offset = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, Origin))
    R.top = R.top - TextPixelOffset
    
Next i
 If R.Left < R.Right Then
If showparagraphonce Then
    cc = "∂"
    SetTextColor mHdc, M2000(14)
    R.Left = Offset
    DrawTextEx mHdc, StrPtr(cc), -1, R, mway, VarPtr(tParam)
    cc = vbNullString
End If
End If
R.Left = rleft
End Sub
Private Sub PrintItem2X(mHdc As Long, ByVal c As String, R As RECT, area As RECT, Col() As Long, ver1 As Boolean)
On Error Resume Next
Dim cc As String, all As Long, Offset As Long, addthis As Long, hRgn As Long
Dim b As RECT, Origin As Long, tabw As Long, dummy As RECT
tabw = tParam.iTabLength * glistN.AveCharWith
Origin = R.Left
dummy = R
dummy.top = R.Bottom + 10
dummy.Bottom = dummy.top + R.Bottom - R.top
Dim ww$(), items&, i, rleft
If UBound(Col()) < 1 Then
    b = area
    hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn mHdc, hRgn
    If ver1 Then
        PrintItem1 mHdc, c, R
    Else
        PrintItem mHdc, c, R
    End If
    SelectClipRgn mHdc, &H0
    DeleteObject hRgn
    Exit Sub
End If
If UBound(Col()) = 2 Then
If Col(0) = 0 And Col(1) = 0 And Col(2) = 0 Then GoTo there
End If
If UBound(Col()) = 1 And (Col(0) = 14 Or Col(0) = 0 Or Col(0) = 7) Then
there:
    b = area
    hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn mHdc, hRgn
    If ver1 Then
        PrintItem1 mHdc, c, R
    Else
        PrintItem mHdc, c, R
    End If
    SelectClipRgn mHdc, &H0
    DeleteObject hRgn
    Exit Sub
End If
rleft = R.Left
all = rleft + glistN.UserControlTextWidthPixels(c)
R.Left = all
ww$() = SplitStringUsingArrayNoTab(c, Col(), items&)

        If showparagraphonce And ver1 Then
            cc = "∂"
            b = area
        hRgn = CreateRectRgnIndirect(b)
        SelectClipRgn mHdc, hRgn
        DrawTextEx mHdc, StrPtr(cc), -1, R, mway, VarPtr(tParam)
        SelectClipRgn mHdc, &H0
        DeleteObject hRgn
            
        End If
    cc = vbNullString
    Offset = rleft
    For i = 0 To items&
    c = ww$(i)
    R.Left = Offset
    R.top = R.top + TextPixelOffset
        b = R

        If b.Right <= area.Left Or b.Left >= area.Right Then
        hRgn = CreateRectRgnIndirect(dummy)
        SelectClipRgn mHdc, hRgn
        Offset = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, Origin))

        SelectClipRgn mHdc, &H0
        DeleteObject hRgn
         
        Else
        hRgn = CreateRectRgnIndirect(area)
        SelectClipRgn mHdc, hRgn
         Offset = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, Origin))

        SelectClipRgn mHdc, &H0
        DeleteObject hRgn
        End If
        If R.Left > area.Right Then Exit For
        R.top = R.top - TextPixelOffset
    Next i
alfa1:
R.Left = rleft
End Sub
Private Sub PrintItem3X(mHdc As Long, ByVal c As String, R As RECT, Col() As Long)
On Error Resume Next
Dim cc As String, all As Long, Offset As Long, addthis As Long
Dim b As RECT, Origin As Long, tabw As Long
tabw = tParam.iTabLength * glistN.AveCharWith
Origin = R.Left
Dim ww$(), items&, i, rleft
If UBound(Col()) < 1 Then
    PrintItem1 mHdc, c, R
    Exit Sub
End If
If UBound(Col()) = 2 Then
If Col(0) = 0 And Col(1) = 0 And Col(2) = 0 Then GoTo there
End If
If UBound(Col()) = 1 And (Col(0) = 14 Or Col(0) = 0 Or Col(0) = 7) Then
there:
    PrintItem1 mHdc, c, R
    Exit Sub
End If
rleft = R.Left
'all = rleft + glistN.UserControlTextWidthPixels(c)
'r.Left = all

ww$() = SplitStringUsingArrayNoTab(c, Col(), items&)

cc = vbNullString
Offset = rleft
For i = 0 To items&
    c = ww$(i)
       R.Left = Offset '+ rleft + addthis
       If R.Left > R.Right Then Exit For
    R.top = R.top + TextPixelOffset
    Offset = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, Origin))
    R.top = R.top - TextPixelOffset
    
Next i
R.Left = rleft
End Sub


Private Sub PrintItem1Xold(mHdc As Long, ByVal c As String, R As RECT, Col() As Long)
On Error Resume Next
    Dim cc As String, all As Long, CM As String, Offset As Long, addthis As Long, hRgn As Long
    Dim b As RECT
    Dim ww$(), items&, i, rleft, cuts() As Long
    If UBound(Col()) <= 1 Or NoColor Then
        PrintItem1 mHdc, c, R
        Exit Sub
    End If
    rleft = R.Left
    all = rleft + glistN.UserControlTextWidthPixels(c)
    R.Left = all
    If InStr(c, vbTab) > 0 Then
    ReDim cuts(UBound(Col()))
    ww$() = SplitStringUsingArray(c, Col(), items&, cuts())
    If showparagraphonce Then
        cc = "∂"
        DrawTextEx mHdc, StrPtr(cc), -1, R, mway, VarPtr(tParam)
    End If
    cc = vbNullString
    
    
    For i = items& To 0 Step -1
        CM = ww$(i)
        R.Left = Offset + rleft + addthis
        R.top = R.top + TextPixelOffset
        b = R
        If i > 0 Then
        
            b.Left = rleft + glistN.UserControlTextWidthPixels(Left$(c, cuts(i - 1))) + 1
          If b.Left < 0 Then b.Left = 0
          
        
        End If

        If Not NoColor Then SetTextColor mHdc, M2000(Col(i * 2))
         hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn mHdc, hRgn
        DrawTextEx mHdc, StrPtr(CM), -1, R, mway, VarPtr(tParam)
          SelectClipRgn mHdc, &H0
   DeleteObject hRgn
           R.top = R.top - TextPixelOffset
   
    Next i
Else
    ww$() = SplitStringUsingArrayNoTab(c, Col(), items&)
    If showparagraphonce Then
        cc = "∂"
        DrawTextEx mHdc, StrPtr(cc), -1, R, mway, VarPtr(tParam)
    End If
    cc = vbNullString
    CM = c
    For i = items& To 0 Step -1
        c = ww$(i)
        CM = Left$(CM, Len(CM) - Len(ww$(i)))
        Offset = glistN.UserControlTextWidthPixels(CM)
        ' for Kenring
        addthis = glistN.UserControlTextWidthPixels(CM + c) - Offset - glistN.UserControlTextWidthPixels(c)
        R.Left = Offset + rleft + addthis
        R.top = R.top + TextPixelOffset
        
        If Not NoColor Then SetTextColor mHdc, M2000(Col(i * 2))
    
        DrawTextEx mHdc, StrPtr(c), -1, R, mway, VarPtr(tParam)
        R.top = R.top - TextPixelOffset
        If R.Left < 0 Then Exit For
    Next i

End If
    R.Left = rleft
    End Sub



Function SplitStringUsingArray(ByVal c$, bb() As Long, wc&, cuts() As Long) As String()
Dim cc$()
ReDim cc$((UBound(bb()) + 1) / 2)
Dim st&, ste&, W&
wc& = 0
ste& = UBound(bb())
For W& = 1 To ste& - 2 Step 2
On Error Resume Next
If bb(W& + 2) - bb(W&) < 1 Then
cc$(wc&) = Mid$(c$, 1, bb(W&))
cuts(wc&) = bb(W& + 2)
Else
If wc& > 0 Then

cc$(wc&) = cc$(wc& - 1) + Mid$(c$, bb(W&), bb(W& + 2) - bb(W&))

Else
cc$(wc&) = Mid$(c$, bb(W&), bb(W& + 2) - bb(W&))
End If
cuts(wc&) = Len(cc$(wc&))

End If
If bb(W& + 2) - bb(W&) = 0 Then
wc& = wc& - 1
SplitStringUsingArray = cc$()
Exit Function
End If
wc& = wc& + 1
Next W&
If bb(W&) > 0 Then
If Trim$(Mid$(c$, bb(W&))) <> vbNullString Then
cc$(wc&) = c$
cuts(wc&) = Len(c$)
Else
 wc& = wc& - 1
End If
End If

SplitStringUsingArray = cc$()
End Function

Function SplitStringUsingArrayNoTab(c$, bb() As Long, wc&) As String()
Dim cc$()
ReDim cc$((UBound(bb()) + 1) / 2)
Dim st&, ste&, W&
wc& = 0
ste& = UBound(bb())
For W& = 1 To ste& - 2 Step 2
On Error Resume Next
If bb(W& + 2) - bb(W&) < 1 Then
cc$(wc&) = Mid$(c$, bb(W&))
Else
cc$(wc&) = Mid$(c$, bb(W&), bb(W& + 2) - bb(W&))
End If
If cc$(wc&) = vbNullString Then
wc& = wc& - 1
SplitStringUsingArrayNoTab = cc$()
Exit Function
End If
wc& = wc& + 1
Next W&
If bb(W&) > 0 Then
cc$(wc&) = Mid$(c$, bb(W&))
End If
If cc$(wc&) = vbNullString Then wc& = wc& - 1
SplitStringUsingArrayNoTab = cc$()
End Function

Function M2000(ch As Long) As Long
Dim b As Integer

Select Case ColorSet
Case 0
HighLightColor = rgb(66, 66, 66)
If ch = 13 Then
M2000 = rgb(&HC5, &HF7, &HF7)
ElseIf ch = 12 Then
M2000 = rgb(&HDD, &HAA, 0)
Else
M2000 = QBColor(ch Mod 16)
End If
Case 1
HighLightColor = rgb(220, 220, 245)
Select Case ch
Case 7
M2000 = rgb(&H68, &H68, &H13)
Case 10
M2000 = 0   '' black  '' basic
Case 11
M2000 = rgb(32, 32, 255)
Case 12
M2000 = rgb(&H80, &H40, &H10)
Case 13
M2000 = rgb(&HAA, &H90, &H20)
Case 14
M2000 = rgb(&HAA, &H60, &H30)
Case 15
M2000 = rgb(&H63, &H30, &H63) '' string

End Select
End Select

End Function
Function htmlM2000(ch As Long) As Long
Dim b As Integer

Select Case Abs(pagiohtml$ = "DARK")
Case 0
If ch = 13 Then
htmlM2000 = rgb(&H33, &HEE, &HEE)
ElseIf ch = 12 Then
htmlM2000 = rgb(0, &HAA, &HDD)

Else
htmlM2000 = (QBColor(ch Mod 16) And &HFF00&) + (QBColor(ch Mod 16) \ &H10000) + ((QBColor(ch Mod 16) And &HFF) * &H10000)
End If
Case 1
Select Case ch
Case 7
htmlM2000 = rgb(&H13, &H68, &H68)
Case 10
htmlM2000 = 0   '' black  '' basic
Case 11
htmlM2000 = rgb(255, 32, 32)
Case 12
htmlM2000 = rgb(&H10, &H40, &H80)
Case 13
htmlM2000 = rgb(&H20, &H37, &H97)
Case 14
htmlM2000 = rgb(&H30, &H60, &HAA)
Case 15
htmlM2000 = rgb(&H63, &H30, &H63) '' string
Case Else
htmlM2000 = rgb(255, 255, 255)
End Select
End Select

End Function
Private Function aheadstatus(A$, Optional srink As Boolean = True, Optional Pos As Long = 1) As String 'ok
Dim b$, part$, W$, pos2 As Long, word$

If A$ = vbNullString Then Exit Function
If Pos = 0 Then Pos = 1
Do While Pos <= Len(A$)

    W$ = Mid$(A$, Pos, 1)
    If part$ = vbNullString And W$ = "0" Then
        If Pos + 2 <= Len(A$) Then
            If LCase(Mid$(A$, Pos, 2)) Like "0[x˜]" Then
            'hexadecimal literal number....
                Pos = Pos + 2
                Do While Pos <= Len(A$)
                If Not Mid$(A$, Pos, 1) Like "[0-9a-fA-F]" Then Exit Do
                Pos = Pos + 1
                Loop
                If Mid$(A$, Pos, 1) = "&" Then
                If Mid$(A$, Pos + 1, 1) = "&" Then Pos = Pos + 2 Else Pos = Pos + 1
                End If
                b$ = b$ & "N"
                If Pos <= Len(A$) Then
                    W$ = Mid$(A$, Pos, 1)
                Else
                    Exit Do
                End If
            End If
        End If
    End If

    If W$ = Chr(34) Then
    
        If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        Pos = Pos + 1
        Do While Pos <= Len(A$)
       ' If Mid$(a$, Pos, 1) = Chr(34) And Mid$(a$, Pos - 1, 1) <> "`" Then Exit Do
       If Mid$(A$, Pos, 1) = Chr(34) Then Exit Do
    If Asc(Mid$(A$, Pos, 1)) < 32 Then Exit Do
   
        Pos = Pos + 1
        Loop
    ElseIf W$ = "(" Or W$ = "[" Then
        If part$ <> "" Then
            ' after
            ElseIf Right$(b$, 1) = "a" Then
            b$ = Left$(b$, Len(b$) - 1)
            part$ = vbNullString
            Else
            part$ = "N"
              
        End If
      Pos = Pos + 1
        If W$ = "(" Then
            If Not BlockParam2(A$, Pos, 1, 0) Then Exit Do
            If Mid$(A$, Pos + 1, 1) <> "." Then b$ = b$ + part$
            part$ = vbNullString
        Else
            If Not BlockParam2(A$, Pos, 0, 1) Then Exit Do
            If Len(word$) = 0 Then
            
            word$ = Mid$(A$, 1, Pos)
            b$ = b$ + part$
            part$ = "N"
            End If
        End If
       
        
        
    ElseIf W$ = "{" Then
If word$ <> "" Then
  Select Case myUcase(word$, True)
  Case "THEN", "ELSE", "‘œ‘≈"
  Pos = Pos - 4
  Exit Do
  Case "ELSE.IF"
  Pos = Pos - 7
  Exit Do
  Case "¡ÀÀ…Ÿ”"
  Pos = Pos - 6
  Exit Do
  Case "¡ÀÀ…Ÿ”.¡Õ"
  Pos = Pos - 9
  Exit Do
  End Select
  End If
    

    If part$ <> "" Then b$ = b$ + part$
        If LenB(b$) = 0 Then
        ElseIf Right(b$, 1) <> "l" Then
        Exit Do
        End If
 
        blockStringAhead A$, Pos
        Pos = Pos + 1
        part$ = "S"
        

            If Pos <= Len(A$) Then
        If Not blockStringAhead(A$, Pos) Then Exit Do
        Pos = Pos - 1
        End If
''Exit Do

    Else
        Select Case W$
        Case ","  ' bye bye
        Exit Do
        Case "%"
            If part$ = vbNullString Then
            ' lathos
            End If
        Case "$"
            If part$ = vbNullString Then
                If b$ = vbNullString Then
                    part$ = "N"
                ElseIf Right$(b$, 1) = "o" Then
                    part$ = "N"
                Else
        
                    aheadstatus = b$
                    Exit Function
                End If
            ElseIf part$ = "N" Then
                    b$ = b$ & "Sa"
                    part$ = vbNullString
            End If

        Case "-", "+", "|"
                    b$ = b$ & part$
                    If b$ = vbNullString Then
                        part$ = "o"
                    Else
                    
                part$ = "o"
                End If
         Case "/"
            If Mid$(A$, Pos + 1, 1) = "/" Then Exit Do
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case "*", "^"
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case " ", ChrW$(160)
          If word$ <> "" Then
  Select Case myUcase(word$, True)
  Case "THEN", "ELSE", "‘œ‘≈"
  Pos = Pos - 4
  Exit Do
  Case "ELSE.IF"
  Pos = Pos - 7
  Exit Do
  Case "¡ÀÀ…Ÿ”"
  Pos = Pos - 6
  Exit Do
  Case "¡ÀÀ…Ÿ”.¡Õ"
  Pos = Pos - 9
  Exit Do
  End Select
  End If
JUMPHERE:
        word$ = vbNullString
            If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            Else
            'skip
            End If
        Case "0" To "9", "."
            If part$ = "N" Then
            If Len(A$) < Pos Then
                If Mid$(A$, Pos + 1, 1) Like "[@#%~]" Then
                    Pos = Pos + 1
                ElseIf Mid$(A$, Pos + 1, 1) = "&" Then
                    If Mid$(A$, Pos + 2, 1) = "&" Then Pos = Pos + 2 Else Pos = Pos + 1
                End If
            End If
            ElseIf part$ = "S" Then
            
            Else
            b$ = b$ & part$
            part$ = "N"
            End If
        Case "&"
        If part$ = vbNullString Then
        part$ = "S"
        Else
        b$ = part$
        part$ = "S"
        End If
        Case "e", "E", "Â", "≈"
        If LenB(word$) = 0 Then part$ = "N"
        
        word$ = word$ + W$
           If part$ = "S" Then
            
            
            ElseIf part$ = "N" Then
            
            b$ = b$ & part$
            'part$ = "N"
            End If
            
        Case "~"
        GoTo there1
        Case ">", "<"
            If Len(A$) > Pos Then
                If Mid$(A$, Pos + 1, 1) = "=" Then
                     b$ = b$ & part$
                     If Len(b$) > 1 Then
                        part$ = "o"
                        If Mid$(A$, Pos, 3) = "<=>" Then
                            Pos = Pos + 2
                        Else
                            Pos = Pos + 1
                        End If
                    End If
                ElseIf Mid$(A$, Pos, 2) = "<>" Then
                    b$ = b$ & part$
                    If Len(b$) > 1 Then
                        part$ = "o"
                        Pos = Pos + 1
                    End If
                ElseIf W$ = ">" And Pos > 1 Then
                    If Mid$(A$, Pos - 1, 2) = "->" Then ' "->"
                        If Right$(b$, 1) = "S" Then
                            b$ = b$ + part$
                            part$ = "N"
                        End If
                    End If
                End If
            End If
            GoTo there1
        Case "="
            If Len(A$) > Pos Then
                If InStr(">=", Mid$(A$, Pos + 1, 1)) > 0 Then
                    Pos = Pos + 2
                    GoTo conthere
                End If
            End If
there1:
                If b$ & part$ <> "" Then
                W$ = Replace(b$ & part$, "a", "")
            part$ = vbNullString
               If srink Then
                  Do
                b$ = W$
                W$ = Replace(b$, "NN", "N")
                Loop While W$ <> b$
                         Do
                        b$ = W$
                          W$ = Replace(b$, "SlS", "N")
                          Loop While W$ <> b$
                            Do
                          b$ = W$
                          W$ = Replace(b$, "NlN", "N")
                          Loop While W$ <> b$
    
                Do
                b$ = W$
                W$ = Replace(b$, "NoN", "N")
                Loop While W$ <> b$
                
                Do
                b$ = W$
                W$ = Replace(b$, "SoS", "S")
                Loop While W$ <> b$
                Else
              b$ = W$
               End If
               
                If Left$(b$, Len(b$) - 1) <> "l" Then part$ = "l"
                Else
                    Exit Do
                End If
        Case ")", "}", Is < " ", ":", ";", "'", "\", "/", "]"
            Exit Do
        Case Else
            If part$ = "N" Then
                word$ = word$ + W$
            ElseIf part$ = "S" Then
            Else
                b$ = b$ & part$
                part$ = "N"
                word$ = word$ + W$
            End If
           
        End Select
        End If
        
    Pos = Pos + 1
conthere:
Loop

    W$ = Replace(b$ & part$, "a", "")
    
    b$ = W$
If srink Then
         Do
  b$ = W$

    W$ = Replace(b$, "SlS", "N")
    Loop While W$ <> b$
      Do
    b$ = W$
    W$ = Replace(b$, "NlN", "N")
    Loop While W$ <> b$
    
    Do
    b$ = W$
    W$ = Replace(b$, "NoN", "N")
    Loop While W$ <> b$
    
    Do
    b$ = W$
    W$ = Replace(b$, "SoS", "S")
    Loop While W$ <> b$
End If
   
   
 
    aheadstatus = b$




End Function
Sub ReColor()
WrapAll
ManualInform
End Sub
Sub ReColorBlock()
Dim order1, order2
order1 = 0: order2 = 0
If BlockEndPara <> BlockStartPara Then
order1 = mDoc.ParagraphOrder(BlockStartPara)
order2 = mDoc.ParagraphOrder(BlockEndPara)
End If
If order1 <= order2 Then
mDoc.WrapAgainBlock BlockStartPara, BlockEndPara
'mDoc.WrapAgainColorBlock BlockStartPara, BlockEndPara
Else
'mDoc.WrapAgainBlock BlockEndPara, BlockStartPara
mDoc.WrapAgainColorBlock BlockEndPara, BlockStartPara
End If


ManualInform
End Sub
Sub ColorBlockOnly()
Dim order1, order2
order1 = 0: order2 = 0
If BlockEndPara <> BlockStartPara Then
order1 = mDoc.ParagraphOrder(BlockStartPara)
order2 = mDoc.ParagraphOrder(BlockEndPara)
End If
If order1 <= order2 Then
mDoc.WrapAgainColorBlock2 BlockStartPara, BlockEndPara
Else
mDoc.WrapAgainColorBlock2 BlockEndPara, BlockStartPara
End If
ManualInform
End Sub
Private Function ISSTRINGA(A$, R$) As Boolean
If Len(A$) < 2 Then ISSTRINGA = False: Exit Function
Dim q$, W As Long
R$ = vbNullString
q$ = Chr(34)
'W = 1
'Do
'W = W + 1
W = InStr(2, A$, q$)
'If W > 0 Then
'If Mid$(a$, W - 1, 1) <> "`" Then Exit Do
'Exit Do
'ElseIf W = 0 Then
'Exit Do
'End If
'Loop
If W = 0 Then ISSTRINGA = False: Exit Function
R$ = Mid$(A$, 2, W - 2)
A$ = NLtrim$(Mid$(A$, W + 1))
ISSTRINGA = True
End Function
Private Sub countbrakcet(A$, k As Long, counter As Long)
Dim i As Long, j As Integer, blind As Boolean
For i = Abs(k) To Len(A$)
j = AscW(Mid$(A$, i, 1))
Select Case j
Case 34
blind = Not blind
Case 123
If Not blind Then counter = counter + 1
Case 125
If Not blind Then counter = counter - 1: If counter = 0 Then k = i: Exit For
End Select
Next i

End Sub
Function blockStringAhead2(s$, pos1 As Long, j As Long) As Boolean
Dim i As Long, c As Long
c = Len(s$)
i = pos1
If i > c Then blockStringAhead2 = False: Exit Function
Do

Select Case AscW(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
Loop
Case 123
j = j - 1
Case 125
j = j + 1
End Select
i = i + 1

Loop Until i > c Or j = 0
pos1 = i - 1
blockStringAhead2 = True


End Function
Function blockStringAhead(s$, pos1 As Long) As Long
Dim i As Long, j As Long, c As Long
c = Len(s$)
i = pos1
If i > c Then blockStringAhead = False: Exit Function
Do

Select Case AscW(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
Loop
Case 123
j = j - 1
Case 125
j = j + 1
End Select
i = i + 1

Loop Until i > c Or j = 0
pos1 = i - 1
blockStringAhead = True


End Function

Function FastSymbol(A$, c$, Optional mis As Boolean = False, Optional cl As Long = 1) As Boolean
Dim i As Long, j As Long
'If Len(c$) <> cl Then Stop  ; only for check
j = Len(A$)
If j = 0 Then Exit Function
i = MyTrimL(A$)
If i > j Then Exit Function  ' this is not good
If j - i < cl - 1 Then
If mis Then MyEr "missing " & c$, "ÎÂﬂÂÈ " + c$
Exit Function
End If
'If c$ = Mid$(a$, i, cl) Then
If InStr(c$, Mid$(A$, i, cl)) > 0 Then
'a$ = Mid$(a$, MyTrimLi(a$, i + cl))
A$ = Mid$(A$, cl + i)
'Mid$(a$, i, cl) = space$(cl)
FastSymbol = True
ElseIf mis Then
MyEr "missing " & c$, "ÎÂﬂÂÈ " + c$
End If
End Function
Private Function Myleft$(A As String, Pos As Long)
If Pos < 1 Then Myleft$ = vbNullString Else Myleft$ = Mid$(A, 1, Pos)
End Function
Private Function MyTrimRNoCr(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimRNoCr = 1: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p4 To p2 Step -2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 9
    Case Else
     MyTrimRNoCr = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimRNoCr = l + 2
End Function
Private Function MyTrimL2NoTab(s$) As Long
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then MyTrimL2NoTab = 0: Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160
    Case Else
     MyTrimL2NoTab = (i - p2) \ 2 + 1
   Exit Function
  End Select
  Next i
 MyTrimL2NoTab = 0
End Function
Private Function IsRTLrun() As Boolean
Dim what$
mark1 = 0
mark2 = 0
what$ = mDoc.TextParagraph(mDoc.MarkParagraphID)
Dim A2() As Integer, WHAT1$, R As Long, ii As Long
Dim wr$
Dim mcharpos As Long
mcharpos = Charpos
If Len(what$) = 0 Then GoTo ex1
If mcharpos > Len(what$) Then mcharpos = Len(what$)
ReDim A2(Len(what$) + 10)
If GetStringTypeExW(&HB, 2, StrPtr(what$), Len(what$), A2(0)) = 0 Then Exit Function

'If A2(mcharpos) = 1 Then goto ex1
For ii = mcharpos To 0 Step -1
If ii < Len(what$) Then
If A2(ii) = 2 Then
mark1 = ii + 1
End If

If mark1 > 0 And A2(ii) = 1 Then Exit For
End If
Next ii
If mark1 = 0 Then GoTo ex1
For ii = mark1 - 1 To Len(what$) - 1
If A2(ii) = 1 Then Exit For
If ii >= 0 Then
If A2(ii) = 7 Then
For R = ii + 1 To Len(what$) - 1
If A2(R) = 1 Then Exit For
If A2(R) <> 7 And A2(R) < 10 Then
If (A2(R) And 2) = 2 Then
ii = R
R = -1
Exit For
End If

If mark2 > mcharpos Then: R = 0: Exit For
End If
Next R
If R = 0 Then
Exit For
ElseIf R <> -1 Then
ii = R - 1
End If
End If
If (A2(ii) And 2) = 2 And A2(ii) <> 10 Then
mark2 = ii + 1
If mark2 > 0 And A2(ii) = 1 Then Exit For
End If
End If
Next ii
If mark2 < mcharpos Then
GoTo ex1
End If
If mark1 <> 0 And mark2 <> 0 Then
IsRTLrun = True
'ManualInform
End If
ex1:
   
End Function
Private Sub glistN_RTL(thisHDC As Long, item As Long, where As Long, mark10 As Long, mark20 As Long, Offset As Long)
Offset = RealX1(thisHDC, where, mDoc.TextLine2(item + 1, showparagraph, NoColor), where, mark10)

End Sub
Function RealX1(mHdc As Long, ByVal ExtSelStart As Long, ByVal that$, Optional nValue As Long = 0, Optional found As Long = 0) As Long
If ExtSelStart <= 0 Then ExtSelStart = 1
Dim ww$(), Col() As Long, items&, i, c As String
Dim rleftacc As Long, original As Long, rx1 As Long, direction As Long
Dim OldValue As Long, orx1 As Long, j As Long, nrtl As Boolean
If Len(that$) = 0 Then
    nValue = 1: Exit Function
End If
If showparagraph Then
        that$ = Replace(that$, vbTab, ChrW(&H21E5))
End If
Col() = mDoc.lastcolordata()
original = ExtSelStart
If UBound(Col()) < 1 Then
GoTo there
ElseIf UBound(Col()) = 2 Then
    If Col(0) = 0 And Col(1) = 0 And Col(2) = 0 Then GoTo there
    ww$() = SplitStringUsingArrayNoTab(that$, Col(), items&)
ElseIf UBound(Col()) = 1 And (Col(0) = 14 Or Col(0) = 0 Or Col(0) = 7) Then
there:
ReDim ww$(0)
ww$(0) = that$
Else
ww$() = SplitStringUsingArrayNoTab(that$, Col(), items&)
End If


Dim M1 As Long, M2 As Long
Dim dummy As RECT, R As RECT, tabw As Long, Offset As Long, hRgn As Long, estX As Long
tabw = tParam.iTabLength * glistN.AveCharWith
glistN.UserControlTextMetricsPixels that$, R.Right, R.Bottom
dummy = R
dummy.top = R.Bottom - 1000
dummy.Bottom = dummy.top + R.Bottom
Offset = 0
estX = 0
hRgn = CreateRectRgnIndirect(dummy)
SelectClipRgn mHdc, hRgn
If found > 0 Then
rx1 = LowWord(GetTabbedTextExtent(mHdc, StrPtr(that$), Len(that$), 1, tabw))
If rx1 < found Then

c = Right$(that$, 1)

nValue = 1
RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
If mark1 <> 0 And mark2 <> 0 Then
nValue = Len(that$)
RTL3 that$, nValue, mark1, mark2, nrtl, M1, M2
nValue = mark1
GoTo finish
Else
nValue = Len(ww(i))
RTL3 ww(i), nValue, mark1, mark2, nrtl, M1, M2
If mark1 <> 0 And mark2 <> 0 And nValue >= mark2 Then
If nValue <= mark2 + 1 Then
nValue = mark1
GoTo finish
Else
nValue = nValue + 1
ExtSelStart = nValue
End If
Else
nValue = Len(that$) + 1
GoTo finish
End If

End If
End If
End If
For i = 0 To items&
    c = ww$(i)
    R.Left = Offset
    rleftacc = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
    If Len(c) >= ExtSelStart - estX Or (found > 0 And rleftacc > found) Then
            If Not Len(c) >= ExtSelStart - estX Then
            ExtSelStart = Len(c) + estX
            End If

        nValue = ExtSelStart - estX

        RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
        If nrtl Then
            If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
        End If
        If mark1 = 0 Or mark2 = 0 Then
        If nValue > 1 Then
            j = nValue - 1
            RTL3 c, j, mark1, mark2, nrtl, M1, M2
            If mark1 <> 0 And mark2 <> 0 And j >= mark1 And j <= mark2 Then
            If i = items& Then
            If found > 0 Then
                rx1 = Offset
                If mark1 > 1 Then
                    Offset = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
                End If
                c = Mid$(ww(i), mark1, mark2 - mark1 + 1)
                If nrtl Then
                    M1 = M1 - mark1 + 1
                    M2 = M2 - mark1 + 1
                    If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
                End If
                If found > Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c$), 1, tabw, 0)) Then
                        nValue = mark1 + estX
                        Exit For
                    Else
                        Offset = rx1
                        c = ww$(i)
                    End If
                nValue = j + estX
                End If
             GoTo here1

            Else
            nValue = mark2 - 1 + estX 'j - (nValue > mark2)
            If nValue <= 0 Then nValue = 1
            If mark1 > 1 Then
            
            rx1 = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark2, 1, tabw, 0))
            If found > 0 Then
            If rx1 <= found Then
            nValue = mark1 + estX
            RealX1 = rx1
            Exit For
            End If

            End If
            RealX1 = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
            
            End If
            Exit For
            
            End If
            End If
        End If
        ElseIf mark1 <> 0 And mark2 <> 0 And nValue = mark2 + 1 Then
        If found > 0 Then
         rx1 = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark2, 1, tabw, 0))
        If found > rx1 Then
         nValue = mark1 + estX
            RealX1 = rx1
        Exit For
        End If
        nValue = mark2
        End If
        GoTo here1
       
        End If

        If mark1 <> 0 And mark2 <> 0 And nValue >= mark1 And nValue <= mark2 Then
here1:
            If mark1 > 1 Then
                
                Offset = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), mark1 - 1, 1, tabw, 0))
            End If
            If mark1 = nValue And found > 0 Then

            If Offset > found Then
            j = mark1 - 2
            If j >= 1 Then
            Do While found < R.Left + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j, 1, tabw, 0))
            j = j - 1
            If j = 0 Then j = 1: Exit Do
            Loop
            Else
                            
                GoTo nortl
            End If
            RTL3 c, j, mark1, mark2, nrtl, M1, M2
            If mark1 <> 0 And mark2 <> 0 Then
            nValue = mark2 - (j - mark1) + estX
            Else
            nValue = j + 1 + estX
            End If

            RealX1 = Offset
            Exit For
            End If
            End If
            c = Mid$(ww(i), mark1, mark2 - mark1 + 1)
            If nrtl Then
                M1 = M1 - mark1 + 1
                M2 = M2 - mark1 + 1
                If M2 - M1 > 0 Then Mid$(c, M1, M2 - M1 + 1) = StrReverse(Mid$(c, M1, M2 - M1 + 1))
            End If
           
            If nValue = mark2 + 1 Then
            If found > 0 Then
            For j = mark2 - mark1 + 1 To 1 Step -1
            If found > Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j, 1, tabw, 0)) Then
            
            If j > 1 Then
            If LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j, 1, tabw, 0)) > LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j - 1, 1, tabw, 0)) < 0 Then
            
                        nValue = mark2 - (j) + estX

            Offset = found
            Exit For
            End If
            End If

            nValue = mark2 - (j - 1) + estX

            Offset = found
            Exit For

            End If
            Next j
            If j = 0 Then
            nValue = mark1 - 1 + estX - 1
            If nValue <= 0 Then nValue = 1
            End If
            Else
            nValue = mark1 - 1 + estX - 1
            If nValue <= 0 Then nValue = 1
            End If
            
          
             RealX1 = Offset
         
            Exit For
            ElseIf found > 0 And mark1 = nValue Then
            If found > Offset Then
             nValue = mark2
             
            GoTo conthere
            Else
            nValue = mark2 + 1 + estX
            RealX1 = Offset
            End If
            
            
            Else
conthere:
            rleftacc = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
             If nValue > mark1 Then
                rx1 = rleftacc - 1 - LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), nValue - mark1, 1, tabw, 0))
            Else
                rx1 = rleftacc - 1
            End If
            End If
        If found > 0 Then
         rx1 = rx1 + 1
            If found < rx1 Then
                original = nValue
                For j = nValue To mark2

                rx1 = rleftacc - LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j - mark1 + 1, 1, tabw, 0))

                    
                    original = original + 1
                    If found >= rx1 Then Exit For
                Next j
                
                j = original
                If nrtl Then
                nValue = M2 - ((j + 1 - mark1) - M1) + mark1 + estX
 
                
                
                Else
                
                
                If j > mark2 Then
      
                  nValue = mark2 + 1 + estX
                  If nValue = 0 Then nValue = 1
                      
                          RealX1 = rx1
                          Exit For
                  Else
                  
                          nValue = j
                          RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
                          If nrtl And j >= M1 And j <= M2 Then
                          nValue = nValue + 2 + estX
                          Else
                          nValue = j + estX
                          End If
            
                  
                  End If
              
                
                    RealX1 = rx1 + 1
        End If
                Else
                  original = nValue
                  For j = nValue - 1 To mark1 + 1 Step -1
                
                      rx1 = rleftacc - LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), j - mark1, 1, tabw, 0))
                      If found <= rx1 Then Exit For
                      original = j
                      
                  Next j
                    j = original
                If nrtl Then
                nValue = M2 - ((j + 1 - mark1) - M1) + mark1 + estX
              Else
                    
                    If j < mark1 Then
                       j = mark1
                    End If
                 
                 
                 
                    If nrtl And j - mark1 >= M1 And j - mark1 <= M2 Then
                    nValue = nValue + estX
                    Else
                    nValue = j + estX
                    End If
     End If

                    RealX1 = rx1

            End If
            
          If found = 0 Then
                If Mid$(that$, nValue, 1) = " " Then
         c = Mid$(that$, nValue + 1, 1)
         If Len(c) > 0 Then
         RTL3 c, (1), mark1, mark2, nrtl, M1, M2
         If mark1 = 0 Or mark2 = 0 Then
         nValue = nValue + 1
         'If nvalue1 < 1 Then nvalue1 = 1
         End If
         End If
         End If
         End If
      Exit For
     End If
            
            If found = 0 And nrtl Then
            'nValue = ExtSelStart
            Else
            nValue = mark2 - (nValue - mark1) + (nValue <> mark2) + estX - 1
            End If
            
If nValue = mark1 And items& = 1 Then
rx1 = 0
           
            
            End If
            If nValue = 0 Then nValue = mark2
       
            
            
            RealX1 = rx1
            
        Else ' NO RTL
nortl:
            c = Left$(c, ExtSelStart - estX - 1)
            rx1 = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
            If found > 0 Then
                If found < rx1 Then
                    nValue = ExtSelStart - 1
                    If nValue - estX - 1 < 1 Then
                        RealX1 = Offset
                    Else
                        c = Left$(c, nValue - estX - 1)
                        j = Len(c)
                        RTL3 c, j, mark1, mark2, nrtl, M1, M2      ' found=0
                        If mark1 <> 0 And mark2 <> 0 Then
                        If j <= mark2 And j >= mark1 Then
                        nValue = mark2 - (j - mark1) + estX
                        Else
                        'nValue = nValue + estX - 1
                        End If
                        Else
                       ' nValue = Len(c) + estX
                        End If
                        rx1 = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
                    End If
                    RealX1 = rx1
                    Exit For
                Else
                    RealX1 = rx1
                    
                    nValue = original
                End If
            Else
                RealX1 = rx1
                nValue = ExtSelStart
            End If
            
            Exit For
        End If
  
         RealX1 = rx1
       ' nValue = ExtSelStart
        
        Exit For
    ElseIf i = items& Then
        nValue = ExtSelStart - estX
        RTL3 c, nValue, mark1, mark2, nrtl, M1, M2
        If mark1 <> 0 And mark2 <> 0 Then
            If mark2 = Len(c) Then
                c = Left$(ww(i), mark1 - 1)
                If Len(c) = 0 Then
                    rleftacc = Offset
                ElseIf nrtl And M2 = mark2 Then
                    rleftacc = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
                    c = ww(i)
                    If M2 - M1 > 0 Then c = Mid$(c, M1, M2 - M1 + 1)
                    rleftacc = rleftacc + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
                    nValue = M1 - 1
                Else
                     rleftacc = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), Len(c), 1, tabw, 0))
                    End If
                ElseIf found > 0 Then
                    rleftacc = Offset + LowWord(TabbedTextOut(mHdc, R.Left, R.top, StrPtr(c), nValue, 1, tabw, 0))
                    If rleftacc < found Then nValue = original
                    RealX1 = found
                    GoTo finish
                End If
            End If
        End If
        Offset = rleftacc
        estX = estX + Len(ww(i))
    Next i


If i = items& + 1 Then
    RealX1 = Offset
    nValue = estX
End If
finish:
SelectClipRgn mHdc, &H0
DeleteObject hRgn
End Function

Private Sub RTL3(what As String, where As Long, mark10 As Long, mark20 As Long, numhandle As Boolean, M1 As Long, M2 As Long)
numhandle = False
M1 = 0
M2 = 0
mark10 = 0
mark20 = 0
Dim A2() As Integer, R As Long, ii As Long
Dim wr$
Dim mcharpos As Long
If Len(what) = 0 Then Exit Sub

If where = -1 Then mcharpos = Len(what) Else mcharpos = where - 1
If mcharpos > Len(what) Then Exit Sub
ReDim A2(Len(what) + 10)


If GetStringTypeExW(&HB, 2, StrPtr(what), Len(what), A2(0)) = 0 Then Exit Sub
For ii = mcharpos To 0 Step -1
    If ii < Len(what) Then
    If A2(ii) = 9 Then
   If mark10 > 0 Then Exit For
    Exit Sub
    End If
    If A2(ii) = 2 Then
    mark10 = ii + 1
    End If
    If mark10 > 0 And A2(ii) = 1 Then
    Exit For
    End If
    End If
Next ii

If mark10 = 0 Then Exit Sub
For ii = mark10 - 1 To Len(what) - 1

    If A2(ii) = 1 Then
    Exit For
    End If
    If ii >= 0 Then
        If A2(ii) = 7 Then
            For R = ii + 1 To Len(what) - 1
                If A2(R) = 1 Then Exit For
                If A2(R) = 9 Then
                Exit For
                End If
                If A2(R) <> 7 And A2(R) < 11 Then
                    If (A2(R) And 2) = 2 Then
                        ii = R
                        R = -1
                        Exit For
                    End If
                   If mark20 > mcharpos Then: R = 0: Exit For
                End If
            Next R
            If R = 0 Then
                Exit For
            ElseIf R <> -1 Then
                ii = R - 1
            End If
        End If
        If (A2(ii) And 2) = 2 And A2(ii) < 11 And A2(ii) <> 10 Then
            mark20 = ii + 1
            If mark20 > 0 And A2(ii) = 1 Then
            Exit For
            End If
        End If
    End If
    If A2(ii) = 9 Then
    If mark20 > 0 Then mark20 = ii
    Exit For
    End If
Next ii

'If mark10 > 0 And mark20 > 0 Then
there11:
If mcharpos > 0 Then
If A2(mcharpos) = 3 And A2(mcharpos - 1) <> 3 Then
M1 = mcharpos
M2 = mcharpos
Else
M1 = mcharpos - 1
M2 = mcharpos - 1
End If
Else
Exit Sub

End If
Dim findme As Long


If A2(M1) <> 3 Then Exit Sub
' find language
For findme = mcharpos To mark1 Step -1
If A2(findme) = 2 Then Exit For
Next findme
findme = AscW(Mid$(what, findme + 1, 1))
If findme >= &H600& And findme <= &H6FF& Then
' arabic language
                    Do While M1 > 0
                        If (A2(M1 - 1) > 3 And A2(M1) < 11) And (A2(M1) > 3 And A2(M1) < 10) Then

                        ElseIf A2(M1 - 1) = 4 Then
                        If M1 > 1 Then
                        If A2(M1 - 2) >= 10 Then
                        M1 = M1 - 1
                        Exit Do
                        End If
                        End If
                        Else
                        If Not (A2(M1) = 3 Or A2(M1) = 7) Then Exit Do
                       
                    End If
                        M1 = M1 - 1
                    Loop
                    Do While M2 < mark2
                    If (A2(M2 + 1) > 3 And A2(M2) < 11) And (A2(M2) > 3 And A2(M2) < 10) Then

                        Else
                    If Not (A2(M2) = 3 Or A2(M2) = 7) Then Exit Do
                    
                    End If
                        M2 = M2 + 1
                    Loop
Else

Do While M1 > 0
    If (A2(M1 - 1) > 3 And A2(M1) < 10) Or A2(M1) = 4 Then
 If A2(M1 - 1) = 4 Then
 If M1 > 1 Then
 If A2(M1 - 2) = 11 Then
 M1 = M1 - 1: Exit Do
 End If
 End If
 ElseIf A2(M1) = 1 Then
Exit Do
End If
    Else
    If Not (A2(M1) = 3 Or A2(M1) = 7) Then Exit Do
   
End If
    M1 = M1 - 1
Loop
Do While M2 < mark2

If A2(M2) > 3 And A2(M2) < 10 Then
If A2(M2) = 4 Then
If A2(M2 + 1) <> 3 Then
Exit Do
End If
End If
    Else
If Not (A2(M2) = 3 Or A2(M2) = 7) Then Exit Do

End If
    M2 = M2 + 1
Loop
End If
M1 = M1 + 2

numhandle = True
If mcharpos = M2 Then
where = M1 - (M1 = 1)

Else
where = M2 - (mcharpos - M1)

End If


End Sub

Function BlockParam2(s$, Pos As Long, j As Long, j1 As Long) As Boolean
' need to be open
Dim i As Long, ii As Long, exit1 As Boolean
exit1 = j <> 0
For i = Pos To Len(s$)
Select Case AscW(Mid$(s$, i, 1))
Case 0
Exit For
Case 34
again:
ii = InStr(i + 1, s$, """")
If ii = 0 Then Exit Function
 i = ii
'If Mid$(s$, ii - 1, 1) = "`" Then GoTo again
Case 91
j1 = j1 + 1
Case 93
j1 = j1 - 1
If Not exit1 Then If j1 = 0 Then Exit For
If j1 < 0 Then Exit For
Case 40
j = j + 1
Case 41
j = j - 1
If exit1 Then If j = 0 Then Exit For
If j < 0 Then Exit For
Case 123
i = i + 1
If blockStringPOS(s$, i) Then
Else
i = 0
End If
If i = 0 Then Exit Function
End Select
Next i
If j = 0 And j1 = 0 Then Pos = i: BlockParam2 = True
End Function

