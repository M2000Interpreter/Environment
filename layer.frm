VERSION 5.00
Begin VB.Form Form5 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   -30000
   ClientTop       =   0
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form5"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&)
Private Const DWL_ANYTHREAD& = 0
Const LOCALE_ILANGUAGE = 1
Private Declare Function SetErrorMode Lib "kernel32" ( _
   ByVal wMode As Long) As Long

Private Const SEM_NOGPFAULTERRORBOX = &H2&
Private Sub Form_Activate()
'If Form1.WindowState <> vbMinimized And Form1.Visible Then Form1.ActiveControl.SetFocus
Me.ZOrder 1
If Form1.Visible Then Form1.SetFocus
End Sub

Private Sub Form_Load()
Set LastGlist = Nothing
form5iamloaded = True
If Not s_complete Then
On Error Resume Next
Me.move -100000
If Err.Number > 0 Then Me.move -30000

If Form1.Visible Then Form1.Hide

End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set LastGlist = Nothing
Set LastGlist2 = Nothing
form5iamloaded = False '
MediaPlayer1.closeMovie
  DisableMidi
 If Not TaskMaster Is Nothing Then TaskMaster.Dispose
  Set TaskMaster = Nothing
Set Basestack1.Owner = Nothing
Set Basestack1 = Nothing
Dim x As Form
If IsWine Then
Modalid = 0

For Each x In Forms
If x.Visible Then x.Visible = False
Next
Set x = Nothing
'Form1.helper1
'MsgBox "quit"
'Exit Sub
Else
For Each x In Forms
If x.Name <> Me.Name Then Unload x
Next
Set x = Nothing
End If

If m_bInIDE Then Exit Sub

SetErrorMode SEM_NOGPFAULTERRORBOX
End Sub
Public Sub BackPort()
If Not IsWine Then Exit Sub
On Error Resume Next
Set LastGlist = Nothing
Set LastGlist2 = Nothing
form5iamloaded = False '
MediaPlayer1.closeMovie
  DisableMidi
 If Not TaskMaster Is Nothing Then TaskMaster.Dispose
  Set TaskMaster = Nothing
  
Dim x As Form
Modalid = 0

For Each x In Forms
If x.Name <> Me.Name Then
Set x.icon = LoadPicture("")
If x.Visible Then x.Visible = False
End If
Next
Set x = Nothing
Form1.helper1

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
INK$ = INK$ & GetKeY(KeyAscii)
End Sub
Public Sub RestoreSizePos()
' calling from form1
Me.move Form1.Left, Form1.Top, Form1.Width, Form1.Height
End Sub
Public Sub RestorePos()
' calling from form1
 'Me.Move Form1.Left, Form1.Top
End Sub
 Function GetKeY(ascii As Integer) As String
    Dim Buffer As String, ret As Long
    Buffer = String$(514, 0)
    Dim R&, k&
      R = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      R = val("&H" & Right(Hex(R), 4))
    ret = GetLocaleInfo(R, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    If ret > 0 Then
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, CLng(val("&h" + Left$(Buffer, ret - 1))))))
    Else
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, 1033)))
    End If
End Function
