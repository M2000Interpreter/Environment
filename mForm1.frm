VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M2000"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3795
   Icon            =   "mForm1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
    Private Const GWL_WNDPROC = -4
    Private Const WM_SETTEXT = &HC
    Private m_Caption As String
Public finished As Boolean
Public m As Object, mm As Object
Public Property Get CaptionW() As String
    If m_Caption = "M2000" Then
        CaptionW = vbNullString
    Else
        CaptionW = m_Caption
    End If

End Property
Public Property Let CaptionW(ByVal NewValue As String)

If LenB(NewValue) = 0 Then NewValue = "M2000"
    m_Caption = NewValue
DefWindowProcW Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue)
If WindowState = 0 Then
   Show
  DoEvents
  If m Is Nothing Then Show: Exit Property
  If m.iamvisible Then
On Error Resume Next
  m.SetFocus
  If Err.Number = 0 Then DoEvents
  End If
  End If
     
End Property

Public Property Let CaptionW2(ByVal NewValue As String)
    Static WndProc As Long, VBWndProc As Long
    If NewValue = "" Then NewValue = "M2000"
    m_Caption = NewValue

    If WndProc = 0 Then
        WndProc = GetProcAddress(GetModuleHandleW(StrPtr("user32")), "DefWindowProcW")
        VBWndProc = GetWindowLongA(hWnd, GWL_WNDPROC)
    End If


    If WndProc <> 0 Then
        SetWindowLongW hWnd, GWL_WNDPROC, WndProc
        SetWindowTextW hWnd, StrPtr(m_Caption)
        SetWindowLongA hWnd, GWL_WNDPROC, VBWndProc
    Else
        Caption = m_Caption
       
    End If
If WindowState = 0 Then
   Show
  DoEvents
  If m Is Nothing Then Show: Exit Property
  If m.iamvisible Then

  m.SetFocus
  DoEvents
  End If
  End If
End Property


Public Sub ShowM2000()
On Error Resume Next
If m Is Nothing Then Exit Sub
m.AsyncShow
End Sub
Public Sub HideM2000()
On Error Resume Next
If m Is Nothing Then Exit Sub
m.WindowState = 1
If Visible Then Refresh
DoEvents
End Sub


Private Sub Form_GotFocus()
ShowM2000
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If m Is Nothing Then Exit Sub
On Error Resume Next
If WindowState = 1 Then Exit Sub
m.SetFocus
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo here
If UnloadMode = vbFormControlMenu Then
        If Not m Is Nothing Then m.ShutDown Cancel
        If Cancel Then Exit Sub
        If Not m Is Nothing Then m.getform Nothing
        Set m = Nothing
Else
    If Not m Is Nothing Then m.ShutDown 2: m.getform Nothing
End If
'Cancel = True
Exit Sub
here:
On Error Resume Next
If Not m Is Nothing Then m.ShutDown 2
End Sub
Private Sub Form_Resize()
If Me.WindowState = 2 Then WindowState = 0: Exit Sub
If m Is Nothing Then Exit Sub
If Not m.IhaveExtForm Then
m.ShutDown 2
Set mm = Nothing
End
Exit Sub
End If
m.WindowState = WindowState
End Sub

Private Sub Form_Terminate()
On Error Resume Next
If Not m Is Nothing Then m.ShutDown 2:   m.getform Nothing
Set m = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set Icon = Nothing
If Cancel Then Exit Sub
If Not m Is Nothing Then m.ShutDown 2
Set mm = Nothing

End Sub
Public Sub ShutDown()
If Not m Is Nothing Then m.getform Nothing
Set mm = Nothing
ShutDownAll
End Sub
Public Sub GetIcon(a As String)
Set Icon = UnPackData(a)
End Sub
Private Function UnPackData(sData As String) As Object
    Dim pbTemp  As PropertyBag
    Dim arData()    As Byte
    Let arData() = sData
    
    Set pbTemp = New PropertyBag
    With pbTemp
        Let .Contents = arData()
        Set UnPackData = .ReadProperty("icon")
    End With
     
    Set pbTemp = Nothing
End Function
