VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RunM2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m As Object, f As Form1
Public Sub doit()
Attribute doit.VB_UserMemId = 0
Dim test
Set f = New Form1
f.Move VirtualScreenWidth() + 2000, VirtualScreenHeight() + 2000
'f.Show


On Error Resume Next
'With New cFIE
'.FEATURE_BROWSER_EMULATION = .InstalledVersion
'End With
Set m = CreateObject("M2000.callback")
'Dim m As New M2000.callback
If Err Then
    MsgBox "Install M2000.dll first", vbCritical
    Unload f
Exit Sub
End If
' check for old m2000.dll
test = m.WindowState
If Err Then
Debug.Assert (InIDECheck = True)
m.run "start"
m.StackMax -12345
a$ = commandW
If Trim$(a$) = "-h" Or Trim$(a$) = "/?" Then frmAbout.Show 1: ExitNow = True: Exit Sub
If m.Status = 0 Then
m.Cli a$, ">"
End If
Set m = Nothing
ExitNow = True
Exit Sub

Else
m.getform f
Debug.Assert (InIDECheck = True)
m.run "start"
m.StackMax -12345
a$ = commandW

If Trim$(a$) = "-h" Or Trim$(a$) = "/?" Then frmAbout.Show 1: ExitNow = True: Exit Sub
If m.Status = 0 Then
Set f.mm = Me
Set f.m = m
'f.Timer1.Interval = 50

Sleep 100
m.AsyncCli a$, ">"
Debug.Print "ok"

Else
Set m = Nothing
ShutDownAll
End If
End If
End Sub

Private Sub Class_Initialize()
Set f = Nothing
'If m Is Nothing Then Exit Sub
'm.ShowGui = False: m.ShutDown: Set m = Nothing

End Sub

Private Sub Class_Terminate()
If m Is Nothing Then Exit Sub
On Error Resume Next
m.ShowGui = False: m.ShutDown 2: Set m = Nothing

Debug.Print "exit now", m Is Nothing
Set m = Nothing
End Sub
