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
Dim link$

If Dir(App.path + "\" + App.ExeName + ".gsb") <> "" Then
link$ = "{dir " + Chr$(34) + App.path + Chr$(34) + " : load " & App.ExeName & "}"
End If

Set f = New Form1
f.Move VirtualScreenWidth() + 2000, VirtualScreenHeight() + 2000

On Error Resume Next

If link$ <> "" Then
    Set m = NewObjectFromActivexDll(App.path + "\m2000.dll", "callback")
    If Err Then
    Err.Clear
    Set m = NewObjectFromActivexDll(App.path + "\lib.bin", "callback")
    End If
Else
Set m = CreateObject("M2000.callback")
    If Err Then
        Err.Clear
        Set m = NewObjectFromActivexDll(App.path + "\m2000.dll", "callback")
    End If
End If
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
If a$ = "" And link$ <> "" Then a$ = link$
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
If a$ = "" And link$ <> "" Then a$ = link$
If m.Status = 0 Then
Set f.mm = Me
Set f.m = m

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
