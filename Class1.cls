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
If Dir(App.path + "\M2000.vbp") <> "" Then
    Set m = CreateObject("M2000.callback")
    If Not m Is Nothing Then GoTo 1000
    Err.Clear
End If

        Set m = NewObjectFromActivexDll(App.path + "\m2000.dll", "callback")
        If Dir(App.path + "\m2000.DLL") <> "" And m Is Nothing Then
            Err.Clear
            GoTo 1235
        End If
        If Err Or m Is Nothing Then
            Err.Clear
            Set m = NewObjectFromActivexDll(App.path + "\lib.bin", "callback")
            If Dir(App.path + "\lib.bin") <> "" And m Is Nothing Then
                Err.Clear
                GoTo 1234
            End If
            If Err Then
                Err.Clear
                Set m = CreateObject("M2000.callback")
            End If
        End If
If Err Then
    Dim thepath$, s$, doslast As Double
    If MsgBox("Install library first. May I try?", vbYesNo) = vbYes Then
    
        If Dir(App.path + "\lib.bin") <> "" Then
1234
        thepath$ = App.path + "\lib.bin"
        ElseIf Dir(App.path + "\M2000.dll") <> "" Then
1235
        thepath$ = App.path + "\M2000.dll"
        Else
            Beep
        End If
        If thepath$ <> "" Then
            s$ = "powershell -WindowStyle Hidden -NoProfile -Command " + Chr$(34) + "Start-Process CMD -WindowStyle Hidden -ArgumentList '/k regsvr32 " + thepath$ + "' -Verb RunAs" + Chr$(34)
            doslast = Shell("CMD /C " + s$, vbMinimizedNoFocus)
            Sleep 5000
            MsgBox "Wait... to finish", vbCritical
            
        End If
    End If
    Unload f
Exit Sub
End If
1000
If m Is Nothing Then

Else
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
    Debug.Assert (InIDECheck = True)
    a$ = commandW
    If Trim$(LCase(a$)) = "-h" Or Trim$(a$) = "/?" Then frmAbout.Show 1: ExitNow = True: Exit Sub
    If a$ = "" And link$ <> "" Then a$ = link$
    m.getform f
    If a$ <> "" Then
        m.run "start"
    Else
        m.run "start desktop"
    End If
    m.StackMax -12345
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
