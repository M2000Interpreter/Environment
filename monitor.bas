Attribute VB_Name = "Module9"
Option Explicit
Global MinMonitorLeft As Long
Global MinMonitorTop As Long
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal Hdc As Long, lprcClip As Any, ByVal lpfnEnum As Long, dwData As Any) As Long
Public Declare Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hmonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function UnionRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Private Type RECT
 '   Left As Long
  '  Top As Long
   ' Right As Long
    'Bottom As Long
'End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
Public Const MONITORINFOF_PRIMARY = &H1
Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1
Dim rcMonitors() As RECT 'coordinate array for all monitors
Dim rcVS         As RECT 'coordinates for Virtual Screen

Public Type Screens
    top As Long
    Left As Long
    Height As Long
    Width As Long
    primary As Boolean
    Handler As Long
End Type
Public ScrInfo() As Screens
Public Console As Long
' MI.dwFlags = MONITORINFOF_PRIMARY

Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79
Private Const SM_CMONITORS = 80
Private Const SM_SAMEDISPLAYFORMAT = 81

Private Declare Function GetSystemMetrics Lib "user32" ( _
   ByVal nIndex As Long) As Long

Public Property Get VirtualScreenWidth() As Long
If IsWine Then
   VirtualScreenWidth = (GetSystemMetrics(SM_CXVIRTUALSCREEN)) * dv15 - 1
Else
   VirtualScreenWidth = (GetSystemMetrics(SM_CXVIRTUALSCREEN)) * dv15
   End If
End Property
Public Property Get VirtualScreenHeight() As Long
If IsWine Then
VirtualScreenHeight = (GetSystemMetrics(SM_CYVIRTUALSCREEN)) * dv15 - 1
Else
   VirtualScreenHeight = (GetSystemMetrics(SM_CYVIRTUALSCREEN)) * dv15
   End If
End Property
Public Property Get DisplayMonitorCount() As Long
   DisplayMonitorCount = GetSystemMetrics(SM_CMONITORS)
End Property
Public Property Get AllMonitorsSame() As Long
   AllMonitorsSame = GetSystemMetrics(SM_SAMEDISPLAYFORMAT)
End Property
Public Sub GetMonitorsNow()
  Dim n As Long
  MinMonitorLeft = &H7FFFFFFF
  MinMonitorTop = &H7FFFFFFF
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, n
    Console = FindMonitorFromMouse
End Sub
Public Sub GetMonitorsAgain()
  Dim n As Long
  MinMonitorLeft = &H7FFFFFFF
  MinMonitorTop = &H7FFFFFFF
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, n
End Sub
Function EnumMonitors(f As Form) As Long
    Dim n As Long
    EnumDisplayMonitors 0, ByVal 0&, AddressOf MonitorEnumProc, n
    With f
        .move .Left, .top, (rcVS.Right - rcVS.Left) * 2 + .Width - .Scalewidth, (rcVS.Bottom - rcVS.top) * 2 + .Height - .Scaleheight
    End With
    f.Scale (rcVS.Left, rcVS.top)-(rcVS.Right, rcVS.Bottom)
    f.Caption = n & " Monitor" & IIf(n > 1, "s", vbNullString)
    f.lblMonitors(0).Appearance = 0 'Flat
    f.lblMonitors(0).BorderStyle = 1 'FixedSingle
    For n = 0 To n - 1
        If n Then
            Load f.lblMonitors(n)
            f.lblMonitors(n).Visible = True
        End If
        With rcMonitors(n)
            f.lblMonitors(n).move .Left, .top, .Right - .Left, .Bottom - .top
            f.lblMonitors(n).Caption = "Monitor " & n + 1 & vbLf & _
                .Right - .Left & " x " & .Bottom - .top & vbLf & _
                "(" & .Left & ", " & .top & ")-(" & .Right & ", " & .Bottom & ")"
        End With
    Next
End Function
Private Function MonitorEnumProc(ByVal hmonitor As Long, ByVal hdcMonitor As Long, lprcMonitor As RECT, dwData As Long) As Long
    Dim mi As MONITORINFO
    ReDim Preserve rcMonitors(dwData)
    ReDim Preserve ScrInfo(dwData)
    rcMonitors(dwData) = lprcMonitor
    mi.cbSize = Len(mi)
    GetMonitorInfo hmonitor, mi
    
    With ScrInfo(dwData)
    'If IsWine And mi.rcMonitor.Left = 0 And mi.rcMonitor.Top = 0 Then
     '   .Left = 1
      '  .Top = 1
        
   ' Else
    .Left = mi.rcMonitor.Left * dv15
    If .Left < MinMonitorLeft Then MinMonitorLeft = .Left
  
    .top = mi.rcMonitor.top * dv15
    If .top < MinMonitorTop Then MinMonitorTop = .top
    'End If
    
    .Height = (mi.rcMonitor.Bottom - mi.rcMonitor.top) * dv15
    .Width = (mi.rcMonitor.Right - mi.rcMonitor.Left) * dv15
    
    .primary = CBool(mi.dwFlags = MONITORINFOF_PRIMARY)
    .Handler = hmonitor
    End With
    UnionRect rcVS, rcVS, lprcMonitor 'merge all monitors together to get the virtual screen coordinates
    dwData = dwData + 1 'increase monitor count
    MonitorEnumProc = 1 'continue
End Function

Sub SavePosition(hWnd As Long)
    Dim rc As RECT
    GetWindowRect hWnd, rc 'save position in pixel units
    SaveSetting "Multi Monitor Demo", "Position", "Left", rc.Left
    SaveSetting "Multi Monitor Demo", "Position", "Top", rc.top
End Sub


Function FindPrimary() As Long
Dim i As Long
For i = 0 To UBound(ScrInfo())
If ScrInfo(i).primary Then FindPrimary = i: Exit Function
Next i
End Function
Function FindFormSScreenCorner(z As Object) As Long
Dim f As Form
If TypeOf z Is Form Then
Set f = z
Else
Set f = z.Parent
End If
FindFormSScreenCorner = FindMonitorFromPixel(f.Left, f.top)

End Function
Function FindFormSScreen(z As Object)
Dim f As Form
If TypeOf z Is Form Then
Set f = z
Else
Set f = z.Parent
End If

On Error Resume Next
Dim thismonitor As Long
If f.Visible Then
thismonitor = MonitorFromWindow(f.hWnd, MONITOR_DEFAULTTONEAREST)

Else
FindFormSScreen = FindMonitorFromMouse()
Exit Function
End If
Dim i As Long
For i = 0 To UBound(ScrInfo())
If thismonitor = ScrInfo(i).Handler Then FindFormSScreen = i:   Exit Function
Next i
FindFormSScreen = FindPrimary
End Function
Function FindMonitorFromPixel(x, y) As Long
Dim x1 As Long, y1 As Long
x1 = x \ dv15
y1 = y \ dv15
Dim i As Long
For i = 0 To UBound(ScrInfo())
If ScrInfo(i).Handler = MonitorFromPoint(x1, y1, MONITOR_DEFAULTTONEAREST) Then FindMonitorFromPixel = i: Exit Function
Next i

End Function
Function FindMonitorFromMouse()
'
   ' - offset
Dim x As Long, y As Long, tp As POINTAPI
GetCursorPos tp
x = tp.x
y = tp.y
Dim i As Long
For i = 0 To UBound(ScrInfo())
If ScrInfo(i).Handler = MonitorFromPoint(x, y, MONITOR_DEFAULTTONEAREST) Then
FindMonitorFromMouse = i: Exit Function
End If
Next i
FindMonitorFromMouse = FindPrimary
End Function
Sub MoveFormToOtherMonitor(f As Form)
Dim k As Long, z As Long
'k = FindMonitorFromPixel(F.Left, F.Top)
z = FindMonitorFromMouse
'If k <> Z Then
' center to z
If f.Width > ScrInfo(z).Width Then
    If f.Height > ScrInfo(z).Height Then
        f.move ScrInfo(z).Left, ScrInfo(z).top
    Else
        f.move ScrInfo(z).Left, ScrInfo(z).top + (ScrInfo(z).Height - f.Height) / 2
    End If
    
ElseIf f.Height > ScrInfo(z).Height Then
    f.move ScrInfo(z).Left + (ScrInfo(z).Width - f.Width) / 2, ScrInfo(z).top
Else
 ' F.Move ScrInfo(Z).Left + (ScrInfo(Z).width - F.width) / 2, ScrInfo(Z).Top + (ScrInfo(Z).Height - F.Height) / 2

End If
'End If
End Sub
Sub MoveFormToOtherMonitorOnly(f As Form, Optional flag As Boolean)
Dim k As Long, z As Long
If DisplayMonitorCount = 1 Then Exit Sub
Dim nowX As Long, nowY As Long
k = FindMonitorFromPixel(f.Left, f.top)
z = FindMonitorFromMouse
If k = z Then
    If flag Then
        Dim tp As POINTAPI
        GetCursorPos tp
        nowX = tp.x * dv15
        nowY = tp.y * dv15
        flag = False
    Else
        flag = False
        If f.Name <> "GuiM2000" Then
        nowX = (ScrInfo(k).Width - f.Width) / 2 + ScrInfo(k).Left
        nowY = (ScrInfo(k).Height - f.Height) / 2 + ScrInfo(k).top
        Else
        If ScrInfo(z).Left = f.Left And f.Left > 0 And Not f.SkipAutoPos Then
        nowX = f.Left + (Forms.Count - 4) * 450 Mod 4500
        nowY = f.top + (Forms.Count - 4) * 450 Mod 4500
        ElseIf ScrInfo(z).top = f.top And f.top > 0 And Not f.SkipAutoPos Then
        nowX = f.Left + (Forms.Count - 4) * 450 Mod 4500
        nowY = f.top + (Forms.Count - 4) * 450 Mod 4500
        Else
        nowX = f.Left
        nowY = f.top
        End If
        End If
    End If
Else
    If f.Name <> "GuiM2000" Then
        nowX = (ScrInfo(z).Width - f.Width) / 2 + ScrInfo(z).Left
        nowY = (ScrInfo(z).Height - f.Height) / 2 + ScrInfo(z).top
    Else
        If ScrInfo(z).Left <> f.Left And f.Left > 0 And Not f.SkipAutoPos Then
            nowX = ScrInfo(z).Left + (Forms.Count - 4) * 450 Mod 4500
            nowY = ScrInfo(z).top + (Forms.Count - 4) * 450 Mod 4500
        ElseIf ScrInfo(z).top <> f.top And f.top > 0 And Not f.SkipAutoPos Then
            nowX = ScrInfo(z).Left + (Forms.Count - 4) * 450 Mod 4500
            nowY = ScrInfo(z).top + (Forms.Count - 4) * 450 Mod 4500
        Else
        nowX = f.Left
        nowY = f.top
        End If
    End If
End If

If nowX > ScrInfo(z).Left + ScrInfo(z).Width Then
    nowX = ScrInfo(z).Left + ScrInfo(z).Width * 2 / 3
End If
If nowX + f.Width > ScrInfo(z).Left + ScrInfo(z).Width Then
    If f.Width < ScrInfo(z).Width Then
    nowX = ScrInfo(z).Left + ScrInfo(z).Width - f.Width
    Else
    nowX = ScrInfo(z).Left
    End If
End If
If nowY > ScrInfo(z).top + ScrInfo(z).Height Then
    nowY = ScrInfo(z).top + ScrInfo(z).Height * 2 / 3
End If
If nowY + f.Height > ScrInfo(z).top + ScrInfo(z).Height Then
    If f.Height < ScrInfo(z).Height Then
    nowY = ScrInfo(z).top + ScrInfo(z).Height - f.Height
    Else
    nowY = ScrInfo(z).top
    End If
End If

If f.Width > ScrInfo(z).Width Then
    If f.Height > ScrInfo(z).Height Then
        nowX = ScrInfo(z).Left
        nowY = ScrInfo(z).top
    Else
        nowX = ScrInfo(z).Left
        nowY = ScrInfo(z).top + (ScrInfo(z).Height - f.Height) / 2
    End If
    
ElseIf f.Height > ScrInfo(z).Height Then
    nowX = ScrInfo(z).Left + (ScrInfo(z).Width - f.Width) / 2
    nowY = ScrInfo(z).top
ElseIf flag Then
    nowX = ScrInfo(z).Left + (ScrInfo(z).Width - f.Width) / 2
    nowY = ScrInfo(z).top + (ScrInfo(z).Height - f.Height) / 2
End If
f.move nowX, nowY
End Sub
Sub MoveFormToOtherMonitorCenter(f As Form)
Dim k As Long, z As Long
z = FindMonitorFromMouse
If f.Width > ScrInfo(z).Width Then
    If f.Height > ScrInfo(z).Height Then
        f.move ScrInfo(z).Left, ScrInfo(z).top
    Else
        f.move ScrInfo(z).Left, ScrInfo(z).top + (ScrInfo(z).Height - f.Height) / 2
    End If
    
ElseIf f.Height > ScrInfo(z).Height Then
    f.move ScrInfo(z).Left + (ScrInfo(z).Width - f.Width) / 2, ScrInfo(z).top
Else
 f.move ScrInfo(z).Left + (ScrInfo(z).Width - f.Width) / 2, ScrInfo(z).top + (ScrInfo(z).Height - f.Height) / 2

End If
End Sub
