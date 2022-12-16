Attribute VB_Name = "Module1"
Option Explicit
' M2000 starter
' We have to give some stack space
Private Declare Function ExtractIconEx Lib "shell32.dll" _
    Alias "ExtractIconExW" (ByVal lpszFile As Long, ByVal nIconIndex As Long, _
    phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Public Declare Function CoAllowSetForegroundWindow Lib "ole32.dll" (ByVal pUnk As Object, ByVal lpvReserved As Long) As Long

Private Declare Function GetProcByName Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal nOrdinal As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Sub DisableProcessWindowsGhosting Lib "user32" ()
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Const LOGPIXELSX = 88

Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79
Private Declare Function GetSystemMetrics Lib "user32" ( _
   ByVal nIndex As Long) As Long
Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Long) As Long
Private Const SEM_NOGPFAULTERRORBOX = &H2&
Public m_bInIDE As Boolean
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" ( _
   ByVal wMode As Long) As Long
Public UnloadForm1 As Boolean, a$
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public dv15 As Long
Public ExitNow As Boolean
Dim cfie As New cfie
Private Type PictDesc
    Size As Long
    Type As Long
    hHandle As Long
    hPal As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Public MyIcon As StdPicture
Public Function HandleToStdPicture(ByVal hImage As Long, ByVal imgType As Long) As IPicture

    ' function creates a stdPicture object from an image handle (bitmap or icon)
    
    Dim lpPictDesc As PictDesc, aGUID(0 To 3) As Long
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = imgType
        .hHandle = hImage
        .hPal = 0
    End With
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc, aGUID(0), True, HandleToStdPicture)
    
End Function
Public Function commandW() As String
Static mm$
If mm$ <> "" Then commandW = mm$: Exit Function
If m_bInIDE Then
mm$ = Command
Else
Dim Ptr As Long: Ptr = GetCommandLineW
    If Ptr Then
        PutMem4 VarPtr(commandW), SysAllocStringLen(Ptr, lstrlenW(Ptr))
     If AscW(commandW) = 34 Then
       commandW = Mid$(commandW, InStr(commandW, """ ") + 2)
       Else
            commandW = Mid$(commandW, InStr(commandW, " ") + 1)
        End If
    End If
    End If
    If mm$ = "" And Command <> "" Then commandW = Command Else commandW = mm$
End Function
Sub Main()
Dim mIcons As Long, there As String, iconHandler As Long
dv15 = 1440 / DpiScrX
DisableProcessWindowsGhosting
If cfie.ReadFeature(cfie.ExeName, cfie.InstalledVersion * 1000) = Empty Then
Debug.Print cfie.InstalledVersion
End If
there = App.path
If Right$(there, 1) <> "\" Then
there = there + "\" + App.ExeName + ".exe"
Else
there = there + App.ExeName + ".exe"
End If
mIcons = ExtractIconEx(StrPtr(there), -1, 0, 0, 0)
Call ExtractIconEx(StrPtr(there), 0, iconHandler, 0, 1)
Set MyIcon = HandleToStdPicture(iconHandler, vbPicTypeIcon)
Dim mm As New RunM2000
Dim o As Object
Set o = mm
CoAllowSetForegroundWindow o, 0

mm.doit
If ExitNow Then ShutDownAll
Sleep 500
End Sub
Public Sub ShutDownAll()
Dim z As Form
If Forms.Count > 0 Then
For Each z In Forms
Set z.Icon = LoadPicture("")
Unload z
Next z
End If

If m_bInIDE Then Exit Sub

Sleep 200
SetErrorMode SEM_NOGPFAULTERRORBOX
Debug.Print "ShutDown"
End Sub

Public Function InIDECheck() As Boolean
    m_bInIDE = True
    InIDECheck = True
End Function
Public Function DpiScrX() As Long
Dim lhWNd As Long, lHDC As Long
    lhWNd = GetDesktopWindow()
    lHDC = GetDC(lhWNd)
    DpiScrX = GetDeviceCaps(lHDC, LOGPIXELSX)
    ReleaseDC lhWNd, lHDC
End Function
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
Function IsWine()
Static www As Boolean, wwb As Boolean
If www Then
Else
Err.Clear
Dim hLib As Long, ntdll As String
On Error Resume Next
ntdll = "ntdll"
hLib = LoadLibrary(StrPtr(ntdll))
wwb = GetProcByName(hLib, "wine_get_version") <> 0
If hLib <> 0 Then FreeLibrary hLib
If Err.Number > 0 Then wwb = False
www = True
End If
IsWine = wwb
End Function
