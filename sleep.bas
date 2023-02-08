Attribute VB_Name = "Module7"
Option Explicit
Private Declare Function GdiFlush Lib "gdi32" () As Long
Private Const MEM_DECOMMIT = &H4000&
Private Const MEM_RELEASE = &H8000&
Private Const MEM_COMMIT = &H1000&
Private Const MEM_RESERVE = &H2000&
Private Const MEM_RESET = &H80000
Private Const MEM_TOP_DOWN = &H100000
Private Const PAGE_READONLY = &H2&
Private Const PAGE_READWRITE = &H4&
Private Const PAGE_EXECUTE = &H10&
Private Const PAGE_EXECUTE_READ = &H20&
Private Const PAGE_EXECUTE_READWRITE = &H40&
Private Const PAGE_GUARD = &H100&
Private Const PAGE_NOACCESS = &H1&
Private Const PAGE_NOCACHE = &H200&
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Declare Function GetSystemTimeAsFileTime Lib "kernel32.dll" (lpSystemTimeAsFileTime As Currency) As Long
Public basictimer As Currency
Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&

Private Const INFINITE = &HFFFF
Private Const ERROR_ALREADY_EXISTS = 183&

Private Const QS_HOTKEY& = &H80&
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4&
Private Const QS_MOUSEMOVE& = &H2&
Private Const QS_PAINT& = &H20&
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const QS_MOUSE& = (QS_MOUSEMOVE _
                            Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE _
                            Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT _
                            Or QS_POSTMESSAGE _
                            Or QS_TIMER _
                            Or QS_PAINT _
                            Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE _
                            Or QS_PAINT _
                            Or QS_TIMER _
                            Or QS_POSTMESSAGE _
                            Or QS_MOUSEBUTTON _
                            Or QS_MOUSEMOVE _
                            Or QS_HOTKEY _
                            Or QS_KEY)

Private Declare Function CreateWaitableTimer Lib "kernel32" _
    Alias "CreateWaitableTimerA" ( _
    ByVal lpSemaphoreAttributes As Long, _
    ByVal bManualReset As Long, _
    ByVal lpName As String) As Long
    
Private Declare Function OpenWaitableTimer Lib "kernel32" _
    Alias "OpenWaitableTimerA" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long
    
Private Declare Function SetWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long, _
    lpDueTime As FILETIME, _
    ByVal lPeriod As Long, _
    ByVal pfnCompletionRoutine As Long, _
    ByVal lpArgToCompletionRoutine As Long, _
    ByVal fResume As Long) As Long
    
Private Declare Function CancelWaitableTimer Lib "kernel32" ( _
    ByVal hTimer As Long)
    
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
    
Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare Function MsgWaitForMultipleObjects Lib "user32" ( _
    ByVal nCount As Long, _
    pHandles As Long, _
    ByVal fWaitAll As Long, _
    ByVal dwMilliseconds As Long, _
    ByVal dwWakeMask As Long) As Long
''DoEvents alternative function.
Private Declare Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, ByVal dwSize As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapSize Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function Wow64EnableWow64FsRedirection Lib "kernel32.dll" (ByVal Enable As Boolean) As Boolean
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
    

Public Function Is64bit() As Boolean
    Static m As Boolean, used As Boolean
If used Then
Else
   If GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process") > 0 Then
      IsWow64Process GetCurrentProcess(), m
      used = True
   End If
End If
   Is64bit = m
End Function

Public Function Blockmalloc(ByVal nBytes As Long, way As Long, Optional NoRun As Boolean = True) As Long
If NoRun Then
    Dim hHeap As Long: hHeap = GetProcessHeap()
    Blockmalloc = HeapAlloc(hHeap, way, nBytes)
Else
    'allocate but not for execution yet
    Blockmalloc = VirtualAlloc(ByVal 0&, nBytes, MEM_COMMIT + MEM_RESERVE, PAGE_READWRITE)
 End If
End Function
Public Sub VirtualBlockCopy(oPtr As Long, mPtr As Long, NewmBytes As Long, oldmBytes As Long)
VirtualLock mPtr, NewmBytes
VirtualLock oPtr, oldmBytes
 If NewmBytes >= oldmBytes Then
            ' copy mBytes to new
            CpyMem ByVal oPtr, ByVal mPtr, oldmBytes
            Else
            CpyMem ByVal oPtr, ByVal mPtr, NewmBytes
            
            End If
VirtualUnlock oPtr, oldmBytes
VirtualUnlock mPtr, NewmBytes
End Sub
Public Function BlockResizeVirtual(ByVal Ptr As Long, ByVal nBytes As Long) As Long
    BlockResizeVirtual = VirtualAlloc(ByVal Ptr, nBytes, MEM_COMMIT, PAGE_READWRITE)
End Function
Public Function BlockResizemalloc(ByVal Ptr As Long, ByVal nBytes As Long, way As Long) As Long
    Dim hHeap As Long: hHeap = GetProcessHeap()
    BlockResizemalloc = HeapReAlloc(hHeap, way, Ptr, nBytes)
End Function
Public Function BlockSize(ByVal Ptr As Long) As Long
    Dim hHeap As Long: hHeap = GetProcessHeap()
    BlockSize = HeapSize(hHeap, 0, Ptr)
End Function
Public Sub BlockFree(ByVal Ptr As Long)

  '  Debug.Print HeapFree(GetProcessHeap(), 0, Ptr)
        HeapFree GetProcessHeap(), 0, Ptr
End Sub
Public Sub SetUpForExecution(ByVal Ptr As Long, ByVal nBytes As Long)
        Dim oldV As Long
        FlushInstructionCache GetCurrentProcess, Ptr, nBytes
        VirtualProtect Ptr, nBytes, PAGE_EXECUTE_READ, oldV ' PAGE_READWRITE
        VirtualLock Ptr, nBytes
End Sub
Public Sub ReleaseExecution(ByVal Ptr As Long, ByVal nBytes As Long)
        Dim oldV As Long
       FlushInstructionCache GetCurrentProcess, Ptr, nBytes
        
        VirtualUnlock Ptr, nBytes
         VirtualProtect Ptr, nBytes, PAGE_READWRITE, oldV '  PAGE_EXECUTE_READ
End Sub
Public Sub BlockFreeVirtual(ByVal Ptr As Long, ByVal nBytes As Long)
      ' VirtualUnlock Ptr, nBytes
     
      If VirtualFree(Ptr, 0&, &H8000&) = 0 Then
      Debug.Print GetLastError()
      End If
End Sub
Public Sub MyRefresh(bstack As basetask)
On Error Resume Next
Dim some As Object
With Prefresh(GetCode(bstack.Owner))
            Call GetSystemTimeAsFileTime(basictimer)
            If basictimer > .k1 Then .RRCOUNTER = 0
            
                If .RRCOUNTER = 0 Then
                    .k1 = basictimer + REFRESHRATE: .RRCOUNTER = 1
                    If Not bstack.toprinter Then
                    Set some = bstack.Owner
                    
                     If some.Visible Then
                    If TypeOf some Is GuiM2000 Then
                        
                        If some.RefreshList > 0 Then some.RefreshALL Else some.Refresh
                        
                    Else
                    some.Refresh
                    End If
                    End If
                    End If
                 End If
                 
                 End With
End Sub
Public Sub SkipRefresh(bstack As basetask)
    With Prefresh(GetCode(bstack.Owner))
        Call GetSystemTimeAsFileTime(basictimer)
        .k1 = basictimer + REFRESHRATE
    End With
End Sub
Public Sub PrintRefresh(bstack As basetask, Scr As Object)
    With Prefresh(GetCode(bstack.Owner))
        Call GetSystemTimeAsFileTime(basictimer)
        If basictimer > .k1 Then
            .k1 = basictimer + REFRESHRATE: .RRCOUNTER = 1
            If Scr.Visible Then Scr.Refresh
            If Not TaskMaster Is Nothing Then
                TaskMaster.StopProcess
                DoEvents
                TaskMaster.StartProcess
            Else
                DoEvents
            End If
        End If
    End With
End Sub

Public Sub MyDoEvents0new(some As Object)
   On Error GoTo procbliah3
With Prefresh(GetCode(some))
    Call GetSystemTimeAsFileTime(basictimer)
    If basictimer > .k1 Then .RRCOUNTER = 0
        If .RRCOUNTER = 0 Then
            .k1 = basictimer + REFRESHRATE: .RRCOUNTER = 1
            If byPassCallback Then Exit Sub
            If some.Visible Then some.Refresh
            If Not TaskMaster Is Nothing Then
                TaskMaster.StopProcess
                DoEvents
                TaskMaster.StartProcess
            Else
                DoEvents
            End If
   End If
End With
Exit Sub
   
procbliah3:
DoEvents
End Sub

Public Sub MyDoEvents0(some As Object)
    On Error GoTo procbliah3
    With Prefresh(GetCode(some))
        Call GetSystemTimeAsFileTime(basictimer)
        If basictimer > .k1 Then .RRCOUNTER = 0
        If .RRCOUNTER = 0 Then
            .k1 = basictimer + REFRESHRATE: .RRCOUNTER = 1
            If byPassCallback Then Exit Sub
            If some.Visible Then
                If TypeOf some Is GuiM2000 Then
                    If some.RefreshList > 0 Then some.RefreshALL Else some.Refresh
                Else
                    some.Refresh
                End If
            End If
        End If
        If Not TaskMaster Is Nothing Then
            TaskMaster.StopProcess
            DoEvents
            TaskMaster.StartProcess
        Else
            DoEvents
        End If
    End With
    Exit Sub
procbliah3:
    DoEvents
End Sub

Public Sub MyDoEvents1(some As Object, Optional DOeVONLY As Boolean = False, Optional ResetK1 As Boolean)
Static once As Boolean
    If some Is Nothing Then
        If TaskMaster Is Nothing Then
            If DOeVONLY Then
                DoEvents
            Else
                Call GetSystemTimeAsFileTime(basictimer)
                If basictimer > k1 Then RRCOUNTER = 0
                    If REFRESHRATE <> 25 And ResetK1 Then RRCOUNTER = 0
                        If RRCOUNTER = 0 Then
                            If ResetK1 Then REFRESHRATE = 25
                            k1 = basictimer + REFRESHRATE: RRCOUNTER = 1
                        End If
                    End If
                Else
                    TaskMaster.rest
                    If DOeVONLY Then
                        If Not once Then
                            once = True
                            TaskMaster.TimerTickNow
                            TaskMaster.StopProcess
                            DoEvents
                            TaskMaster.StartProcess
                            once = False
                        Else
                            TaskMaster.TimerTickNow
                            TaskMaster.StopProcess
                            DoEvents
                        End If
                    Else
                        If basictimer > k1 Then RRCOUNTER = 0
                        If REFRESHRATE <> 25 And ResetK1 Then RRCOUNTER = 0
                        If RRCOUNTER = 0 Then
                            If ResetK1 Then REFRESHRATE = 25
                            k1 = basictimer + REFRESHRATE: RRCOUNTER = 1
                            TaskMaster.Dispose
                            If Not once Then
                                once = True
                                TaskMaster.TimerTickNow
                                TaskMaster.StopProcess
                                DoEvents
                                TaskMaster.StartProcess
                                once = False
                            Else
                                TaskMaster.TimerTickNow
                                TaskMaster.StopProcess
                                DoEvents
                            End If
                        End If
                    End If
                    TaskMaster.RestEnd
                End If
            Exit Sub
        End If
        On Error Resume Next
        If some Is Nothing Then Set some = Form1
        With Prefresh(GetCode(some))
        If TaskMaster Is Nothing Then
            If DOeVONLY Then
                DoEvents
            Else
                Call GetSystemTimeAsFileTime(basictimer)
                If basictimer > .k1 Then .RRCOUNTER = 0
                If REFRESHRATE <> 25 And ResetK1 Then .RRCOUNTER = 0
                If .RRCOUNTER = 0 Then
                    If ResetK1 Then REFRESHRATE = 25
                    .k1 = basictimer + REFRESHRATE: .RRCOUNTER = 1
                    If byPassCallback Then Exit Sub
                    If some.Visible Then
                        If TypeOf some Is GuiM2000 Then
                            If some.RefreshList > 0 Then some.RefreshALL Else some.Refresh
                        Else
                            some.Refresh
                        End If
                    End If
                End If
            End If
        Else
            TaskMaster.rest
            If DOeVONLY Then
                If Not once Then
                    once = True
                    TaskMaster.TimerTickNow
                    TaskMaster.StopProcess
                    DoEvents
                    TaskMaster.StartProcess
                    once = False
                Else
                    TaskMaster.TimerTickNow
                    TaskMaster.StopProcess
                    DoEvents
                End If
            Else
                Call GetSystemTimeAsFileTime(basictimer)
                If basictimer > .k1 Then .RRCOUNTER = 0
                If REFRESHRATE <> 25 And ResetK1 Then .RRCOUNTER = 0
                If .RRCOUNTER = 0 Then
                    If ResetK1 Then REFRESHRATE = 25
                    .k1 = basictimer + REFRESHRATE: .RRCOUNTER = 1
                    If byPassCallback Then Exit Sub
                    If some.Visible Then
                        If TypeOf some Is GuiM2000 Then
                            If some.RefreshList > 0 Then some.RefreshALL Else some.Refresh
                        Else
                            some.Refresh
                        End If
                    End If
                    If Not once Then
                        once = True
                        TaskMaster.TimerTickNow
                        TaskMaster.StopProcess
                        DoEvents
                        TaskMaster.StartProcess
                        once = False
                    Else
                        TaskMaster.TimerTickNow
                        TaskMaster.StopProcess
                        DoEvents
                    End If
                End If
            End If
            TaskMaster.RestEnd
        End If
    End With
End Sub

Public Sub MyDoEvents2(Optional obj As Object)
On Error GoTo endevents
If TaskMaster.PlayMusic Then
    If k1 = 0 Then Call GetSystemTimeAsFileTime(basictimer): k1 = basictimer: RRCOUNTER = 1
    TaskMaster.OnlyMusic = True
    TaskMaster.TimerTick
    TaskMaster.OnlyMusic = False
End If
If TaskMaster.Processing Then
    If Not extreme Then
        If Not obj Is Nothing Then
              With Prefresh(GetCode(obj))
                Call GetSystemTimeAsFileTime(basictimer)
                If k1 = 0 Then k1 = basictimer: RRCOUNTER = 1
                If basictimer > .k1 Then .RRCOUNTER = 0
                If RRCOUNTER = 0 Then
                    If obj.Visible Then
                        If Kform Then
                            Kform = False
                            .k1 = 0
                            TaskMaster.rest
                            UpdateWindow obj.hWnd
                            DoEvents
                            TaskMaster.RestEnd
                        Else
                            MyDoEvents1 obj
                        End If
                       .RRCOUNTER = 1
                    End If
                End If
            End With
        Else
            If k1 = 0 Then Call GetSystemTimeAsFileTime(basictimer): k1 = basictimer: RRCOUNTER = 1
        End If
    Else
        Call GetSystemTimeAsFileTime(basictimer)
        If k1 = 0 Then k1 = basictimer: RRCOUNTER = 1
        If basictimer > k1 Then RRCOUNTER = 0
        If RRCOUNTER = 0 Then k1 = basictimer + REFRESHRATE: RRCOUNTER = 1
    End If
Else
    Call GetSystemTimeAsFileTime(basictimer)
    If k1 = 0 Then: k1 = basictimer: RRCOUNTER = 1
    If basictimer > k1 Then RRCOUNTER = 0
    If RRCOUNTER = 0 Then
        k1 = basictimer + REFRESHRATE: RRCOUNTER = 1
        If QRY Then
             DoEvents
        Else
            If Kform Then
                Kform = False
                TaskMaster.rest
                DoEvents
                TaskMaster.RestEnd
            Else
                DoEvents
            End If
        End If
    End If
End If
Exit Sub
endevents:
DoEvents
End Sub


Public Sub SleepWait3(lNumberOf10ThmiliSeconds As Long)


Do
 MyDoEvents1 Form3
Sleep 1
lNumberOf10ThmiliSeconds = lNumberOf10ThmiliSeconds - 25

Loop Until lNumberOf10ThmiliSeconds < 0

End Sub

Public Sub SleepWaitNO(ByVal A As Long)
Exit Sub
 Dim b As New clsProfiler
Dim l As Boolean, k
l = NOEDIT
  b.MARKONE
While A > b.MARKTWO And l = NOEDIT
MyDoEvents2 Form1
If Not TaskMaster Is Nothing Then If TaskMaster.Processing Then TaskMaster.TimerTick Else Sleep 0

 A = A \ 3
Wend
End Sub
Private Sub SleepWaitNew(A As Long)
 Dim b As New clsProfiler
  b.MARKONE
Do
 MyDoEvents
Loop Until A > b.MARKTWO
End Sub
Public Sub SleepWaitEdit(bstack As basetask, lNumberOf10ThmiliSeconds As Long)
On Error Resume Next
If Forms.count < 3 Then
Sleep 1
 DoEvents
Exit Sub
End If
If TaskMaster Is Nothing Then
Set TaskMaster = New TaskMaster
TaskMaster.Interval = 5
End If
TaskMaster.rest

    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    Dim dblUnits As Double
    Dim hTimer As Long
    If lNumberOf10ThmiliSeconds = 0 Then lNumberOf10ThmiliSeconds = 1
    hTimer = CreateWaitableTimer(0, True, "")
    
    If Err.LastDllError = ERROR_ALREADY_EXISTS Then
        ' If the timer already exists, it does not hurt to open it
        ' as long as the person who is trying to open it has the
        ' proper access rights.
    Else
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, 0)
    End If
    
    ' Convert the Units to nanoseconds.
    dblUnits = CDbl(&H10000) * CDbl(&H10000)
    
    dblDelay = CDbl(lNumberOf10ThmiliSeconds) * 1000
    
    ' By setting the high/low time to a negative number, it tells
    ' the Wait (in SetWaitableTimer) to use an offset time as
    ' opposed to a hardcoded time. If it were positive, it would
    ' try to convert the value to GMT.
    ft.dwHighDateTime = -CLng(dblDelay / dblUnits) - 1
    dblDelayLow = -dblUnits * (dblDelay / dblUnits - _
    Fix(dblDelay / dblUnits))
    
    If dblDelayLow < CDbl(&H80000000) Then
        ' &H80000000 is MAX_LONG, so you are just making sure
        ' that you don't overflow when you try to stick it into
        ' the FILETIME structure.
        dblDelayLow = dblUnits + dblDelayLow
        ft.dwHighDateTime = ft.dwHighDateTime + 1
    End If
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, False)
    With Prefresh(GetCode(bstack.Owner))
        Do
            ' QS_ALLINPUT means that MsgWaitForMultipleObjects will
            ' return every time the thread in which it is running gets
            ' a message. If you wanted to handle messages in here you could,
            ' but by calling Doevents you are letting DefWindowProc
            ' do its normal windows message handling---Like DDE, etc.
            lBusy = MsgWaitForMultipleObjects(1, hTimer, False, _
                INFINITE, QS_ALLINPUT&)
            Call GetSystemTimeAsFileTime(basictimer)
            If basictimer > .k1 Then .RRCOUNTER = 0
            If .RRCOUNTER = 0 Then
                    .k1 = basictimer + REFRESHRATE: .RRCOUNTER = 1
                If TaskMaster Is Nothing Then
                    DoEvents
                Else
                    TaskMaster.StopProcess
                    DoEvents
                    TaskMaster.StartProcess
                End If
            End If
        Loop Until lBusy = WAIT_OBJECT_0
    End With
    ' Close the handles when you are done with them.
    CloseHandle hTimer
If Not TaskMaster Is Nothing Then TaskMaster.RestEnd
End Sub
        
Public Sub SleepWaitEdit2(lNumberOf10ThmiliSeconds As Long)
On Error Resume Next
If Forms.count < 3 Then
Sleep 1
DoEvents
Exit Sub
End If
If TaskMaster Is Nothing Then
Set TaskMaster = New TaskMaster
TaskMaster.rest
TaskMaster.Interval = 5
Else
TaskMaster.rest
End If
If IsWine Then
    lNumberOf10ThmiliSeconds = 30
End If
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    Dim dblUnits As Double
    Dim hTimer As Long
    If lNumberOf10ThmiliSeconds = 0 Then lNumberOf10ThmiliSeconds = 1
    hTimer = CreateWaitableTimer(0, True, "")
    
    If Err.LastDllError = ERROR_ALREADY_EXISTS Then

        DoEvents
        Exit Sub
        
    Else
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, 0)
    End If
    
    ' Convert the Units to nanoseconds.
    dblUnits = CDbl(&H10000) * CDbl(&H10000)
    
    dblDelay = CDbl(lNumberOf10ThmiliSeconds) * 1000
    
    ft.dwHighDateTime = -CLng(dblDelay / dblUnits) - 1
    dblDelayLow = -dblUnits * (dblDelay / dblUnits - _
        Fix(dblDelay / dblUnits))
    
    If dblDelayLow < CDbl(&H80000000) Then
        ' &H80000000 is MAX_LONG, so you are just making sure
        ' that you don't overflow when you try to stick it into
        ' the FILETIME structure.
        dblDelayLow = dblUnits + dblDelayLow
        ft.dwHighDateTime = ft.dwHighDateTime + 1
    End If
    
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(hTimer, ft, 0, 0, 0, False)
    Dim handlepopup As Boolean, lastpopup As Long
    Do
        lBusy = MsgWaitForMultipleObjects(1, hTimer, False, _
        INFINITE, QS_ALLINPUT&)
        
        DoEvents
        If Not Screen.ActiveForm Is Nothing Then
            If TypeOf Screen.ActiveForm Is GuiM2000 Then
                If Not handlepopup Then
                    If Screen.ActiveForm.PopUpMenuVal Then
                        lastpopup = Screen.ActiveForm.hDC
                        handlepopup = True
                    End If
                ElseIf GetForegroundWindow <> Screen.ActiveForm.hWnd Then
                    If handlepopup Then
                        handlepopup = False
                        SetVisibleByHDC lastpopup, False
                    End If
                End If
            Else
                If GetForegroundWindow <> Screen.ActiveForm.hWnd Then
                    Exit Do
                End If
            End If
        End If
    Loop Until lBusy = WAIT_OBJECT_0
CloseHandle hTimer
If Not TaskMaster Is Nothing Then TaskMaster.RestEnd
End Sub
Sub SetVisibleByHDC(whatHDC As Long, setit As Long)
Dim k As Form
On Error Resume Next
For Each k In Forms
    If k.hDC = whatHDC Then k.Visible = setit
Next k

End Sub
