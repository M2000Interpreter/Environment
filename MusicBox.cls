VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MusicBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TaskInterface
Private instrum As Long
Private octava As Integer
Private Channel As Long, Channel1plus As Long
Private volume As Long
Private lNote As Integer
Private mNote As Integer, feedback As Long
Private m_duration As Long, m_now As Long, m_gateleft As Long, m_gate As Long
Private m_scale As Double, gl_volume As Long
Private o_gate As Double, gl_octava As Integer, RVOL As Integer, LVOL As Integer
Private Tie(1 To 20) As Variant  ' maximum 20 notes ties???
Private TieLevel As Long, TieLen As Long
Dim jobnameID As Long
Dim mybusy As Boolean
Private myBase As TaskInterface
'Private Declare Function timeGetTime Lib "kernel32.dll" Alias "GetTickCount" () As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private myProcess As basetask
Private Sub setVolume()
Dim ParkVolume As Long
ParkVolume = LVOL
MemInt(VarPtr(ParkVolume) + 2) = RVOL
midiOutSetVolume hmidi, ParkVolume
End Sub
Private Sub Class_Terminate()
Dim midimsg As String

If mNote >= 0 Then
 midimsg = &H80 + (mNote * &H100) + Channel
 midiOutShortMsg hmidi, midimsg
End If
MusicBoxNote(Channel1plus) = mNote
MusicBoxValue(Channel1plus) = -2
MusicBoxState(Channel1plus) = False
MusicBoxVolume(Channel1plus) = 127
TieLen = 0
TieLevel = 0
If TaskMaster Is Nothing Then Exit Sub
MusicTaskNum = MusicTaskNum - 1
End Sub

Private Property Get TaskInterface_Process() As basetask
Set TaskInterface_Process = myProcess
End Property
Private Property Set TaskInterface_Process(aProcess As basetask)
Set myProcess = aProcess
Set myProcess.Process = myBase
End Property
Private Property Let TaskInterface_ID(ByVal RHS As Long)
If RHS = Channel1plus Then
myBase.Done = True
End If
End Property

Private Property Get TaskInterface_ID() As Long
TaskInterface_ID = Channel1plus
End Property

Private Property Let TaskInterface_interval(ByVal RHS As Currency)
m_duration = Signed(RHS)
End Property

Private Property Get TaskInterface_interval() As Currency
TaskInterface_interval = m_duration
End Property

' IMPLEMENTED PROPERTIES

Private Property Set TaskInterface_Owner(RHS As Object)
' Usage: Private Property Set TaskInterface_Owner(RHS As Form)
  
  'Validating type
 ' If TypeOf RHS Is Form Then
    Set myBase.Owner = RHS
 ' Else
   ' Error tmTypeMisMatch
  'End If
End Property


' IMPLEMENTED METHODS

Private Sub TaskInterface_Parameters(ParamArray Values() As Variant)
'' Usage: Private Sub TaskInterface_Parameters(Color As Long, Count As Long)
On Error GoTo poulos
  ' Verifing parameter count
  If UBound(Values) = 1 Then
    
    On Error Resume Next
   jobnameID = CLng(Values(0)) ' this is the channel name
   instrum = CLng(Values(1))
    Channel = jobnameID - 1
    Channel1plus = 1 + Channel
    volume = 127
    MusicBoxVolume(Channel1plus) = 127
    m_duration = 0&
    
    MusicBoxState(Channel1plus) = True
    If Err.Number > 0 Then
      On Error GoTo 0

    End If
    On Error GoTo 0
    Dim midimsg As Long
    instrument instrum, Channel
  Else
poulos:

  End If
  
End Sub


Private Function TaskInterface_Tick() As Boolean
Dim midimsg As Long, Pad As Long, padtot As Long, BypassGate As Long
 
Dim check As Long, needmute As Boolean
check = UnsignedSub(timeGetTime, m_now)
If mNote >= 0 Then
If m_gate > 0 And check >= m_gate Then
    TieLevel = TieLen
    GoTo cont1
End If
End If
If mute Or check >= m_duration Then
    If TieLevel < TieLen Then
        TieLevel = TieLevel + 1
        m_now = timeGetTime
        m_duration = Tie(TieLevel)(0)
        MusicBoxValue(Channel1plus) = Tie(TieLevel)(1)
        MusicBoxDots(Channel1plus) = Tie(TieLevel)(2)
        TaskInterface_Tick = True
        Exit Function
    End If
    If mNote >= 0 Then
cont1:
        midimsg = &H80 + (mNote * &H100) + Channel
        
        If m_gateleft > 0 Then
            midiOutShortMsg hmidi, midimsg
            mNote = -1
            MusicBoxNote(Channel1plus) = -1
            MusicBoxGate(Channel1plus) = 100 - MusicBoxGate(Channel1plus)
            m_now = timeGetTime
            m_duration = m_gateleft
            m_gateleft = 0
            m_gate = 0
            TaskInterface_Tick = True
            Exit Function
        Else
            ' midiOutShortMsg hmidi, midimsg
            needmute = True
        End If
    End If
    TaskInterface_Tick = True
 Dim BE As Long, BE2 As Long, rvol1 As Integer, lvol1 As Integer
again:
    rvol1 = RVOL
    lvol1 = LVOL
    volume = 0
    BypassGate = -1
    If PlayTuneMIDI(Channel, octava, lNote, BE, BE2, volume, BypassGate, m_scale, feedback, rvol1, lvol1) And Not mute Then

      '      Debug.Print Channel, octava, lNote, BE, BE2, volume, BypassGate, m_scale, feedback
        If feedback = 1 Then
            gl_volume = volume ' we have to restore that
            gl_octava = octava
            MusicBoxScale(Channel1plus) = m_scale
            MusicBoxVolume(Channel1plus) = volume
            o_gate = GatePerCent(Channel)
             If BypassGate <> -1 Then
                If BypassGate < 100 Then
                    GatePerCent(Channel) = BypassGate / 100
                End If
             End If
            GoTo again
        ElseIf feedback = 2 Then
            MusicBoxScale(Channel1plus) = 1
            MusicBoxVolume(Channel1plus) = gl_volume
            GatePerCent(Channel) = o_gate
            octava = gl_octava
            m_scale = 1
            feedback = 0
            GoTo again
        End If

        If lNote = 24 Or mute Then
            mNote = -1
        Else
            mNote = 0
        End If
        If BE < 1 Then BE = 1
        If BE > 6 Then BE = 6
        m_now = timeGetTime
        MusicBoxValue(Channel1plus) = 2# ^ (BE - 1)
        
        m_duration = Signed(BEATS(Channel) / MusicBoxValue(Channel1plus))
        
        If BE2 = 2 Then
            m_duration = m_duration + m_duration \ 2 + m_duration \ 4
            MusicBoxDots(Channel1plus) = 2
        ElseIf BE2 = 1 Then
            m_duration = m_duration + m_duration \ 2
            MusicBoxDots(Channel1plus) = 1
        Else
            MusicBoxDots(Channel1plus) = 0
        End If
        padtot = m_duration
        TieLen = 0
        TieLevel = 0
        m_gate = 0
        If volume < 0 Then
            Do
                volume = 0
                Pad = m_duration
                If Not PlayTuneMIDItie(Channel, BE, BE2, volume, BypassGate) Then
Err1:
                    voices(Channel) = vbNullString
                    volume = 0
                    m_duration = 0
                    GoTo here
                End If
                If BE < 1 Then BE = 1
                If BE > 6 Then BE = 6
                BE = 2# ^ (BE - 1)
                Pad = Signed(BEATS(Channel) / BE)
                If BE2 = 2 Then
                    Pad = Pad + m_duration \ 2 + m_duration \ 4
                ElseIf BE2 = 1 Then
                    Pad = Pad + Pad \ 2
                End If
                padtot = padtot + Pad
                TieLen = TieLen + 1
                If TieLen > 20 Then GoTo Err1
                Tie(TieLen) = Array(Pad, BE, BE2)
            Loop Until volume <> -1
            
        End If
        If feedback = 3 Then
            padtot = padtot * m_scale
        End If
        m_duration = padtot
        If volume > 127 Then volume = 127
        If volume < 1 Then volume = MusicBoxVolume(Channel1plus)
        
        If mNote <> -1 Then
            If BypassGate <> -1 Then
                If BypassGate < 100 Then
                    m_gateleft = padtot * (1 - BypassGate / 100#)
                    m_gate = padtot - m_gateleft
                    MusicBoxGate(Channel1plus) = BypassGate
                End If
            ElseIf GatePerCent(Channel) < 1 Then
                m_gateleft = padtot * (1 - GatePerCent(Channel))
                m_gate = padtot - m_gateleft
                MusicBoxGate(Channel1plus) = GatePerCent(Channel) * 100
            End If
            
            mNote = GetNote(octava, lNote \ 2)
            MusicBoxNote(Channel1plus) = mNote
            MusicBoxVolume(Channel1plus) = volume
            If needmute Then midiOutShortMsg hmidi, midimsg
            If rvol1 <> RVOL Or lvol1 <> LVOL Then
                RVOL = rvol1: LVOL = lvol1
                setVolume
            End If
            midimsg = &H90 + (mNote * &H100) + (volume * &H10000) + Channel
            midiOutShortMsg hmidi, midimsg
        Else
            If needmute Then midiOutShortMsg hmidi, midimsg
            If rvol1 <> RVOL Or lvol1 <> LVOL Then
                RVOL = rvol1: LVOL = lvol1
                setVolume
            End If
            m_gateleft = 0
            m_gate = padtot
            MusicBoxNote(Channel1plus) = -1 ' pause
            MusicBoxGate(Channel1plus) = 100
        End If
        
    Else
here:
        If needmute Then midiOutShortMsg hmidi, midimsg
        If rvol1 <> RVOL Or lvol1 <> LVOL Then
            RVOL = rvol1: LVOL = lvol1
            setVolume
        End If
        mNote = -2
        MusicBoxNote(Channel1plus) = -2
        MusicBoxDots(Channel1plus) = 0
        MusicBoxState(Channel1plus) = False
        MusicBoxGate(Channel1plus) = 100
        MusicBoxScale(Channel1plus) = 1
        TieLen = 0
        TieLevel = 0
        myBase.Done = True
    End If
End If
If needmute Then midiOutShortMsg hmidi, midimsg
'Form1.OwnEvet = jobnameID
End Function



' DELEGATED PROPERTIES



Private Property Let TaskInterface_Done(ByVal RHS As Boolean)
  myBase.Done = RHS
End Property

Private Property Get TaskInterface_Done() As Boolean
 TaskInterface_Done = myBase.Done
End Property


Private Property Get TaskInterface_Owner() As Object
 Set TaskInterface_Owner = myBase.Owner
End Property

Public Property Let TaskInterface_Priority(ByVal Value As PriorityLevel)
  myBase.Priority = Value
End Property

Public Property Get TaskInterface_Priority() As PriorityLevel
  TaskInterface_Priority = myBase.Priority
End Property
Private Property Let TaskInterface_busy(ByVal RHS As Boolean)
mybusy = RHS
End Property

Private Property Get TaskInterface_busy() As Boolean
TaskInterface_busy = mybusy
End Property
Public Property Get TaskInterface_CodeData() As String
TaskInterface_CodeData = vbNullString
End Property
' DELEGATED METHODS

Private Sub TaskInterface_Dispose(ByVal Action As DisposeAction)
If myBase Is Nothing Then Exit Sub
  myBase.Dispose Action
  Set myBase = Nothing
End Sub



' PRIVATE ROUTINES

Private Sub Class_Initialize()
  Set myBase = New TaskBase
octava = 4
volume = 127
mNote = -2
m_scale = 1
End Sub




