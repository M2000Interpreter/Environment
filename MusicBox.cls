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
Private Channel As Long
Private volume As Long
Private lNote As Integer
Private mNote As Integer
Private m_duration As Long
Dim jobnameID As Long
Dim metronome As Long
Dim mybusy As Boolean
Private myBase As TaskInterface
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private myProcess As basetask

Private Sub Class_Terminate()
Dim midimsg As String
TaskMaster.MusicTaskNum = TaskMaster.MusicTaskNum - 1
If mNote > 0 Then
 midimsg = &H80 + (mNote * &H100) + Channel
 midiOutShortMsg hmidi, midimsg
 End If
End Sub

Private Property Get TaskInterface_Process() As basetask
Set TaskInterface_Process = myProcess
End Property
Private Property Set TaskInterface_Process(aProcess As basetask)
Set myProcess = aProcess
Set myProcess.Process = myBase
End Property
Private Property Let TaskInterface_ID(ByVal RHS As Long)
If RHS = Channel + 1 Then
myBase.Done = True
End If
End Property

Private Property Get TaskInterface_ID() As Long
TaskInterface_ID = Channel + 1
End Property

Private Property Let TaskInterface_interval(ByVal RHS As Double)
m_duration = CLng(RHS)
End Property

Private Property Get TaskInterface_interval() As Double
TaskInterface_interval = CDbl(m_duration)
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
volume = 127
    m_duration = 0 'MasterTimer + 1600
    
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
 Dim midimsg As Long
If mute Or m_duration <= uintnew(CDbl(timeGetTime)) Then
TaskInterface_Tick = True
If mNote > 0 Then
 midimsg = &H80 + (mNote * &H100) + Channel
 midiOutShortMsg hmidi, midimsg
 End If

 Dim BE As Long
 
    If PlayTuneMIDI(voices(Channel), octava, lNote, BE, volume) And Not mute Then
    If volume > 127 Then volume = 127
    If volume < 1 Then volume = 1
    If lNote = 24 Then
    
    mNote = -1
    Else
    mNote = 0
    End If
    
    If BE < 1 Then BE = 1
    If BE > 6 Then BE = 6
        m_duration = uintnew(CDbl(timeGetTime)) + BEATS(Channel) / 2 ^ (BE - 1)
   If mNote <> -1 Then
    mNote = GetNote(octava, lNote \ 2)
      midimsg = &H90 + (mNote * &H100) + (volume * &H10000) + Channel
    midiOutShortMsg hmidi, midimsg
    End If
    
    Else
    mNote = 0
    myBase.Done = True
    End If


End If
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
End Sub




