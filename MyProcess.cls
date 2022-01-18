VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TaskInterface

Private mybuf$
' load code or just use it as a counter
' load code first time
Private myPriority As PriorityLevel
Private myOwner As Object
Private myDone As Boolean
Private m_duration As Currency
Private mYspace As Currency
Private myHere$
Private myid As Long
Private mycode As String
Private MyMode As Long
Private mybusy As Boolean
Public hPipe As Long  ' τι δουλειά έχει εδώ;;
Private pipename As String  '= "\\.\pipe\bigtest"
Private myholdtime As Currency
Private myBUFvar As String, res As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private myProcess As basetask  ' myprocess.process.id is the handler
Private myProcessPureParent As basetask
Private mNostretch As Boolean
Private mStep As Boolean


Private Property Get TaskInterface_Process() As basetask
Set TaskInterface_Process = myProcess
End Property

Private Property Set TaskInterface_Process(aProcess As basetask)
If aProcess Is Nothing Then
    Set myOwner = Nothing
   myProcessPureParent.ThrowOne myid
    TaskInterface_Done = True
Else
Set myProcess = aProcess
Set myProcessPureParent = aProcess.Parent
While myProcessPureParent.IamThread = True
Set myProcessPureParent = myProcessPureParent.Parent
Wend

If myProcess.Process Is Me Then
myDone = True
Else
Set myProcess.Process = Me
Set myOwner = myProcess.Parent.Owner
MyMode = 1
End If
End If
End Property
Private Sub Class_Terminate()
If hPipe <> 0 Then
res = DisconnectNamedPipe(hPipe)
 CloseHandle hPipe
End If
If MyMode > 0 Then
myProcess.ThrowOne myid
    Set myProcess.Process = Nothing ' αφαίρεση...για να εξαφανιστεί!

    End If
    If Not myProcess Is Nothing Then
     Set myProcess.Owner = Nothing
     Set myProcess = Nothing
     End If
        
    Set myProcessPureParent = Nothing
End Sub

Private Property Let TaskInterface_interval(ByVal RHS As Currency)
mYspace = RHS
m_duration = 0
mStep = False
If mYspace = 0 Then

Else
m_duration = uintnew(CCur(timeGetTime)) + mYspace
End If
End Property

Private Property Get TaskInterface_interval() As Currency
'

TaskInterface_interval = mYspace
End Property

' IMPLEMENTED PROPERTIES

Private Property Set TaskInterface_Owner(RHS As Object)

  Set myOwner = RHS
End Property

' IMPLEMENTED METHODS

Private Sub TaskInterface_Parameters(ParamArray Values() As Variant)
'' Usage: Private Sub TaskInterface_Parameters(Color As Long, Count As Long)
On Error GoTo poulos
  ' Verifing parameter count
  If UBound(Values) = 5 Then
    
    On Error Resume Next
    myid = CLng(Values(0))  ' this is the hanlde
   mYspace = CCur(Values(1)) '
   mycode = CStr(Values(2))
   myholdtime = CCur(Values(3)) ' για το after μονο
    m_duration = uintnew(CCur(timeGetTime)) + myholdtime
   myHere$ = CStr(Values(4))  ' change from 5 to 4
   mNostretch = CBool(Values(5))
    If Err.Number > 0 Then
      On Error GoTo 0
    
    End If
    On Error GoTo 0
    ' do something
  Else
poulos:
    ' not defined yet
  End If
  
End Sub



Private Function TaskInterface_Tick() As Boolean
On Error Resume Next
 Dim x As Long, mycnt As Long, ohere$, look As Boolean, ok As Boolean
 Dim NowProcess As basetask
 Set NowProcess = myProcess

If MyMode = 0 Then
myDone = True
End If
If NowProcess Is Nothing Then
myDone = True
End If
         
 If myDone Then GoTo there
 ok = m_duration <= uintnew(CCur(timeGetTime))
 If ok Or mStep Then
 
            If ok Then m_duration = uintnew(CCur(timeGetTime)) + mYspace

            If mybuf$ = vbNullString Then
            If Not ok Then mStep = False: Exit Function
            mybuf$ = mycode: NowProcess.IFCTRL = 0: NowProcess.jump = False: NowProcess.UseofIf = 0
            End If
            TaskInterface_Tick = True
            ohere$ = here$
            here$ = myHere$
            x = 1
            look = True
            
Set NowProcess.Owner = myOwner
If Interrupted Then SetTextBasketBack myOwner, players(GetCode(myOwner)): mStep = True
 Call executeblock(x, NowProcess, mybuf$, False, look, Interrupted, True)
  If Not (NowProcess Is Nothing) Then Set myOwner = NowProcess.Owner
        If Len(mybuf$) < 3 Then
            mybuf$ = vbNullString
            mStep = False
            
        Else
      
    If Not look Then GoTo conthere
     End If
       If LastErNum <> 0 Or x = 0 Then
       mStep = False
       myDone = True
       End If
        If x = 3 Then x = 1: look = False
        If ProcessEnd Then Exit Function
        
        If (x = 1 Or (x = 2 And mybuf$ = vbNullString)) And look = True Then
        
            If NowProcess.Parent.exist(myid, "_multi") Then
                mStep = False
              TaskMaster.RestEnd1
              TaskMaster.RestEnd
              
            myProcessPureParent.ThrowThreads

            MyMode = 0
            myDone = True
           ''
            GoTo there
            Else
            GoTo there
            End If
        End If
conthere:
If NOEXECUTION Then GoTo there
                here$ = ohere$
        If x = 0 Then
        myDone = True
        
        End If
If MOUT Then GoTo there

   


   End If
    If myDone Then GoTo there
    
   Exit Function
there:
On Error GoTo 0

    Set myOwner = Nothing

   TaskInterface_Dispose tmThisTask
   
End Function
Public Property Get TaskInterface_CodeData() As String
TaskInterface_CodeData = mycode
End Property

' DELEGATED PROPERTIES
Private Property Let TaskInterface_Done(ByVal RHS As Boolean)
myDone = RHS
If MyMode > 0 Then
If Not myProcess.IamRunning Then TaskInterface_Dispose tmAllTasks
End If
End Property

Private Property Get TaskInterface_Done() As Boolean
On Error Resume Next
 TaskInterface_Done = myDone
End Property


Private Property Get TaskInterface_Owner() As Object
 Set TaskInterface_Owner = myOwner
End Property

Public Property Let TaskInterface_Priority(ByVal Value As PriorityLevel)
  myPriority = Value
End Property


Public Property Get TaskInterface_Priority() As PriorityLevel
  TaskInterface_Priority = myPriority
End Property
Private Property Let TaskInterface_ID(ByVal RHS As Long)
If RHS = myid Then
myDone = True
Else
' δεν αλλάζει
End If
End Property

Private Property Get TaskInterface_ID() As Long
TaskInterface_ID = myid
End Property
Private Property Let TaskInterface_busy(ByVal RHS As Boolean)
mybusy = RHS
End Property

Private Property Get TaskInterface_busy() As Boolean
TaskInterface_busy = mybusy
End Property

' DELEGATED METHODS

Private Sub TaskInterface_Dispose(ByVal Action As DisposeAction)
On Error Resume Next
myDone = True
On Error Resume Next
If Not TaskMaster Is Nothing Then TaskMaster.rest
    Set myOwner = Nothing
If MyMode > 0 Then
 If Not myProcess.IamRunning Then
MyMode = 0

   Set myProcess.Process = Nothing

   If Not ProcessEnd Then Set myProcess = Nothing
       
    End If
   End If

If Not TaskMaster Is Nothing Then TaskMaster.RestEnd
End Sub
Private Function ProcessEnd() As Boolean
On Error Resume Next
If myProcess Is Nothing Then
ProcessEnd = True
Else
myProcess.EndRun
End If
End Function





