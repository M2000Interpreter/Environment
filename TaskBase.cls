VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TaskInterface
' Task Base provides common implementation for
' the TaskInterface interface.

Private myOwner As Object
Private myPriority As PriorityLevel
Private myDone As Boolean
Private myid As Long
Private mybusy As Boolean
Private myInterval As Double
Private myProcess As basetask

Private Sub Class_Initialize()
myid = -1
End Sub


Private Property Get TaskInterface_CodeData() As String

End Property

Private Property Get TaskInterface_Process() As basetask
Set TaskInterface_Process = myProcess
End Property

Public Property Set TaskInterface_Process(Value As basetask)
Set myProcess = Value

End Property



Private Property Let TaskInterface_busy(ByVal RHS As Boolean)
mybusy = RHS
End Property

Private Property Get TaskInterface_busy() As Boolean
TaskInterface_busy = mybusy
End Property


' IMPLEMENTED PROPERTIES




Private Property Let TaskInterface_Done(ByVal RHS As Boolean)
  myDone = RHS
End Property


Private Property Get TaskInterface_Done() As Boolean
  TaskInterface_Done = myDone
End Property


Private Property Let TaskInterface_ID(ByVal RHS As Long)
myid = RHS
End Property

Private Property Get TaskInterface_ID() As Long
TaskInterface_ID = myid
End Property

Private Property Let TaskInterface_interval(ByVal RHS As Double)
myInterval = RHS
End Property

Private Property Get TaskInterface_interval() As Double
TaskInterface_interval = myInterval
End Property



Private Property Get TaskInterface_Owner() As Object
 Set TaskInterface_Owner = myOwner
End Property


Private Property Set TaskInterface_Owner(RHS As Object)
  Set myOwner = RHS
End Property


Public Property Let TaskInterface_Priority(ByVal Value As PriorityLevel)
  myPriority = Value
End Property


Public Property Get TaskInterface_Priority() As PriorityLevel
  TaskInterface_Priority = myPriority
End Property


' IMPLEMENTED METHODS

Private Sub TaskInterface_Dispose(ByVal Action As DisposeAction)
' Releasing references, and chain if requested
  Set myOwner = Nothing
''
''

End Sub






' NOT IMPLEMENTED - Required for interface

Private Sub TaskInterface_Parameters(ParamArray Values() As Variant)
  myDone = True
End Sub


Private Sub TaskInterface_Tick()
  myDone = True
End Sub


