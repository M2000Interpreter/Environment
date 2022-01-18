VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' The interface class simply gives names
' to the required properties, methods,
' and parameters.


' OBJECTS


Public Property Set Owner(ByRef RHS As Object)
End Property

Public Property Get Owner() As Object
End Property


' SCALARS


Public Property Let Done(ByVal Value As Boolean)
End Property

Public Property Get Done() As Boolean
End Property

Public Property Let Priority(ByVal Value As PriorityLevel)
End Property

Public Property Get Priority() As PriorityLevel
End Property
Public Property Let id(ByVal Value As Long)
End Property
Public Property Get id() As Long
End Property
Public Property Let Busy(ByVal Value As Boolean)
End Property
Public Property Get Busy() As Boolean
End Property
Public Property Get Interval() As Currency
End Property
Public Property Let Interval(ByVal Value As Currency)
End Property
Public Property Get CodeData() As String
End Property
' METHODS



Public Sub Dispose(ByVal Action As DisposeAction)
End Sub

Public Sub Parameters(ParamArray Values())
End Sub
Public Property Get Process() As basetask
End Property
Public Property Set Process(Value As basetask)
End Property
Public Function Tick() As Boolean

End Function

