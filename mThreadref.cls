VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mThreadref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mReference As TaskInterface
Private mThreadVar As String
Private mThreadID As Long

 Property Let VarName(ByRef RHS As String)
mThreadVar = RHS
End Property

 Property Get VarName() As String
VarName = mThreadVar
End Property
 Property Set thread(ByRef RHS As TaskInterface)
Set mReference = RHS
End Property

 Property Get thread() As TaskInterface
Set thread = mReference
End Property
 Property Let Id(ByVal RHS As Long)
mThreadID = RHS
End Property

Property Get Id() As Long
Id = mThreadID
End Property
