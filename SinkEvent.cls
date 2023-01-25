VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SinkEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event fire()
Public Event GetString(a$)
Public Event GetVariantRef(ByRef a As Variant)
Public Event GetValueRef(a As Variant)
Public Event GetData()
Private resp()
'Mark as default member via Tools|Procedure Attributes... dialog.
Public Sub FireMe()
Attribute FireMe.VB_UserMemId = 0
    RaiseEvent fire
End Sub
Public Sub SendString(a)


    On Error Resume Next
    Dim b$
    If Not myVarType(a, vbString) Then a = vbNullString
    SwapString2Variant b$, a
    RaiseEvent GetString(b$)
    SwapString2Variant b$, a
End Sub
Public Sub SendVariantRef(ByRef a As Variant)
    On Error Resume Next
    Dim k
  '  k = a
    RaiseEvent GetVariantRef(a)
'    a = k
End Sub
Public Sub SendValueRef(a)
    On Error Resume Next
    RaiseEvent GetValueRef(a)
End Sub
Public Sub SendData(ParamArray flds())
    resp() = flds()
    RaiseEvent GetData
End Sub
Public Function GetData() As Object
Dim m As New mArray, H As New mHandler
    m.LoadTuple resp
    H.t1 = 3
    Set H.objref = m
    Set GetData = H
End Function
