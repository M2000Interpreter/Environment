VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Constant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mValue As Variant
Private mflag As Integer, mSubCall As Integer
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
If flag Then
Set Value = mValue
Else
Value = mValue
End If
End Property
Public Property Let Value(RHS As Variant)
MyEr "Constant value", "������� ����"
Err.Raise 20
End Property
Public Property Set Value(RHS As Variant)
MyEr "Constant value", "������� ����"
Err.Raise 20
End Property
Public Property Get flag() As Boolean
    flag = (mflag And 1) <> 0
End Property
Public Property Get vType() As Boolean
    vType = (mflag And 2) <> 0
End Property
Public Property Get SUBPord() As Boolean
    SUBPord = mSubCall
End Property
Public Sub DefineOnce(RHS, Optional ByVal vType, Optional fromsubcall As Boolean)
mSubCall = fromsubcall
If myVarType(mValue, vbEmpty) Then
If MyIsObject(RHS) Then
    mflag = mflag Or 1
    Set mValue = RHS
Else
    If Not IsMissing(vType) Then
    mflag = mflag Or 2
    End If
    mValue = RHS
End If
End If
End Sub
