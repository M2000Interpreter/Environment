VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public mItem As Variant
Public slot As Byte
Public mTypeName As Byte
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Friend Sub GrabItem(ByRef RHS)
    If MemInt(VarPtr(mItem)) = 36 Then
        SwapVariant mItem, RHS
    ElseIf slot = 1 Then
        Set RHS = mItem
    ElseIf slot = 4 Then
        SwapVariant mItem, RHS
    Else
        RHS = mItem
    End If
End Sub
Property Let Typename(ByRef RHS As String)
    mTypeName = Asc(RHS)
End Property
Property Get Typename() As String
    Typename = Chr(mTypeName)
End Property
Property Let ItemStrVar(ByRef RHS As Variant)
    slot = 4
    mItem = RHS
End Property
Property Let ItemStr(ByRef RHS As String)
    slot = 4
    MoveStringToVariant RHS, mItem
End Property
Public Sub PopString(S$)
    If slot = 4 Then
        If Not myVarType(mItem, vbString) Then
            mItem = 0
            S$ = vbNullString
        Else
            SwapString2Variant S$, mItem
        End If
    ElseIf slot = 5 Then
        S$ = mItem
    ElseIf slot = 2 Then
        If VarType(mItem) = vbString Then
            S$ = mItem
        Else
            S$ = vbNullString
        End If
    Else
        S$ = vbNullString
    End If
End Sub
Property Get ItemStr() As String
    If slot > 3 Then ItemStr = mItem
End Property
Property Let ItemDouble(ByVal RHS As Variant)
    slot = 2
    mItem = RHS
End Property

Property Get ItemDouble() As Variant
    If slot = 2 Then ItemDouble = mItem
End Property
Property Let ItemVariant(RHS As Variant)
    slot = 2
    mItem = RHS
End Property
Property Get ItemVariantOnly() As Variant
    If slot = 1 Then
        Set ItemVariantOnly = mItem
    Else
        ItemVariantOnly = mItem
    End If
End Property

Property Get ItemVariant() As Variant
    If slot = 2 Then ItemVariant = mItem
End Property

Property Get ItemNoObject() As Variant
    If slot <> 1 Then ItemNoObject = mItem
End Property
Property Get NoValue() As Boolean
    If slot = 1 Then NoValue = mItem Is Nothing Else NoValue = slot = 0
End Property
Public Property Set ObjectRef(RHS As Variant)
    slot = 1
    Set mItem = RHS
End Property
Public Property Get ObjectRef() As Variant
    If slot = 1 Then Set ObjectRef = mItem Else Set ObjectRef = Nothing
End Property
Public Sub Recycle()
    Dim obj As IUnknown, GetRefCount As Long
    Set obj = Me
    CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
    If GetRefCount > 4 Then Exit Sub
    On Error Resume Next
    mItem = vbEmpty
    slot = 0
    mTypeName = 0
    TrushCount = TrushCount + 1
    If TrushCount > UBound(Trush) Then
        ReDim Preserve Trush(UBound(Trush) * 2)
    End If
    Set Trush(TrushCount) = Me
End Sub

