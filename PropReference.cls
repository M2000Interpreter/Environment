VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ObjectRef As Long
Dim vtref As Long
Dim mIndex As Variant
Public UseIndex As Boolean
Dim Init As Boolean
Dim IamMethod  As Boolean
Dim hardlink As Object, useobj As Boolean
Private mLastObj As Object
Public Arr As Boolean
Dim oEnum As IUnknown
Private IndexList As mIndexes
Private IndexNew As Boolean
Function isObj()
    isObj = useobj
End Function
Property Get ObjectType() As String
On Error Resume Next
If Not hardlink Is Nothing Then
    ObjectType = "*" + Typename(ObjectRef)
Else
    ObjectType = "[" + Typename(var(ObjectRef)) + "]"
End If
End Property
Sub Construct(ref As Long, vtrefnumber As Long, Optional indirect As Boolean = False)
If useobj Then Exit Sub
Init = True

ObjectRef = ref
If indirect Then
vtref = -vtrefnumber
Else
vtref = vtrefnumber
End If
End Sub
Sub ConstructObj(obj As Object, vtrefnumber As Long)
Init = True
Set hardlink = obj
useobj = True
'ObjectRef = ref
vtref = vtrefnumber
End Sub
Public Property Get lastobj() As Object
Set lastobj = mLastObj
End Property
Public Property Get lastobjfinal() As Object
Set lastobjfinal = mLastObj
Set mLastObj = Nothing
End Property
Public Sub clearlastobject()
Set mLastObj = Nothing
End Sub
Public Property Get value() As Variant
Dim v
Dim RETVAR As Variant, retvar2, obj As Object
Dim check As Boolean
Set mLastObj = Nothing
If Init Then
If UseIndex Then
If Not IndexNew Then
 value = 0
Set mLastObj = Me
Exit Property
ElseIf useobj Then
    IndexNew = False
    value = ReadPropIndexObj(hardlink, vtref, mIndex)
Else
    IndexNew = False
    value = ReadPropIndex(ObjectRef, vtref, mIndex, check, RETVAR)
    If check And IsObject(RETVAR) Then
    On Error Resume Next
    Set obj = RETVAR
    If ReadOneParameter(obj, (0), vbNullString, retvar2) Then
        value = retvar2
       Else
          Set mLastObj = RETVAR
    End If

 End If
End If

Else


If useobj Then
If ReadPropObj(hardlink, vtref, RETVAR) Then
If IsObject(RETVAR) Then
 value = 0
          Set mLastObj = RETVAR

Else
value = RETVAR
End If
End If
Else
 If ReadProp(ObjectRef, vtref, RETVAR) Then
 If IsObject(RETVAR) Then
         value = 0
          Set mLastObj = RETVAR

Else
On Error Resume Next
If TypeOf RETVAR Is IUnknown Then
value = 0
Set mLastObj = RETVAR: Exit Property
End If
Err.Clear
value = RETVAR
End If
 End If
 End If
 End If
Else
' ??????????

End If
'If Not IsObject(Value) Then Debug.Print Value
End Property
Public Property Let value(vNewValue As Variant)
If Init Then
If Not UseIndex Then
If useobj Then
WritePropObj hardlink, vtref, vNewValue
Else
WriteProp ObjectRef, vtref, vNewValue
End If
ElseIf Not IndexNew Then
    MyEr "Missing index", "Έχασα τον δείκτη"
    Exit Property
ElseIf useobj Then
IndexNew = False
WritePropIndexObj hardlink, vtref, vNewValue, mIndex
Else
IndexNew = False
WritePropIndex ObjectRef, vtref, vNewValue, mIndex
End If
End If

End Property
Public Property Let ValueStr(s As String)
Dim vNewValue
vNewValue = vbNullString
SwapString2Variant s, vNewValue
If Init Then
If Not UseIndex Then
If useobj Then
    WritePropObj hardlink, vtref, vNewValue
Else
    WriteProp ObjectRef, vtref, vNewValue
End If
ElseIf Not IndexNew Then
' do nothing
    MyEr "Missing index", "Έχασα τον δείκτη"
    Exit Property
ElseIf useobj Then
    IndexNew = False
    WritePropIndexObj hardlink, vtref, vNewValue, mIndex
Else
    IndexNew = False
    WritePropIndex ObjectRef, vtref, vNewValue, mIndex
End If
End If


End Property
Public Property Get isMethod() As Boolean
isMethod = IamMethod
End Property

Public Property Let isMethod(ByVal vNewValue As Boolean)
IamMethod = vNewValue
End Property

Private Sub Class_Terminate()
Set hardlink = Nothing
End Sub
Public Sub IndexOpt()
IndexNew = True
mIndex = 0
OptVariant mIndex
End Sub
Public Property Get Index() As Variant
Index = mIndex
End Property
Public Sub ArrayIndex(ParamArray RHS())
'IndexNew = True
    mIndex = RHS
End Sub

Public Property Let Index(ByVal RHS As Variant)
IndexNew = True
mIndex = RHS
End Property
Friend Sub PushIndexes(idx As mIndexes)
If idx Is Nothing Then Exit Sub
IndexNew = True
UseIndex = True
If IndexList Is Nothing Then
    If idx.count = 1 Then
       idx.Dump 0, mIndex
    Else
    idx.GetArr mIndex
    End If
End If
End Sub
Sub IndexAgain()
    IndexNew = True
End Sub
