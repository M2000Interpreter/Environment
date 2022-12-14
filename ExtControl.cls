VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents mycontrol As VBControlExtender
Attribute mycontrol.VB_VarHelpID = -1
Dim another As Object
Dim myname$, mytypedef$
Dim Callback As GuiM2000
Dim mIndex As Long
Friend Property Get GetCallBack() As GuiM2000
Set GetCallBack = Callback
End Property
Public Property Get Index() As Long
Index = mIndex
End Property
Friend Property Let Index(ByVal RHS As Long)
mIndex = RHS
End Property
Friend Property Get ControlName() As String
ControlName = myname
End Property
Friend Property Get TypeDef() As String
TypeDef = mytypedef$
End Property
Friend Sub Attach(b As Control, aName$, typ$, mform As GuiM2000, Optional indx As Long = -1)
Err.Clear
On Error Resume Next
Set mycontrol = b
If Err.Number Then
Err.Clear
Set another = b
End If
mIndex = indx
Set Callback = mform
myname$ = aName$
End Sub
Property Get FixEvent() As Boolean
FixEvent = mycontrol Is Nothing
End Property
Property Get Value() As Object
Attribute Value.VB_UserMemId = 0
On Error Resume Next
If Not another Is Nothing Then
Set Value = another
Else
Set Value = mycontrol
End If
End Property
Public Sub deconstruct()
Set Callback = Nothing
Set mycontrol = Nothing
Set another = Nothing
End Sub

Private Sub mycontrol_ObjectEvent(Info As EventInfo)
Dim Values(), m As Long
If Info.EventParameters.Count > 0 Then
ReDim Values(0 To Info.EventParameters.Count)
Dim evinf As EventParameter
m = 0
For Each evinf In Info.EventParameters
    Values(m) = evinf.Value
    m = m + 1
Next evinf
If mIndex <> -1 Then
    Callback.CallbackNow myname$ + "." + Info.Name + "(" & mIndex & ")", Values()
Else
    Callback.CallbackNow myname$ + "." + Info.Name + "()", Values()
End If
With Info.EventParameters
For m = m - 1 To 0
    .item(m).Value = Values(m)
Next m
End With
ElseIf mIndex <> -1 Then
    Callback.Callback myname$ + "." + Info.Name + "(" & mIndex & ")"
Else
    Callback.Callback myname$ + "." + Info.Name + "()"
End If
End Sub