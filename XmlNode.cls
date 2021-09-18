VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Node class
Private node, iamobject As Boolean
Friend Sub CreateNodeArray(v)
    node = v
    iamobject = False
End Sub

Friend Sub putnode(ByRef v)
If IsObject(v) Then
Set node = v
iamobject = True
Else
node = v
iamobject = False
End If
End Sub
Friend Sub getnode(ByRef v)
If iamobject Then
    Set v = node
Else
    v = node
End If
End Sub
