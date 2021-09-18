VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlMono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ver 1.1
Dim oXml As XmlMonoInternal

Friend Sub createTree(t As XmlMonoInternal)
Set oXml = t
End Sub
Friend Sub getTree(t As XmlMonoInternal)
Set t = oXml
End Sub
' status as property
Property Get status()
    status = oXml.status
End Property
Property Let beautify(b As Long)
    oXml.beautify = b
End Property
Property Get beautify() As Long
    beautify = oXml.beautify
End Property
Function HasNode() As Long
If oXml Is Nothing Then
    HasNode = IsArray(node)
Else
    HasNode = True
End If
End Function
Property Get CountProcess() As Long
    CountProcess = oXml.CountProcess
End Property
Property Get ProcessInstruction(i As Long) As XmlNode
    Dim nod, m As New XmlNode
    nod = oXml.ProcessInstruction(i)
    m.putnode nod
    Set ProcessInstruction = m
End Property

Sub PlaceProcessingInstructions(m As XmlNode)
    Dim nod
    m.getnode nod
    oXml.PlaceProcessingInstructions nod
End Sub
 Function PrepareNodeSimple(tagname As String) As XmlNode
    Dim m As New XmlNode
    m.CreateNodeArray Array(tagname, Nothing, Nothing, CVar(0), vbNullString, vbNullString)
    Set PrepareNodeSimple = m
 End Function
Function PrepareNode(tagname As String, Optional val$, Optional cdata As Boolean = False) As XmlNode
    Dim m As New XmlNode
    m.CreateNodeArray Array(tagname, Nothing, Nothing, CVar(1 - 2 * cdata), val$, vbNullString)

    Set PrepareNode = m
End Function
Function AlterProcessAttr(id, t$, NewValue As String, Optional OldValue As String) As Boolean
    AlterProcessAttr = oXml.AlterProcessAttr(id, t$, NewValue, OldValue)
End Function

Sub PlaceAttribute(attrname$, attrvalue$, Optional usesingle As Boolean = False)
    oXml.PlaceAttribute attrname$, attrvalue$, usesingle
End Sub
Sub PlaceAttributeToNode(xnode As XmlNode, attrname$, attrvalue$, Optional usesingle As Boolean = False)
    Dim node
    xnode.getnode node
    oXml.PlaceAttributeToNode node, attrname$, attrvalue$, usesingle
    xnode.putnode node
End Sub
Sub insertnode(xnode As XmlNode)
    Dim node
    xnode.getnode node
    oXml.insertnode node
End Sub
Sub AppendChild(xnode As XmlNode)
    Dim node
    xnode.getnode node
    oXml.AppendChild node
End Sub
Sub AppendChildObj(RHS As XmlMono)
    Dim m As XmlMonoInternal
    RHS.getTree m
    oXml.AppendChildObj m
End Sub
Property Get FirstChild() As XmlMono
    Dim m As XmlMonoInternal
    Set m = oXml.FirstChild
    If Not m Is Nothing Then
        Dim zz As New XmlMono
        zz.createTree m
        Set FirstChild = zz
    End If
End Property
Property Get lastchild() As XmlMono
    Dim m As XmlMonoInternal
    Set m = oXml.lastchild
    If Not m Is Nothing Then
        Dim zz As New XmlMono
        zz.createTree m
        Set lastchild = zz
    End If
End Property
Property Get Nextchild() As XmlMono
    Dim m As XmlMonoInternal
    Set m = oXml.Nextchild
    If Not m Is Nothing Then
        Dim zz As New XmlMono
        zz.createTree m
        Set Nextchild = zz
    End If
End Property
Property Get Anychild(p As Long) As XmlMono
    Dim m As XmlMonoInternal
    Set m = oXml.Anychild(p)
    If Not m Is Nothing Then
        Dim zz As New XmlMono
        zz.createTree m
        Set Anychild = zz
    End If
End Property
Function ScanForTag(Tag$, ok) As XmlMono
    Dim m As XmlMonoInternal
    Set m = oXml.ScanForTag(Tag$, ok)
    If Not m Is Nothing Then
        Dim zz As New XmlMono
        zz.createTree m
        Set ScanForTag = zz
    End If
End Function
Property Get textFromChildDeep(Tag$) As String
    textFromChildDeep = oXml.textFromChildDeep(Tag$)
End Property
Function AttrFromTagDeep(Tag$, attr$) As String
    AttrFromTagDeep = oXml.AttrFromTagDeep(Tag$, attr$)
End Function
Function AttrFromChild(Tag$, attr$) As String
    AttrFromChild = oXml.AttrFromChild(Tag$, attr$)
End Function
Property Get textFromChild(Tag$) As String
textFromChild = oXml.textFromChild(Tag)
End Property
Function ScanForTagDeep(Tag$, child As XmlMono) As Boolean
    Dim m As XmlMonoInternal
    If oXml.ScanForTagDeep(Tag$, m) Then
        If Not m Is Nothing Then
            Dim zz As New XmlMono
            zz.createTree m
            Set child = zz
            ScanForTagDeep = True
        End If
    End If
End Function
Function GetListByTag(Tag$, Optional many As Long = -1) As mHandler
    Dim v As XmlMono, myCollection As New mStiva
    Dim ret As New mHandler
    ret.t1 = 3
    Set ret.objref = myCollection
    For Each v In oXml.GetListByTag(Tag$, many, True)
        myCollection.DataObj v
    Next
    Set GetListByTag = ret
End Function

Sub ResetChildPointer()
    oXml.ResetChildPointer
End Sub

Property Get ChildCount() As Long
    ChildCount = oXml.ChildCount
End Property
Property Get node(Optional Remove As Boolean = False) As XmlNode
    Dim m As New XmlNode
    m.putnode oXml.node(Remove)
    Set node = m
End Property
Property Get Text() As String
Text = oXml.Text
End Property
Property Let Text(t$)
    oXml.Text = t$
End Property
Property Let cdata(t$)
    oXml.cdata = t$
End Property
Function EndOffChilds(child As XmlMono) As Boolean
    Dim m As XmlMonoInternal
    Set child = New XmlMono
   If Not oXml.EndOffChilds(m) Then
   child.createTree m: EndOffChilds = True
   End If
    End Function
Property Get Tag() As String
On Error Resume Next

    Tag = oXml.Tag
End Property
Property Get Xml() As String
    Xml = oXml.Xml(oXml)
End Property
Property Let attr(t$, RHS As String)
    oXml.attr(t$) = RHS
End Property
Property Get attr(t$) As String
    attr = oXml.attr(t$)
End Property
Property Let Xml(t$)
' use original
    oXml.Xml(oXml) = t$
End Property
Sub RemoveChild()
    oXml.RemoveChild
End Sub
Function AttrExist(t$, Optional attr As String) As Boolean
    AttrExist = oXml.AttrExist(t$, attr)
End Function
Sub RemoveChildren()
    oXml.RemoveChildren
End Sub
Sub RemoveAttributes()
    oXml.RemoveAttributes
End Sub
Sub RemoveProcess()
    oXml.RemoveProcess
End Sub

