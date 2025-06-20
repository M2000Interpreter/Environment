VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlMonoInternal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ver 1.1
Private root
Public status As String
Public beautify As Long
Public skip As Boolean
' this is false by default
Public NumericCharacterEntities As Boolean
Private ProcessBag As Collection
Private achild As Long, thelast As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Function HasNode() As Boolean
HasNode = IsArray(root)
End Function
Private Function packtext(t1$) As Variant
    Dim t$
    t$ = Replace$(t1$, "&", "&amp;")
    t$ = Replace$(t$, Chr(34), "&quot;")
    t$ = Replace$(t$, "'", "&apos;")
    t$ = Replace$(t$, "<", "&lt;")
    t$ = Replace$(t$, ">", "&qt;")
    
    packtext = CVar(t$)
End Function
Private Function packattrib(t1$, usesingle As Boolean) As Variant
    Dim t$
    t$ = Replace$(t1$, "&", "&amp;")
    If Not usesingle Then
        t$ = Replace$(t$, Chr(34), "&quot;")
    Else
        t$ = Replace$(t$, "'", "&apos;")
    End If
    t$ = Replace$(t$, "<", "&lt;")
    t$ = Replace$(t$, ">", "&qt;")
    If usesingle Then
    packattrib = CVar("'" + t$ + "'")
    Else
    packattrib = CVar(Chr$(34) + t$ + Chr$(34))
    End If
End Function
Private Function unpackattrib(t1 As Variant) As String
    Dim t$, usesingle As Boolean
    t$ = t1
    usesingle = Left(t$, 1) = "'"
    If Not usesingle Then
    t$ = Replace$(t$, "&quot;", Chr(34))
    Else
    t$ = Replace$(t$, "&apos;", "'")
    End If
    t$ = Replace$(t$, "&lt;", "<")
    t$ = Replace$(t$, "&qt;", ">")
    t$ = Mid$(Replace$(t$, "&amp;", "&"), 2)
    unpackattrib = Left$(t$, Len(t$) - 1)
        If NumericCharacterEntities Then
        convertNumericEntities unpackattrib
    End If
End Function

Private Function unpacktext(t1 As Variant) As String
    Dim t$
    t$ = t1
    t$ = Replace$(t$, "&quot;", Chr(34))
    t$ = Replace$(t$, "&apos;", "'")
    t$ = Replace$(t$, "&lt;", "<")
    t$ = Replace$(t$, "&qt;", ">")
    unpacktext = Replace$(t$, "&amp;", "&")
    If NumericCharacterEntities Then
        convertNumericEntities unpacktext
    End If
End Function
Private Sub convertNumericEntities(a$)
Dim b() As String, i As Long, j As Long
b = split(a$, "&#")
If UBound(b) <= 0 Then Exit Sub
For i = 1 To UBound(b)
j = InStr(b(i), ";")
Select Case Left$(b(i), 1)
Case "x"
If j - 2 > 0 Then
    b(i) = ChrW(val("&h" + Mid$(b(i), 2, j - 2))) + Mid$(b(i), j + 1)
    End If
Case Else
    b(i) = ChrW(val(Left$(b(i), j - 1))) + Mid$(b(i), j + 1)
End Select
Next
a$ = Join(b, "")
End Sub
Sub PlaceCdata()
' <![CDATA["'<>&]]>  ' no escaped
End Sub
Sub PlaceProcessingInstructions(node)
If IsArray(node) Then
    If node(3) = True Or Not node(2) Is Nothing Then
        MyEr "Not childs for processing instructions", "��� ������ ��� ������� �������"
        Exit Sub
    End If
    If ProcessBag Is Nothing Then
    Set ProcessBag = New Collection
    ProcessBag.Add node
    Else
    ProcessBag.Add node, , , ProcessBag.Count
    End If
Else
    MyEr "Not a valid node", "��� ������� ������"
End If
End Sub
Function PrepareNodeSimple(tagname As String)
    PrepareNodeSimple = Array(tagname, Nothing, Nothing, CVar(0), vbNullString, vbNullString)
End Function
Function PrepareNode(tagname As String, Optional val$, Optional cdata As Boolean = False) As Variant()
    PrepareNode = Array(tagname, Nothing, Nothing, CVar(1 - 2 * cdata), val$, vbNullString)
End Function
Sub PlaceAttribute(attrname$, attrvalue$, Optional usesingle As Boolean = False)
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
Else
Dim c As Collection, p
If root(1) Is Nothing Then
    Set c = New Collection
    p = packattrib(attrvalue$, usesingle)
    c.Add p, attrname$
    Set root(1) = c
Else
    p = packattrib(attrvalue$, usesingle)
    Set c = root(1)
    On Error Resume Next
    c.Add p, attrname$
    If Err.Number <> 0 Then MyEr "Attribute exist", "� �������� �������"
End If
End If
End Sub
Sub PlaceAttributeToNode(node, attrname$, attrvalue$, Optional usesingle As Boolean = False)
If Not IsArray(node) Then
    MyEr "This isn't a node", "��� ������ �����"
Else
Dim c As Collection, p
If node(1) Is Nothing Then
    Set c = New Collection
    p = packattrib(attrvalue$, usesingle)
    node(5) = " " + attrname$ + " "
    c.Add p, attrname$
    Set node(1) = c
Else
    p = packattrib(attrvalue$, usesingle)
    
    Set c = node(1)
    On Error Resume Next
    
    c.Add p, attrname$
    If Err.Number <> 0 Then MyEr "Attribute exist", "� �������� �������": Exit Sub
    node(5) = node(5) + attrname$ + " "
End If
End If
End Sub
Sub insertnode(node)
' tuple
If Not IsArray(root) Then
    Dim t$
    t$ = node(4)
    If (node(3) And 2) = 2 Then
        root = Array(node(0), node(1), node(2), node(3), "<![CDATA[" + t$ + "]]>", node(5))
    Else
        root = Array(node(0), node(1), node(2), node(3), packtext(t$), node(5))
    End If
Else
    MyEr "I have node", "��� �����"
End If
End Sub
Sub AppendChild(node)
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (node(3) And &H1) Then
    Dim val
    Set val = root(2)
    Dim p As XmlMonoInternal, c As Collection
    Set p = New XmlMonoInternal
    p.NumericCharacterEntities = NumericCharacterEntities
    p.insertnode node
    If val Is Nothing Then
        Set c = New Collection
        c.Add p
        Set root(2) = c
    Else
        Set c = root(2)
        c.Add p, , , c.Count
    End If
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
End Sub
Public Sub AppendChildObj(p As XmlMonoInternal)
    root(3) = root(3) Or 1
    Dim val
    Set val = root(2)
    Dim c As Collection
    If val Is Nothing Then
        Set c = New Collection
        c.Add p
        Set root(2) = c
    Else
        Set c = root(2)
        c.Add p, , , c.Count
    End If
End Sub
Property Get FirstChild() As XmlMonoInternal
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And 1) Then
    Dim val, c As Collection
    Set val = root(2)
    If val Is Nothing Then
        MyEr "I have't childs yet", "��� ��� ������ �����"
    Else
        Set c = val
        Set FirstChild = c.item(1)
        thelast = achild = c.Count - 1
        achild = 0
    End If
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
End Property
Property Get lastchild() As XmlMonoInternal
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And 1) Then
    Dim val, c As Collection
    Set val = root(2)
    If val Is Nothing Then
        MyEr "I have't childs yet", "��� ��� ������ �����"
    Else
        Set c = val
        achild = c.Count - 1
        thelast = True
        Set lastchild = c.item(achild + 1)
    End If
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
End Property
Property Get Nextchild() As XmlMonoInternal
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And 1) Then
    Dim val, c As Collection
    Set val = root(2)
    If val Is Nothing Then
        MyEr "I have't childs yet", "��� ��� ������ �����"
    Else
        Set c = val
        
        achild = achild + 1
        If achild > c.Count - 1 Then
            MyEr "No more childs", "��� �������� ���� ������"
        Else
            thelast = achild = c.Count - 1
            Set Nextchild = c(achild + 1)
        End If
    End If
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
End Property
Property Get Anychild(p As Long) As XmlMonoInternal
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And 1) Then
    Dim val, c As Collection
    Set val = root(2)
    If val Is Nothing Then
        MyEr "I have't childs yet", "��� ��� ������ �����"
    Else
        Set c = val

        If p < 0 Or p > c.Count - 1 Then
            MyEr "index fault", "������� ����� �����"
        Else
            Set Anychild = c(p + 1)
        End If
    End If
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
End Property
Function ScanForTag(Tag$, ok) As XmlMonoInternal
ok = False
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And 1) Then
    Dim val, c As Collection, aXml As XmlMonoInternal
    Set val = root(2)
    If val Is Nothing Then
        MyEr "I have't childs yet", "��� ��� ������ �����"
    Else
        Set c = val
        Dim i As Long
        For i = 1 To c.Count
        Set aXml = c(i)
        If aXml.Tag = Tag$ Then Set ScanForTag = aXml: ok = True: Exit Function
        Next i


    End If
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
Set ScanForTag = New XmlMonoInternal
ScanForTag.NumericCharacterEntities = NumericCharacterEntities
End Function
Property Get textFromChildDeep(Tag$) As String
Dim tmp As XmlMonoInternal
If ScanForTagDeep(Tag$, tmp) Then
    textFromChildDeep = tmp.Text
Else
    MyEr "Tag not exist", "��� ������� �� ��������"
End If
End Property
Function AttrFromTagDeep(Tag$, attr$) As String
Dim tmp As XmlMonoInternal
If ScanForTagDeep(Tag$, tmp) Then
    AttrFromTagDeep = tmp.attr(attr$)
Else
    MyEr "Tag not exist", "��� ������� �� ��������"
End If
End Function
Function AttrFromChild(Tag$, attr$) As String
Dim tmp As XmlMonoInternal, ok
Set tmp = ScanForTag(Tag$, ok)
If ok Then
    AttrFromChild = tmp.Text
Else
    MyEr "Tag not exist", "��� ������� �� ��������"
End If

End Function
Property Get textFromChild(Tag$) As String
Dim tmp As XmlMonoInternal, ok
Set tmp = ScanForTag(Tag$, ok)
If ok Then
    textFromChild = tmp.Text
Else
    MyEr "Tag not exist", "��� ������� �� ��������"
End If

End Property
Function ScanForTagDeep(Tag$, child As XmlMonoInternal) As Boolean
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And 1) Then
    beautify = CLng(100000 * Rnd) + 1
    If deep(Tag$, child, beautify) Then
        beautify = 0
        ScanForTagDeep = True
        Exit Function
    End If
    beautify = 0
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
Set child = New XmlMonoInternal
child.NumericCharacterEntities = NumericCharacterEntities
End Function
Friend Function deep(Tag$, child As XmlMonoInternal, guard As Long) As Boolean
   Dim val, c As Collection, aXml As XmlMonoInternal
    Set val = root(2)
    If val Is Nothing Then
        MyEr "I have't childs yet", "��� ��� ������ �����"
    Else
        Set c = val
        Dim i As Long
        For i = 1 To c.Count
        Set aXml = c(i)
        If aXml.beautify = guard Then Exit Function ' safety
        If aXml.Tag = Tag$ Then Set child = aXml: deep = True: Exit Function
                    Dim p
            p = aXml.node
            If (p(3) And 1) = 1 Then
        If aXml.ChildCount > 0 Then
            If aXml.deep(Tag$, child, guard) Then deep = True: Exit Function
        End If
        End If
        Next i
    End If
End Function
Friend Sub adddeep(Tag$, bag As Collection, Optional many = -1)
   Dim val, c As Collection, aXml As XmlMonoInternal
    Set val = root(2)
    If val Is Nothing Then
        MyEr "I have't childs yet", "��� ��� ������ �����"
    Else
        Set c = val
        Dim i As Long
        For i = 1 To c.Count
        Set aXml = c(i)
            If aXml.Tag = Tag$ Then
                If bag.Count = 0 Then
                    bag.Add aXml
                Else
                    bag.Add aXml, , , bag.Count
                End If
                If many > 0 Then many = many - 1
                If many = 0 Then Exit Sub
            End If
            Dim p
            p = aXml.node
            If (p(3) And 1) = 1 Then
            If aXml.ChildCount > 0 Then
                aXml.adddeep Tag$, bag, many
                If many = 0 Then Exit Sub
            End If
            End If
        Next i
    End If
End Sub
Friend Sub adddeepexport(Tag$, bag As Collection, Optional many = -1)
   Dim val, c As Collection, aXml As XmlMonoInternal, bXml As XmlMono
    Set val = root(2)
    If val Is Nothing Then
        MyEr "I have't childs yet", "��� ��� ������ �����"
    Else
        Set c = val
        Dim i As Long
        For i = 1 To c.Count
        Set aXml = c(i)
            If aXml.Tag = Tag$ Then
                Set bXml = New XmlMono
                bXml.createTree aXml
                If bag.Count = 0 Then
                    bag.Add bXml
                Else
                    bag.Add bXml, , , bag.Count
                End If
                If many > 0 Then many = many - 1
                If many = 0 Then Exit Sub
            End If
            Dim p
            p = aXml.node
            If (p(3) And 1) = 1 Then
            If aXml.ChildCount > 0 Then
                aXml.adddeepexport Tag$, bag, many
                If many = 0 Then Exit Sub
            End If
            End If
        Next i
    End If
End Sub
Function GetListByTag(Tag$, Optional many As Long = -1, Optional export As Boolean = False) As Collection
Dim mList As New Collection
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And 1) Then
    If export Then
        adddeepexport Tag$, mList, many
    Else
        adddeep Tag$, mList, many
    End If
Else
  '  MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
Set GetListByTag = mList
End Function

Sub ResetChildPointer()
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
Else
achild = -1
If root(2) Is Nothing Then
thelast = True
Else
thelast = False
End If
End If
End Sub

Property Get ChildCount() As Long
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And 1) Then
    Dim val, c As Collection
    Set val = root(2)
    If val Is Nothing Then
        ChildCount = 0
    Else
        Set c = val

        ChildCount = c.Count
      
    End If
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
End Property
Property Get node(Optional Remove As Boolean = False) As Variant()
    If Not IsArray(root) Then
        MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        node = root
        If Remove Then root = Empty
    End If
End Property
Property Get Text() As String
    If Not IsArray(root) Then
        MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        If (root(3) And 2) = 2 Then
            Text = Mid$(root(4), 10, Len(root(4)) - 12)
        Else
            Text = unpacktext(root(4))
        End If
    End If
End Property
Property Let Text(t$)
    If Not IsArray(root) Then
        MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        If (root(3) And 2) = 2 Then
            root(4) = "<![CDATA[" + t$ + "]]>"
        Else
            root(4) = packtext(t$)
        End If
    End If
End Property
Property Let cdata(t$)
    If Not IsArray(root) Then
        MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        root(3) = root(3) Or 2
        root(4) = "<![CDATA[" + t$ + "]]>"
    End If
End Property
Function EndOffChilds(child As XmlMonoInternal) As Boolean
EndOffChilds = False
If Not thelast Then
    Set child = Nextchild
    child.ResetChildPointer
Else
EndOffChilds = True
End If
End Function
Property Get Tag() As String
    If Not IsArray(root) Then
        MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        Tag = root(0)
    End If
End Property
Private Function AnyNodeAttr(node) As String
    Dim c As Collection, M() As String, i As Long
        If Not node(1) Is Nothing Then
            Set c = node(1)
            M() = split(node(5))
            For i = 1 To c.Count
                AnyNodeAttr = AnyNodeAttr + " " + M(i) + "=" + c(M(i))
            Next i
        End If
End Function

Friend Property Get Xml(guard As XmlMonoInternal) As String
On Error GoTo 1000
status = ""
If skip Then
    beautify = 0
Else
    If Not skip And beautify = 0 Then skip = True
End If
    If Not IsArray(root) Then
     '   MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        Dim i As Long, c As Collection, p, j As Long, D As XmlMonoInternal, t$
        Dim id$
        
        If beautify > 0 Then id$ = space(beautify)
        Xml = id$
        If Not ProcessBag Is Nothing Then
        For i = 1 To ProcessBag.Count
        p = ProcessBag(i)
        If p(1) Is Nothing And Left$(p(0), 1) = "!" Then
        Xml = Xml + "<" + p(5)
        Else
        Xml = Xml + "<?" + p(0) + AnyNodeAttr(p) + "?>"
        End If
        If beautify Then Xml = Xml + vbCrLf + id$
        Next i
        End If
        If (root(3) And 1) = 0 Then
            Xml = Xml + "<" + root(0) + AnyNodeAttr(root) + "/>"
            If beautify Then Xml = Xml + vbCrLf
        Else
            Xml = Xml + "<" + root(0) + AnyNodeAttr(root) + ">"
            If Not root(2) Is Nothing Then
                    Set c = root(2)
                    If beautify Then Xml = Xml + vbCrLf
                    For i = 1 To c.Count
                    Set D = c(i)
                    If D Is guard Then Set D = Nothing: Exit Property
                    D.skip = skip
                    If beautify < 0 Then
                        D.beautify = -beautify
                        Xml = Xml + D.Xml(guard)
                    Else
                      D.beautify = beautify + Abs(guard.beautify)
                         Xml = Xml + D.Xml(guard)
                    End If
                    If guard Is Nothing Then Set root(2) = Nothing ' weave untik top
                    If Len(D.status) > 0 Then status = "<" + root(0) + " child='" + Trim$(str$(i)) + "'>" + D.status + "</" + root(0) + ">": GoTo 1000
                    Next i
                    Xml = Xml + id$
            Else
                    Xml = Xml + root(4)
            End If
            Xml = Xml + "</" + root(0) + ">"
            If beautify Then Xml = Xml + vbCrLf
        End If
    End If
1000:    beautify = 0
        skip = False
End Property

Property Get attr(t$) As String
    If Not IsArray(root) Then
        MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        Dim c As Collection
        If root(1) Is Nothing Then NoSuchAttr t$: Exit Property
        Set c = root(1)
        On Error Resume Next
        attr = unpackattrib(c.item(t$))
        If Err.Number > 0 Then Err.Clear: NoSuchAttr t$
    End If
End Property
Property Let attr(t$, RHS As String)
    If Not IsArray(root) Then
        MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        Dim c As Collection
        If root(1) Is Nothing Then NoSuchAttr t$: Exit Property
        Set c = root(1)
        On Error Resume Next
        c.Remove t$
        c.Add CVar(packattrib(RHS, InStr(RHS, "'") > 0)), t$
        If Err.Number > 0 Then Err.Clear: NoSuchAttr t$
    End If
End Property
Function AttrExist(t$, Optional attr As String) As Boolean
    If Not IsArray(root) Then
        MyEr "node not inserted yet", "��� ���� �������� ������"
    Else
        Dim c As Collection
        If root(1) Is Nothing Then NoSuchAttr t$: Exit Function
        Set c = root(1)
        On Error Resume Next
        attr = unpackattrib(c.item(t$))
        AttrExist = Err.Number = 0
    End If
End Function
Function AlterProcessAttr(id, t$, NewValue As String, Optional OldValue As String) As Boolean
If ProcessBag.Count < id Or id < 1 Then MyEr "index fault", "������� ����� �����": Exit Function
Dim node, i As Long, c As Collection
node = ProcessBag.item(id)
If Typename(node(1)) = "Nothing" Then
    i = InStrRev(node(5), ">")
    OldValue = Mid$(node(5), i - 1)
    node(5) = NewValue + Mid$(node(5), i)
Else
On Error GoTo er1
Set c = node(1)
If c.Count = 0 Then GoTo er1
OldValue = c(t$)
c.Remove t$
c.Add CVar(NewValue), t$
End If
Exit Function
er1:
NoSuchAttr t$
End Function
Property Get CountProcess() As Long
    If ProcessBag Is Nothing Then
        CountProcess = 0
    Else
        CountProcess = ProcessBag.Count
    End If
End Property
Property Get ProcessInstruction(i As Long) As Variant()
If ProcessBag Is Nothing Then
    MyEr "there are no instructions in this node", "��� �������� ������� �� ����� �� �����"
ElseIf i < 0 Or i >= ProcessBag.Count Then
    MyEr "index fault", "������� ����� �����"
Else
    ProcessInstruction = ProcessBag.item(i)
End If
End Property
Private Sub MyEr1(a$, b$)
    status = a$
   
   ' Err.Raise 5000, , a$
End Sub
Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal VarPtr(b), 16
   CopyMemory ByVal VarPtr(b), t(0), 16
End Sub
Private Function Advance2Lt(t$, pc As Long, st As Long) As Boolean
Do While Mid$(t$, pc, 1) <> "<"
pc = pc + 1
If pc > st Then Exit Do
Loop
pc = pc + 1
Advance2Lt = pc <= st
End Function
Private Function Advance2Gt(t$, pc As Long, st As Long) As Boolean
Do While Mid$(t$, pc, 1) <> ">"
pc = pc + 1
If pc > st Then Exit Do
Loop
Advance2Gt = pc <= st
End Function
Private Function AdvanceAnySpaceorGT_NameSpace(t$, pc As Long, st As Long) As Boolean
Dim c$
c$ = Mid$(t$, pc, 1)
If c$ <> ">" And (c$ >= "A") And c$ <> "/" Then
    pc = pc + 1
    If Not pc > st Then
        c$ = Mid$(t$, pc, 1)
        Do While c$ <> ">" And (c$ >= "A" Or InStr(":1234567890.-_", c$) > 0) And c$ <> "/"
            pc = pc + 1
            If pc > st Then Exit Do
            c$ = Mid$(t$, pc, 1)
        Loop
    End If
End If
AdvanceAnySpaceorGT_NameSpace = pc <= st
End Function

Private Function AdvanceAnySpaceorGT(t$, pc As Long, st As Long) As Boolean
Dim c$
c$ = Mid$(t$, pc, 1)
Do While c$ <> ">" And c$ >= "A" And c$ <> "/"
pc = pc + 1
If pc > st Then Exit Do
c$ = Mid$(t$, pc, 1)
Loop
AdvanceAnySpaceorGT = pc <= st
End Function
Private Function AdvanceSpaceorGT(t$, pc As Long, st As Long) As Boolean
Dim c$
c$ = Mid$(t$, pc, 1)
Do While Not (c$ = ">" Or c$ > " ")
pc = pc + 1
If pc > st Then Exit Do
c$ = Mid$(t$, pc, 1)
Loop
AdvanceSpaceorGT = pc <= st And Mid$(t$, pc, 1) = ">"
End Function
Private Function AdvanceAnySpaceorLT(t$, pc As Long, st As Long) As Boolean
Dim c$
c$ = Mid$(t$, pc, 1)
Do While c$ <> "<" And c$ < "A"
pc = pc + 1
If pc > st Then Exit Do
c$ = Mid$(t$, pc, 1)
Loop
AdvanceAnySpaceorLT = pc <= st
End Function
Private Function AdvanceSpace(t$, pc As Long, st As Long) As Boolean
Dim c$
c$ = Mid$(t$, pc, 1)
Do While AscW(c$) < 33
pc = pc + 1
If pc > st Then Exit Do
c$ = Mid$(t$, pc, 1)
Loop
AdvanceSpace = pc <= st
End Function
Private Function AdvanceScope(t$, pc As Long, st As Long, scope As String) As Boolean
Dim c$
c$ = Mid$(t$, pc, 1)
Do While c$ <> scope
pc = pc + 1
If pc > st Then Exit Do
c$ = Mid$(t$, pc, 1)
Loop
AdvanceScope = pc <= st
End Function

Private Sub createprocessxml(t$, pc As Long, st As Long)
Dim n1 As Long, n2 As Long, Tag$, node, val$
pc = pc + 1
n1 = pc
If AdvanceAnySpaceorGT_NameSpace(t$, pc, st) Then
    Tag$ = Mid$(t$, n1, pc - n1)
    node = Me.PrepareNodeSimple(Tag$)
    If Mid$(t$, pc, 1) < "A" Then
lookfortagagain:
        If AdvanceSpace(t$, pc, st) Then
            n1 = pc

            If AdvanceAnySpaceorGT_NameSpace(t$, pc, st) Then
                Tag$ = Mid$(t$, n1, pc - n1)
                If Len(Tag$) > 0 Then
                    If AdvanceSpace(t$, pc, st) Then
                        If Mid$(t$, pc, 1) = "=" Then
                            pc = pc + 1
                            If AdvanceSpace(t$, pc, st) Then
                                    n2 = 0
                                If Mid$(t$, pc, 1) = Chr$(34) Then
                                    n2 = 34
                                ElseIf Mid$(t$, pc, 1) = Chr$(39) Then
                                    n2 = 39
                                End If
                                If n2 <> 0 Then
                                    pc = pc + 1
                                    n1 = pc
                                    If AdvanceScope(t$, pc, st, Chr$(n2)) Then
                                        val$ = Mid$(t$, n1, pc - n1)
                                        PlaceAttributeToNode node, Tag$, val$, n2 = 39
                                        pc = pc + 2

                                        GoTo lookfortagagain
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        PlaceProcessingInstructions node
        If Mid$(t$, pc, 2) = "?>" Then pc = pc + 1
        If Mid$(t$, pc, 1) <> ">" Then pc = pc + 1
    End If
End If
End Sub
Private Sub CheckThis(t$, pc As Long, st As Long)
Dim n1 As Long, n2 As Long, Tag$, node, val$
'
Dim n0 As Long
n0 = pc
n1 = pc
pc = pc + 1
If AdvanceAnySpaceorGT_NameSpace(t$, pc, st) Then
    Tag$ = Mid$(t$, n1, pc - n1)
    node = Me.PrepareNodeSimple(Tag$)
    node(5) = "!"
    If Mid$(t$, pc, 1) < "A" Then
lookfortagagain:
        If AdvanceSpace(t$, pc, st) Then
            n1 = pc
againtag:
            If AdvanceAnySpaceorGT_NameSpace(t$, pc, st) Then
                Tag$ = Mid$(t$, n1, pc - n1)
                If Len(Tag$) > 0 Then
                    If AdvanceSpace(t$, pc, st) Then
                     n1 = pc
                            
                            If Mid$(t$, pc, 1) = Chr$(34) Then
                                n2 = 34
                            ElseIf Mid$(t$, pc, 1) = Chr$(39) Then
                                n2 = 39
                            Else
                            GoTo againtag
                            End If
                            If n2 <> 0 Then
                                pc = pc + 1
                                n1 = pc
                                If AdvanceScope(t$, pc, st, Chr$(n2)) Then
                                 '   val$ = Mid$(t$, n1, pc - n1)
                                  '  PlaceAttributeToNode node, Tag$, val$, n2 = 39
                                    pc = pc + 2

                                    GoTo lookfortagagain
                                End If
                            End If
                        
                    End If
                End If
            End If
        End If
        For n1 = n0 To pc - 1
            If AscW(Mid$(t$, n1, 1)) > 31 Then val$ = val$ + Mid$(t$, n1, 1)
        Next n1
        node(5) = val$
        PlaceProcessingInstructions node
    End If
End If
End Sub
Private Sub insertnode2(node)
' tuple
If Not IsArray(root) Then
    If (node(3) And 2) = 2 Then
        root = Array(node(0), node(1), node(2), node(3), "<![CDATA[" + node(4) + "]]>", node(5))
    Else
        root = Array(node(0), node(1), node(2), node(3), node(4), node(5))
    End If
Else
    MyEr "I have node", "��� �����"
End If
End Sub
Friend Sub createfromtext(t$, pc As Long, st As Long)
Dim n1 As Long, n2 As Long, Tag$, node, val$
again:
If Advance2Lt(t$, pc, st) Then
    n1 = pc
    If Mid$(t$, pc, 1) = "?" Then createprocessxml t$, pc, st: GoTo again
    If Mid$(t$, pc, 1) = "!" Then
    ' special???
    If Mid$(t$, pc, 3) = "!--" Then
        pc = InStr(pc + 3, t$, "-->")
        If pc = 0 Then pc = st: Exit Sub
        pc = pc + 3
        GoTo again
        Else
         CheckThis t$, pc, st: GoTo again
    End If
    ElseIf AdvanceAnySpaceorGT_NameSpace(t$, pc, st) Then
        Tag$ = Mid$(t$, n1, pc - n1)
        node = Me.PrepareNode(Tag$)
        If Mid$(t$, pc, 1) < "A" Then
lookfortagagain:
            If AdvanceSpace(t$, pc, st) Then
                n1 = pc

                If AdvanceAnySpaceorGT_NameSpace(t$, pc, st) Then
                    Tag$ = Mid$(t$, n1, pc - n1)
                    If Tag$ = "" Then
                        If Mid$(t$, pc, 1) = ">" Then
                        pc = pc + 1
                        n1 = pc
                        If Advance2Lt(t$, pc, st) Then

                            node(4) = Mid$(t$, n1, pc - n1 - 1)
                            If Mid$(t$, pc, 1) = "/" Then
entry2:
                                If Mid$(t$, pc + 1, Len(node(0)) + 1) = node(0) + ">" Then
                                    insertnode2 node
                                    pc = pc + Len(node(0)) + 2
                                    Exit Sub
                                End If
                            Else
                                pc = pc - 1
                                GoTo newentry
                            End If
                        End If
                    Else
                    If Mid$(t$, pc, 2) = "/>" Then
                            node(3) = node(3) And &HFFFFFFFE  ' FALSE BIT 0
                            insertnode2 node
                            pc = pc + 1
                            
                        End If
                    End If
                ElseIf AdvanceSpace(t$, pc, st) Then
                    If Mid$(t$, pc, 1) = "=" Then
                        pc = pc + 1
                        If AdvanceSpace(t$, pc, st) Then
                            n2 = 0
                            If Mid$(t$, pc, 1) = Chr$(34) Then
                                n2 = 34
                            ElseIf Mid$(t$, pc, 1) = Chr$(39) Then
                                n2 = 39
                            End If
                            If n2 <> 0 Then
                                pc = pc + 1
                                n1 = pc
                                If AdvanceScope(t$, pc, st, Chr$(n2)) Then
                                    val$ = Mid$(t$, n1, pc - n1)
                    ' attribute
                                    PlaceAttributeToNode node, Tag$, val$, n2 = 39
                                    pc = pc + 1
                                    n1 = pc
                                    n2 = 0
                                    If AdvanceAnySpaceorGT(t$, pc, st) Then
                                        If Mid$(t$, pc - 1, 1) = ">" Then
newentry:
                                            insertnode2 node
                                            ' check for <
                                            If AdvanceAnySpaceorLT(t$, pc, st) Then
                                                If Mid$(t$, pc + 1, 1) = "/" Then pc = pc + 1: GoTo here
                                                If Mid$(t$, pc, 2) = "<!" Then
                                                    If Mid$(t$, pc, 4) = "<!--" Then
                                                        n1 = InStr(pc + 5, t$, "-->")
                                                        If n1 > pc Then
                                                            pc = n1 + 4
                                                            GoTo newentry
                                                        End If
                                                            pc = st + 4
                                                        Exit Sub
                                                    ElseIf Mid$(t$, pc, 9) = "<![CDATA[" Then
                                                    n1 = InStr(pc + 11, t$, "]]>") + 3
                                                    If n1 > pc Then
                                                        root(3) = node(3) Or 2
                                                        root(4) = Mid$(t$, pc, n1 - pc)
                                                        pc = n1 + 1
                                                        Exit Sub
                                                    End If
                                                        pc = st + 1
                                                        Exit Sub
                                                    End If
                                                End If
                                                
                                                
                                                Dim child As XmlMonoInternal
                                                Set child = New XmlMonoInternal
                                                child.NumericCharacterEntities = NumericCharacterEntities
                                                child.createfromtext t$, pc, st
                                                If child.HasNode Then AppendChildObj child
                                                While AdvanceAnySpaceorLT(t$, pc, st)
                                                    If Mid$(t$, pc + 1, 1) = "/" Then pc = pc + 1: GoTo here
                                                    Set child = New XmlMonoInternal
                                                    child.NumericCharacterEntities = NumericCharacterEntities
                                                    child.createfromtext t$, pc, st
                                                    If child.HasNode Then AppendChildObj child
                                                Wend
                                            End If
here:
                                            If Mid$(t$, pc + 1, Len(node(0)) + 1) = node(0) + ">" Then
                                                pc = pc + Len(node(0)) + 2
                                                Exit Sub
                                            End If
                                        Else
                                            n1 = pc
                                            GoTo lookfortagagain
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    
                    ElseIf n2 <> 0 Or Tag$ = "" Then
                        If Mid$(t$, pc, 1) = ">" Then
                            node(3) = node(3) And &HFFFFFFFE
                            insertnode2 node
                            pc = pc + 1
                            End If
                        End If
                    End If
             
             End If

            End If

        End If
    End If
End If
End Sub

Friend Property Let Xml(guard As XmlMonoInternal, t$)
Dim other
status = ""
Set ProcessBag = Nothing
achild = 0
thelast = False
SwapVariant root, other
other = Array(0)
If Len(t$) > 0 Then
Dim i As Long
i = 1
createfromtext t$, i, Len(t$)
If Len(t$) > i Then
' missing some data ???
status = Mid$(t$, i)
End If
End If
End Property

Sub RemoveChild()
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
ElseIf (root(3) And &H1) Then
    Dim val
    Set val = root(2)
    Dim c As Collection
    If val Is Nothing Then
       MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
    Else
        Set c = root(2)
        c.Remove achild + 1
        If achild >= c.Count Then achild = achild - 1
        If c.Count = 0 Then Set root(2) = Nothing: achild = -1
    End If
Else
    MyEr "No childs for this node, only attributes", "��� ������ ��� ����� ��� �����, ���� ���������"
End If
End Sub
Sub RemoveChildren()
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
Else  'If (root(3) And &H1) Then
    Set root(2) = Nothing
    achild = -1
    thelast = True
End If
End Sub

Sub RemoveAttributes()
If Not IsArray(root) Then
    MyEr "I have't node", "��� ��� �����"
Else
    Set root(1) = Nothing
    root(5) = vbNullString
End If
End Sub
Sub RemoveProcess()
    Set ProcessBag = Nothing
End Sub
Friend Sub SetNumericCharacterEntities(RHS)
    Dim M As XmlMonoInternal
    ' lastchild
    Dim mLast As XmlMonoInternal
    NumericCharacterEntities = CBool(RHS)
    If Not Me.HasNode Then Exit Sub
    If Me.ChildCount = 0 Then Exit Sub
    Set mLast = lastchild
    Set M = FirstChild
    While Not M Is Nothing
        M.SetNumericCharacterEntities RHS
        If Not M Is mLast Then Set M = Nextchild Else Set M = Nothing
    Wend
    Set M = FirstChild
End Sub

