VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "basetask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Basetask is a module task of M2000
' This is a utility soros for passing parammeters to functions
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, retval As Integer)
Public LoadOnly As Boolean, NoRun As Boolean
Public CheckUnReferenceGui As Long
Public callx1 As Long, tpointer As Long
Public RRCOUNTER As Long
Public callohere As String
Public addlen As Long, NoFuncError As Boolean, CallW As String, fHere$, strg As Boolean
Public nokillvars As Boolean
Public finalFlag As Boolean
Public UseofIf As Integer
Public IFCTRL As Boolean, jump As Boolean
Public Vars As Long, Vname As Long, ByName As Long
Public ErrVars As Long, ErrVname As Long
Public usestack As Boolean
Public commnum As Long
Public strnum As Long
Public numnum As Long
Public strfunnum As Long
Public numfunnum As Long
Public connectnow As Boolean
Public LastExecute As Long
Public LastOnce As Boolean
Public priveflag As Boolean, uniflag As Boolean
Private parentSoros As mStiva ' reference to parent..stiva
Private MyParent As basetask  'reference to parent
Private myOwner As Object
Private mtoback As Boolean
Private mtoprinter As Boolean
Private mitalic As Boolean   ' todo
Private mCharSet As Integer  ' todo
Private mBold As Boolean   ' todo
Private Mchild As Boolean
Private mtolayer As Long
Private myRun As Boolean
Private mprocess As TaskInterface ' reference to current process maybe same to parrent...
Private mprocessOk As Boolean ' if i have a process class
Private threadsnames As ThreadsClass
Public TaskMain As Boolean
Public originalname$
Private mGroupName$
Public UseGroupname$
Public LastClick As Long
Public Look2Parent As Boolean
Private Const dotcolmax = 10
Private dotcol(dotcolmax) As String, curdotcol As Long
Public StaticCollection As FastCollection
Private topdotlevel As Long
Public IndexSub As Long
Public lastobj As Object
Public FuncObj As Object
Public FuncRec As String
Public FuncValue As Variant
Public ErrorOriginal As Long
Public LastValue As Variant
Public LastState As Boolean
Public LastEnum As Variant
Private RetSoros As mStiva2
Public SubLevel As Long
Public OriginalCode As Long
Public StaticInUse$
Private mIamlambda As Boolean
Private mIamAnEvent As Boolean
Public NoError As Boolean, lastpointer As Object
Public tmpstr As String
Public CopyInOutCol As Collection
Public CallLocalLast As Boolean
Public IamGlobal As Boolean
Public numnumlimit As Long
Public SkipClear As Long
Public Prof As clsProfiler
Public modfuncall As Boolean
Public restpart$


Public Sub SetSkip()
    SkipClear = SkipClear + 1
End Sub
Public Sub ResetSkip()
    SkipClear = SkipClear - 1
End Sub

Public Sub SetV()
Vars = var2used  ' these are globals
Vname = varhash.count
End Sub
Public Sub SetV2()
Vars = var2used + ByName
Vname = varhash.count + ByName
End Sub
Public Sub CopyStrip(bst As basetask)
With bst
Set .Owner = Owner
.StaticInUse = StaticInUse
.OriginalCode = OriginalCode
Set .StaticCollection = StaticCollection
Set .Sorosref = soros
End With

End Sub
Public Sub CopyStrip2(bst As basetask)
Dim i As Long
With bst
Set .Owner = Owner
.StaticInUse = StaticInUse
.OriginalCode = OriginalCode
Set .StaticCollection = StaticCollection
Set .Sorosref = soros
.DropNdot 20
For i = 1 To curdotcol
.MoveNameDot dotcol(i)
Next i
.DropRetStack
.UseGroupname = UseGroupname
.GroupName = GroupName
End With


End Sub
Public Sub SetVar(ByVal Name$, var As Variant)

    If Not StaticCollection.Find(Name$) Then
    StaticCollection.AddKey Name$
    End If
   StaticCollection.Value = var
End Sub
Public Property Get ForLevel() As Long
ForLevel = curdotcol
End Property
Public Sub SetVarobJ(ByVal Name$, var As Variant)
    If StaticCollection Is Nothing Then
    Set StaticCollection = New FastCollection
    If Parent Is Nothing Then
    Else
    Parent.SetVarobJ "%_" + StaticInUse, StaticCollection
    End If
    End If
    If Me.Parent Is Nothing Then
    
    Else
    If Not StaticCollection.Find(Name$) Then StaticCollection.AddKey Name$
    Set StaticCollection.ValueObj = var
    End If
End Sub
Public Sub clearbasket()

    Dim oldname$, Name$
    Name$ = "%_" + StaticInUse
    If StaticCollection Is Nothing Then Exit Sub
    If Me.IamThread Or Me.IamAnEvent Or Not Me.IamChild Then
        Set StaticCollection = Nothing
        Exit Sub
    End If
    
    If Not Parent Is Nothing Then
        oldname$ = Name$
        
        If StaticCollection Is Nothing Then Exit Sub
        Name$ = "%_" + Parent.StaticInUse
        If oldname$ = Name$ Then Exit Sub
        Set StaticCollection = Nothing
        If Parent.IamThread Or Parent.IamAnEvent Then Exit Sub
        If Parent.ExistVar(oldname$) Then Parent.StaticCollection.RemoveWithNoFind
    End If
    
End Sub
Public Sub SetBacket(ByVal Name$)
    Dim paren As basetask, var As Variant
    If StaticCollection Is Nothing Then
        Set StaticCollection = New FastCollection
        
    End If
    Set var = StaticCollection
    If Me.IamThread Then Exit Sub
    If Parent Is Nothing Then
    
    Else
        Set paren = Me
        Do While paren.StaticInUse = paren.Parent.StaticInUse
        Set paren = paren.Parent
        
        If paren.Parent Is Nothing Then Exit Do
        
        Loop
        If Not paren.Parent Is Nothing Then Set paren = paren.Parent
        If paren.StaticCollection Is Nothing Then
            Set paren.StaticCollection = New FastCollection
        End If
        If Not paren.StaticCollection.Find(Name$) Then paren.StaticCollection.AddKey Name$
        
        Set paren.StaticCollection.ValueObj = var
        If paren.IamThread Or paren.IamAnEvent Then Exit Sub
        Dim oldcol As FastCollection
        Set oldcol = paren.StaticCollection
        Do While Not paren.Parent Is Nothing
        
        Name$ = "%_" + paren.StaticInUse
        Set paren = paren.Parent
        If paren.IamThread Or paren.IamAnEvent Then Exit Do
        If paren.StaticCollection Is Nothing Then
        
            Set paren.StaticCollection = New FastCollection
            paren.StaticCollection.AddKey Name$
            Set paren.StaticCollection.ValueObj = oldcol
            
        ElseIf paren.StaticCollection.Find(Name$) Then
        
            Exit Do
            ' problem this exist
        Else
            paren.StaticCollection.AddKey Name$
            Set paren.StaticCollection.ValueObj = oldcol
        End If
        Set oldcol = paren.StaticCollection
        Loop
    End If
    
End Sub
Public Sub SetVarobJvalue(ByVal Name$, var As Variant)
    If Not StaticCollection.Find(Name$) Then
    StaticCollection.AddKey Name$
    End If
    Dim myobject As Object, usehandler As mHandler
    
    If MyIsObject(var) Then
      Set StaticCollection.ValueObj = var
    ElseIf StaticCollection.IsObj Then

    Set myobject = StaticCollection.ValueObj
    If TypeOf myobject Is Group Then
        StaticCollection.Value = var
    ElseIf TypeOf myobject Is mHandler Then
    Set usehandler = myobject
    If usehandler.t1 = 4 Then
        Dim ok As Boolean
        Set var = usehandler.objref.SearchValue(var, ok)
        If ok Then
            Set StaticCollection.ValueObj = var
        Else
            ExpectedEnumType
        End If
    Else
        NoNumberAssign
    End If
    End If
    Else
    StaticCollection.Value = var
    
    End If
End Sub

Function AlterVar(ByVal Name$, p As Variant, ss$, inttype As Boolean) As Boolean
Dim sp As Variant, usehandler As mHandler
If Not StaticCollection.Find(Name$) Then Exit Function

  With StaticCollection
  If .IsObj Then
    If Typename(.ValueObj) = "mHandler" Then
    Set usehandler = .ValueObj
        If usehandler.t1 = 3 Then
            Dim mm As mArray
            Set mm = usehandler.objref
            If mm.Arr Then
                mm.Compute0 p, ss$
            End If
            Set mm = Nothing
        ElseIf usehandler.t1 = 4 Then
                If ss$ = "++" Then
                        If usehandler.index_start < usehandler.objref.count - 1 Then
                            usehandler.index_start = usehandler.index_start + 1
                            usehandler.objref.index = usehandler.index_start
                            usehandler.index_cursor = usehandler.objref.Value
                        End If
        ElseIf ss$ = "--" Then
                    If usehandler.index_start > 0 Then
                            usehandler.index_start = usehandler.index_start - 1
                            usehandler.objref.index = usehandler.index_start
                            usehandler.index_cursor = usehandler.objref.Value
                        End If
        ElseIf ss$ = "-!" Then
                        usehandler.sign = -usehandler.sign
                        Else
                        NoOperatorForThatObject ss$
        
        End If
End If
    End If
  Else
  Dim isLong As Boolean
            isLong = VarType(.Value) = vbLong
            
            Select Case ss$
            Case "!!", "~"
     
             Select Case VarType(.Value)
                        Case vbBoolean
                            .Value = Not CBool(.Value)
                        Case vbCurrency
                            .Value = CCur(Not CBool(.Value))
                        Case vbDecimal
                            .Value = CDec(Not CBool(.Value))
                        Case Else
                            .Value = CDbl(Not CBool(.Value))
                        End Select
            Case "++"
            .Value = .Value + 1
            Case "--"
            .Value = .Value - 1
            Case "-!"
            .Value = -.Value
            Case "=", "g"
            .Value = p
            Case "+", "+="
            .Value = p + .Value
            Case "*", "*="
            .Value = p * .Value
            Case "-", "-="
            .Value = .Value - p
            Case "/", "/="
                If p = 0 Then
                DevZero
                Exit Function
            End If
            .Value = .Value / p
            Case "!"
            .Value = -1 - (.Value <> 0)
                
            
            Case "DIV", "ΔΙΑ"
            If p = 0 Then
                DevZero
                Exit Function
            End If
                .Value = Fix(.Value / p)
            Case "DIV#", "ΔΙΑ#"
                If p = 0 Then
                    DevZero
                    Exit Function
                End If
                If p < 0 Then
                    .Value = Int((.Value - Abs(.Value - Abs(p) * Int(.Value / Abs(p)))) / p)
                Else
                    .Value = Int(.Value / p)
                End If
            Case "MOD", "ΥΠΟΛ", "ΥΠΟΛΟΙΠΟ"
                If p = 0 Then
                    DevZero
                    Exit Function
                End If
                sp = .Value - Fix(.Value / p) * p
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                .Value = sp
             Case "MOD#", "ΥΠΟΛ#", "ΥΠΟΛΟΙΠΟ#"
                 If p = 0 Then
                    DevZero
                    Exit Function
                End If
                sp = Abs(.Value - Abs(p) * Int(.Value / Abs(p)))
                If Abs(sp) >= Abs(p) Then sp = sp - sp
                .Value = sp
            Case Else
                WrongOperator
            End Select
            If isLong Then
                On Error Resume Next
                .Value = CLng(MyRound(.Value, 0))
                If Err.Number = 6 Then .Value = CDbl(MyRound(.Value, 0)): Err.Clear
            ElseIf inttype Then
                .Value = MyRound(.Value, 0)
            Else
                If VarType(.Value) <> vbDecimal Then .Value = MyRound(.Value, 13)
            End If
            End If
            End With
   AlterVar = True
End Function
Static Function MyRound(a As Variant, Optional ByVal i As Integer = 0)
Dim c As Variant, j As VbVarType
Dim n(1 To 28) As Single, D(1 To 28) As Double, cur(1 To 28) As Currency, dec(1 To 28) As Variant
Dim SG(-1 To 1) As Single, sg4(-1 To 1) As Double, sg8(-1 To 1) As Currency
j = VarType(a)
If j < vbSingle Then MyRound = a: Exit Function
On Error GoTo there
If n(1) = 0 Then
    SG(-1) = CSng(-0.5)
    SG(1) = CSng(0.5)
    sg4(-1) = -0.5
    sg4(1) = 0.5
    sg8(-1) = CCur(-0.5)
    sg8(1) = CCur(0.5)
    For c = 1& To 6&
        n(c) = CSng(10 ^ c)
    Next c
    For c = 7& To 27&
        n(c) = CSng(-1)
    Next c
        n(c) = CSng(10 ^ 5)
        n(10) = CSng(10 ^ 5)
    For c = 1& To 13&
        D(c) = CDbl(10 ^ c)
    Next c
    For c = 14& To 27&
        D(c) = CDbl(-1)
    Next c
    D(c) = CDbl(10 ^ 13)
    
    For c = 1& To 3&
        cur(c) = CCur(10 ^ c)
    Next c
    For c = 4& To 28&
        cur(c) = CCur(-1)
    Next c
    For c = 1& To 28&
        dec(c) = CDec(10 ^ c)
    Next c
End If
    If i = 0 Then
        MyRound = Sgn(a) * Int(Abs(a) + 0.5)
        Exit Function
    Else
        c = Fix(a)
        Select Case j
        Case vbSingle
            If n(i) > 0 Then
                c = SG(Sgn(a))
                MyRound = Fix(a * n(i) + c) / n(i)
            Else
                MyRound = a
            End If
            Exit Function
        Case vbDouble
            c = sg4(Sgn(a))
 
            If D(i) > 0 Then
                MyRound = Fix(a * D(i) + c) / D(i)
            Else
                MyRound = a
            End If
            Exit Function
        Case vbCurrency
            If cur(i) > 0 Then
                c = sg8(Sgn(a))
                MyRound = Fix(a) + Fix((a - Fix(a)) * cur(i) + c) / cur(i)
            Else
                MyRound = a
            End If
            Exit Function
        Case vbDecimal
            c = sg8(Sgn(a))
            MyRound = Fix(a) + Fix((a - Fix(a)) * dec(i) + c) / dec(i)
            Exit Function
        Case Else
            MyRound = a
            Exit Function
        End Select
    End If
there:
Err.Clear
MyRound = a
If i < 0 Or i > 28 Then MyEr "Round Place exit range 0 to 28", "Η θέση στρογγυλοποίησης εκτός περιοχής από 0 έως 28"
End Function

Public Function ExistVar(ByVal Name$, Optional IsObj As Boolean = False) As Boolean
        ExistVar = StaticCollection.ExistKey(Name$)
        IsObj = StaticCollection.IsObj
End Function

Public Function ExistVar2(ByVal Name$) As Boolean
        If StaticCollection Is Nothing Then Exit Function
        ExistVar2 = StaticCollection.ExistKey(Name$)
End Function
Public Sub ReadVar(ByVal Name$, var As Variant)
    Dim i As Long, vv As Variant
    If StaticCollection Is Nothing Then Exit Sub
    With StaticCollection
        If .Find(Name$) Then
            If .IsObj Then
                    Set var = .ValueObj
                    
            If Left$(Name$, 1) = "%" Then Set .ValueObj = Nothing
            Else
                    var = .Value
            End If
        End If
    End With
End Sub
Public Sub ReadStaticCol(ByVal Name$, var As Variant)
    Dim i As Long, vv As Variant
    If StaticCollection Is Nothing Then
    Set var = Nothing
    Exit Sub
    End If
    With StaticCollection
        If .Find(Name$) Then
            If .IsObj Then
                    Set var = .ValueObj
              Else
                   Set var = Nothing
            End If
        End If
    End With
End Sub
Public Sub ReadStaticCol2(Name$, var As Variant)
    Dim i As Long, vv As Variant
    If StaticCollection Is Nothing Then
    Set var = Nothing
    Exit Sub
    End If
    With StaticCollection
        If .Find(Name$) Then
            If .IsObj Then
                    Set var = .ValueObj
              Else
                   Set var = Nothing
            End If
        End If
    End With
End Sub
Public Sub DropNdot(n As Long)
If n = 0 Then Exit Sub
If curdotcol - n < 0 Then
    curdotcol = 0
Else
    curdotcol = curdotcol - n
End If
End Sub

Public Sub MoveNameDot(a$)
If a$ = vbNullString Then Exit Sub
If curdotcol < dotcolmax Then
    curdotcol = curdotcol + 1
     dotcol(curdotcol) = a$
Else
    MyEr "Max count for dot names", "Μέγιστος αριθμός τελείας"
End If
End Sub
Public Function GetDotNew(a$, dotnum&, Optional Clear As Boolean = False) As Boolean
If Look2Parent Then GetDotNew = Parent.GetDotNew(a$, dotnum&): Exit Function

If curdotcol = 0 Then
If Len(UseGroupname) > 0 Then
    If Right(a$, 1) = "(" Then
        If AscW(Left$(a$, 1)) > 255 Then
       
        a$ = "ΑΥΤΟ." + a$
       Else
        a$ = "THIS." + a$
       End If
    ElseIf Len(a$) = 0 Then
       If tpointer > 0 Then
           a$ = var(tpointer).Patch + "."
       Else
           a$ = "THIS"
       End If
    Else
       a$ = UseGroupname + a$
    End If
 
 End If
Else
newrun:
If curdotcol < dotnum Then
'' wrong
If UseGroupname = vbNullString Then

Else

            
            a$ = UseGroupname + a$
 End If
Else

  If dotcol(dotnum) = "THIS" Then
    If Len(UseGroupname) > 0 Then
      a$ = UseGroupname + a$
    Else
    a$ = "THIS." + a$
    End If
  Else
  If Len(a$) = 4 Then
  If a$ = "THIS" Then
    a$ = dotcol(dotnum)
    GoTo aaa
  ElseIf a$ = "ΑΥΤΟ" Then
    a$ = dotcol(dotnum)
    GoTo aaa
  End If
  End If
  If Left$(dotcol(dotnum), 1) = "*" And Not Clear Then
    a$ = Mid$(dotcol(dotnum), 2) + "." + a$
  Else

    a$ = dotcol(dotnum) + "." + a$

    End If
    
If Left$(a$, 1) = "." Then
a$ = Me.UseGroupname + Mid$(a$, 2)
GoTo aaa
End If
    End If
    If AscW(a$) = 46 Then
    dotnum = 1
    While AscW(Mid$(a$, dotnum + 1) + " ") = 46
     dotnum = dotnum + 1
    Wend
    a$ = Mid$(a$, dotnum + 1)
    dotnum = dotnum + 1
    GoTo newrun
    End If
    End If
    End If
aaa:
    GetDotNew = True
'If Len(here$) > 0 Then
'If Left$(a$, Len(here$) + 1) = here$ + "." Then a$ = Mid$(a$, Len(here$) + 2)
'End If
End Function

Public Function GetDot(a$, st&) As Boolean
Dim i&
If curdotcol = 0 Then a$ = Mid$(a$, 2): Exit Function
Do While Mid$(a$, st&, 1) = " "
st& = st& + 1
If st& > Len(a$) Then Exit Function
Loop

For i& = st& + 1 To Len(a$)
If Mid$(a$, i&, 1) <> "." Then i& = i& + 1: Exit For
Next i&
i& = i& - 2
If st& > 1 Then a$ = Left$(a$, st& - 1) + dotcol(i& - st& + 1) + Mid$(a$, i&) Else Exit Function
If Not Err Then GetDot = True
End Function



Public Property Let myitalic(ByVal RHS As Boolean)
mitalic = RHS
End Property

Public Property Get myitalic() As Boolean
myitalic = mitalic
End Property
Public Property Let myCharSet(ByVal RHS As Integer)
mCharSet = RHS
End Property

Public Property Get myCharSet() As Integer
myCharSet = mCharSet
End Property
Public Property Let myBold(ByVal RHS As Boolean)
mBold = RHS
End Property

Public Property Get myBold() As Boolean
myBold = mBold
End Property
Public Property Let toprinter(ByVal RHS As Boolean)
mtoprinter = RHS
End Property

Public Property Get toprinter() As Boolean
toprinter = mtoprinter
End Property
Public Property Let tolayer(ByVal RHS As Long)
mtolayer = RHS
End Property

Public Property Get tolayer() As Long
tolayer = mtolayer
End Property

Public Property Let toback(ByVal RHS As Boolean)
mtoback = RHS
End Property

Public Property Get toback() As Boolean
toback = mtoback
End Property


Public Property Get IamRunning() As Boolean
IamRunning = myRun
End Property
Public Property Get IamChild() As Boolean
IamChild = Mchild
End Property

Public Property Set Parent(dady As basetask)
Set MyParent = dady
Mchild = True
Me.toback = dady.toback
Me.toprinter = dady.toprinter
Me.tolayer = dady.tolayer
' αν θέλω μαζί με την σύνδεση του Parent
' δίνω και την Sorosref από την Soros του Parent (την δική του ή από τον πατέρα του)
End Property
Public Property Get Parent() As basetask
Set Parent = MyParent

End Property
Public Property Set Sorosref(a As mStiva)
Set parentSoros = a
End Property
Public Property Get RetStackTotal() As Long
If RetSoros Is Nothing Then
Else
RetStackTotal = RetSoros.Total
End If
End Property

Public Property Get RetStack() As mStiva2
If RetSoros Is Nothing Then Set RetSoros = New mStiva2
Set RetStack = RetSoros
End Property
Public Function SorosNothing() As Boolean
SorosNothing = parentSoros Is Nothing
End Function
Public Sub DropRetStack()
Set RetSoros = Nothing
End Sub

Public Property Get soros() As mStiva
If parentSoros Is Nothing Then
Set parentSoros = New mStiva
End If
Set soros = parentSoros
End Property
Public Property Set Owner(a As Object)
Set myOwner = a
End Property
Public Property Get Owner() As Object
If mtoprinter Then
Set Owner = Form1.PrinterDocument1
Else
Set Owner = myOwner
End If

On Error Resume Next
Err.Clear
If Not Owner Is Nothing Then
mitalic = Owner.Font.Italic
mBold = Owner.Font.bold
mCharSet = Owner.Font.charset
myRun = True
Else
myRun = False
End If
Err.Clear
End Property
Public Property Get IamThread() As Boolean
IamThread = mprocessOk
End Property
Public Sub EndRun()
' μόνο από το Myprocess.tick
myRun = False
End Sub
Public Property Set Process(a As TaskInterface)
If mprocess Is a Then
Set mprocess = Nothing
mprocessOk = True
Else
Set mprocess = a

mprocessOk = True
End If
End Property
Public Property Get Process() As TaskInterface
Set Process = mprocess
End Property
Private Sub checkialive()
If threadsnames Is Nothing Then Set threadsnames = New ThreadsClass
End Sub

Public Property Set LinkThread(Threadid As Long, a As TaskInterface)
Dim aa As mThreadref
checkialive
If threadsnames.Find(Threadid) Then
Set aa = threadsnames.ValueObj
Set aa.thread = a
End If
End Property

Public Sub PushThread(Threadid As Long, threadvar As String)
Dim aa As New mThreadref
checkialive
Set aa.thread = Nothing  ' no thread
aa.Id = Threadid
aa.VarName = threadvar
If Not threadsnames.Find(Threadid) Then threadsnames.AddKey Threadid
Set threadsnames.ValueObj = aa

End Sub
Public Function exist(Threadid As Long, ThreadName As String) As Boolean
'Dim aa As mThreadref
If threadsnames Is Nothing Then exist = False: Exit Function
If threadsnames.Find(Threadid) Then
 With threadsnames.ValueObj
    exist = .VarName = ThreadName
End With
End If

End Function

Public Function ThrowOne(Threadid As Long) As Boolean
Dim aa As mThreadref
On Error Resume Next
If threadsnames Is Nothing Then Exit Function
' this is mandatory
If TaskMaster Is Nothing Then Exit Function
If TaskMaster.ThrowOne(Threadid) Then ThrowOne = True
If Err.Number > 0 Then Err.Clear: Exit Function
' maybe the task isn't on our list - because we send message from other thread!
If threadsnames.Find(Threadid) Then
    Set aa = threadsnames.ValueObj

    threadsnames.Remove Threadid
    threadsnames.PrevIndex
    aa.thread.Dispose tmAllTasks
    ThrowOne = True
    Set aa.thread = Nothing
Else
    Err.Clear
    If Not Me.Parent Is Nothing Then
    ThrowOne = Me.Parent.ThrowOne(Threadid)
    End If
End If
Set aa = Nothing
End Function
Public Function ThreadsStr(Optional ByVal Lang As Long = 0)
'' use threadsnames

Dim aa As mThreadref, a$
If threadsnames Is Nothing Then Exit Function

If threadsnames.count > 0 Then
threadsnames.ToStart
While threadsnames.Done
' For Each aa In threadsnames
Set aa = threadsnames.ValueObj
If Left$(aa.VarName, 1) = "_" Then
If Lang = 0 Then
a$ = a$ + "ΕΣΩΤΕΡΙΚΟ ΝΗΜΑ "
Else
a$ = a$ + "INTERNAL THREAD "
End If
Else
a$ = a$ + aa.VarName + Str$(aa.Id) + " "
End If
threadsnames.NextIndex
Wend
Set aa = Nothing
'Next aa
End If
If Me.IamChild Then
If Not Me.Parent Is Nothing Then
ThreadsStr = Trim$(a$ & Me.Parent.ThreadsStr(Lang))
Else
ThreadsStr = vbNullString
End If
Else
ThreadsStr = Trim$(a$)
End If
End Function
Public Function ThreadsNumber() As Long
'' use threadsnames
Dim aa As mThreadref, kk
If threadsnames Is Nothing Then ThreadsNumber = kk: Exit Function
If threadsnames.count > 0 Then
threadsnames.ToStart
While threadsnames.Done
Set aa = threadsnames.ValueObj
If Left$(aa.VarName, 5) <> "_multi" Then kk = kk + 1
threadsnames.NextIndex
Wend
Set aa = Nothing
End If
ThreadsNumber = kk
End Function
Public Sub ThrowThreads(Optional idexclude As Long = -1)
If threadsnames Is Nothing Then Exit Sub
Dim aa As mThreadref, Tid As Long
Dim thatProcess As TaskInterface
Dim thatBasetask As basetask
cont:
If threadsnames.count > 0 Then
threadsnames.ToStart
While threadsnames.Done
Set aa = threadsnames.ValueObj
If idexclude <> aa.Id Then
If Not TaskMaster Is Nothing Then
        If TaskMaster.ThrowOne(aa.Id) Then
        ThrowOne aa.Id
        ElseIf Not IamThread Then
        threadsnames.RemoveWithNoFind
        GoTo cont
        End If
Else
ThrowOne aa.Id
End If
End If
threadsnames.NextIndex

Wend
Set aa = Nothing
End If

End Sub

'Private Sub Class_Initialize()
'Set Owner = Nothing
'Set lastobj = Nothing
'Set FuncObj = Nothing
'End Sub
Public Function IsOptionalForArray(check As Boolean) As Boolean
If parentSoros.LookNull Then
    IsOptionalForArray = False
ElseIf parentSoros.PopType = ">" Then
    parentSoros.drop 1
    IsOptionalForArray = check
ElseIf parentSoros.count = 0 Then
  IsOptionalForArray = check
End If
End Function
Public Function IsOptional() As Boolean
If parentSoros.LookNull Then
    IsOptional = False
ElseIf parentSoros.PopType = ">" Then
    parentSoros.drop 1
    IsOptional = True
ElseIf parentSoros.count = 0 Then
  IsOptional = True
End If
End Function
Public Function IsNumber(R As Variant) As Boolean
If parentSoros.LookNull Then
R = parentSoros.PopVal
IsNumber = True
ElseIf parentSoros.PopType = "N" Then
R = parentSoros.PopVal
IsNumber = True
ElseIf parentSoros.PopType = "L" Then
R = parentSoros.PopValLong
IsNumber = True
End If
End Function
Public Function IsVariant(R As Variant) As Boolean
If parentSoros.LookNull Then
R = parentSoros.PopVal
IsVariant = True
ElseIf parentSoros.PopType = "D" Then
R = parentSoros.PopVariant
IsVariant = True
End If
End Function
Public Sub RemoveOptionals()
If Parent Is Nothing Then Exit Sub
Do While Not Parent.soros.IsEmpty
If parentSoros.PopType = ">" Then Parent.soros.drop 1 Else Exit Do
Loop
End Sub

Public Function IsObjectRef(R As Object) As Boolean
If parentSoros.LookNull Then
Set R = Nothing
IsObjectRef = True
ElseIf parentSoros.PopType = "*" Then
Set R = parentSoros.PopObj
IsObjectRef = True
End If
End Function
Public Function IsString(R As String) As Boolean
If parentSoros.LookNull Then
parentSoros.PopStr2 R
IsString = True
ElseIf parentSoros.PopType = "S" Then
parentSoros.PopStr2 R
IsString = True
End If
End Function
Public Function IsInStackObject(R As Object) As Boolean  ' not used yet
If soros.LookNull Then
Set R = Nothing
IsInStackObject = True
ElseIf soros.PopType = "*" Then
Set R = parentSoros.PopObj
IsInStackObject = True
End If
End Function
Public Function RetStackLookTopVal()
If RetSoros Is Nothing Then Exit Function
RetStackLookTopVal = RetSoros.LookTopVal
End Function
Public Function IsInRetStackNumberLong(R As Long) As Boolean
If RetSoros Is Nothing Then Exit Function
If RetSoros.IsDecimal Then Exit Function

Dim v As Variant
v = RetSoros.LookTopVal

If MyIsNumeric(v) Then
IsInRetStackNumberLong = True

R = CLng(v)
RetSoros.drop 1
End If
End Function
Public Sub PushSecondThird(a As Long, b As Long)
    RetStack.PushSecondThird a, b
End Sub
Public Function isPop3Long(a As Long, b As Long) As Long
    isPop3Long = RetStack.isPop3Long(a, b)
End Function
Public Function Pop3Long(a As Long, b As Long) As Long
    Pop3Long = RetStack.Pop3Long(a, b)
End Function
Public Function IsDecimal() As Boolean
    IsDecimal = RetStack.IsDecimal
End Function
Public Sub RetStackDrop(a As Long)
    RetStack.drop a
End Sub
Public Sub Flush()
If Not RetStack Is Nothing Then RetStack.Flush
End Sub
Public Function IsInRetStackNumber(R As Variant) As Boolean
If RetSoros Is Nothing Then Exit Function
Dim v As Variant
v = RetSoros.LookTopVal
If MyIsNumeric(v) Then
IsInRetStackNumber = True
SwapVariant R, v
RetSoros.drop 1
End If
End Function
Public Function IsInRetStackString(R As String) As Boolean
If RetSoros Is Nothing Then Exit Function
Dim v As Variant
v = RetSoros.LookTopVal
If Not MyIsNumeric(v) Then
R = v
IsInRetStackString = True
RetSoros.drop 1
End If


End Function
Public Function IsInStackLong(l As Long) As Boolean
If soros.PopType = "L" Then
    l = CLng(soros.PopValLong)
    IsInStackLong = True
End If
End Function
Public Function IsInStackNumber(R As Variant) As Boolean
If soros.LookNull Then
R = soros.PopVal
IsInStackNumber = True
ElseIf soros.PopType = "N" Then
R = soros.PopVal
IsInStackNumber = True
End If
End Function
Public Function IsInStackString(R As String) As Boolean
If soros.LookNull Then
R = soros.PopStr
IsInStackString = True
ElseIf soros.PopType = "S" Then
R = soros.PopStr
IsInStackString = True
End If
End Function

Private Sub Class_Terminate()
Me.ThrowThreads
If IamChild Then Set MyParent.lastobj = lastobj
Set LastEnum = Nothing
Set lastobj = Nothing
Set myOwner = Nothing
Set mprocess = Nothing
Set parentSoros = Nothing
Set MyParent = Nothing
Set StaticCollection = Nothing
End Sub



Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Let GroupName(ByVal vNewValue As String)

If vNewValue <= "." Then
mGroupName = vbNullString
Else
mGroupName = vNewValue
End If
End Property

Public Property Get IamAnEvent() As Variant
IamAnEvent = mIamAnEvent
End Property

Public Property Let IamAnEvent(ByVal vNewValue As Variant)
mIamAnEvent = vNewValue
End Property

Friend Function lastobjIndirect(var()) As Variant
Dim usehandler As mHandler
If lastobj Is Nothing Then
ElseIf TypeOf lastobj Is mHandler Then
Set usehandler = lastobj
With usehandler
Select Case .indirect
    Case 0 To var2used
        On Error Resume Next
        Set lastobjIndirect = var(.indirect)
    Case Else
    
       Set lastobjIndirect = .objref
    End Select
End With
Else
Set lastobjIndirect = lastobj
End If

End Function


Public Property Get IamLambda() As Variant
IamLambda = mIamlambda
End Property

Public Property Let IamLambda(ByVal RHS As Variant)
mIamlambda = RHS
End Property


