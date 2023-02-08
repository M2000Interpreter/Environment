VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "callback"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'class for call back
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function timeBeginPeriod Lib _
"winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib _
"winmm.dll" (ByVal uPeriod As Long) As Long
Private objref As basetask
Private mineHere$
Private setstacklimit As Boolean
Private par1$, par2$
Public Event ErrorAtt(ByRef recover As Variant)
Public Finished As Boolean, mylogfile$
Private hTimer As Long, hTimer2
Dim locked As Boolean
Dim ExtForm As Object
Private lasttitle As String

Friend Sub NoPublic(bstack As basetask, where$)

        Set objref = New basetask
        Set objref.Sorosref = bstack.soros
        Set objref.Owner = bstack.Owner
        mineHere$ = where$
End Sub
Friend Sub GetIcon(A As StdPicture)
On Error Resume Next
Dim k As StdPicture
If IhaveExtForm Then
ExtForm.GetIcon PackData(A)
End If
End Sub
Friend Function GetAppIcon() As StdPicture
On Error Resume Next
Dim k As StdPicture
If IhaveExtForm Then
Set GetAppIcon = UnPackData(ExtForm.GetAppIcon)
End If
End Function
Private Function UnPackData(sData As String) As Object
    Dim pbTemp  As PropertyBag
    Dim arData()    As Byte
    Let arData() = sData
    
    Set pbTemp = New PropertyBag
    With pbTemp
        Let .Contents = arData()
        Set UnPackData = .ReadProperty("icon")
    End With
     
    Set pbTemp = Nothing
End Function
Private Function PackData(obj As Object) As String
    Dim pbTemp  As PropertyBag
    Set pbTemp = New PropertyBag
    With pbTemp
        Call .WriteProperty("icon", obj)
        Let PackData = .Contents
    End With
    
    Set pbTemp = Nothing
End Function
Property Get IhaveExtForm() As Boolean
IhaveExtForm = Not ExtForm Is Nothing
End Property
Public Sub SetExtCaption(A$)
On Error Resume Next
If ExtForm Is Nothing Then Exit Sub
If LenB(A$) = 0 Then
Hide
Form1.CaptionW = ExtForm.CaptionW
Else
Show
lasttitle = A$
ExtForm.CaptionW = A$
End If
End Sub
Public Sub GetForm(f As Object)

Set ExtForm = f
On Error Resume Next
If f Is Nothing Then Exit Sub
Set Form3AppIcon = UnPackData(ExtForm.GetAppIcon())
Err.Clear
'If F Is Nothing Then Exit Sub
End Sub
Public Sub StackMax(Stack As Long)
If IsWine Then
stacksize = 32000000
findstack = stacksize - 100000
Exit Sub
End If
If m_bInIDE Then
ClaimStack
If findstack - 100000 > 0 Then
    stacksize = findstack - 100000
End If
Exit Sub
End If
If Stack = -12345 Then stacksize = 32000000: Exit Sub
If Stack > 1000000 Then
 stacksize = Stack - 100000
 findstack = Stack
Else
ClaimStack
If findstack - 100000 > 0 Then
    stacksize = findstack - 100000
End If
End If
End Sub
Public Sub Cli(LoadFileAndSwitches$, Optional Prompt$ = ">")
par1$ = LoadFileAndSwitches$
par2$ = Prompt$
CliRun
End Sub
Public Sub Shutdown(Cancel As Integer)

If ASKINUSE Or loadfileiamloaded Then
NOEXECUTION = Cancel = 2
If NOEXECUTION Then
ASKINUSE = False: loadfileiamloaded = False: DoEvents
Else
Cancel = -1
Exit Sub
End If
End If
If Cancel < 1 Then
If Not MayQuit Then Cancel = -1: Exit Sub
Dim f As Form, i As Long, F1 As GuiM2000, nook As Long

NERR = True: NOEXECUTION = True
ExTarget = True: INK$ = Chr(27): UKEY$ = Chr$(27)
If Cancel = -100 Then Exit Sub
Set Form3AppIcon = LoadPicture("")
For i = Forms.count - 1 To 0 Step -1
Set f = Forms(i)
Set f.icon = LoadPicture("")
If TypeOf f Is GuiM2000 Then
    Set F1 = f
    
    If Modalid <> 0 Then
        If F1.Modal = Modalid Then
            If F1.myname$ <> "" Then
                
                    
                    F1.ByeBye2 nook
                
            End If
        End If
    Else
        If F1.myname$ <> "" Then
        
        
        F1.ByeBye2 nook
        
        End If
    End If

End If
If nook Then Cancel = True: Exit For
Next i

Set f = Nothing



If Cancel Then Exit Sub
End If

byPassCallback = True
NOEXECUTION = True
If Cancel > 0 Then
Check2SaveModules = False
NERR = True: NOEXECUTION = True
ExTarget = True: INK$ = Chr(27): UKEY$ = Chr$(27)  ': UINK$ = Chr(27)
End If
If Not TaskMaster Is Nothing Then
If TaskMaster.Processing Then
TaskMaster.StopProcess

End If
TaskMaster.Dispose
End If
If exWnd <> 0 Then
MyDoEvents
    nnn$ = "bye bye"
    exWnd = 0

    End If
For i = Forms.count - 1 To 0 Step -1
Set f = Forms(i)
Unload f
Next i
Set f = Nothing

If Not ExtForm Is Nothing Then

ExtForm.Shutdown

Set ExtForm = Nothing
End If

'If form5iamloaded Then Unload Form5
Finished = True
End Sub
Public Sub AsyncCli(LoadFileAndSwitches$, Optional Prompt$ = ">")
Dim pid As Long
If strTemp = vbNullString Then SetTmpPath
Sleep 100
par1$ = LoadFileAndSwitches$
par2$ = Prompt$
Set UseMe = Me

If Not Form3AppIcon Is Nothing Then
    Set Form1.icon = Form3AppIcon
    
End If
pid = GetCurrentProcessId()
If strTemp = vbNullString Then strTemp = GetTempPathgg()
Dim m As Long
mylogfile$ = strTemp + LTrim$(Str$(pid)) + ".M2000"
m = 1
Open mylogfile$ For Output As #m
Close #m
tempList2delete = Sput(mylogfile$) + tempList2delete

hTimer = SetTimer(0&, 0&, 2000, AddressOf RunMe)
End Sub
Public Sub AsyncShow()
If hTimer2 Then Exit Sub
Set UseMe = Me
hTimer2 = SetTimer(0&, 0&, 100, AddressOf ShowForm1)
End Sub
Friend Sub CliRun()
Dim LoadFileAndSwitches$, Prompt$
LoadFileAndSwitches$ = par1$
par1$ = ""
Prompt$ = par2$
Dim ohere$, s$
If findstack = 0 Then ClaimStack
If findstack - 100000 > 0 Then
    stacksize = findstack - 100000
End If
If elevatestatus Then Exit Sub
If Basestack1.Owner Is Nothing Then Exit Sub
elevatestatus = 1
ohere$ = here$
here$ = mineHere$
If LoadFileAndSwitches$ = vbNullString Then newshow objref
Form3.skiptimer = True
NoAction = True
NOEXECUTION = False
objref.toprinter = False
MOUT = False
ResetBreak
LastErNum = 0: LastErNum1 = 0
LastErName = vbNullString
LastErNameGR = vbNullString
s$ = LoadFileAndSwitches$
If Not ISSTRINGA(s$, cLine) Then
cLine = mylcasefILE(Trim(s$))
Else
cLine = mylcasefILE(cLine)
End If
While Left$(cLine, 1) = Chr(34) And Right$(cLine, 1) = Chr(34) And Len(cLine) > 2
cLine = Mid$(cLine, 2, Len(cLine) - 2)
Wend
If ExtractType(cLine) <> "gsb" Then cLine = vbNullString
If Left$(cLine, 1) Like "[+-]" Then
Dim ss$
para$ = ""
While Left$(cLine, 1) Like "[+-]"
ss$ = Split(cLine, " ")(0)
cLine = LTrim$(Mid$(cLine, Len(ss$) + 1))
para$ = para$ + ss$ + " "
Wend
ss$ = Trim$(para$)
para$ = ""
End If
If cLine <> "" Then
para$ = ExtractPath(cLine) + ExtractName(cLine)
para$ = Trim$(Mid$(cLine, Len(para$) + 1))
's$ = cLine + " " + s$
cLine = Left$(cLine, Len(cLine$) - Len(para$))
If ss$ = "" And para$ = "" Then
s$ = LTrim(s$)
If Left$(s$, 1) Like "[+-]" Then
While Left$(s$, 1) Like "[+-]"
ss$ = Split(s$, " ")(0)
s$ = LTrim$(Mid$(s$, Len(ss$) + 1))
para$ = para$ + ss$ + " "
Wend
ss$ = Trim$(para$)
para$ = ""
If ExtractType(s$) = "gsb" Then
If ExtractType(cLine) = "gsb" Then
' replace file
cLine = ExtractPath(s$) + ExtractName(s$)
End If
Else
s$ = ""
End If
End If

End If

ElseIf s$ <> "" Then
para$ = Trim$(s$)
End If

If ss$ <> "" Then
Switches ss$, True
ElseIf para$ = "" And s$ <> "" Then
Switches s$, True
Else
Switches para$, True
End If
If Trim$(para$) <> "" And Not Left$(para$, 1) Like "[+-]" Then interpret Basestack1, para$
If Not INK$ = Chr(27) Then
Form1.MyPrompt cLine, Prompt$
End If
finalize:
here$ = ohere$
elevatestatus = 0
If IsWine Then
DoEvents
Sleep 5
End If
On Error Resume Next
Dim i As Long
For i = UBound(var()) To 0 Step -1: var(i) = 0&: Next i
Finished = True
End Sub
Property Let ShowGui(RHS As Boolean)
If Forms.count = 0 Then Exit Property
If RHS Then
Form1.Visible = True
Form1.WindowState = 0
If Form1.Top >= VirtualScreenHeight Then Form1.move ScrInfo(Console).Left, ScrInfo(Console).Top
Else
If Form1.WindowState = 0 Then
If Not (Forms.count > 5) Then Form1.Visible = False
End If
End If
End Property
Public Sub ExecuteStatement(ByVal Code$)
Dim what$, Pos As Long, func As Boolean, param$, drop$
If IsLabelOnly(Code$, what$) > 4 Then
    func = True
ElseIf what$ = vbNullString Then
    Exit Sub
ElseIf LTrim$(Code$) = vbNullString Then
' one call
Call Run(what$)
Exit Sub
End If
Do
 Pos = 1
  drop$ = aheadstatus(Code$, True, Pos)
  If drop$ = vbNullString And Pos > Len(Code$) Then Exit Sub
  If param$ <> "" Then
  param$ = param$ + ", " + Left$(Code$, Pos - 1)
  Else
  param$ = Left$(Code$, Pos - 1)
  End If
  Code$ = Mid$(Code$, Pos)
  Loop Until Not FastSymbol(Code$, ",")
  If func Then
  Call Run(what$ + " " + param$ + ")")
  Else
  Call Run(what$ + " " + param$)
  End If
End Sub
Public Function Eval(Code$) As Variant
Dim Pos As Long, p As Double, s As String, drop$
 Pos = 1
  drop$ = aheadstatus(Code$, True, Pos)
  If drop$ = vbNullString And Pos > Len(Code$) Then Exit Function
  If InStr(drop$, "N") Then
  If IsExp(objref, Left$(Code$, Pos), p) Then
  If Not objref.lastobj Is Nothing Then
  Set Eval = objref.lastobj
  
  Else
    Eval = p
   End If
  End If
  Else
  If IsStrExp(objref, Left$(Code$, Pos), s$) Then
  If Not objref.lastobj Is Nothing Then
      Set Eval = objref.lastobj
  Else
    Eval = s$
  End If
  End If
  End If
Set objref.lastobj = Nothing
End Function
Public Sub AddExecCode(Code$)
If locked Then Exit Sub
Call Run(Code$, False)
End Sub
Public Function Run(ByVal Code$, Optional clean As Boolean = True) As Boolean
Dim ohere$, stay As Variant
If elevatestatus Then Exit Function
byPassCallback = False
 If Basestack1.Owner Is Nothing Then Exit Function
elevatestatus = 1
If clean Then PushStage objref, 0
ohere$ = here$
here$ = mineHere$
NoAction = True
NOEXECUTION = False
objref.toprinter = False
MOUT = False
ResetBreak
Run = True
        LastErNum = 0: LastErNum1 = 0
        LastErName = vbNullString
        LastErNameGR = vbNullString
If Execute(objref, Code$, False) <> 1 Then
        Run = False
                 MOUT = True
End If
If clean Or MOUT Or ExTarget Then
ClearLoadedForms
If NERR Then

RaiseEvent ErrorAtt(stay)
If Not stay Then GoTo finalize
End If
If Not objref Is Nothing Then
    objref.toprinter = False
Else
    Run = False: GoTo finalize
End If
    If MOUT Then
            NOEXECUTION = False
            ResetBreak
            MOUT = interpret(objref, "new: flush: clear")
            MOUT = False
        Else
        
        If NOEXECUTION Then
                closeAll
        End If

        End If
        
MOUT = True
NoAction = False

If ExTarget Then Run = False
End If
finalize:
here$ = ohere$

If clean Then
If Not objref Is Nothing Then PopStage objref Else Run = False
End If
elevatestatus = 0
End Function

Property Get ErrorGr() As String
ErrorGr = LastErNameGR
End Property
Property Get ErrorEn() As String
ErrorEn = LastErName
End Property


Property Get status() As Long
If elevatestatus Then
status = -elevatestatus
Else
status = LastErNum
End If
End Property
Public Sub Reset()
elevatestatus = 0
    NERR = False
        LastErNum = 0: LastErNum1 = 0
        LastErName = vbNullString
        LastErNameGR = vbNullString
 'If Not App.StartMode = vbSModeStandalone Then
 
NeoSubMain
 
' End If

        
End Sub
Public Sub Show()
On Error Resume Next
If UseMe Is Nothing Then Exit Sub
Form1.CaptionW = vbNullString
If ExtForm Is Nothing Then Exit Sub
'If ExtForm.WindowState = 1 Then
'ExtForm.Visible = True
'End If
ExtForm.CaptionW = lasttitle

End Sub

Public Sub Hide()
On Error Resume Next
If UseMe Is Nothing Then Exit Sub
ExtForm.Hide
End Sub
Public Function IamVisible() As Boolean
On Error Resume Next
If UseMe Is Nothing Then Exit Function
If Not IhaveExtForm Then
IamVisible = Form1.TrueVisible: Exit Function
End If
IamVisible = ExtForm.Visible
If Err.Number > 0 Then
Set ExtForm = Nothing
IamVisible = Form1.Visible
End If
End Function
Private Sub Class_Initialize()
Dim anyname As Long
startaddress = VarPtr(anyname)
stacksize = 900000
timeBeginPeriod 1
 If Not l_complete Then NeoSubMain
        Set objref = New basetask
        Set objref.Owner = Form1.DIS
        mineHere$ = vbNullString 'HERE$
End Sub

Private Sub Class_Terminate()
Dim Zero As basket, i As Long, zerocounter As Counters
StopTimer
If objref Is Nothing Then Exit Sub
Set objref.Owner = Nothing
Set objref.Sorosref = Nothing
Set objref = Nothing
globalstack.Flush
TerminateM2000
terminatefinal
For i = -2 To 1000
    players(i) = Zero
    Prefresh(i) = zerocounter
Next i
If Forms.count > 0 Then
For i = Forms.count - 1 To 0
Set Forms(i).icon = LoadPicture("")
Unload Forms(i)
Next i
End If
ReleaseMem
l_complete = False
timeEndPeriod 1
End Sub
Public Sub StopTimer()
  If hTimer <> 0 Then
  '' MasterTimer = 0
    KillTimer 0&, hTimer
    hTimer = 0
  End If
End Sub
Public Sub StopTimer2()
  If hTimer2 <> 0 Then
  '' MasterTimer = 0
    KillTimer 0&, hTimer2
    hTimer2 = 0
  End If
End Sub
Public Sub SetFocus()
On Error Resume Next
If WindowState = 1 Then WindowState = 0
If Form1.Visible Then
Form1.Show , Form5
If Form1.enabled Then Form1.SetFocus
End If

   
End Sub
Public Property Get WindowState() As Variant
If Not IhaveExtForm Then Exit Property
WindowState = 1 - CLng(Form1.Visible Or Form1.TrueVisible)
End Property
Public Property Let ExtWindowState(ByVal RHS As Variant)
If Not ExtForm Is Nothing Then

ExtForm.WindowState = RHS
End If
End Property
Public Property Let WindowState(ByVal vNewValue As Variant)
If Not IhaveExtForm Then
StopTimer
StopTimer2
Class_Terminate
Exit Property
End If
If vNewValue = 1 Then
    If Form1.Visible Or Form1.TrueVisible Then
    Form1.TrueVisible = Form1.Visible Or Form1.TrueVisible
    Form1.Visible = False
    If Form4Loaded Then If Form4.Visible Then Form4.Visible = False: reopen4 = True
    If Form2.Visible Then If trace Then Form2.Visible = False: reopen2 = True
    End If
Else
        
       Form1.Visible = Form1.Visible Or Form1.TrueVisible 'Or Forms.count < 5
        If Form1.Visible Then
        If Not trace Then reopen2 = False
        If vH_title$ = vbNullString Then reopen4 = False
        If reopen4 Then Form4.Show , Form1: Form4.Visible = True
        If reopen2 Then Form2.Show , Form1: Form2.Visible = True
        Form1.Show
        
        End If
        
End If

End Property

Public Property Get AppTitle() As Variant
AppTitle = lasttitle
End Property

Public Property Get MayQuit() As Boolean
MayQuit = Not Check2Save
End Property
Public Sub GetMess(A$)
    If Not ExtForm Is Nothing Then
        ExtForm.GetMessNow A$
    End If
End Sub
Public Sub GetMessOneChar(A$)
    If Not ExtForm Is Nothing Then
        ExtForm.GetMessNow A$
    End If
End Sub
Public Sub SendMess(A$)
    If Not ExtForm Is Nothing Then
    On Error Resume Next
        ExtForm.SendMessNow A$
        If Err Then
        Me.Shutdown True
        
        End If
        
    End If
End Sub
Public Property Get Code() As String
' GetCode
    If Not ExtForm Is Nothing Then
        On Error Resume Next
        Code = ExtForm.GetCode
        
        
    End If
End Property

