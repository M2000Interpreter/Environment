VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents pbox As PictureBox
Attribute pbox.VB_VarHelpID = -1
Private mEnabled As Boolean
Public Relax As Boolean
Private Once As Boolean
Dim mIndex As Long, MoveOnEnter As Boolean
Public ctrlName As String
Private Targets As Boolean, q() As target
Dim Callback As GuiM2000
Public MY_BACK As cDIBSection
Private mesCtrlName As String, mprive As Long, mtext As String
Friend Property Get GetCallBack() As GuiM2000
Set GetCallBack = Callback
End Property
Friend Sub Construct(there As GuiM2000, Name$)
 mIndex = -1
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set pbox = there.Controls.Add("vb.picturebox", ctrlName)
 there.AddGuiControl Me
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Friend Sub ConstructArray(there As GuiM2000, Name$, i As Long)
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set pbox = there.Controls.Add("vb.picturebox", ctrlName)
 mIndex = i
 there.AddGuiControl Me
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = Name$
End Sub
Public Sub SetUp()
Dim ff As New StdFont
With Callback
ff.Name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With
MoveOnEnter = True
With pbox
.AutoRedraw = True
.BorderStyle = 0
.backcolor = rgb(200, 120, 60)
Set .Font = ff
.backcolor = Callback.Controls(1).backcolor
.forecolor = Callback.Controls(1).forecolor
End With
End Sub
Public Property Let Text(ByVal RHS As String)
mtext = RHS
pbox.Cls
pbox.currentX = 0
pbox.currentY = 0
pbox.Print mtext
pbox.Refresh
End Property
Public Property Get Text() As String
    Text = mtext
End Property
Public Sub deconstruct()
Set Callback = Nothing
Set pbox = Nothing
End Sub
Public Sub move(x, y, Optional w, Optional h)
If IsMissing(w) Then
pbox.move CLng(x), CLng(y), 0, 0
ElseIf IsMissing(h) Then
pbox.move CLng(x), CLng(y), CLng(w), 0
Else
pbox.move CLng(x), CLng(y), CLng(w), CLng(h)
End If
End Sub

Public Property Get locked() As Boolean
locked = Not pbox.enabled
End Property

Public Property Let locked(ByVal RHS As Boolean)
pbox.enabled = Not RHS
End Property
Public Property Let prive(ByVal RHS As Long)
    If Not pbox Is Nothing Then
        pbox.Tag = RHS
        players(RHS).controlname = myUcase(Callback.modulename + "." + mesCtrlName, True)
       
    End If
        mprive = RHS
End Property
Public Property Get prive() As Long
    If Not pbox Is Nothing Then
        pbox.Tag = mprive
    End If
    prive = mprive
End Property

Private Sub Class_Terminate()
Dim clearbasket As basket
If prive <> 0 Then
    players(prive) = clearbasket
End If
End Sub
Property Let Visible(ByVal RHS As Boolean)
    If Not pbox Is Nothing Then
        pbox.Visible = RHS
    End If
End Property
Property Get Visible() As Boolean
    If Not pbox Is Nothing Then
        Visible = pbox.Visible
    End If
End Property

Private Sub pbox_Click()
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
    Else
      Callback.Callback mesCtrlName + ".Click()"
    End If
End Sub
Private Sub pbox_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
Dim bstack As basetask, oldhere$

If Not Relax Then
Relax = True
Dim sel&

If Button > 0 And Targets Then
    sel& = ScanTarget(q(), CLng(x), CLng(y), prive)
    If sel& >= 0 Then
        If Button = 1 Then
            Select Case q(sel&).id Mod 100
            Case Is < 10
                SwapStrings here$, oldhere$
                here$ = Callback.modulename()
                Set bstack = New basetask
                Set bstack.Owner = pbox
                Set bstack.Sorosref = New mStiva
                If Execute(bstack, (q(sel&).Comm), False) = 0 Then Beep
                SwapStrings here$, oldhere$
                
            Case Else
            
            If mIndex > -1 Then
                Callback.Callback mesCtrlName + ".Target" + "(" + CStr(index) + "," + Str(sel& + prive * 10000) + ")"
            Else
                Callback.Callback mesCtrlName + ".Target" + "(" + Str(sel& + prive * 10000) + ")"
            End If

            End Select
            
        End If
        
        Button = 0
        Relax = False
        Exit Sub
    End If
End If



If mIndex > -1 Then
    Callback.Callback mesCtrlName + ".MouseDown(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
    Callback.Callback mesCtrlName + ".MouseDown(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If
End Sub
Private Sub pbox_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then
Relax = True
If mIndex > -1 Then
Callback.Callback mesCtrlName + ".MouseUp(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback.Callback mesCtrlName + ".MouseUp(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If
End Sub
Private Sub pbox_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then
Relax = True

If mIndex > -1 Then
Callback.Callback mesCtrlName + ".MouseMove(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback.Callback mesCtrlName + ".MouseMove(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If

End Sub
Private Sub pbox_DblClick()
    If mIndex >= 0 Then
        Callback.Callback mesCtrlName + ".DblClick(" + Str(mIndex) + ")"
    Else
        Callback.Callback mesCtrlName + ".DblClick()"
    End If
End Sub

Private Sub pbox_Resize()
If mprive = 0 Then Exit Sub
Set MY_BACK = Nothing
If Not Once Then
Once = True
Dim VR(2)
VR(0) = Width
VR(1) = Height
    If mIndex >= 0 Then
        Callback.CallbackNow mesCtrlName + ".Resize(" + Str(mIndex) + "," + Str(mIndex) + ")", VR()
    Else
        Callback.CallbackNow mesCtrlName + ".Resize()", VR()
    End If
    If VR(0) <> Width Or VR(1) <> Height Then
        pbox.move pbox.Left, pbox.top, VR(0), VR(1)
    End If
Once = False
End If
End Sub
Public Sub CopyBack()
If pbox Is Nothing Then Exit Sub
pbox.PaintPicture pbox.Parent.Image, 0, 0, , , pbox.Left, pbox.top
End Sub

Public Sub DisAllTargets()
DisableTargets q(), prive
End Sub
Friend Sub RenderTarget(bstack As basetask, rest$, Lang As Long, tHandle As Variant)
Dim p, w$, x
If tHandle \ 10000 <> prive Then
MyEr "target not for this image", "Ô ÛÙ¸˜ÔÚ ‰ÂÌ ÂﬂÌ·È „È· ·ıÙﬁ ÙÁÌ ÂÈÍ¸Ì·"
Exit Sub
End If
p = tHandle Mod 10000
If p >= 0 And p < UBound(q()) Then
     
              '
While FastSymbol(rest$, ",")
x = Empty
If IsLabelSymbolNew(rest$, "÷—¡”«", "TEXT", Lang) Then
If IsStrExp(bstack, rest$, w$) Then q(p).Tag = w$
ElseIf IsLabelSymbolNew(rest$, "–≈Õ¡", "PEN", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).pen = x
ElseIf IsLabelSymbolNew(rest$, "÷œÕ‘œ", "BACK", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).back = x
ElseIf IsLabelSymbolNew(rest$, "–À¡…”…œ", "BORDER", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).fore = x
ElseIf IsLabelSymbolNew(rest$, "œƒ«√…¡", "COMMAND", Lang) Then
If IsStrExp(bstack, rest$, w$) Then q(p).Comm = w$
ElseIf IsLabelSymbolNew(rest$, "‘…Ã«", "VALUE", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).topval = Int(x * 100)
ElseIf IsLabelSymbolNew(rest$, "¬¡”«", "BASE", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).botval = Int(x * 100):
ElseIf IsLabelSymbolNew(rest$, "◊—ŸÃ¡", "COLOR", Lang) Then
If IsExp(bstack, rest$, x, , True) Then q(p).barC = x
ElseIf IsLabelSymbolNew(rest$, "Ã≈√≈»œ”", "SIZE", Lang) Then
If IsExp(bstack, rest$, x, , True) Then
If x > 100 Then x = 100
If x < -100 Then x = -100
q(p).imagesize = Int(x)
End If
' " ¡»≈‘«", "PORTRAIT"
ElseIf IsLabelSymbolNew(rest$, " ¡»≈‘«", "PORTRAIT", Lang) Then
    If IsExp(bstack, rest$, x, , True) Then q(p).Vertical = Int(x) <> 0
ElseIf IsLabelSymbolNew(rest$, "≈… œÕ¡", "IMAGE", Lang) Then
If IsExp(bstack, rest$, x) Then
    If bstack.lastobj Is Nothing Then
    Set q(p).drawimage = Nothing
    ElseIf TypeOf bstack.lastobj Is mHandler Then
        Dim usehandler As mHandler
        Set usehandler = bstack.lastobj
        Set bstack.lastobj = Nothing
        If usehandler.t1 = 2 Then
        Set q(p).drawimage = usehandler.objref
        Else
        GoTo err123
        End If
     
    Else
err123:
        WrongObject
        Exit Sub
    End If
    End If
End If
Wend
RTarget bstack, q(p)
End If
End Sub
Friend Function AddTarget(t As target) As Long
            If UBound(q()) < 9999 Then
                Targets = False
                ReDim Preserve q(UBound(q()) + 1)
                q(UBound(q()) - 1) = t
                AddTarget = prive * 10000 + UBound(q()) - 1
                Targets = True
            End If
End Function
Friend Sub EnableTarget(bstack As basetask, ByVal tHandle As Variant, p As Variant)
        If tHandle \ 10000 = prive Then
        tHandle = tHandle Mod 10000
        q(tHandle).Enable = Not (p = 0)
        RTarget bstack, q(tHandle)
        End If
End Sub
Public Sub ClearTargets()
    Targets = False
    ReDim q(0) As target
End Sub

Public Property Get index() As Long
index = mIndex
End Property

Friend Property Let index(ByVal RHS As Long)
mIndex = RHS
End Property

Public Property Get Height() As Single
If Not pbox Is Nothing Then
    Height = pbox.Height
End If
End Property
Public Property Get Width() As Single
If Not pbox Is Nothing Then
    Width = pbox.Width
End If
End Property

