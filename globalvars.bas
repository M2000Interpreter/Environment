Attribute VB_Name = "globalvars"
' This is for selectors..
Option Explicit
Public AskTitle$, AskText$, AskCancel$, AskOk$, AskDIB$, ASKINUSE As Boolean
Public AskInput As Boolean, AskResponse$, AskStrInput$, AskDIBicon$
Public UseAskForMultipleEntry As Boolean
Public BreakMe As Boolean
Public CancelDialog As Boolean
Public SizeDialog As Single, helpSizeDialog As Single
Public textinformCaption As String
Public FileTypesShow As String
Public ReturnFile As String
Public ReturnListOfFiles As String  ' # between
Public Settings As String
Public TopFolder As String
Public AskLastX As Long, AskLastY As Long
Public selectorLastX As Long, selectorLastY As Long
Public FolderOnly As Boolean
Public AskCancelGR As String
Public AskOkGR As String
Public LoadFileCaptionGR As String
Public SaveFileCaptionGR As String
Public SelectFolderCaptionGR As String
Public SelectFolderButtonGR As String
Public FontSelectorGr As String
Public ColorSelectorGr As String
Public SetUpGR As String
Public AskCancelEn As String
Public AskOkEn As String
Public SetUpEn As String
Public LoadFileCaptionEn As String
Public SaveFileCaptionEn As String
Public SelectFolderCaptionEn As String
Public SelectFolderButtonEn As String
Public FontSelectorEn As String
Public ColorSelectorEn As String
Public SetUp As String
Public LoadFileCaption As String
Public SaveFileCaption As String
Public SelectFolderCaption As String
Public SelectFolderButton As String
Public FontSelector As String
Public ColorSelector As String
Public SaveDialog As Boolean
Public DialogPreview As Boolean, LastWidth As Long, HelpLastWidth As Long, PopUpLastWidth As Long
Public ExpandWidth As Boolean, lastfactor As Single, Helplastfactor As Single, Pouplastfactor As Single
Public NewFolder As Boolean, multifileselection As Boolean
Public FileExist As Boolean
Public UserFileName As String
Private inUse As Boolean
Public ReturnColor As Double
Public ReturnFontName As String
Public ReturnBold As Boolean
Public ReturnItalic As Boolean
Public ReturnCharset As Integer
Public ReturnSize As Single
Public DialogLang As Long
Public Sub DialogSetupLang(Lang As Long)
DialogLang = Lang
If Lang = 0 Then
AskCancel$ = AskCancelGR
AskOk$ = AskOkGR
 LoadFileCaption = LoadFileCaptionGR
 SaveFileCaption = SaveFileCaptionGR
 SelectFolderCaption = SelectFolderCaptionGR
 SelectFolderButton = SelectFolderButtonGR
  FontSelector = FontSelectorGr
ColorSelector = ColorSelectorGr
 SetUp = SetUpGR
Else
AskCancel$ = AskCancelEn
AskOk$ = AskOkEn
 LoadFileCaption = LoadFileCaptionEn
 SaveFileCaption = SaveFileCaptionEn
 SelectFolderCaption = SelectFolderCaptionEn
 SelectFolderButton = SelectFolderButtonEn
  FontSelector = FontSelectorEn
ColorSelector = ColorSelectorEn
 SetUp = SetUpEn
End If
End Sub
Public Function IsSelectorInUse() As Boolean
IsSelectorInUse = inUse
End Function
Public Function OpenColor(bstack As basetask, thiscolor As Long) As Boolean
If inUse Then OpenColor = False: Exit Function
inUse = True
Dim thisform As Object
If Not Screen.ActiveForm Is Nothing Then
If TypeOf Screen.ActiveForm Is GuiM2000 Then
    Set thisform = Screen.ActiveForm
ElseIf Screen.ActiveForm Is MyPopUp Then
    Set thisform = MyPopUp.LASTActiveForm
    MyPopUp.Hide
Else
    Set thisform = Screen.ActiveForm
End If

End If
ExpandWidth = True
ReturnColor = thiscolor

If thisform Is Nothing Then
ColorDialog.Show
Else
ColorDialog.Show , thisform
End If
Dim Scr As Object
Set Scr = bstack.Owner
If TypeOf Scr Is GuiM2000 Then
    Scr.UNhookMe
ElseIf val("0" + bstack.Owner.Tag) > 32 Then
    Set Scr = bstack.Owner.Parent
    If Scr Is Nothing Then Exit Function
    While Not TypeOf Scr Is GuiM2000
        Set Scr = Scr.Parent
        If Scr Is Nothing Then Exit Function
    Wend
    Scr.UNhookMe
ElseIf Not Screen.ActiveForm Is Nothing Then
    Screen.ActiveForm.UNhookMe
End If

If Not Screen.ActiveForm Is Nothing Then
If Not Screen.ActiveForm Is ColorDialog Then
ColorDialog.Show , Screen.ActiveForm
End If
End If
MoveFormToOtherMonitorOnly ColorDialog, Scr.Name = "GuiM2000"
CancelDialog = False
If Not ColorDialog.Visible Then
    ColorDialog.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If Not thisform Is Nothing Then
        If thisform.Visible Then
            If Typename(thisform) = "GuiM2000" Then
                thisform.Enablecontrol = True
                'Thisform.ShowmeALL
            End If
            thisform.SetFocus
        End If
End If
OpenColor = Not CancelDialog
thiscolor = ReturnColor
ExpandWidth = False
inUse = False
End Function
Public Function OpenFont(bstack As basetask) As Boolean
If inUse Then OpenFont = False: Exit Function
inUse = True
Dim thisform As Object
If Not Screen.ActiveForm Is Nothing Then
If TypeOf Screen.ActiveForm Is GuiM2000 Then
    Set thisform = Screen.ActiveForm
ElseIf Screen.ActiveForm Is MyPopUp Then
    Set thisform = MyPopUp.LASTActiveForm
    MyPopUp.Hide
Else
    Set thisform = Screen.ActiveForm
End If

End If
ExpandWidth = True
Dim Scr As Object
Set Scr = bstack.Owner
If TypeOf Scr Is GuiM2000 Then
    Scr.UNhookMe
ElseIf val("0" + bstack.Owner.Tag) > 32 Then
    Set Scr = bstack.Owner.Parent
    If Scr Is Nothing Then Exit Function
    While Not TypeOf Scr Is GuiM2000
        Set Scr = Scr.Parent
        If Scr Is Nothing Then Exit Function
    Wend
    Scr.UNhookMe
ElseIf Not Screen.ActiveForm Is Nothing Then
    Screen.ActiveForm.UNhookMe
End If

If thisform Is Nothing Then
FontDialog.Show
Else
FontDialog.Show , thisform
End If
If Not Screen.ActiveForm Is Nothing Then
If Not Screen.ActiveForm Is FontDialog Then
FontDialog.Show , Screen.ActiveForm
End If
End If
MoveFormToOtherMonitorOnly FontDialog, Scr.Name = "GuiM2000"
CancelDialog = False
If Not FontDialog.Visible Then
    FontDialog.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If Not thisform Is Nothing Then
        If thisform.Visible Then
            If Typename(thisform) = "GuiM2000" Then
                thisform.Enablecontrol = True
                'Thisform.ShowmeALL
            End If
            thisform.SetFocus
        End If
End If
If ReturnFontName <> "" Then OpenFont = Not CancelDialog
ExpandWidth = False
inUse = False
End Function
Public Function OpenImage(bstack As basetask, TopDir As String, lastname As String, thattitle As String, TypeList As String) As Boolean
If inUse Then OpenImage = False: Exit Function
inUse = True
Dim thisform As Object
If Not Screen.ActiveForm Is Nothing Then
If TypeOf Screen.ActiveForm Is GuiM2000 Then
    Set thisform = Screen.ActiveForm
ElseIf Screen.ActiveForm Is MyPopUp Then
    Set thisform = MyPopUp.LASTActiveForm
    MyPopUp.Hide
Else
    Set thisform = Screen.ActiveForm
End If

End If
' do something with multifiles..
ReturnFile = lastname
If ReturnFile <> "" Then If ExtractPath(lastname) = vbNullString Then ReturnFile = mcd + lastname
SaveDialog = False
FileExist = True
FolderOnly = False
''If TopDir <> "" Then TopFolder = TopDir
If TopDir = vbNullString Then
TopFolder = mcd
ReturnFile = mcd
ElseIf TopDir = "\" Then
TopFolder = vbNullString
ReturnFile = mcd
ElseIf TopDir = "*" Then
TopFolder = vbNullString
ReturnFile = vbNullString

Else
TopFolder = TopDir
End If
ReturnListOfFiles = vbNullString
If TypeList = vbNullString Then FileTypesShow = "BMP|JPG|GIF|WMF|EMF|DIB|ICO|CUR|PNG|TIF" Else FileTypesShow = TypeList
DialogPreview = True
If thattitle <> "" Then
LoadFileCaption = thattitle
If InStr(Settings, ",expand") = 0 Then
Settings = Settings & ",expand"
End If
End If
Dim Scr As Object
Set Scr = bstack.Owner
If TypeOf Scr Is GuiM2000 Then
    Scr.UNhookMe
ElseIf val("0" + bstack.Owner.Tag) > 32 Then
    Set Scr = bstack.Owner.Parent
    If Scr Is Nothing Then Exit Function
    While Not TypeOf Scr Is GuiM2000
        Set Scr = Scr.Parent
        If Scr Is Nothing Then Exit Function
    Wend
    Scr.UNhookMe
ElseIf Not Screen.ActiveForm Is Nothing Then
    Screen.ActiveForm.UNhookMe
End If

If thisform Is Nothing Then
LoadFile.Show
Else
LoadFile.Show , thisform
End If
If Not Screen.ActiveForm Is Nothing Then
If Not Screen.ActiveForm Is LoadFile Then
LoadFile.Show , Screen.ActiveForm
End If
End If
MoveFormToOtherMonitorOnly LoadFile, Scr.Name = "GuiM2000"
CancelDialog = False
If Not LoadFile.Visible Then
    LoadFile.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If Not thisform Is Nothing Then
        If thisform.Visible Then
            If Typename(thisform) = "GuiM2000" Then
                thisform.Enablecontrol = True
                'Thisform.ShowmeALL
            End If
            thisform.SetFocus
        End If
End If
If ReturnListOfFiles <> "" Or ReturnFile <> "" Then OpenImage = Not CancelDialog
inUse = False

' read files
End Function
Public Function OpenDialog(bstack As basetask, TopDir As String, lastname As String, thattitle As String, TypeList As String, OpenNew As Boolean, MULTFILES As Boolean) As Boolean
If inUse Then OpenDialog = False: Exit Function
Dim foundmulti As Boolean, Scr As Object
inUse = True
Dim thisform As Object
If Not Screen.ActiveForm Is Nothing Then
If TypeOf Screen.ActiveForm Is GuiM2000 Then
    Set thisform = Screen.ActiveForm
ElseIf Screen.ActiveForm Is MyPopUp Then
    Set thisform = MyPopUp.LASTActiveForm
    MyPopUp.Hide
Else
    Set thisform = Screen.ActiveForm
End If

End If
' do something with multifiles..
ReturnFile = lastname
If ReturnFile <> "" Then If ExtractPath(lastname) = vbNullString Then ReturnFile = mcd + lastname
SaveDialog = False
FileExist = OpenNew
FolderOnly = False
' If TopDir <> "" Then TopFolder = TopDir
If TopDir = vbNullString Then
TopFolder = mcd
ReturnFile = mcd
ElseIf TopDir = "\" Then
TopFolder = vbNullString
ReturnFile = mcd
ElseIf TopDir = "*" Then
TopFolder = vbNullString
ReturnFile = vbNullString

Else
TopFolder = TopDir
End If
ReturnListOfFiles = vbNullString
FileTypesShow = TypeList
DialogPreview = False
If thattitle <> "" Then
LoadFileCaption = thattitle
If InStr(Settings, ",expand") = 0 Then
Settings = Settings & ",expand"
End If
End If
If MULTFILES Then

If InStr(Settings, ",multi") = 0 Then
Settings = Settings & ",multi"
Else
foundmulti = True
End If
End If
Set Scr = bstack.Owner
If TypeOf Scr Is GuiM2000 Then
    Scr.UNhookMe
ElseIf val("0" + bstack.Owner.Tag) > 32 Then
    Set Scr = bstack.Owner.Parent
    If Scr Is Nothing Then Exit Function
    While Not TypeOf Scr Is GuiM2000
        Set Scr = Scr.Parent
        If Scr Is Nothing Then Exit Function
    Wend
    Scr.UNhookMe
ElseIf Not Screen.ActiveForm Is Nothing Then
    Screen.ActiveForm.UNhookMe
End If
If thisform Is Nothing Then
LoadFile.Show
Else
LoadFile.Show , thisform
End If
If Not Screen.ActiveForm Is Nothing Then
If Not Screen.ActiveForm Is LoadFile Then
LoadFile.Show , Screen.ActiveForm
End If
End If
MoveFormToOtherMonitorOnly LoadFile, Scr.Name = "GuiM2000"
CancelDialog = False
If Not LoadFile.Visible Then
    LoadFile.Visible = True
    MyDoEvents
    End If
'Hook3 LoadFile.hWnd, Nothing
WaitDialog bstack
If Not thisform Is Nothing Then
        If thisform.Visible Then
            If Typename(thisform) = "GuiM2000" Then
                thisform.Enablecontrol = True
                'Thisform.ShowmeALL
            End If
            thisform.SetFocus
        End If
End If
If ReturnListOfFiles <> "" Or ReturnFile <> "" Then OpenDialog = Not CancelDialog
If MULTFILES And Not foundmulti Then

Settings = Replace(Settings, ",multi", "")

End If
inUse = False
' read files
End Function
Public Function SaveAsDialog(bstack As basetask, lastname As String, TopDir As String, thattitle As String, TypeList As String) As Boolean
Dim Scr As Object
If inUse Then SaveAsDialog = False: Exit Function
inUse = True
Dim thisform As Object
If Not Screen.ActiveForm Is Nothing Then
If TypeOf Screen.ActiveForm Is GuiM2000 Then
    Set thisform = Screen.ActiveForm
ElseIf Screen.ActiveForm Is MyPopUp Then
    Set thisform = MyPopUp.LASTActiveForm
    MyPopUp.Hide
Else
    Set thisform = Screen.ActiveForm
End If

End If
DialogPreview = False
FileExist = False
NewFolder = False
FolderOnly = False
SaveDialog = True
UserFileName = lastname
'ReturnFile = ExtractPath(LastName)
ReturnFile = lastname
If ReturnFile <> "" Then If ExtractPath(lastname) = vbNullString Then ReturnFile = mcd + lastname
FileTypesShow = TypeList
''If TopDir <> "" Then TopFolder = TopDir
If TopDir = vbNullString Then
TopFolder = mcd
ReturnFile = mcd
ElseIf TopDir = "\" Then
TopFolder = vbNullString
ReturnFile = mcd
ElseIf TopDir = "*" Then
TopFolder = vbNullString
ReturnFile = vbNullString

Else
TopFolder = TopDir
End If
If ReturnFile = vbNullString Then ReturnFile = TopDir + ExtractName(lastname)
If thattitle <> "" Then
SaveFileCaption = thattitle
If InStr(Settings, ",expand") = 0 Then
Settings = Settings & ",expand"
End If
End If
Set Scr = bstack.Owner
If TypeOf Scr Is GuiM2000 Then
    Scr.UNhookMe
ElseIf val("0" + bstack.Owner.Tag) > 32 Then
    Set Scr = bstack.Owner.Parent
    If Scr Is Nothing Then Exit Function
    While Not TypeOf Scr Is GuiM2000
        Set Scr = Scr.Parent
        If Scr Is Nothing Then Exit Function
    Wend
    Scr.UNhookMe
ElseIf Not Screen.ActiveForm Is Nothing Then
    Screen.ActiveForm.UNhookMe
End If



If thisform Is Nothing Then
LoadFile.Show
Else
LoadFile.Show , thisform
End If
If Not Screen.ActiveForm Is Nothing Then
If Not Screen.ActiveForm Is LoadFile Then
LoadFile.Show , Screen.ActiveForm
End If
End If
MoveFormToOtherMonitorOnly LoadFile, Scr.Name = "GuiM2000"
 CancelDialog = False
 If Not LoadFile.Visible Then
    LoadFile.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If Not thisform Is Nothing Then
        If thisform.Visible Then
            If Typename(thisform) = "GuiM2000" Then
                thisform.Enablecontrol = True
                'Thisform.ShowmeALL
            End If
            thisform.SetFocus
        End If
End If
If ReturnFile <> "" Then SaveAsDialog = Not CancelDialog
inUse = False
End Function
Public Function GetFile(bstack As basetask, thistitle As String, thisfolder As String, onetype As String, Optional multifiles As Boolean = False) As String
    If OpenDialog(bstack, thisfolder, "", thistitle, onetype, False, multifiles) Then
        GetFile = ReturnFile
    End If
End Function

Public Function FolderSelector(bstack As basetask, thatfolder As String, TopDir As String, thattitle As String, newflag As Boolean) As Boolean
If inUse Then FolderSelector = False: Exit Function
inUse = True
Dim thatform As Object
If Not Screen.ActiveForm Is Nothing Then
If TypeOf Screen.ActiveForm Is GuiM2000 Then
    Set thatform = Screen.ActiveForm
ElseIf Screen.ActiveForm Is MyPopUp Then
    Set thatform = MyPopUp.LASTActiveForm
    MyPopUp.Hide
Else
    Set thatform = Screen.ActiveForm
End If
End If

DialogPreview = False
ReturnFile = thatfolder
SaveDialog = False
NewFolder = newflag
FolderOnly = True
FileExist = True
If thattitle <> "" Then
SelectFolderCaption = thattitle
If InStr(Settings, ",expand") = 0 Then
Settings = Settings & ",expand"
End If
End If
If NewFolder Then FileExist = False
If TopDir = vbNullString Then
TopFolder = mcd
ReturnFile = mcd
ElseIf TopDir = "\" Then
TopFolder = vbNullString
ReturnFile = mcd
ElseIf TopDir = "*" Then
TopFolder = vbNullString
ReturnFile = vbNullString

Else
TopFolder = TopDir
End If
Dim Scr As Object
Set Scr = bstack.Owner
If TypeOf Scr Is GuiM2000 Then
    Scr.UNhookMe
ElseIf val("0" + bstack.Owner.Tag) > 32 Then
    Set Scr = bstack.Owner.Parent
    If Scr Is Nothing Then Exit Function
    While Not TypeOf Scr Is GuiM2000
        Set Scr = Scr.Parent
        If Scr Is Nothing Then Exit Function
    Wend
    Scr.UNhookMe
ElseIf Not Screen.ActiveForm Is Nothing Then
    Screen.ActiveForm.UNhookMe
End If

If thatform Is Nothing Then
LoadFile.Show
Else
LoadFile.Show , thatform
End If
If Not Screen.ActiveForm Is Nothing Then
If Not Screen.ActiveForm Is LoadFile Then
LoadFile.Show , Screen.ActiveForm
End If
End If
MoveFormToOtherMonitorOnly LoadFile, Scr.Name = "GuiM2000"
CancelDialog = False
If Not LoadFile.Visible Then
    LoadFile.Visible = True
    MyDoEvents
    End If
WaitDialog bstack
If Not thatform Is Nothing Then
        If thatform.Visible Then
            If Typename(thatform) = "GuiM2000" Then
                thatform.Enablecontrol = True
                'ThatForm.ShowmeALL
            End If
            thatform.SetFocus
        End If
End If
If ReturnFile <> "" Then FolderSelector = Not CancelDialog
inUse = False
End Function
Sub ReleaseSelector()
inUse = False
End Sub
Function ConCat(ParamArray aa() As Variant) As String
Dim all$, i As Long
For i = 0 To UBound(aa)
    all$ = all$ & CStr(aa(i))
Next i
ConCat = all$
End Function
