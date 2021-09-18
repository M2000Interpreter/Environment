VERSION 5.00
Begin VB.Form TweakForm 
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   ClientHeight    =   6300
   ClientLeft      =   3000
   ClientTop       =   3000
   ClientWidth     =   7485
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "tweakprive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin M2000.gList DIS 
      Height          =   2265
      Left            =   285
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3465
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   3995
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin M2000.gList gList11 
      Height          =   1875
      Left            =   3840
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   3307
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      Backcolor       =   3881787
      ForeColor       =   14737632
   End
   Begin M2000.gList command1 
      Height          =   525
      Index           =   0
      Left            =   3720
      TabIndex        =   16
      Top             =   5790
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   926
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      ForeColor       =   16777215
   End
   Begin M2000.gList gList1 
      Height          =   315
      Left            =   330
      TabIndex        =   0
      Top             =   840
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList glist3 
      Height          =   1545
      Left            =   5325
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2895
      Visible         =   0   'False
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   2725
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
   End
   Begin M2000.gList gList4 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList gList5 
      Height          =   660
      Left            =   1560
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   1164
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
   End
   Begin M2000.gList gList6 
      Height          =   315
      Left            =   375
      TabIndex        =   6
      Top             =   2205
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList gList7 
      Height          =   315
      Left            =   2580
      TabIndex        =   7
      Top             =   2205
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList gList8 
      Height          =   315
      Left            =   4755
      TabIndex        =   8
      Top             =   2235
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList gList9 
      Height          =   315
      Left            =   330
      TabIndex        =   1
      Top             =   1305
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      ForeColor       =   16777215
   End
   Begin M2000.gList gList10 
      Height          =   315
      Left            =   2865
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      ForeColor       =   16777215
   End
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   873
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Backcolor       =   3881787
      ForeColor       =   16777215
      CapColor        =   16777215
   End
   Begin M2000.gList command1 
      Height          =   525
      Index           =   1
      Left            =   195
      TabIndex        =   14
      Top             =   5775
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   926
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      ForeColor       =   16777215
   End
   Begin M2000.gList command1 
      Height          =   525
      Index           =   2
      Left            =   3630
      TabIndex        =   15
      Top             =   5310
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   926
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
      ForeColor       =   16777215
   End
   Begin M2000.gList gList12 
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Top             =   1800
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList gList13 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   1800
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBar         =   0   'False
   End
   Begin M2000.gList gList14 
      Height          =   660
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   1164
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "TweakForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements InterPress
Private ex As Boolean
Private mcd As String
Private pen As Long
Public textbox2 As myTextBox
Public WithEvents combo1 As dropdownlist
Attribute combo1.VB_VarHelpID = -1
Public textbox3 As myTextBox
Public combo2 As dropdownlist
' new combo box
Public combo3 As dropdownlist
Public textbox4 As myTextBox
''
Public WithEvents tbPaper As myTextBox
Attribute tbPaper.VB_VarHelpID = -1
Public WithEvents tbPen As myTextBox
Attribute tbPen.VB_VarHelpID = -1
Public WithEvents tbSize As myTextBox
Attribute tbSize.VB_VarHelpID = -1
Public WithEvents tbLineSpacing As myTextBox
Attribute tbLineSpacing.VB_VarHelpID = -1
Public WithEvents checkbox1 As myCheckBox
Attribute checkbox1.VB_VarHelpID = -1
Public WithEvents checkbox2 As myCheckBox
Attribute checkbox2.VB_VarHelpID = -1
Dim myCommand As myButton
Dim myUnicode As myButton
Dim myCancel As myButton
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Dim Mysize As Single
Dim setupxy As Single
Dim Lx As Long, ly As Long, dr As Boolean, drmove As Boolean
Dim prevx As Long, prevy As Long
Dim a$
Dim bordertop As Long, borderleft As Long
Dim allheight As Long, allwidth As Long, itemWidth As Long, itemwidth3 As Long, itemwidth2 As Long
Dim height1 As Long, width1 As Long

Private Sub closeme_Click()
Unload Me
End Sub


Private Sub checkbox1_Changed(state As Boolean)
DIS.Font.bold = state
playall
End Sub



Private Sub combo1_AutoCompleteDone(ByVal this$)

playfontname this$
End Sub

Private Sub combo1_PickOther(ByVal this As String)
playfontname this$

End Sub


Private Sub Form_Activate()
Hook hWnd, Nothing
End Sub

Private Sub Form_Load()
DIS.enabled = True
AutoRedraw = True
Form_Load1
'AS A LABEL ONLY
gList11.TabStop = False
DIS.TabStop = False
gList2.TabStop = False
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)

If Button = 1 Then
    
    If lastfactor = 0 Then lastfactor = 1

    If bordertop < 150 Then
    If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then
    dr = True
    mousepointer = vbSizeNWSE
    Lx = x
    ly = y
    End If
    
    Else
    If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then
    dr = True
    mousepointer = vbSizeNWSE
    Lx = x
    ly = y
    End If
    End If

End If
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addX As Long, addy As Long, factor As Single, once As Boolean
If once Then Exit Sub
If Button = 0 Then dr = False: drmove = False
If bordertop < 150 Then
If (y > Height - 150 And y < Height) And (x > Width - 150 And x < Width) Then mousepointer = vbSizeNWSE Else If Not (dr Or drmove) Then mousepointer = 0
 Else
 If (y > Height - bordertop And y < Height) And (x > Width - borderleft And x < Width) Then mousepointer = vbSizeNWSE Else If Not (dr Or drmove) Then mousepointer = 0
End If
If dr Then



If bordertop < 150 Then

        If y < (Height - 150) Or y > Height Then addy = (y - ly)
     If x < (Width - 150) Or x > Width Then addX = (x - Lx)
     
Else
    If y < (Height - bordertop) Or y > Height Then addy = (y - ly)
        If x < (Width - borderleft) Or x > Width Then addX = (x - Lx)
    End If
    

    
   If Not ExpandWidth Then addX = 0
        If lastfactor = 0 Then lastfactor = 1
        factor = lastfactor

        
  
        once = True
        If Height > VirtualScreenHeight() Then addy = -(Height - VirtualScreenHeight()) + addy
        If Width > VirtualScreenWidth() Then addX = -(Width - VirtualScreenWidth()) + addX
        If (addy + Height) / height1 > 0.4 And ((Width + addX) / width1) > 0.4 Then
   
        If addy <> 0 Then SizeDialog = ((addy + Height) / height1)
        lastfactor = ScaleDialogFix(SizeDialog)


        If ((Width * lastfactor / factor + addX) / Height * lastfactor / factor) < (width1 / height1) Then
        addX = -Width * lastfactor / factor - 1
      
           End If

        If addX = 0 Then
        If lastfactor <> factor Then ScaleDialog lastfactor, Width
        Lx = x
        
        Else
        Lx = x * lastfactor / factor
         ScaleDialog lastfactor, (Width + addX) * lastfactor / factor
         End If

        
         
        
        LastWidth = Width
      gList2.HeadlineHeight = gList2.HeightPixels
        gList2.PrepareToShow
      gList3.PrepareToShow
      gList5.PrepareToShow
      DIS.PrepareToShow
        ly = ly * lastfactor / factor
        End If
        Else
        Lx = x
        ly = y
   
End If
once = False
End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)

If dr Then Me.mousepointer = 0
dr = False
End Sub
Sub ScaleDialog(ByVal factor As Single, Optional NewWidth As Long = -1)

lastfactor = factor
setupxy = 20 * factor

bordertop = 10 * dv15 * factor

borderleft = bordertop
allwidth = width1 * factor
allheight = height1 * factor
itemWidth = allwidth - 2 * borderleft
itemwidth3 = (itemWidth - 2 * borderleft) / 3
itemwidth2 = (itemWidth - borderleft) / 2
move Left, top, allwidth, allheight
FontTransparent = False  ' clear background  or false to write over
gList2.move borderleft, bordertop, itemWidth, bordertop * 3
gList2.FloatLimitTop = VirtualScreenHeight() - bordertop - bordertop * 3
gList2.FloatLimitLeft = VirtualScreenWidth() - borderleft * 3
gList1.move borderleft, bordertop * 6, itemWidth, bordertop * 3
gList3.move borderleft + itemWidth * 2 / 5, bordertop * 9, itemWidth * 3 / 5, bordertop * 18
gList9.move borderleft, bordertop * 10, itemWidth * 2 / 5 - borderleft, bordertop * 3
gList10.move borderleft + itemWidth * 2 / 5, bordertop * 10, itemWidth * 3 / 5, bordertop * 3
'gList4.Move borderleft, bordertop * 14, itemwidth2, bordertop * 3
gList4.move borderleft, bordertop * 14, itemWidth * 2 / 5 - borderleft * 2, bordertop * 3

' FOR NEW COMBO3
gList13.move itemWidth * 4 / 10, bordertop * 14, itemWidth * 3 / 10 + borderleft, bordertop * 3
''gList12.Move borderleft * 2 + itemwidth2, bordertop * 14, itemwidth2, bordertop * 3
gList12.move borderleft * 3 + itemWidth * 7 / 10 - borderleft, bordertop * 14, itemWidth * 3 / 10 - borderleft, bordertop * 3
gList12.ShowBar = True
gList5.move borderleft + itemWidth / 6, bordertop * 17, itemWidth / 5, bordertop * 6
' FOR NEW COMBO3
gList14.move borderleft + itemWidth / 2, bordertop * 17, itemWidth / 5, bordertop * 6
gList6.move borderleft, bordertop * 18, itemwidth3, bordertop * 3
gList6.ShowBar = True
gList7.move borderleft * 2 + itemwidth3, bordertop * 18, itemwidth3, bordertop * 3
gList7.ShowBar = True
gList8.move borderleft * 3 + itemwidth3 * 2, bordertop * 18, itemwidth3, bordertop * 3
gList8.ShowBar = True

DIS.move borderleft, bordertop * 22, itemwidth2, bordertop * 16
DIS.ScrollTo 0
DIS.ShowBar = False
gList11.move borderleft + itemwidth2 + borderleft, bordertop * 22, itemwidth2, bordertop * 16

command1(1).move borderleft, bordertop * 39, itemwidth3, bordertop * 3
command1(2).move borderleft + itemwidth3 + borderleft, bordertop * 39, itemwidth3, bordertop * 3
command1(0).move borderleft + itemwidth3 * 2 + borderleft * 2, bordertop * 39, itemwidth3, bordertop * 3
TwoColorsGradient Me, 0, rgb(100, 100, 100), rgb(100, 0, 0)
Set Me.Picture = Me.Image
gList11.ShowMe2
End Sub
Function ScaleDialogFix(ByVal factor As Single) As Single
gList2.FontSize = 14.25 * factor * dv15 / 15
factor = gList2.FontSize / 14.25 / dv15 * 15
gList1.FontSize = 11.25 * factor * dv15 / 15
factor = gList1.FontSize / 11.25 / dv15 * 15
gList3.FontSize = gList1.FontSize
gList4.FontSize = gList1.FontSize
gList5.FontSize = gList1.FontSize
gList6.FontSize = gList1.FontSize
gList7.FontSize = gList1.FontSize
gList8.FontSize = gList1.FontSize
gList9.FontSize = gList1.FontSize
command1(0).FontSize = gList1.FontSize
command1(1).FontSize = gList1.FontSize
command1(2).FontSize = gList1.FontSize
gList10.FontSize = gList1.FontSize
gList11.FontSize = gList1.FontSize
gList12.FontSize = gList1.FontSize
gList13.FontSize = gList1.FontSize
gList14.FontSize = gList1.FontSize
ScaleDialogFix = factor
End Function
Public Sub MyReFill()

Dim cc As Object
Set cc = New cRegistry
cc.ClassKey = HKEY_CURRENT_USER
    cc.SectionKey = basickey
    cc.ValueKey = "FONT"
        cc.ValueType = REG_SZ
   cc.Value = combo1.Text
    cc.ValueKey = "LINESPACE"
        cc.ValueType = REG_DWORD
      cc.Value = tbLineSpacing.Value * 2
    cc.ValueKey = "SIZE"
        cc.ValueType = REG_DWORD
      cc.Value = tbSize.Value

    cc.ValueKey = "PEN"
        cc.ValueType = REG_DWORD
       cc.Value = CLng(tbPen)
            cc.ValueKey = "BOLD"
        cc.ValueType = REG_DWORD
       cc.Value = CLng(DIS.Font.bold)
    cc.ValueKey = "PAPER"
        cc.ValueType = REG_DWORD
        cc.Value = CLng(tbPaper)
    
        cc.ValueKey = "COMMAND"
        cc.ValueType = REG_SZ
        cc.Value = UCase(combo2.Text)
        cc.ValueKey = "HTML"
        cc.ValueType = REG_SZ
        cc.Value = UCase(combo3.Text)
        cc.ValueKey = "CASESENSITIVE"
        cc.ValueType = REG_SZ
        If checkbox2.checked Then
        casesensitive = True
         cc.Value = "YES"
        Else
        casesensitive = False
        cc.Value = "NO"
        End If
        Set cc = Nothing
       
End Sub

Public Sub MyFill()
On Error Resume Next
Dim cc As Object
Set cc = New cRegistry
cc.ClassKey = HKEY_CURRENT_USER
    cc.SectionKey = basickey
    cc.ValueKey = "FONT"
        cc.ValueType = REG_SZ
        If cc.Value = vbNullString Then
        cc.Value = "Verdana"
        End If
        MyFont = cc.Value
    combo1.Text = cc.Value
Err.Clear
DIS.Font.Name = MyFont
If Err.Number > 0 Then
Err.Clear
MyFont = "Verdana"

End If
If DIS.Font.charset <> 161 Then
    Font.charset = myCharSet
    DIS.Font.charset = myCharSet
End If
'MsgBox "lets open the main screen"
'Show
    cc.ValueKey = "LINESPACE"
        cc.ValueType = REG_DWORD
        If cc.Value >= 0 And cc.Value <= 120 * dv15 Then
    tbLineSpacing = Int(cc.Value / 2)
    Else
    tbLineSpacing = 0
    End If
    
    cc.ValueKey = "SIZE"
        cc.ValueType = REG_DWORD
        If cc.Value = 0 Then
        cc.Value = 15
        SzOne = 15
        Else
        If cc.Value >= 8 And cc.Value <= 48 Then
        tbSize = cc.Value
        Else
        cc.Value = 15
        tbSize = 15
        End If
        End If
    cc.ValueKey = "BOLD"
        cc.ValueType = REG_DWORD
        checkbox1.CheckReset = CStr(cc.Value)
    cc.ValueKey = "PEN"
        cc.ValueType = REG_DWORD
        pen = cc.Value
    tbPen.enabled = False
        tbPen = CStr(pen)
        tbPen.Value = CStr(pen)
tbPen.enabled = True
      DIS.forecolor = QBColor(tbPen)
    cc.ValueKey = "PAPER"
        cc.ValueType = REG_DWORD
        tbPaper = CStr(cc.Value)
        tbPaper.Value = cc.Value
        DIS.backcolor = QBColor(cc.Value)
        cc.ValueKey = "COMMAND"
        cc.ValueType = REG_SZ

        combo2.additem "GREEK"
        combo2.additem "LATIN"
        
        
        
         If cc.Value = vbNullString Then
        cc.Value = "GREEK"
        
        End If
        combo2.Text = cc.Value
        If combo2.Text = "GREEK" Then
    DIS.Font.charset = 161
Else
DIS.Font.charset = 0
End If
     combo3.additem "BRIGHT"
        combo3.additem "DARK"
        cc.ValueKey = "HTML"
        cc.ValueType = REG_SZ
        If cc.Value = vbNullString Then
        cc.Value = "DARK"
        End If
        combo3.Text = cc.Value
DIS.Font.bold = checkbox1.checked
      cc.ValueKey = "CASESENSITIVE"
        cc.ValueType = REG_SZ
        
     
         Me.checkbox2.CheckReset = cc.Value = "YES"
        
        Set cc = Nothing
      
End Sub




Private Sub Combo1_Click()
On Error Resume Next
'DIS.Font.name = combo1.List(combo1.listindex)

If Err.Number > 0 Then
'combo1.Text = DIS.Font.name
End If
'DIS.Font.Size = MySize.Value
'DoEvents
'Combo2_Click
End Sub
Private Sub Command111_Click()
notweak = False
MyReFill
If Not IsSupervisor Then
If CFname(userfiles & "desktop.inf") <> "" Then
RenameFile userfiles & "desktop.inf", Format(Date, "YYYYMMDD") + Format$(Timer Mod 10000, "0000") + ".inf"
End If
End If
ShutMe
End Sub
Private Sub ShutMe()
myCommand.Shutdown
myUnicode.Shutdown
myCancel.Shutdown
combo1.Shutdown
combo2.Shutdown
combo3.Shutdown
Sleep 200
tbPaper.locked = True
tbPen.locked = True
tbSize.locked = True
tbLineSpacing.locked = True
checkbox1.Shutdown
checkbox2.Shutdown

Unload Me
End Sub
Private Sub Form_Load1()
Dim CD As String, dummy As Long, q$


Dim i, a$
DIS.NoCaretShow = True
DIS.LeftMarginPixels = 10
  
' Combobox1 SetUp
gList3.restrictLines = 6
Set textbox2 = New myTextBox
Set textbox2.Container = gList1

Set combo1 = New dropdownlist
combo1.UseOnlyTheList = True


Set combo1.Textbox = textbox2
Set combo1.Container = gList3
combo1.locked = False
combo1.AutoComplete = True
combo1.Textbox.ShowAlways = True
If TweakLang = 0 Then
combo1.Label = "Όνομα Γραμματοσειράς"
Else
combo1.Label = "Font name"
End If
'Mode edit but exist in list
textbox2.Retired
textbox2.enabled = True:   combo1.UseOnlyTheList = True

' Combobox2 SetUp
gList5.restrictLines = 2
Set textbox3 = New myTextBox
Set textbox3.Container = gList4
Set combo2 = New dropdownlist
combo2.UseOnlyTheList = True

textbox3.enabled = False
Set combo2.Textbox = textbox3
Set combo2.Container = gList5
combo2.locked = False
combo2.AutoComplete = True
If TweakLang = 0 Then
combo2.Label = "Τύπος γραμμάτων"
Else
combo2.Label = "Char type"
End If
textbox3.Retired
' Combobox3 SetUp
gList14.restrictLines = 2
Set textbox4 = New myTextBox
Set textbox4.Container = gList13
Set combo3 = New dropdownlist
combo3.UseOnlyTheList = True

textbox4.enabled = False
Set combo3.Textbox = textbox4
Set combo3.Container = gList14
combo3.locked = False
combo3.AutoComplete = True
If TweakLang = 0 Then
combo3.Label = "Χρώμα Html"
Else
combo3.Label = "Color Html"
End If
textbox4.Retired
'' continue for others

Dim aa As New Document, bb As New FastCollection

    For i = 0 To Screen.FontCount - 1
        If bb.ExistKey(Trim(Screen.Fonts(i))) Then
        Else
        aa.AppendParagraphOneLine Screen.Fonts(i)
        bb.AddKey Trim(Screen.Fonts(i))
        End If
       If ex Then Exit For
    Next i
    aa.SortDoc 1, 1, aa.DocLines
gList3.Text = aa.textDoc
 gList11.enabled = True
gList11.BackStyle = 1
gList11.FontSize = 10.25
gList11.NoCaretShow = True
gList11.restrictLines = 6
gList11.CenterText = True
gList11.VerticalCenterText = True

gList11.Text = "Warning: There is no " & vbCrLf & "warning about this " & vbCrLf & "software except that" & vbCrLf & "is given AS-IS" & vbCrLf & vbCrLf & "George Karras 1999-2020 ©"

height1 = 6450 * DYP / 15
width1 = 9900 * DXP / 15
lastfactor = 1
If ExpandWidth Then
If LastWidth = 0 Then LastWidth = -1
Else
LastWidth = -1
End If
FontName = "Arial"
gList2.enabled = True
gList2.CapColor = rgb(255, 160, 0)
gList2.FloatList = True
gList2.MoveParent = True
Form1.AutoRedraw = True
a$ = ListenUnicode(915, 953, 974, 961, 947, 959, 962, 32, 922, 945, 961, 961, 940, 962)
lastfactor = ScaleDialogFix(SizeDialog)
ScaleDialog lastfactor, LastWidth
gList2.HeadLine = vbNullString
If TweakLang = 0 Then
gList2.HeadLine = "Ρυθμίσεις"
Else
gList2.HeadLine = "Settings"
End If
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
Set checkbox1 = New myCheckBox

With checkbox1
If TweakLang = 0 Then
.Caption = "Φαρδιά"
Else
.Caption = "Bold"
End If
.CheckReset = True
Set .Container = gList9
End With

Set checkbox2 = New myCheckBox
With checkbox2
If TweakLang = 0 Then
.Caption = "Πεζά/κεφαλαία διαφορετικά σε αρχεία"
Else
.Caption = "Case Sensitive Filenames"
End If

.CheckReset = False
Set .Container = gList10
End With

Set tbPaper = New myTextBox
Set tbPaper.Container = gList6
If TweakLang = 0 Then
tbPaper.Prompt = "Χρώμα φόντου: "
Else
tbPaper.Prompt = "Paper Color: "
End If
tbPaper.Spinner True, 0, 15, 1
tbPaper.Value = 0
tbPaper.Retired
tbPaper.enabled = True

Set tbPen = New myTextBox
Set tbPen.Container = gList7
If TweakLang = 0 Then
tbPen.Prompt = "Χρώμα γραφής: "
Else
tbPen.Prompt = "Pen Color: "
End If
tbPen.Spinner True, 0, 15, 1
tbPen.Value = 15
tbPen.Retired
tbPen.enabled = True

Set tbSize = New myTextBox
Set tbSize.Container = gList8
If TweakLang = 0 Then
tbSize.Prompt = "Μέγεθος Γραμμάτων: "
Else
tbSize.Prompt = "Font Size: "

End If
tbSize.ThisKind = "pt"
tbSize.Spinner True, 8, 48, 1
tbSize.Value = 15
tbSize.Retired
tbSize.enabled = True

Set tbLineSpacing = New myTextBox
Set tbLineSpacing.Container = gList12
If TweakLang = 0 Then
tbLineSpacing.Prompt = "Διάστιχο: "

Else
tbLineSpacing.Prompt = "Line spacing: "
End If
tbLineSpacing.ThisKind = "twips"
tbLineSpacing.Spinner True, 0, 60 * dv15, 2 * dv15
tbLineSpacing.Value = 0
tbLineSpacing.Retired
tbLineSpacing.enabled = True

Set myCommand = New myButton
Set myCommand.Container = command1(0)
If TweakLang = 0 Then
myCommand.Caption = "Εντάξει"
Else
myCommand.Caption = "OK"
End If
  Set myCommand.Callback = Me
  myCommand.Index = 1
myCommand.enabled = True
Set myUnicode = New myButton
Set myUnicode.Container = command1(1)
If TweakLang = 0 Then
myUnicode.Caption = "Προεπισκόπηση Ansi"
Else
myUnicode.Caption = "Ansi Preview"
End If
  Set myUnicode.Callback = Me
myUnicode.enabled = True
Set myCancel = New myButton
Set myCancel.Container = command1(2)
myCancel.Index = 2
If TweakLang = 0 Then
myCancel.Caption = "ΑΚΥΡΟ"
Else
myCancel.Caption = "CANCEL"
End If
  Set myCancel.Callback = Me
myCancel.enabled = True
MyFill
 playall
 End Sub



Public Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT, b As Long
b = 2
CopyFromLParamToRect a, thatRect
a.Left = b
a.Right = setupxy - b
a.top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), 0
b = 5
a.Left = b
a.Right = setupxy - b
a.top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), rgb(255, 160, 0)


End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT
CopyFromLParamToRect a, thatRect
FillBack thathDC, a, thatbgcolor
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set LastGlist = Nothing
UnHook hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set myCommand = Nothing
Set myUnicode = Nothing
End Sub




Private Sub gList1_ChangeListItem(item As Long, content As String)
If combo1.ListText <> "" Then DIS.Font.Name = combo1.ListText: playall
End Sub

Private Sub gList11_SpecialColor(RGBcolor As Long)
RGBcolor = rgb(255, 200, 100)
End Sub



Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then
FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
skip = True
End If
End Sub

Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList2.DoubleClickCheck(Button, item, x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
                      ShutMe
End If
End Sub




Private Sub glist3_LostFocus()
playall
End Sub

Private Sub glist3_ScrollSelected(item As Long, y As Long)
playall
End Sub

Private Sub glist3_selected(item As Long)
playall
End Sub


Private Sub gList4_ChangeListItem(item As Long, content As String)
If gList11.enabled Then playall
End Sub
Private Sub playall()


On Error Resume Next

Dim cc$
If combo1.ListText <> "" Then cc$ = combo1.ListText Else cc$ = combo1.Text
playfontname cc$
End Sub
Private Sub playfontname(c$)
On Error Resume Next
Dim b$
If combo2.ListText <> "" Then b$ = combo2.ListText Else b$ = combo2.Text
DIS.Font.Italic = False
If b$ = "GREEK" Then
DIS.Font.Size = tbSize.Value
DIS.Font.charset = 161
DIS.Font.Name = c$
DIS.Font.Italic = False
DIS.Font.Size = tbSize.Value
DIS.Font.charset = 161
DIS.Font.Name = c$
Else
DIS.Font.Size = tbSize.Value
DIS.Font.charset = 0
DIS.Font.Name = c$
DIS.Font.Italic = False
DIS.Font.Size = tbSize.Value
DIS.Font.charset = 0
DIS.Font.Name = c$

End If

DIS.addpixels = (tbLineSpacing.Value * 2 \ dv15)
If InStr(myUnicode.Caption, "Unicode") = 0 Then
    DIS = Convert2Ansi("Test " & vbCrLf & "Latin" & vbCrLf + ListenUnicode(917, 955, 955, 951, 957, 953, 954, 940), IIf(DIS.charset = 161, 1032, 1033))
Else
    DIS = "Test " & vbCrLf & "Latin" & vbCrLf + ListenUnicode(917, 955, 955, 951, 957, 953, 954, 940)
End If
DIS.ShowMe2

End Sub

Private Sub InterPress_Press(Index As Long)
If Index = 1 Then

Command111_Click
'Unload Me  'remove this line
ElseIf Index = 2 Then
ShutMe
Else
If myUnicode.Caption = "Unicode Preview" Then
myUnicode.Caption = "Ansi Preview"
ElseIf myUnicode.Caption = "Προεπισκόπηση Unicode" Then
myUnicode.Caption = "Προεπισκόπηση Ansi"
ElseIf myUnicode.Caption = "Προεπισκόπηση Ansi" Then
myUnicode.Caption = "Προεπισκόπηση Unicode"
Else
myUnicode.Caption = "Unicode Preview"
End If
playall


End If
End Sub





Private Sub TBLineSpacing_SpinnerValue(ThisValue As Long)
tbLineSpacing = CStr(ThisValue)
End Sub

Private Sub TBLineSpacing_ValidString(ThatString As String, setpos As Long)
Dim a As Long, k As String
On Error Resume Next
k = tbLineSpacing
If ThatString = vbNullString Then ThatString = "0"
a = CLng(ThatString)

If Err.Number > 0 Then
tbLineSpacing.Value = CLng(tbLineSpacing)
ThatString = k: setpos = 1: tbLineSpacing.ResetPan
Exit Sub
End If
tbLineSpacing.Value = a
a = tbLineSpacing.Value  ' cut max or min

DIS.addpixels = (a * 2 \ dv15)
DIS.ShowMe2
'TBLineSpacing.Info = "This is info box" + vbCrLf + "X = " + CStr(a)
ThatString = CStr(a)
If a = 0 Then setpos = 2: tbLineSpacing.ResetPan
End Sub



Private Sub tbPaper_SpinnerValue(ThisValue As Long)
tbPaper = CStr(ThisValue)
End Sub

Private Sub tbPaper_ValidString(ThatString As String, setpos As Long)
Dim a As Long, k As String
On Error Resume Next
k = tbPaper
If ThatString = vbNullString Then ThatString = "0"
a = CLng(ThatString)
If a = CLng(tbPen) Or Err.Number > 0 Then

''tbPaper.Value = CLng(tbPaper)
ThatString = k: setpos = 1: tbPaper.ResetPan
If Abs(tbPaper.Value - CLng(k)) > 2 Then tbPaper.Value = CLng(k)
Exit Sub
End If
tbPaper.Value = a
a = tbPaper.Value  ' cut max or min
tbPaper.Value = a
DIS.backcolor = mycolor(a)
DIS.ShowMe2
'tbPaper.Info = "This is info box" + vbCrLf + "X = " + CStr(a)
ThatString = CStr(a)
If a = 0 Then setpos = 2: tbPaper.ResetPan
End Sub

Private Sub tbPen_SpinnerValue(ThisValue As Long)
tbPen = CStr(ThisValue)
End Sub

Private Sub tbpen_ValidString(ThatString As String, setpos As Long)
Dim a As Long, k As String
On Error Resume Next
k = tbPen
If ThatString = vbNullString Then ThatString = "0"
a = CLng(ThatString)
If a = CLng(tbPaper) Or Err.Number > 0 Then
ThatString = k: setpos = 1: tbPen.ResetPan
If Abs(tbPen.Value - CLng(k)) > 2 Then tbPen.Value = CLng(k)
Exit Sub
End If
tbPen.Value = a
a = tbPen.Value  ' cut max or min
tbPen.Value = a
DIS.forecolor = mycolor(a)
DIS.ShowMe2
'tbpen.Info = "This is info box" + vbCrLf + "X = " + CStr(a)
ThatString = CStr(a)
If a = 0 Then setpos = 2: tbPen.ResetPan
End Sub
Private Sub tbsize_SpinnerValue(ThisValue As Long)
tbSize = CStr(ThisValue)
End Sub

Private Sub tbsize_ValidString(ThatString As String, setpos As Long)
Dim a As Long, k As String
On Error Resume Next
k = tbSize
If ThatString = vbNullString Then Exit Sub  'special here
a = CLng(ThatString)

If Err.Number > 0 Then
tbSize.Value = CLng(tbSize)
ThatString = k: setpos = 1: tbSize.ResetPan
Exit Sub
End If

tbSize.Value = a
If a <> tbSize.Value And a <= 2 And a > 0 Then
Exit Sub
End If
a = tbSize.Value  ' cut max or min

DIS.FontSize = CLng(a)
playall
ThatString = CStr(a)
If a = 0 Then setpos = 2: tbSize.ResetPan
End Sub
Public Sub hookme(this As gList)
Set LastGlist = this
End Sub
Private Sub gList2_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub
