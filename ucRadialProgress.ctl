VERSION 5.00
Begin VB.UserControl ucRadialProgress 
   CanGetFocus     =   0   'False
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ScaleHeight     =   2010
   ScaleWidth      =   2160
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucRadialProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'simple, minimal (DPI-aware) implementation of a "pure GDI-based" circular Progress-Control
                'Olaf Schmidt, in May 2020

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC&, ByVal nIndex&) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC&, ByVal nStretchMode&) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDC&, ByVal x&, ByVal y&, Optional ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&) As Long
Private Declare Function TextOutW Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&, ByVal lpString&, ByVal nCount&) As Long

Private mPercVal#, mBB As VB.PictureBox, NumElm&, GapRatio!, InnerRadiusPerc!, TopColor&, BottomColor&

Private Sub UserControl_Initialize()
  ClipBehavior = 0: BackStyle = 0: FillStyle = 1: ScaleMode = vbPixels 'ensure transparent behaviour
  
  NumElm = 20: GapRatio = 0.15: InnerRadiusPerc = 0.75 'initialize a few internal "math-vars"
  TopColor = &HAA00&: BottomColor = vbWhite: FontName = "Arial" 'plus a few "formatting-vars"
End Sub

Public Sub ChangeDefaults(NumElements, StripeTopColor, Optional ByVal ElementGapRatio# = 0.15, Optional ByVal InnerRadiusRatio# = 0.75, Optional StripeBottomColor, Optional FontName, Optional TextColor)
  NumElm = NumElements: TopColor = mycolor(StripeTopColor): GapRatio = ElementGapRatio: InnerRadiusPerc = InnerRadiusRatio
  If Not IsMissing(StripeBottomColor) Then BottomColor = mycolor(StripeBottomColor)
  If Not IsMissing(FontName) Then UserControl.FontName = FontName
  If Not IsMissing(TextColor) Then UserControl.ForeColor = mycolor(TextColor)
  UserControl.Refresh
End Sub

Public Property Get Value() As Long
  Value = mPercVal * 100
End Property
Public Property Let Value(ByVal RHS As Long)
  If RHS < 0 Then RHS = 0 Else If RHS > 100 Then RHS = 100
  mPercVal = RHS / 100: UserControl.Refresh
End Property

Public Sub Refresh()
  UserControl.Refresh
End Sub

Private Sub UserControl_HitTest(x As Single, y As Single, HitResult As Integer)
  If UserControl.enabled Then HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Show()
  If Not (Ambient.UserMode And mBB Is Nothing) Then Exit Sub
  Set mBB = Parent.Controls.Add("VB.PictureBox", "mBB" & ObjPtr(Me))
      mBB.BorderStyle = 0: mBB.AutoRedraw = True: mBB.ScaleMode = vbPixels
End Sub

Private Sub UserControl_Paint()
Const D2R# = 1.74532925199433E-02, sc& = 6
Static a&, Cs#, Sn#, x1&(720), y1&(720), x2&(720), y2&(720) 'statics, to avoid re-allocs

  Dim cx#: cx = ScaleWidth / 2
  Dim cy#: cy = ScaleHeight / 2
  Dim r1#: r1 = IIf(cx < cy, cx, cy) - GetDeviceCaps(hDC, 88) / 96
  Dim r2#: r2 = r1 * InnerRadiusPerc
  Dim sL$: sL = Format$(mPercVal, "0%")
  
  FontSize = r2 * 0.4 / (GetDeviceCaps(hDC, 88) / 96)
  TextOutW hDC, cx - TextWidth(sL) * 0.48, cy - TextHeight(sL) * 0.5, StrPtr(sL), Len(sL)
 
  If mBB Is Nothing Then Circle (cx, cy), r1 - 1, TopColor: Exit Sub
  If mBB.Width <> ScaleX(cx * 2 * sc, 3, Parent.ScaleMode) Or mBB.Height <> ScaleY(cy * 2 * sc, 3, Parent.ScaleMode) Then _
     mBB.move 0, 0, ScaleX(cx * 2 * sc, 3, Parent.ScaleMode), ScaleY(cy * 2 * sc, 3, Parent.ScaleMode)
     mBB.DrawWidth = sc * 0.018 * Atn(1) * r1
     
  StretchBlt mBB.hDC, 0, 0, cx * 2 * sc, cy * 2 * sc, hDC, 0, 0, cx * 2, cy * 2, vbSrcCopy
    For a = 0 To 720 - 1
        Cs = Cos((a / 2 - 90 + GapRatio * 180 / NumElm) * D2R)
        Sn = Sin((a / 2 - 90 + GapRatio * 180 / NumElm) * D2R)
        x1(a) = sc * (cx + r1 * Cs): y1(a) = sc * (cy + r1 * Sn)
        x2(a) = sc * (cx + r2 * Cs): y2(a) = sc * (cy + r2 * Sn)
    Next

    mBB.ForeColor = TopColor 'first draw the circular strip up to the current Perc-Value
    For a = 0 To mPercVal * 720 - 1
      If (a Mod 720 / NumElm) < 720 / NumElm * (1 - GapRatio) Then _
         MoveTo mBB.hDC, x2(a), y2(a): LineTo mBB.hDC, x1(a), y1(a)
    Next
    mBB.ForeColor = BottomColor 'and the remaining percent-circle with the bottom-color
    For a = a To 720 - 1
      If (a Mod 720 / NumElm) < 720 / NumElm * (1 - GapRatio) Then _
         MoveTo mBB.hDC, x2(a), y2(a): LineTo mBB.hDC, x1(a), y1(a)
    Next
  SetStretchBltMode hDC, 4 '<- ensures good HalfTone-quality for the StretchBlt-call below
  StretchBlt hDC, 0, 0, cx * 2, cy * 2, mBB.hDC, 0, 0, cx * 2 * sc, cy * 2 * sc, vbSrcCopy
End Sub
