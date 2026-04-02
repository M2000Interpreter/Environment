VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' From Tools make changed property as default

Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Dim check1focus As Boolean
Event Changed(state As Boolean)
Dim mCaptext As String
Dim mValue As Boolean
Public Property Set Container(glistNN As gList)
Set glistN = glistNN
glistN.Clear
glistN.additem "  " + mCaptext
glistN.NoCaretShow = True
glistN.VerticalCenterText = True
glistN.restrictLines = 1
glistN.MenuItem 1, True, False, mValue
glistN.ListIndex = 0
glistN.LeftMarginPixels = glistN.HeightPixels
glistN.ShowMe
glistN.enabled = True
End Property
Public Property Let Caption(myCap As String)
'
mCaptext = myCap
If Not glistN Is Nothing Then
glistN.List(0) = "  " + mCaptext
glistN.ShowMe
End If
End Property
Public Property Let CheckReset(RHS As Boolean)
' without event
mValue = RHS
If Not glistN Is Nothing Then
glistN.ListSelected(0) = mValue
glistN.ShowMe
End If
End Property



Private Sub Class_Terminate()
Set glistN = Nothing
End Sub

Private Sub glistN_CheckGotFocus()

check1focus = True
End Sub

Private Sub gListN_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
'
If check1focus Then
glistN.FillThere thisHDC, thisrect, rgb(100, 200, 160)
Else
glistN.FillThere thisHDC, thisrect, rgb(200, 120, 60)
End If
skip = False
End Sub


Private Sub glistN_CheckLostFocus()
check1focus = False
glistN.ShowMe
End Sub
Private Sub glistN_MenuChecked(item As Long)
' always item is 1  so 1-1..0 (base 0)
RaiseEvent Changed(glistN.ListSelected(0))
End Sub
Public Property Get checked() As Boolean
Attribute checked.VB_UserMemId = 0
If Not glistN Is Nothing Then
checked = glistN.ListSelected(0)
Else
' error
End If
End Property

Sub Shutdown()
glistN.Shutdown
glistN.enabled = False
End Sub

Private Sub glistN_OnResize()
glistN.LeftMarginPixels = glistN.HeightPixels
End Sub
