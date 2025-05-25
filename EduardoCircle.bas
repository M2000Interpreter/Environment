Attribute VB_Name = "Module11"
'' Addition


Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long, ByVal nXStartArc As Long, ByVal nYStartArc As Long, ByVal nXEndArc As Long, ByVal nYEndArc As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long, ByVal nXStartArc As Long, ByVal nYStartArc As Long, ByVal nXEndArc As Long, ByVal nYEndArc As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nDrawStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Type TRIVERTEX
    X As Long
    Y As Long
    r1 As Byte
    Red As Byte 'Ushort value
    G1 As Byte
    Green As Byte 'Ushort value
    b1 As Byte
    Blue As Byte 'ushort value
    Alpha As Integer 'ushort
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type
Private Type GRADIENT_TRIANGLE
  Vertex1 As Long
  Vertex2 As Long
  Vertex3 As Long
End Type
Public Const GRADIENT_FILL_RECT_H As Long = &H0
Public Const GRADIENT_FILL_RECT_V  As Long = &H1
Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, _
         hRgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Byte)
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal addr As Long, retval As Byte)

' Sub Circle(Step As Integer, iX As Single, iY As Single, Radius As Single, Color As Long, StartArc As Single, EndArc As Single, Aspect As Single)
' When an arc or a partial circle or ellipse is drawn, StartArc and EndArc specify (in radians) the beginning and end positions of the arc.
' The range for both is 2 pi radians to 2 pi radians. The default value for StartArc is 0 radians; the default for EndArc is 2 * pi radians.
' Sub Circle(Step As Integer, iX As Single, iY As Single, Radius As Single, Color As Long, StartArc As Single, EndArc As Single, Aspect As Single)
' When an arc or a partial circle or ellipse is drawn, StartArc and EndArc specify (in radians) the beginning and end positions of the arc.
' The range for both is 2 pi radians to 2 pi radians. The default value for StartArc is 0 radians; the default for EndArc is 2 * pi radians.
Public Sub DrawCircleApi(Scr As Object, X As Single, Y As Single, Radius As Single, Optional color, Optional Aspect As Single = 1, Optional StartArc, Optional EndArc)
    Dim iXStartArc As Long, iYStartArc As Long, iXEndArc As Long, iYEndArc As Long
    Dim iAspectX As Single
    Dim iAspectY As Single
    Dim iStartArc As Single
    Dim iEndArc As Single
    Dim iFilledFigure As Boolean
    Dim iColor As Long
    Dim iPen As Long
    Dim iPenPrev As Long
    Dim iX As Long
    Dim iY As Long
    Dim iStartArcIsNegative As Boolean
    Dim iEndArcIsNegative As Boolean
    Dim iPoints(1) As POINTAPI

    iX = X
    iY = Y
    
    If IsMissing(color) Then
        iColor = Scr.ForeColor
    Else
        iColor = color
    End If
    TranslateColor iColor, 0, iColor
    
    
    ' API
    If Aspect > 1 Then
        iAspectX = 1 / Aspect
        iAspectY = 1
    Else
        iAspectX = 1
        iAspectY = 1 * Aspect
    End If
    
    If IsMissing(StartArc) Then
        iStartArc = 0
    Else
        iStartArcIsNegative = StartArc < 0
        iStartArc = Abs(StartArc)
    End If
    If IsMissing(EndArc) Then
        iEndArc = 0
        ' Note: 0 (zero) for EndArc seems to be handled as 2 * Pi by the API (in fact they are the same point)
    Else
        iEndArcIsNegative = EndArc < 0
        iEndArc = Abs(EndArc)
    End If
  
    If (IsMissing(StartArc) And IsMissing(EndArc)) Or (iStartArcIsNegative And iEndArcIsNegative) Then
        If Scr.FillStyle = vbFSSolid Then
            iFilledFigure = True
        End If
    End If
  
    
    
    iXStartArc = Radius * iAspectX * Cos(iStartArc) + iX
    iYStartArc = Radius * iAspectY * Sin(iStartArc) * -1 + iY
    iXEndArc = Radius * iAspectX * Cos(iEndArc) + iX
    iYEndArc = Radius * iAspectY * Sin(iEndArc) * -1 + iY
    
   'If iColor <> Scr.forecolor Then  ' not used in M2000
        iPen = CreatePen(Scr.DrawStyle, Scr.DrawWidth, iColor)
        iPenPrev = SelectObject(Scr.hDC, iPen)
    'End If
    
    If iFilledFigure Then
        If iStartArcIsNegative Then
            Pie Scr.hDC, iX - Radius * iAspectX, iY - Radius * iAspectY, iX + Radius * iAspectX, iY + Radius * iAspectY, iXStartArc, iYStartArc, iXEndArc, iYEndArc
        Else
            Ellipse Scr.hDC, iX - Radius * iAspectX, iY - Radius * iAspectY, iX + Radius * iAspectX, iY + Radius * iAspectY
        End If
    Else
        Arc Scr.hDC, iX - Radius * iAspectX, iY - Radius * iAspectY, iX + Radius * iAspectX, iY + Radius * iAspectY, iXStartArc, iYStartArc, iXEndArc, iYEndArc
        If iStartArcIsNegative Then
            iPoints(0).X = iX
            iPoints(0).Y = iY
            iPoints(1).X = iXStartArc
            iPoints(1).Y = iYStartArc
            Polyline Scr.hDC, iPoints(0), 2
        End If
        If iEndArcIsNegative Then
            iPoints(0).X = iX
            iPoints(0).Y = iY
            iPoints(1).X = iXEndArc
            iPoints(1).Y = iYEndArc
            Polyline Scr.hDC, iPoints(0), 2
        End If
    End If

    If iPenPrev <> 0 Then
        Call SelectObject(Scr.hDC, iPenPrev)
    End If
    If iPen <> 0 Then
        DeleteObject iPen
    End If
    

End Sub
Public Sub TwoColorsGradient(Scr As Object, ByVal typegrad As Long, ByVal col1 As Long, ByVal Col2 As Long)
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    Dim gTriangle(1) As GRADIENT_TRIANGLE
    Dim bt As Byte
    With vert(0)
        GetMem1 VarPtr(col1), bt
        PutMem1 VarPtr(.Blue), bt
        GetMem1 VarPtr(col1) + 1, bt
        PutMem1 VarPtr(.Green), bt
        GetMem1 VarPtr(col1) + 2, bt
        PutMem1 VarPtr(.Red), bt
        GetMem1 VarPtr(col1) + 3, bt
        PutMem1 VarPtr(.Alpha), bt
    End With

    
    With vert(1)
        .X = Scr.ScaleWidth \ dv15
        .Y = Scr.ScaleHeight \ dv15
        GetMem1 VarPtr(Col2), bt
        PutMem1 VarPtr(.Blue), bt
        GetMem1 VarPtr(Col2) + 1, bt
        PutMem1 VarPtr(.Green), bt
        GetMem1 VarPtr(Col2) + 2, bt
        PutMem1 VarPtr(.Red), bt
        GetMem1 VarPtr(Col2) + 3, bt
        PutMem1 VarPtr(.Alpha), bt
    End With
    
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    With gTriangle(0)
        .Vertex1 = 0
        .Vertex2 = 1
    End With
    GradientFillRect Scr.hDC, vert(0), 2, gTriangle(0), 1, typegrad
    
End Sub
Public Sub TwoColorsGradientPart(Scr As Object, ByVal all As Boolean, ByVal typegrad As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, ByVal col1 As Long, ByVal Col2 As Long)
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT
    Dim gTriangle(1) As GRADIENT_TRIANGLE
    Dim bt As Byte, scNow As Integer
    scNow = Scr.ScaleMode
    
    If all Then
    With vert(0)
        .X = Scr.ScaleX(x1, scNow, 3)
        .Y = Scr.ScaleX(y1, scNow, 3)
        GetMem1 VarPtr(col1), bt
        PutMem1 VarPtr(.Red), bt
        GetMem1 VarPtr(col1) + 1, bt
        PutMem1 VarPtr(.Green), bt
        GetMem1 VarPtr(col1) + 2, bt
        PutMem1 VarPtr(.Blue), bt
     '   PutMem1 VarPtr(.Alpha),  transp And &HFF
    End With

    
    With vert(1)
        .X = Scr.ScaleX(x2, scNow, 3)
        .Y = Scr.ScaleX(y2, scNow, 3)
        GetMem1 VarPtr(Col2), bt
        PutMem1 VarPtr(.Red), bt
        GetMem1 VarPtr(Col2) + 1, bt
        PutMem1 VarPtr(.Green), bt
        GetMem1 VarPtr(Col2) + 2, bt
        PutMem1 VarPtr(.Blue), bt
    End With
gRect.UpperLeft = 0
    gRect.LowerRight = 1
    With gTriangle(0)
        .Vertex1 = 0
        .Vertex2 = 1
    End With
    GradientFillRect Scr.hDC, vert(0), 2, gTriangle(0), 1, typegrad
    Else
    With vert(0)
  '      .x = 0
   '     .y = 0
        GetMem1 VarPtr(col1), bt
        PutMem1 VarPtr(.Red), bt
        GetMem1 VarPtr(col1) + 1, bt
        PutMem1 VarPtr(.Green), bt
        GetMem1 VarPtr(col1) + 2, bt
        PutMem1 VarPtr(.Blue), bt
     '   PutMem1 VarPtr(.Alpha),  transp And &HFF
    End With

    
    With vert(1)
        .X = Scr.ScaleX(Scr.ScaleWidth, scNow, 3)
        .Y = Scr.ScaleX(Scr.ScaleHeigh, scNow, 3)
        GetMem1 VarPtr(Col2), bt
        PutMem1 VarPtr(.Red), bt
        GetMem1 VarPtr(Col2) + 1, bt
        PutMem1 VarPtr(.Green), bt
        GetMem1 VarPtr(Col2) + 2, bt
        PutMem1 VarPtr(.Blue), bt
    End With
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    With gTriangle(0)
        .Vertex1 = 0
        .Vertex2 = 1
    End With
     Dim nRect As RECT, nRgn As Long, oldbRgn As Long
     x1 = CLng(Scr.ScaleX(x1, scNow, 3))
     y1 = CLng(Scr.ScaleY(y1, scNow, 3))
     x2 = CLng(Scr.ScaleX(x2, scNow, 3))
     y2 = CLng(Scr.ScaleY(y2, scNow, 3))
    nRgn = CreateRectRgn(x1, y1, x2, y2)
    Debug.Print GetClipRgn(Scr.hDC, oldbRgn)
    SelectClipRgn Scr.hDC, nRgn
    GradientFillRect Scr.hDC, vert(0), 2, gTriangle(0), 1, typegrad
    SelectClipRgn Scr.hDC, oldbRgn
    DeleteObject nRgn
    End If
    
End Sub



