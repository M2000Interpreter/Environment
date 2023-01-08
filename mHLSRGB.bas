Attribute VB_Name = "mHLSRGB"
Option Explicit

Public Sub RGBToHLS( _
      ByVal R As Long, ByVal G As Long, ByVal b As Long, _
      H As Single, s As Single, l As Single _
   )
Dim Max As Single
Dim Min As Single
Dim delta As Single
Dim rr As Single, rG As Single, rB As Single
If R < 0 Then R = 0 Else If R > 255 Then R = 255
If G < 0 Then G = 0 Else If G > 255 Then G = 255
If b < 0 Then b = 0 Else If b > 255 Then b = 255

   rr = R / 255: rG = G / 255: rB = b / 255

        Max = Maximum(rr, rG, rB)
        Min = Minimum(rr, rG, rB)
     l = (Max + Min) / 2    '{This is the lightness}

        If Max = Min Then
            'begin {Acrhomatic case}
            s = 0
            H = 0
           'end {Acrhomatic case}
        Else
           'begin {Chromatic case}
                '{First calculate the saturation.}
           If l <= 0.5 Then
               s = (Max - Min) / (Max + Min)
           Else
               s = (Max - Min) / (2 - Max - Min)
            End If
            '{Next calculate the hue.}
            delta = Max - Min

           If rr = Max Then
                H = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
           ElseIf rG = Max Then
                H = 2 + (rB - rr) / delta '{Resulting color is between cyan and yellow}
           ElseIf rB = Max Then
                H = 4 + (rr - rG) / delta '{Resulting color is between magenta and cyan}
            End If
      End If
'end {RGB_to_HLS}
End Sub

Public Function rgbconv(H As Long) As String
Dim mr As Long, mg As Long, mb As Long
Dim mg1 As Long, mb1 As Long
Dim hh As Single, ss As Single, LL As Single
rgbconv = "000000"
For mr = 0 To 255
For mg = 0 To 255
For mb = 0 To 255
RGBToHLS mr, mg, mb, hh, ss, LL
If H = CLng(Int(hh * 60) Mod 360) Then
rgbconv = Hex$(mb + mg * 256 + mr * 256 * 256)
mb = 255
mg = 255
mr = 255
End If
Next mb
Next mg
Next mr


End Function
Public Function hueconvSpecial(hr As Variant) As Long
Dim mr As Long, mg As Long, mb As Long
Dim ba$
Dim hh As Single, ss As Single, LL As Single
If VarType(hr) <> vbString Then ba$ = Hex$(hr) Else ba$ = hr
ba$ = Right$("00000" & ba$, 6)
mr = val("&h" & Mid$(ba$, 1, 2))
mg = val("&h" & Mid$(ba$, 3, 2))
mb = val("&h" & Mid$(ba$, 5, 2))
RGBToHLS mr, mg, mb, hh, ss, LL
'Debug.Print hh, sS, ll
hueconvSpecial = CLng(Int(hh * 60) Mod 360)

End Function
Public Function hueconv(hr As Variant) As Long
Dim mr As Long, mg As Long, mb As Long
Dim ba$
Dim hh As Single, ss As Single, LL As Single
If VarType(hr) <> vbString Then ba$ = Hex$(hr) Else ba$ = hr
ba$ = Right$("00000" & ba$, 6)
mb = val("&h" & Mid$(ba$, 1, 2))
mg = val("&h" & Mid$(ba$, 3, 2))
mr = val("&h" & Mid$(ba$, 5, 2))
RGBToHLS mr, mg, mb, hh, ss, LL
'Debug.Print hh, sS, ll

hueconv = Int((360 + hh * 60) Mod 360)

End Function
Public Function lightconv(hr As Variant) As Long
Dim mr As Long, mg As Long, mb As Long
Dim ba$
Dim hh As Single, ss As Single, LL As Single
If VarType(hr) <> vbString Then ba$ = Hex$(hr) Else ba$ = hr
ba$ = Right$("00000" & ba$, 6)
mb = val("&h" & Mid$(ba$, 1, 2))
mg = val("&h" & Mid$(ba$, 3, 2))
mr = val("&h" & Mid$(ba$, 5, 2))
RGBToHLS mr, mg, mb, hh, ss, LL
'Debug.Print hh, sS, ll
lightconv = CLng(LL * 255)

End Function
Public Function satconv(hr As Variant) As Long
Dim mr As Long, mg As Long, mb As Long
Dim ba$
Dim hh As Single, ss As Single, LL As Single
If VarType(hr) <> vbString Then ba$ = Hex$(hr) Else ba$ = hr
ba$ = Right$("00000" & ba$, 6)
mb = val("&h" & Mid$(ba$, 1, 2))
mg = val("&h" & Mid$(ba$, 3, 2))
mr = val("&h" & Mid$(ba$, 5, 2))
RGBToHLS mr, mg, mb, hh, ss, LL
'Debug.Print hh, sS, ll
satconv = CLng(LL * 255)

End Function
Public Function HSL(ByVal H, ByVal s, ByVal l) As Double
Dim R As Long, G As Long, b As Long
HLSToRGB CSng((H * 100&) Mod 36000) / 6000!, s / 100, l / 100, R, G, b
HSL = R + (G + b * 256#) * 256#
End Function

Public Sub HLSToRGB( _
      ByVal H As Single, ByVal s As Single, ByVal l As Single, _
      R As Long, G As Long, b As Long _
   )
Dim rr As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single
Dim Minl As Long, Maxl As Long, MidL As Long, dif As Single

   If s = 0 Then
      ' Achromatic case:
      rr = l: rG = l: rB = l
   Else
      ' Chromatic case:
      ' delta = Max-Min
      If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = l * (1 - s)
      Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = l - s * (1 - l)
      End If
      ' Get the Max value:
      Max = 2 * l - Min
      
      ' Now depending on sector we can evaluate the h,l,s:
      If (H < 1) Then
         rr = Max
         If (H < 0) Then
            rG = Min
            rB = rG - H * (Max - Min)
         Else
            rB = Min
            rG = H * (Max - Min) + rB
         End If
      ElseIf (H < 3) Then
         rG = Max
         If (H < 2) Then
            rB = Min
            rr = rB - (H - 2) * (Max - Min)
         Else
            rr = Min
            rB = (H - 2) * (Max - Min) + rr
         End If
      Else
         rB = Max
         If (H < 4) Then
            rr = Min
            rG = rr - (H - 4) * (Max - Min)
         Else
            rG = Min
            rr = (H - 4) * (Max - Min) + rG
         End If
         
      End If
            
   End If
   R = rr * 255: G = rG * 255: b = rB * 255
   If R < 0 Or G < 0 Or b < 0 Or R > 255 Or G > 255 Or b > 255 Then
   Maxl = Maximuml(R, G, b)
   Minl = Minimuml(R, G, b)
   MidL = R + G + b - Maxl - Minl
   If Maxl > Minl Then
   If Minl < 0 Then Maxl = Maxl - Minl: MidL = MidL - Minl: Minl = 0
   If Maxl > 255 Then
   dif = (255 - Minl) / (Maxl - Minl)
   Maxl = (Maxl - Minl) * dif + Minl
   MidL = (MidL - Minl) * dif + Minl
   End If
   If Maximuml(R, G, b) = R Then
   R = Maxl
            If Minimuml(R, G, b) = G Then
            G = Minl: b = MidL
            Else
            G = Minl: b = MidL
            End If
   ElseIf Maximuml(R, G, b) = G Then
            G = Maxl
            If Minimuml(R, G, b) = R Then
            R = Minl: b = MidL
            Else
            b = Minl: R = MidL
            End If
   Else
   b = Maxl
   If Minimuml(R, G, b) = R Then
            R = Minl: G = MidL
            Else
            G = Minl: R = MidL
            End If

   End If
   Else
   R = 0: b = 0: G = 0
   End If
   End If
   
   
End Sub
Private Function Maximum(rr As Single, rG As Single, rB As Single) As Single
   If (rr > rG) Then
      If (rr > rB) Then
         Maximum = rr
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function
Private Function Minimum(rr As Single, rG As Single, rB As Single) As Single
   If (rr < rG) Then
      If (rr < rB) Then
         Minimum = rr
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function

Private Function Maximuml(rr As Long, rG As Long, rB As Long) As Long
   If (rr > rG) Then
      If (rr > rB) Then
         Maximuml = rr
      Else
         Maximuml = rB
      End If
   Else
      If (rB > rG) Then
         Maximuml = rB
      Else
         Maximuml = rG
      End If
   End If
End Function
Private Function Minimuml(rr As Long, rG As Long, rB As Long) As Long
   If (rr < rG) Then
      If (rr < rB) Then
         Minimuml = rr
      Else
         Minimuml = rB
      End If
   Else
      If (rB < rG) Then
         Minimuml = rB
      Else
         Minimuml = rG
      End If
   End If
End Function

Public Function ChooseByHue(orig As Long, dark As Long, light As Long) As Long
'' break to R,G,B
Dim orR As Long, orG As Long, orB As Long, toplight As Long
' first with read
orB = orig Mod 256
orG = orig \ 256 Mod 256
orR = orig \ 65536
toplight = Maximuml(orR, orG, orB)
Dim orR1 As Long, orG1 As Long, orB1 As Long, toplight1 As Long
' first with read
orB1 = dark Mod 256
orG1 = dark \ 256 Mod 256
orR1 = dark \ 65536
toplight1 = Maximuml(orR1, orG1, orB1)
Dim orR2 As Long, orG2 As Long, orB2 As Long, toplight2 As Long
' first with read
orB2 = light Mod 256
orG2 = light \ 256 Mod 256
orR2 = light \ 65536
toplight2 = Maximuml(orR2, orG2, orB2)
If Abs(toplight - toplight1) > Abs(toplight - toplight2) Then
ChooseByHue = dark
Else
ChooseByHue = light
End If




End Function

