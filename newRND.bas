Attribute VB_Name = "Module8"
'Author: Merri of vbforums.com
'http://www.vbforums.com/showthread.php?t=499661
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Type rndvars
     lngX   As Long
     lngY   As Long
     lngZ   As Long
     blnInit As Boolean
End Type
Public Sub RandomizeIt(m As rndvars, NeoNumber As Long)
    Dim d As Double
    If NeoNumber = 0 Then m.blnInit = False
    d = RndM(m, NeoNumber)
End Sub
Public Function RndM(m As rndvars, Optional ByVal Number As Long) As Double
' Static lngX As Long, lngY As Long, lngZ As Long, blnInit As Boolean
With m
    Dim dblRnd As Double
    ' if initialized and no input number given
    If .blnInit And Number = 0 Then
        ' lngX, lngY and lngZ will never be 0
        .lngX = (171 * .lngX) Mod 30269
        .lngY = (172 * .lngY) Mod 30307
        .lngZ = (170 * .lngZ) Mod 30323
    Else
        ' if no initialization, use Timer, otherwise ensure positive Number
        If Number = 0 Then Number = timeGetTime And &H7FFFFFFF Else Number = Number And &H7FFFFFFF
        .lngX = (Number Mod 30269)
        .lngY = (Number Mod 30307)
        .lngZ = (Number Mod 30323)
        ' lngX, lngY and lngZ must be bigger than 0
        If .lngX > 0 Then Else .lngX = 171
        If .lngY > 0 Then Else .lngY = 172
        If .lngZ > 0 Then Else .lngZ = 170
        ' mark initialization state
        .blnInit = True
    End If
    ' generate a random number
    dblRnd = CDbl(.lngX) / 30269# + CDbl(.lngY) / 30307# + CDbl(.lngZ) / 30323#
    ' return a value between 0 and 1
    RndM = dblRnd - Int(dblRnd)
    End With
End Function


