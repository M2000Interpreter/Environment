VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lexar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' FOR COMMA MAY I HAVE TO CHANGE COLOR???
Option Explicit
Public Event addone(n As Long, t As Long)
Public Event addval(v As Variant)
Public Event addstr(v As String, UseCase As Boolean)
Public Event NewLine(ByVal i As Long)
Public Event abort(ByRef ok As Boolean)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Dim WithEvents mDoc As Document
Attribute mDoc.VB_VarHelpID = -1
Private Col() As Long, lastdata As String, thispara As Long
Dim var As New Document
Private Sub Add(s As String)
If var.IsEmpty Then
var.textDoc = s
Else
var.InsertDoc var.LastParagraph, var.TextParagraphLen(var.LastParagraph) + 1, s
End If

End Sub

Public Sub Construct2(z As String)
Set mDoc = New Document
Dim num As Variant, dummy As Long, abortnow As Boolean
mDoc.textDoc = z
mDoc.ColorEvent = True
mDoc.WrapAgainColor2
Dim i As Long, j As Long, j0 As Long, j1 As Long, s$, LastCode As Long, counter As Long
Dim noucase As Boolean, symb$, nclev As Integer, emptystr As Boolean, zz As Long, br As Long
For i = 1 To mDoc.DocParagraphs
RaiseEvent NewLine(i)
s$ = mDoc.TextParagraph(i)
j0 = 0
j1 = mDoc.Paracolordata(i)(1)
For j = 3 To UBound(mDoc.Paracolordata(i)) Step 2
    emptystr = False
    j0 = j1
    j1 = mDoc.Paracolordata(i)(j)
    If j1 < j0 Then
    If Len(s$) < j0 Then GoTo conthere1
    j1 = j0
    GoTo conthere
    ElseIf j1 = j0 Then
    j1 = j0 + 1
    End If
    If j0 = 0 Then j0 = 1
    If NLtrim$(Mid$(s$, j0, j1 - j0)) = "" Then
    If LastCode = 16 Then
       ' If j = 3 Then j1 = 2
        RaiseEvent addstr(vbCrLf + Mid$(s$, j0, j1 - j0), False)
        
    End If
    Else
    If mDoc.Paracolordata(i)(j - 3) = 7 Then
    ElseIf mDoc.Paracolordata(i)(j - 3) = 15 Then
    If LastCode = 15 Or LastCode = 16 Then
        If j0 = 1 Then
            RaiseEvent addstr(vbCrLf + Mid$(s$, 1, j1 - 1), False)
        Else
            RaiseEvent addstr(Mid$(s$, j0, j1 - j0), False)
        End If
    Else
        RaiseEvent addone(counter + j0, 15)
        LastCode = mDoc.Paracolordata(i)(j - 3)
        RaiseEvent addstr(Mid$(s$, j0 + 16 - LastCode, j1 - j0 - 1), False) ' +1 no { or "
        LastCode = 16
    End If
    Else
    
    LastCode = mDoc.Paracolordata(i)(j - 3)
    ' Str$(i) + "," + Str$(j0)
    RaiseEvent abort(abortnow)
    If abortnow Then Exit Sub
    Select Case LastCode
    Case 3
    If j0 = 1 Then
    RaiseEvent addstr(vbCrLf, False)
    Else
    'RaiseEvent addstr(Mid$(s$, j0, j1 - j0), False)
    End If
    LastCode = 0
    Case 16
    
        RaiseEvent addone(counter + j0, LastCode)
        RaiseEvent addstr(Mid$(s$, j0 + 1, j1 - j0 - 1), False)
    Case 100, 13
    symb$ = Mid$(s$, j0, j1 - j0)
    If IsNumberOnly(symb$, 1, num, dummy) Then
        RaiseEvent addone(counter + j0, LastCode)
        RaiseEvent addval(num)
    End If
    Case 19, 14, 12, 9, 8, 5

        symb$ = MyTrimRW(Mid$(s$, j0, j1 - j0))
        zz = MyTrimL(symb$)
        If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
        If symb$ = "{" Then
            symb$ = ""
            LastCode = 140
            RaiseEvent addone(counter + j0 + zz, LastCode)
        ElseIf symb$ = "}" Then
            symb$ = ""
            LastCode = 141
            RaiseEvent addone(counter + j0 + zz, LastCode)
        ElseIf symb$ = ")" Then
        If noucase Then nclev = 0
            RaiseEvent addone(counter + j0 + zz, LastCode)
            RaiseEvent addstr(symb$, False)

        ElseIf symb$ = ":" Then
            RaiseEvent addone(counter + j0 + zz, LastCode)
            RaiseEvent addstr(symb$, False)
            noucase = False
        Else
            RaiseEvent addone(counter + j0 + zz, LastCode)
            RaiseEvent addstr(symb$, False)
        End If
    Case 20
        RaiseEvent addone(counter + j0, LastCode)
        RaiseEvent addstr(Mid$(s$, j0, j1 - j0), False)
    Case Else
        
        If noucase And nclev > 0 Then
            symb$ = MyTrimRW(Mid$(s$, j0, j1 - j0))
            zz = MyTrimL(symb$)
            If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
            If Right$(symb$, 1) = "(" Then nclev = nclev - 1
            RaiseEvent addone(counter + j0 + zz, LastCode)
            If LastCode = 165 Then
            RaiseEvent addstr(UCase(symb$), False)
            Else
            RaiseEvent addstr(symb$, False)
            End If
        Else
            symb$ = MyTrimRW(Mid$(s$, j0, j1 - j0))
            zz = MyTrimL(symb$)
            If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
            RaiseEvent addone(counter + j0 + zz, LastCode)
            RaiseEvent addstr(symb$, True)
            
        End If
        If (LastCode = 111 Or LastCode = 112) And Not noucase Then noucase = True: nclev = 1
    
    End Select
    LastCode = 0
    End If
    End If
conthere:

Next j

If j0 = 0 Then
    If LastCode = 16 Then
        RaiseEvent addstr(vbCrLf, False)
    End If
'  Debug.Print -i, mDoc.Paracolordata(i)(1), Len(MyTrim$(s$))
Else
j0 = j1
If Len(s$) < j1 Then
Else
If mDoc.Paracolordata(i)(j - 3) = 7 Then
ElseIf mDoc.Paracolordata(i)(j - 3) = 15 Then
If LastCode = 16 Or LastCode = 15 Then
    If j0 = 1 Then
    RaiseEvent addstr(vbCrLf + Mid$(s$, 1, j1 - 1), False)
    Else
    RaiseEvent addstr(Mid$(s$, j0, j1 - j0), False)
    End If
Else
If Mid$(s$, j0, 1) = Chr$(34) Then

RaiseEvent addone(counter + j0, 15)
RaiseEvent addstr(Mid$(s$, j0 + 1, Len(s$) - j0 - 1), False)
LastCode = 0
Else
RaiseEvent addone(counter + j0, 16)
RaiseEvent addstr(Mid$(s$, j0 + 1), False)
LastCode = 16
End If
End If
Else
LastCode = mDoc.Paracolordata(i)(j - 3)
Select Case LastCode
Case 3
    LastCode = 0
Case 100, 13
    symb$ = Mid$(s$, j0)
    If IsNumberOnly(symb$, 1, num, dummy) Then
        RaiseEvent addone(counter + j0, LastCode)
        RaiseEvent addval(num)
        num = vbEmpty
    End If
Case 14
    symb$ = MyTrimRW(Mid$(s$, j0))
    zz = MyTrimL(symb$)
    If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
    If symb$ = "{" Then
        LastCode = 140
        RaiseEvent addone(counter + j0 + zz, 140)
    ElseIf symb$ = "}" Then
        
        
        LastCode = 141
        RaiseEvent addone(counter + j0 + zz, 141)
    Else
    RaiseEvent addone(counter + j0 + zz, LastCode)
    RaiseEvent addstr(symb$, False)
    End If
    
Case 20
    RaiseEvent addone(counter + j0, LastCode)
    RaiseEvent addstr(Mid$(s$, j0), False)
        
Case Else
    symb$ = MyTrimRW(Mid$(s$, j0))
    zz = MyTrimL(symb$)
    If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
    
    If noucase And nclev > 0 Then
        RaiseEvent addone(counter + j0 + zz, LastCode)
        RaiseEvent addstr(symb$, False)
    Else
        RaiseEvent addone(counter + j0 + zz, LastCode)
        RaiseEvent addstr(symb$, True)
    End If
End Select
LastCode = 0
End If
End If
End If

conthere1:
If LastCode = 0 Then
If Not emptystr Then
RaiseEvent addone(counter + Len(s$), 0)
emptystr = True
End If
Else
emptystr = False
End If
counter = counter + Len(s$) + 2
noucase = False
Next i
End Sub
Public Sub Construct(z As String)
Set mDoc = New Document
Dim ms As New mStiva2
mDoc.textDoc = z
mDoc.ColorEvent = True
mDoc.WrapAgainColor2
Dim i As Long, j As Long, j0 As Long, j1 As Long, s$, LastCode As Long, counter As Long
Dim noucase As Boolean, symb$, nclev As Integer, emptystr As Boolean, zz As Long, br As Long
For i = 1 To mDoc.DocParagraphs
s$ = mDoc.TextParagraph(i)
j0 = 0
j1 = mDoc.Paracolordata(i)(1)
For j = 3 To UBound(mDoc.Paracolordata(i)) Step 2
    emptystr = False
    j0 = j1
    j1 = mDoc.Paracolordata(i)(j)
    If j1 <= j0 Then
    If Len(s$) < j0 Then GoTo conthere1
    j1 = j0
    GoTo conthere
    End If
    If NLtrim$(Mid$(s$, j0, j1 - j0)) = "" Then
    Else
    If mDoc.Paracolordata(i)(j - 3) = 7 Then
    ElseIf mDoc.Paracolordata(i)(j - 3) = 15 Then
    If LastCode = 15 Or LastCode = 16 Then
    If Left$(RTrim$(Mid$(s$, j0, j1 - j0)), 1) = "}" Then
    Add "\n" + Mid$(s$, 1, j1 - 1) + vbCrLf
    LastCode = 0
    Else
    Add "\n" + Mid$(s$, j0, j1 - j0)
    LastCode = 15
    End If
    Else
    Add str$(counter + j0) + ", 15, " + Mid$(s$, j0, j1 - j0) + vbCrLf
    End If
    Else
    
    LastCode = mDoc.Paracolordata(i)(j - 3)
    ' Str$(i) + "," + Str$(j0)
    Select Case LastCode
    Case 19, 14, 13, 12, 100
        symb$ = MyTrimRW(Mid$(s$, j0, j1 - j0))
        zz = MyTrimL(symb$)
        If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
        If symb$ = "{" Then
            symb$ = "START_BRACKET" + str$(var.DocParagraphs)
            ms.PushVal CVar(var.DocParagraphs)
            LastCode = 140
        ElseIf symb$ = "}" Then
            br = CLng(ms.PopVal)
            var.ReWritePara br, var.TextParagraph(br) + ", " + str$(var.DocParagraphs)
            symb$ = "END_BRACKET" + str$(br)
            LastCode = 141
        ElseIf symb$ = ")" Then
        If noucase Then nclev = nclev + 1
        ElseIf symb$ = ":" Then
        noucase = False
        End If
        Add str$(counter + j0 + zz) + "," + str$(LastCode) + ", " + symb$ + vbCrLf
    Case 20
        Add str$(counter + j0 + zz) + "," + str$(LastCode) + ", " + Mid$(s$, j0, j1 - j0) + vbCrLf
    Case Else
        If noucase And nclev > 0 Then
        symb$ = MyTrimRW(Mid$(s$, j0, j1 - j0))
        zz = MyTrimL(symb$)
        If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
        If Right$(symb$, 1) = "(" Then nclev = nclev - 1
        Add str$(counter + j0 + zz) + "," + str$(LastCode) + ", " + symb$ + vbCrLf
        Else
        symb$ = MyTrimRW(Mid$(s$, j0, j1 - j0))
        zz = MyTrimL(symb$)
        If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0

        Add str$(counter + j0 + zz) + "," + str$(LastCode) + ", " + myUcase(symb$) + vbCrLf
        End If
        If (LastCode = 111 Or LastCode = 112) And Not noucase Then noucase = True: nclev = 1
    
    End Select
    LastCode = 0
    End If
    End If
conthere:

Next j

If j0 = 0 Then
  '  Debug.Print -i, mDoc.Paracolordata(i)(1), Len(MyTrim$(s$))
Else
j0 = j1
If Len(s$) < j1 Then
Else
If mDoc.Paracolordata(i)(j - 3) = 7 Then
ElseIf mDoc.Paracolordata(i)(j - 3) = 15 Then
If Mid$(s$, j0, 1) = Chr$(34) Then
Add str$(counter + j0) + ", 15, " + Mid$(s$, j0) + vbCrLf
LastCode = 0
Else
Add str$(counter + j0) + ", 16, " + Mid$(s$, j0)
LastCode = 16
End If
Else
LastCode = mDoc.Paracolordata(i)(j - 3)
Select Case LastCode
Case 14
    symb$ = MyTrimRW(Mid$(s$, j0))
    zz = MyTrimL(symb$)
    If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
    If symb$ = "{" Then
        symb$ = "START_BRACKET" + str$(var.DocParagraphs)
        ms.PushVal CVar(var.DocParagraphs)
        LastCode = 140
    ElseIf symb$ = "}" Then
        
        br = CLng(ms.PopVal)
        var.ReWritePara br, var.TextParagraph(br) + ", " + str$(var.DocParagraphs)
        symb$ = "END_BRACKET" + str$(br)
        LastCode = 141
    End If
    Add str$(counter + j0 + zz) + "," + str$(LastCode) + ", " + symb$ + vbCrLf
Case 20
    Add str$(counter + j0) + "," + str$(LastCode) + ", " + Mid$(s$, j0) + vbCrLf
Case Else
    symb$ = MyTrimRW(Mid$(s$, j0))
    zz = MyTrimL(symb$)
    If zz > 1 Then symb$ = Mid$(symb$, zz): zz = zz - 1 Else zz = 0
    
    If noucase And nclev > 0 Then
        Add str$(counter + j0 + zz) + "," + str$(LastCode) + ", " + symb$ + vbCrLf
    Else
        Add str$(counter + j0 + zz) + "," + str$(LastCode) + ", " + myUcase(symb$) + vbCrLf
    End If
End Select
LastCode = 0
noucase = False
End If
End If
End If
conthere1:

If LastCode = 0 Then
If Not emptystr Then Add str$(counter + j0) + ", 0" + vbCrLf: emptystr = True
Else
emptystr = False
End If
counter = counter + Len(s$) + 2
Next i
s$ = var.textDoc
SetTextData CF_UNICODETEXT, s$    'set as unicode text
End Sub

Private Function aheadstatus(A$, Optional srink As Boolean = True, Optional Pos As Long = 1) As String 'ok
Dim b$, part$, W$, pos2 As Long, word$, item As Integer
Dim OLDPOS As Long

If A$ = vbNullString Then Exit Function
If Pos = 0 Then Pos = 1
OLDPOS = Pos
Do While Pos <= Len(A$)

    W$ = Mid$(A$, Pos, 1)
    If part$ = vbNullString And W$ = "0" Then
        If Pos + 2 <= Len(A$) Then
            If LCase(Mid$(A$, Pos, 2)) Like "0[xВ]" Then
            'hexadecimal literal number....
                Pos = Pos + 2
                Do While Pos <= Len(A$)
                If Not Mid$(A$, Pos, 1) Like "[0-9a-fA-F]" Then Exit Do
                Pos = Pos + 1
                Loop
                b$ = b$ & "N"
                item = item + 1
                If Pos <= Len(A$) Then
                    W$ = Mid$(A$, Pos, 1)
                Else
                    Exit Do
                End If
            End If
        End If
    End If

    If W$ = Chr(34) Then
    
        If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        item = item + 1
        Pos = Pos + 1
        Do While Pos <= Len(A$)
        'If Mid$(a$, Pos, 1) = Chr(34) And Mid$(a$, Pos - 1, 1) <> "`" Then Exit Do
        If Mid$(A$, Pos, 1) = Chr(34) Then Exit Do
    If Asc(Mid$(A$, Pos, 1)) < 32 Then Exit Do
   
        Pos = Pos + 1
        Loop
    ElseIf W$ = "(" Then
        If part$ <> "" Then
            ' after
            ElseIf Right$(b$, 1) = "a" Then
            b$ = Left$(b$, Len(b$) - 1)
            part$ = vbNullString
            Else
            part$ = "N"
              
        End If
      Pos = Pos + 1
        If Not BlockParam2p(A$, Pos) Then Exit Do
        item = item + 1
       If Mid$(A$, Pos + 1, 1) <> "." Then b$ = b$ & part$
        part$ = vbNullString
        
    ElseIf W$ = "{" Then
If word$ <> "" Then
  Select Case myUcase(word$, True)
  Case "THEN", "ELSE", "тоте"
  Pos = Pos - 4
  Exit Do
  Case "ELSE.IF"
  Pos = Pos - 7
  Exit Do
  Case "аккиыс"
  Pos = Pos - 6
  Exit Do
  Case "аккиыс.ам"
  Pos = Pos - 9
  Exit Do
    
  End Select
  End If
    

    If part$ <> "" Then b$ = b$ & part$
        If LenB(b$) = 0 Then
        ElseIf Right(b$, 1) <> "l" Then
        Exit Do
        End If
 
        blockStringAhead A$, Pos
        Pos = Pos + 1
        part$ = "S"
        item = item + 1

            If Pos <= Len(A$) Then
        If Not blockStringAhead(A$, Pos) Then Exit Do
        Pos = Pos - 1
        End If
''Exit Do

    Else
        Select Case W$
        Case ","  ' bye bye
        Exit Do
        Case "%"
            If part$ = vbNullString Then
            ' lathos
            End If
        Case "$"
            If part$ = vbNullString Then
                If b$ = vbNullString Then
                    part$ = "N"
                ElseIf Right$(b$, 1) = "o" Then
                    part$ = "N"
                Else
        
                    aheadstatus = b$
                    Exit Function
                End If
            ElseIf part$ = "N" Then
                    b$ = b$ & "Sa"
                    item = item + 1
                    part$ = vbNullString
            End If

        Case "-", "+", "|"
                    b$ = b$ & part$
                    If b$ = vbNullString Then
                        part$ = "o"
                    Else
                    
                part$ = "o"
                End If
         Case "/"
            If Mid$(A$, Pos + 1, 1) = "/" Then Exit Do
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case "*", "^"
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case " ", ChrW(160)
          If word$ <> "" Then
          
  Select Case myUcase(word$, True)
  
  Case "THEN", "ELSE", "тоте", "GOTO", "пяос"
  Pos = Pos - 4
  Exit Do
  Case "диалесоу"
  Pos = Pos - 8
  Exit Do
  Case "GOSUB"
  Pos = Pos - 6
  Exit Do
  Case "ELSE.IF"
  Pos = Pos - 7
  Exit Do
  Case "аккиыс"
  Pos = Pos - 6
  Exit Do
  Case "аккиыс.ам"
  Pos = Pos - 9
  Exit Do
  End Select
  End If
JUMPHERE:
        word$ = vbNullString
            If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            Else
            'skip
            End If
        Case "0" To "9", "."
            If part$ = "N" Then
            If Len(A$) < Pos Then
                If Mid$(A$, Pos + 1, 1) Like "[@#%~]" Then
                    Pos = Pos + 1
                ElseIf Mid$(A$, Pos + 1, 1) = "&" Then
                    If Mid$(A$, Pos + 2, 1) = "&" Then Pos = Pos + 2 Else Pos = Pos + 1
                End If
            End If
            ElseIf part$ = "S" Then
            
            Else
            b$ = b$ & part$
            item = item + 1
            part$ = "N"
            End If
        Case "&"
        If part$ = vbNullString Then
        part$ = "S"
        Else
        b$ = part$
        part$ = "S"
        item = item + 1
        End If
        Case "e", "E", "Е", "е"
        If LenB(word$) = 0 Then part$ = "N"
        
        word$ = word$ + W$
        
           If part$ = "S" Then
            
            
            ElseIf part$ = "N" Then
            item = item + 1
            b$ = b$ & part$
            'part$ = "N"
            End If
            
            
        Case ">", "<", "~"
            If Len(A$) >= Pos + 1 Then
            If Mid$(A$, Pos, 2) = Mid$(A$, Pos, 1) Then
                b$ = b$ & part$
                If b$ = vbNullString Then
                        Else
                        
                    part$ = "o"
                    Pos = Pos + 1
                    End If
                ElseIf W$ = ">" And Pos > 1 Then
                    If Mid$(A$, Pos - 1, 2) = "->" Then ' "->"
                   If Right$(b$, 1) = "S" Then
                    b$ = b$ + part$
                    part$ = "N"
                    Else
                      '  part$ = vbNullString
                    End If
                        
                    End If
                End If
            End If
            GoTo there1
        Case "="
            If Mid$(A$, Pos + 1, 1) = ">" Then
                Pos = Pos + 2
                GoTo conthere
                End If
there1:
                If b$ & part$ <> "" Then
                W$ = Replace(b$ & part$, "a", "")
            part$ = vbNullString
               If srink Then
                  Do
                b$ = W$
                W$ = Replace(b$, "NN", "N")
                Loop While W$ <> b$
                         Do
                        b$ = W$
                          W$ = Replace(b$, "SlS", "N")
                          Loop While W$ <> b$
                            Do
                          b$ = W$
                          W$ = Replace(b$, "NlN", "N")
                          Loop While W$ <> b$
    
                Do
                b$ = W$
                W$ = Replace(b$, "NoN", "N")
                Loop While W$ <> b$
                
                Do
                b$ = W$
                W$ = Replace(b$, "SoS", "S")
                Loop While W$ <> b$
                Else
              b$ = W$
               End If
               
                If Left$(b$, Len(b$) - 1) <> "l" Then
                If b$ = "S" Then part$ = "o" Else part$ = "l"
                End If
                Else
                Exit Do
                End If
        
        Case ")", "}", Is < " ", ":", ";", "'", "\"
        Exit Do
        Case Else
        
        If part$ = "N" Then
        word$ = word$ + W$
        ElseIf part$ = "S" Then
        word$ = word$ + W$
        item = item + 1
        Else
        
     b$ = b$ & part$
     part$ = "N"
     word$ = word$ + W$
     item = item + 1
            End If
        End Select
        End If
        
    Pos = Pos + 1
conthere:
Loop

    W$ = Replace(b$ & part$, "a", "")
    
    b$ = W$
If srink Then
         Do
  b$ = W$

    W$ = Replace(b$, "SlS", "N")
    Loop While W$ <> b$
      Do
    b$ = W$
    W$ = Replace(b$, "NlN", "N")
    Loop While W$ <> b$
    
    Do
    b$ = W$
    W$ = Replace(b$, "NoN", "N")
    Loop While W$ <> b$
    
    Do
    b$ = W$
    W$ = Replace(b$, "SoS", "S")
    Loop While W$ <> b$
End If
   If item = 0 Then
   Pos = OLDPOS
   Else
   
 
    aheadstatus = b$
End If



End Function


Private Sub makebase(basecolor As Long)
ReDim Col(1)
Col(0) = basecolor
Col(1) = 1

End Sub
Private Sub addone(ByVal thispos As Long, ByVal thiscolor As Long)
Dim K As Long

K = UBound(Col())
If VarType(Col(0)) = 0 Then
makebase 14
If thispos <> 1 Then
K = UBound(Col())
Else
Col(0) = thiscolor
Exit Sub
End If
End If
' LOOK THIS SECOND TIME
If K > 1 Then
'If col(k - 1) = thiscolor Then Exit Sub
If Col(K) = thispos Then
Col(K - 1) = thiscolor
Exit Sub
End If
End If


''''If col(k - 1) = thiscolor And k > 1 Then Exit Sub

ReDim Preserve Col(K + 2)
Col(K + 1) = thiscolor
Col(K + 2) = thispos


End Sub


Private Sub mDoc_ColorDataEvent(thispara1 As Long, Data As String, feedback As Long)

thispara = thispara1
Col() = mDoc.Paracolordata(thispara)
If feedback = 0 Then
ReDim Col(1)
End If
lastdata = Data
ColorizeInnerStatements (Data), Len(Data), (0), feedback
If feedback = -1 Then
If thispara <> 0 Then
mDoc.Paracolordata(thispara) = Col()
thispara1 = thispara
End If
Else
mDoc.Paracolordata(thispara) = Col()
thispara1 = thispara
End If
End Sub
Private Sub ColorizeInnerStatements(A$, K As Long, pos2 As Long, feedback As Long, Optional jump As Boolean = False)
Dim drop$, pos1 As Long, onetime As Boolean, kind As Long, lastkind As Long, feedthis As Long, counter As Long
Dim lastcol As Long, firstonly As Boolean
firstonly = True
A$ = Replace(A$, vbTab, " ")
If jump Then
kind = IsLabelOnly((A$), drop$)
If kind = 0 Then
addone pos2, 14
lastcol = 14
If FastSymbol(A$, "(") Then
 pos2 = K - Len(A$) + 1
End If
GoTo again2
End If
GoTo fromjump

End If
onetime = True

If K < 1 Then
makebase 14
mDoc.Paracolordata(thispara) = Col()
If mDoc.startfrom(thispara) = thispara Then
feedback = thispara
ReDim Col(1)
Do While mDoc.Advance(thispara)
If mDoc.startfrom(thispara) = feedback Then
mDoc.startfrom(thispara) = 0
mDoc.Paracolordata(thispara) = Col()
End If
Loop
feedback = feedback <> thispara
Exit Sub
End If

End If
If feedback > 0 Then
'lastcol = 15
A$ = Mid$(A$, feedback + 1)
'pos2 = feedback + 1
 addone feedback + 1, lastcol
feedback = 0
GoTo checkthistoo
Else
lastcol = 14
End If
again:
If RTrim$(A$) = vbNullString Then
If pos2 > 0 Then
'If Col((UBound(Col()) \ 2) * 2) <> lastcol Then
'If pos2 < k Then
'addone pos2 + 1, lastcol
'End If
'ElseIf UBound(Col()) < 2 Then
addone pos2, lastcol
'End If
  Else
  makebase 14
  End If
mDoc.Paracolordata(thispara) = Col()
Exit Sub
' FOR TESTS
  pos1 = 0
  For K = 1 To UBound(Col()) Step 2
    If VarType(Col(K)) = 0 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
    If Col(K) > pos1 Then pos1 = Col(K)
    If Col(K) < pos1 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
  Next K
    For K = 0 To UBound(Col()) Step 2
   If Col(K) > 15 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
  Next K
  
End If
If MaybeIsSymbol(A$, ":", "\") Then
A$ = Mid$(A$, 2)
If AscW(Left$(A$, 1) + " ") = 92 Then
If pos2 > 1 Then
    If Col(UBound(Col())) = 1 Or InStr(" #@.+-~^%$,!({'?|/&)}" + Chr$(34), Mid$(lastdata, pos2 - 1, 1)) > 0 Then

    If pos2 > 0 Then
        addone pos2, 14
    Else
        addone 1, 14
    End If
 pos2 = K - Len(A$) + 1
GoTo novar
End If
End If
A$ = Mid$(A$, 2)
End If

If Len(A$) > 1 Then
If AscW(A$) <> 32 Then
lastkind = 1
GoTo again3
End If
End If
Else
novar:
While FastSymbol(A$, ":")
If pos2 > 0 Then
addone pos2, 9   ' 14
Else
addone 1, 9  '14
End If
 pos2 = K - Len(A$) + 1
Wend
End If
If onetime And MaybeIsSymbol(A$, "(") Then
GoTo CheckThis
End If
drop$ = "@(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then
If IsSymbol(A$, "@") Then
        If pos2 > 0 Then
        addone pos2, 14
        Else
        addone 1, 14
        End If
     pos2 = K - Len(A$) + 1
End If
End If
again500:
kind = IsLabelDot("", A$, drop$, firstonly)
If kind = -500 Then
    addone K - Len(A$) - Len(drop$) + 1, 20
     pos2 = K - Len(A$) + 1
    kind = 0
ElseIf kind = 0 Then
    drop$ = "$(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
    If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
    If kind = 0 Then drop$ = "->": kind = Abs(FastSymbol(A$, drop$, , 2))
End If
fromjump:
        If kind = 0 Then drop$ = vbNullString
        If kind > 0 Then
            If ismine(drop$) Then
                If ismine6(drop$) Then
                    lastcol = 111
                    If pos2 > 0 Then
                       addone pos2, 111
                    Else
                       addone K - Len(A$) - Len(drop$) + 1, 111
                   End If
                ElseIf ismine60(drop$) Then
                    lastcol = 112
                    If pos2 > 0 Then
                        addone pos2, 112
                    Else
                        addone K - Len(A$) - Len(drop$) + 1, 112
                    End If
                Else
                    lastcol = 170 + kind
                    If pos2 > 0 Then
                        addone pos2, lastcol
                    Else
                        addone K - Len(A$) - Len(drop$) + 1, lastcol
                    End If
                End If
                If ismine2(drop$) Then
                    kind = -2
                    If Trim$(A$) <> "" Then
                        pos2 = K - Len(A$) + 1
                        If ismine22(drop$) Then
                        ElseIf MaybeIsSymbol(A$, "{") Then
                            addone pos2, 14
                            pos2 = K - Len(A$) + 1
                            onetime = True
                            GoTo again
                        ElseIf ismine5(drop$) Then
                            GoTo again500
                        ElseIf ismine3(drop$) Then
                            onetime = True
                            GoTo again
                        End If
                    End If
                ElseIf ismine1(drop$) Then
                    kind = -1
                    If Trim$(A$) <> "" Then
                        pos2 = K - Len(A$) + 1
                        If FastSymbol(A$, "{") Then
                            addone pos2, 14
                            pos2 = K - Len(A$) + 1
                            onetime = True
                            GoTo again
                        Else
                            mDoc.Paracolordata(thispara) = Col()
                            Exit Sub
                        End If
                    End If
                ElseIf ismine33(drop$) Then
                    ' fix for case >100, =50...
                    pos2 = K - Len(A$) + 1
                    Do
                        lastcol = 14
                        addone pos2, 14
                        If FastSymbol(A$, "<") Then
                            If Not FastSymbol(A$, "=") Then
                                FastSymbol A$, ">"
                            End If
                        ElseIf FastSymbol(A$, ">") Then
                            FastSymbol A$, "="
                        ElseIf FastSymbol(A$, "=") Then
                            FastSymbol A$, "="
                        ElseIf FastSymbol(A$, "~") Then
                        
                        End If
                        pos1 = 1
                        drop$ = aheadstatus(A$, False, pos1)
                        If pos1 > 0 Then
                            drop$ = Mid$(A$, 1, pos1 - 1)
                            ColorizeExpressions drop$, K - Len(A$) + Len(drop$), pos2, feedthis
                            A$ = Mid$(A$, pos1)
                        End If
                        If MaybeIsSymbol(A$, ",") Then
                            lastcol = 8
                            addone pos2, 8
                            FastSymbol A$, ","
                            pos2 = K - Len(A$) + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    onetime = True
                    GoTo here1
                End If
            Else
                lastcol = 10
                If Len(A$) > 0 Then
                    If AscW(A$) = 92 Then
                        lastkind = 1
                    End If
                End If
                If pos2 > 0 Then
                    addone pos2, 16 * lastcol + kind  ' 10
                Else
                    lastkind = 0
                    'addone k - Len(a$) - Len(drop$) + 1, 10
                    addone K - Len(A$) - Len(drop$) + 1, 16 * lastcol + kind
                 End If
            End If
            pos2 = K - Len(A$) + 1
            If kind > 3 Then
CheckThis:
                pos1 = 1
                drop$ = aheadstatus(A$, False, pos1)
marvel:
                If pos1 > 1 Then
                    feedthis = 0  ' NEW FROM COLORSTATEMENTS11
marvel2:
                    drop$ = Mid$(A$, 1, pos1 - 1)
                    ColorizeExpressions drop$, K - Len(A$) + Len(drop$), pos2, feedthis
                    If feedthis = 11 Or feedthis = 14 Or feedthis = 10 Then
                    'If kind < 5 Then lastcol = feedthis
                    feedthis = 0
                End If
newjump:
                feedback = feedthis
                If Col((UBound(Col()) \ 2) * 2) = 7 Then
                    A$ = vbNullString
                Else
                    A$ = Mid$(A$, pos1)
                End If
checkthistoo:
                If RTrim$(A$) = vbNullString Then
                    If pos2 > 0 Then
                        If feedthis <> 15 Then
                            If Col((UBound(Col()) \ 2) * 2) <> lastcol And pos2 < pos1 Then addone pos2, lastcol
                        End If
                    Else
                        makebase 14
                    End If
                    mDoc.Paracolordata(thispara) = Col()
                    If feedthis = 15 Then GoTo use15
                    Exit Sub
                End If
                If kind > 4 And onetime Then
                    If MaybeIsSymbol(A$, ",") Then
                        lastcol = 8
                        addone pos2, 8
                        FastSymbol A$, ","
                        pos2 = K - Len(A$) + 1
                            'If FastSymbol(a$, ",") Then
                            ' addone pos2, 14
                            '  pos2 = k - Len(a$) + 1
                        pos1 = 1
                        GoTo CheckThis
                    End If
                End If
                If FastSymbol(A$, "->", , 2) Then
                    addone pos2, 14
                    pos2 = K - Len(A$) + 1
                    If FastSymbol(A$, "{") Then
                        pos2 = K - Len(A$) + 1
                        onetime = True
                        GoTo again
                    End If
                ElseIf MaybeIsSymbol(A$, ",") Then
                    lastcol = 8
                    addone pos2, 8
                    FastSymbol A$, ","
                    pos2 = K - Len(A$) + 1
  'If FastSymbol(a$, ",") Then
  'addone pos2, 14
   ' pos2 = k - Len(a$) + 1
                    pos1 = 1
                End If
            ElseIf kind > 4 And onetime Then
                If MaybeIsSymbol(A$, ",") Then
                            'lastcol = 8
                    addone pos2, 8
                    FastSymbol A$, ","
                    pos2 = K - Len(A$) + 1
                            'If FastSymbol(a$, ",") Then
                              'addone pos2, 14
                              '  pos2 = k - Len(a$) + 1
                    pos1 = 1
                    GoTo CheckThis
                End If
            End If
            If FastSymbol(A$, ")") Then
                'pos2 = k - Len(a$)
                'If lastcol = 0 Then
                'addone pos2, 14
                'Else
                'addone pos2, lastcol
                'End If
                addone pos2, 19 ' lastcol
                pos2 = K - Len(A$) + 1
                ' NEW CODE START HERE
            ElseIf feedthis = 15 Then
                If FastSymbol(A$, "+") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "=") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "~") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, ">") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                GoTo CheckThis
                ElseIf FastSymbol(A$, "<") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                    GoTo CheckThis
                ElseIf FastSymbol(A$, "}") Then
                    If pos2 > 1 Then
                        addone pos2, 3
                    Else
                        Col(0) = 3
                    End If
                    pos2 = K - Len(A$) + 1
                    GoTo CheckThis
                ElseIf MaybeIsSymbol(A$, Chr(34)) Then
                    If Not ISSTRINGA(A$, drop$) Then
                        A$ = vbNullString
                        addone pos2, 10
                    Else
                        addone pos2, 15
                    End If
                        addone pos2, 14  ' for string
                        pos2 = K - Len(A$) + 1
                        ' Chr(34))
                    Else
                        feedthis = 0
                    End If
                    '  If feedthis = 15 Then GoTo CheckThis
                    feedthis = 0
                    ' END OF NEW CODE
                End If
            End If
here3:
            If kind > 0 Then
                If MaybeIsSymbol(A$, ".") Then
                    If FastSymbol(A$, ".(", , 2) Then
                        addone pos2, 14
                        A$ = vbNullString
                        pos2 = K + 1
                    ElseIf FastSymbol(A$, "..", , 2) Then
                        lastcol = 10
                        lastkind = 1
                        addone pos2, lastcol
                        pos2 = K - Len(A$) + 1
                        GoTo here3
                    Else
                        pos1 = pos1 + 1
                        lastcol = 14
                        GoTo CheckThis
                    End If
                End If
            End If
            If FastSymbol(A$, "<=", , 2) Then
                addone pos2, 24
                pos2 = K - Len(A$) + 1
            Else
again4:
                If MaybeIsSymbolNoSpace(A$, "~", True) Then
                    addone pos2, 18
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "=>", , 2) Then
                    addone pos2, 14
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "->", , 2) Then
                    addone pos2, 11
                    pos2 = K - Len(A$) + 1
                ElseIf MaybeIsSymbolLot(A$, "/*-+~^|<>", "/*-+=~^|<>") Then
                    addone pos2, 18
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "=") Then
                    addone pos2, 14
                    pos2 = K - Len(A$) + 1
              ''ElseIf FastSymbol(a$, "//", , 2) Then
                  ''  GoTo here1
              ''ElseIf MaybeIsSymbol(a$, "+-*/", "=", True) Then
                    ''addone pos2, 15
                    '' pos2 = k - Len(a$) + 1
                    ''ElseIf FastSymbol(a$, "/*-+=~^|<>") Then
                    '''addone pos2, 14
                    '' pos2 = k - Len(a$) + 1
                    ''If FastSymbol(a$, "/*-+=~^&|<>") Then
                    '''addone pos2, 14
                    ''pos2 = k - Len(a$) + 1
                    ''End If
                End If
            End If
again2:
            Do
again3:
                pos1 = 1
                drop$ = aheadstatus(A$, False, pos1)
    ' new code start here
                If Len(drop$) = 0 And MaybeIsSymbol(A$, "(") Then
                    FastSymbol A$, "("
                    addone pos2, 17  ' change from 15 no white =
                    pos2 = K - Len(A$) + 1
                    GoTo again3
                    ' new CodeBlock end here
                ElseIf pos1 > 1 Or drop$ = "S" Then
                    If pos1 = 1 Then
                        drop$ = A$
                        pos1 = Len(A$) + 1
                        counter = 1
                    Else
                        drop$ = Mid$(A$, 1, pos1 - 1)
                        If drop$ = vbNullString Then drop$ = " "
                    End If
                    feedback = 0
                    ColorizeExpressions drop$, K - Len(A$) + Len(drop$), pos2, feedback
                    If feedback <> 15 Then
                        If feedback = 0 And lastcol = 0 Then feedback = 14
                        If drop$ = "(" Then
                            If feedback = 100 Then
                                lastcol = 14
                            Else
                                lastcol = feedback
                            End If
                        Else 'If lastcol <> 11 And lastcol <> 10 Then
                            If feedback = 100 Then
                                lastcol = 14
                            Else
                                If feedback > 0 And kind > 0 Then
                                    lastcol = feedback
                                    feedback = 0
                                    A$ = Mid$(A$, pos1)
                                    ColorizeInnerStatements A$, K, pos2, feedback, True
                                    If MaybeIsSymbol(A$, ")") Then
                                        addone pos2, 19 'lastcol
                                        FastSymbol A$, ")"
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        feedback = 0
                    End If
                    If feedback = 100 Then
                        A$ = Mid$(A$, pos1)
                        GoTo again112
                    ElseIf feedback = 15 Then
use15:
                        counter = 1
nextline:
                        feedback = 0
                        drop$ = mDoc.NextTextParagraph(thispara, feedback)
                        If feedback = 0 Then
                            A$ = Mid$(A$, pos1)
                            mDoc.Paracolordata(thispara) = Col()
                            Exit Sub
                        End If
                        pos1 = 1
                        countbrakcet drop$, pos1, counter
                        mDoc.Paracolordata(thispara) = Col()
                        A$ = drop$
                        feedthis = 15  ' new this
                            'If feedthis = 15 Then
                        K = Len(A$)
                            'Else
                            'k = Len(a$) - 1
                            'End If
                        If mDoc.startfrom(thispara) = 0 Then
                            mDoc.startfrom(thispara) = thispara
                        End If
                        mDoc.startfrom(feedback) = mDoc.startfrom(thispara)
                        thispara = feedback
                        lastdata = A$
                        pos2 = 1
                        ' If Len(a$) > 0 Then  ' WAS LIKE THIS
                        If counter > 0 Then ' CHANGE TO THIS
                            ReDim Col(3)
                            Col(0) = 15
                            Col(1) = 1
                            Col(2) = 15
                            Col(3) = K + 1
                        Else
                            ReDim Col(1)
                            Col(0) = 15
                            Col(1) = 1
                        End If
                        If counter = 0 Then
                            If UBound(Col()) > 1 Then
                                Col(3) = pos1
                                If Col(3) = Col(1) Then
                                    pos1 = pos1 + 1
                                    Col(3) = pos1
                                    Col(2) = 14
                                End If
                                '    Else
                                        'Col(1) = pos1  ' new here
                            End If
                                        'If Left(a$, 1) <> "}" Then
                            If pos1 > 0 Then ' new here
                                A$ = Mid$(A$, pos1 + 1)
                                pos2 = K - Len(A$) + 1
                            End If
                            feedback = 0 'lastcol
    'feedthis = 0  ' ALTER
                                    'If MaybeIsSymbol(a$, "<>=~") Then
                                     '   onetime = False
                                      '          pos2 = k - Len(a$) + 1
                                       '         addone pos2, 14
                                    
                                        'If FastSymbol(a$, "<") Then
                                         '   If FastSymbol(a$, ">") Then
                                          '  ElseIf FastSymbol(a$, "=") Then
                                           ' End If
                                            'pos2 = k - Len(a$) + 1
                                            'addone pos2, 14
                                        'ElseIf FastSymbol(a$, ">") Then
                                            
                                         '   If FastSymbol(a$, "=") Then
                                           ' End If
                                          '
                                            
                                        'ElseIf FastSymbol(a$, "=") Then
                                        'ElseIf FastSymbol(a$, "~") Then
                                        'End If
                                    'End If
                        Else
                            GoTo nextline
                        End If
                            'kind = 0
                            ' feedback = 0
                        pos1 = 1
                        GoTo CheckThis
                    ElseIf feedback = 0 And pos1 > 1 And kind > 0 Then
                        GoTo newjump
                    End If
                    If pos1 > 0 Then
                        If Mid$(A$, pos1 - 1, 1) = " " Then lastkind = 0
                        A$ = Mid$(A$, pos1)
                    End If
                    If MaybeIsSymbol(A$, ":{") Then
                        onetime = True
                        GoTo again
                    End If
                    If InStr(drop$, " ") = 0 Or kind >= 0 Then
                        GoTo CheckThis
                    End If
                End If
again112:
                If MaybeIsSymbol(A$, ",;", , True) Then
                    addone pos2, 8 '14
                    pos2 = K - Len(A$) + 1
                    lastkind = 0
                    onetime = False
                ElseIf FastSymbol(A$, "}") Then
                    lastkind = 0
                    addone pos2, 14
                    pos2 = K - Len(A$) + 1
                    onetime = True ' new addition
                    If MaybeIsSymbol(A$, ".;+(<>=~") Then
                        onetime = False  ' not in Col..Sta..11
                        If FastSymbol(A$, "<") Then
                            If FastSymbol(A$, ">") Then
                            ElseIf FastSymbol(A$, "=") Then
                            End If
                        ElseIf FastSymbol(A$, ">") Then
                            If FastSymbol(A$, "=") Then
                            End If
                        ElseIf FastSymbol(A$, "=") Then
                        ElseIf FastSymbol(A$, "~") Then
                        End If
                        GoTo CheckThis
                    End If
                    GoTo again112
                ElseIf MaybeIsSymbol(A$, "'") Then
                    lastkind = 0
                    pos2 = K - Len(A$) + 1
                    FastSymbol A$, "'"
                    If pos2 > 0 Then
                        addone pos2, 7
                    Else
                        addone 1, 7
                    End If
                    A$ = vbNullString
                    pos2 = K
                    If pos2 > 0 Then
                        addone pos2 + 1, 14
                    End If
                    mDoc.Paracolordata(thispara) = Col()
                    Exit Sub
                Else
                    onetime = True
                    mDoc.Paracolordata(thispara) = Col()
                    Exit Do
                End If
            Loop
                            'If a$ = vbNullString Then
                            '
                            'mDoc.Paracolordata(thispara) = col()
                            'Exit Sub
                            
                            'End If
            GoTo again
        ElseIf FastSymbol(A$, "//", , 2) Then
            GoTo here1
        ElseIf MaybeIsSymbol(A$, "\", , True) Then
            If lastkind = 1 Then
                pos2 = K - Len(A$) + 1
                If AscW(A$ + ".") <> 32 Then
                    If MaybeIsSymbol(A$, ".", ".", True) Then
                    End If
                    If MaybeIsSymbolNoSpace(A$, "\") Then GoTo again
                    pos2 = K - Len(A$) + 1
                    lastkind = 0
                    If MaybeIsSymbol(A$, "\", , True) Then GoTo here1
                    GoTo again
                End If
            Else
here1:
                If pos2 > 0 Then
                    addone pos2, 7
                Else
                    addone 1, 7
                End If
                pos2 = K - Len(A$) + 1
                mDoc.Paracolordata(thispara) = Col()
                Exit Sub
            End If
        ElseIf MaybeIsSymbol(A$, "'") Then
            If pos2 > 1 Then
                addone pos2, 7
            Else
                addone 1, 7
            End If
            pos2 = K - Len(A$) + 1
            mDoc.Paracolordata(thispara) = Col()
            Exit Sub
        ElseIf Not onetime Then
            If FastSymbol(A$, ",") Then
                mDoc.Paracolordata(thispara) = Col()
                Exit Sub  ' stop
            ElseIf FastSymbol(A$, ":") Then
                onetime = True
            ElseIf FastSymbol(A$, ")") Then
                addone pos2, 19 ' lastcol
                pos2 = K - Len(A$) + 1
                GoTo again4
            ElseIf FastSymbol(A$, ";") Then
                addone pos2, 14
                pos2 = K - Len(A$) + 1
                GoTo again2
            ElseIf FastSymbol(A$, "}") Then
                addone pos2, 14
                addone K - Len(A$), 11
                pos2 = K - Len(A$) + 1
                GoTo again3
            ElseIf FastSymbol(A$, "{") Then
                onetime = True
            ElseIf FastSymbol(A$, "?") Then
                addone pos2, 14
                addone K - Len(A$), 11
                pos2 = K - Len(A$) + 1
                GoTo again2
            Else
                addone pos2, lastcol
                mDoc.Paracolordata(thispara) = Col()
                Exit Sub
            End If
            addone pos2, 14
            pos2 = K - Len(A$) + 1
            GoTo again
        ElseIf IsSymbol(A$, "?") Then
            addone 1, 11
            pos2 = K - Len(A$) + 1
            GoTo again2
        ElseIf IsSymbol(A$, "=") Then
            If pos2 > 0 Then
                addone pos2, 14 ' 11
            Else
                addone 1, 14 ' 11
            End If
            pos2 = K - Len(A$) + 1
            onetime = False
            GoTo again2
        ElseIf FastSymbol(A$, "{") Then
            onetime = True
            If pos2 > 0 Then
                If Col((UBound(Col()) \ 2) * 2) <> 14 Then addone pos2, 14
                                    'addone pos2, 14
                pos2 = K - Len(A$) + 1
                If FastSymbol(A$, "&") Then
                    addone pos2, 10
                    pos2 = K - Len(A$) + 1
                End If
            Else
                addone 1, 14
                pos2 = K - Len(A$) + 1
            End If
            GoTo again
        ElseIf FastSymbol(A$, "}") Then
            If pos2 > 0 Then
                addone pos2, 14
            Else
                addone 1, 14
                If MaybeIsSymbol(A$, "(") Then
                    GoTo CheckThis
                ElseIf MaybeIsSymbol(A$, ",;") Then
                    A$ = Mid$(A$, 2)
                    GoTo CheckThis
                End If
            End If
            pos2 = K - Len(A$) + 1
            GoTo again
        ElseIf pos2 = 0 Then ' new pos2=0 here
            If IsNumberLabel(A$, drop$) Then
                addone K - Len(A$) - Len(drop$) + 1, 20
                pos2 = K - Len(A$) + 1
                GoTo again
            End If
        End If
        If pos2 > 0 Then addone pos2, 14
                        'If Col(1) = 0 Then
                        'makebase 0
                        'End If
End Sub


Private Sub ColorizeInnerStatements11(A$, K As Long, pos2 As Long, feedback As Long, Optional jump As Boolean = False)
' delete this
Dim drop$, pos1 As Long, onetime As Boolean, kind As Long, lastkind As Long, feedthis As Long, counter As Long
Dim lastcol As Long, firstonly As Boolean
firstonly = True
A$ = Replace(A$, vbTab, " ")
If jump Then
kind = IsLabelOnly((A$), drop$)
If kind = 0 Then
addone pos2, 14
lastcol = 14
If FastSymbol(A$, "(") Then
 pos2 = K - Len(A$) + 1
End If
GoTo again2
End If
GoTo fromjump

End If
onetime = True

If K < 1 Then
makebase 14
mDoc.Paracolordata(thispara) = Col()
If mDoc.startfrom(thispara) = thispara Then
feedback = thispara
ReDim Col(1)
Do While mDoc.Advance(thispara)
If mDoc.startfrom(thispara) = feedback Then
mDoc.startfrom(thispara) = 0
mDoc.Paracolordata(thispara) = Col()
End If
Loop
feedback = feedback <> thispara
Exit Sub
End If

End If
If feedback > 0 Then
'lastcol = 15
A$ = Mid$(A$, feedback + 1)
'pos2 = feedback + 1
 addone feedback + 1, lastcol
feedback = 0
GoTo checkthistoo
Else
lastcol = 14
End If
again:
If RTrim$(A$) = vbNullString Then
If pos2 > 0 Then
'If col((UBound(col()) \ 2) * 2) <> lastcol Then
'addone pos2 + 1, lastcol
'Else
'If UBound(col()) < 2 Then
addone pos2, lastcol
'End If
  Else
  makebase 14
  End If
mDoc.Paracolordata(thispara) = Col()
Exit Sub
' FOR TESTS
  pos1 = 0
  For K = 1 To UBound(Col()) Step 2
    If VarType(Col(K)) = 0 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
    If Col(K) > pos1 Then pos1 = Col(K)
    If Col(K) < pos1 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
  Next K
    For K = 0 To UBound(Col()) Step 2
   If Col(K) > 15 Then mDoc.Paracolordata(thispara) = Col(): Exit Sub
  Next K
  
End If
If MaybeIsSymbol(A$, ":", "\") Then
A$ = Mid$(A$, 2)
If AscW(Left$(A$, 1) + " ") = 92 Then
If pos2 > 1 Then
    If Col(UBound(Col())) = 1 Or InStr(" #@.+-~^%$,!({'?|/&)}" + Chr$(34), Mid$(lastdata, pos2 - 1, 1)) > 0 Then

    If pos2 > 0 Then
        addone pos2, 14
    Else
        addone 1, 14
    End If
 pos2 = K - Len(A$) + 1
GoTo novar
End If
End If
A$ = Mid$(A$, 2)
End If
If Len(A$) > 1 Then
If AscW(A$) <> 32 Then
lastkind = 1
GoTo again3
End If
End If
Else
novar:
While FastSymbol(A$, ":")
If pos2 > 0 Then
addone pos2, 9
Else
addone 1, 9
End If
 pos2 = K - Len(A$) + 1
Wend
End If
drop$ = "@(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then
If IsSymbol(A$, "@") Then
        If pos2 > 0 Then
        addone pos2, 14
        Else
        addone 1, 14
        End If
     pos2 = K - Len(A$) + 1
End If
End If
again500:
kind = IsLabelDot("", A$, drop$, firstonly)
If kind = -500 Then
    addone K - Len(A$) - Len(drop$) + 1, 20
     pos2 = K - Len(A$) + 1
    kind = 0
ElseIf kind = 0 Then
drop$ = "$(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "->": kind = Abs(FastSymbol(A$, drop$, , 2))
End If
fromjump:
        If kind = 0 Then drop$ = vbNullString
        If kind > 0 Then
            If ismine(drop$) Then
                If ismine6(drop$) Then
                    lastcol = 111
                    If pos2 > 0 Then
                       addone pos2, 111
                    Else
                       addone K - Len(A$) - Len(drop$) + 1, 111
                    End If
                ElseIf ismine60(drop$) Then
                    lastcol = 112
                    If pos2 > 0 Then
                       addone pos2, 112
                    Else
                       addone K - Len(A$) - Len(drop$) + 1, 112
                    End If
                Else
                    lastcol = 170 + kind
                    If pos2 > 0 Then
                        addone pos2, lastcol
                    Else
                        addone K - Len(A$) - Len(drop$) + 1, lastcol
                    End If
                End If
                If ismine2(drop$) Then
                    kind = -2
                    If Trim$(A$) <> "" Then
                        pos2 = K - Len(A$) + 1
                        If ismine22(drop$) Then
                        ElseIf MaybeIsSymbol(A$, "{") Then
                            addone pos2, 14
                            pos2 = K - Len(A$) + 1
                            onetime = True
                            GoTo again
                        ElseIf ismine5(drop$) Then
                            GoTo again500
                        ElseIf ismine3(drop$) Then
                            onetime = True
                            GoTo again
                        End If
                    End If
                ElseIf ismine1(drop$) Then
                    kind = -1
                    If Trim$(A$) <> "" Then
                        pos2 = K - Len(A$) + 1
                        If FastSymbol(A$, "{") Then
                            addone pos2, 14
                            pos2 = K - Len(A$) + 1
                            onetime = True
                            GoTo again
                        Else
                            mDoc.Paracolordata(thispara) = Col()
                            Exit Sub
                        End If
                    End If
                ElseIf ismine33(drop$) Then
                    ' fix for case >100, =50...
                    pos2 = K - Len(A$) + 1
                    Do
                        lastcol = 14
                        addone pos2, 14
                        If FastSymbol(A$, "<") Then
                            If Not FastSymbol(A$, "=") Then
                                FastSymbol A$, ">"
                            End If
                        ElseIf FastSymbol(A$, ">") Then
                            FastSymbol A$, "="
                        ElseIf FastSymbol(A$, "=") Then
                            FastSymbol A$, "="
                        ElseIf FastSymbol(A$, "~") Then
                        
                        End If
                        pos1 = 1
                        drop$ = aheadstatus(A$, False, pos1)
                        If pos1 > 0 Then
                            drop$ = Mid$(A$, 1, pos1 - 1)
                            ColorizeExpressions drop$, K - Len(A$) + Len(drop$), pos2, feedthis
                            A$ = Mid$(A$, pos1)
                        End If
                    Loop Until Not FastSymbol(A$, ",")
                    onetime = True
                    GoTo here1
                End If
            Else
                lastcol = 10
                If Len(A$) > 0 Then
                    If AscW(A$) = 92 Then
                        lastkind = 1
                    End If
                End If
                If pos2 > 0 Then
                    addone pos2, 16 * lastcol + kind
                Else
                    lastkind = 0
                    addone K - Len(A$) - Len(drop$) + 1, 16 * lastcol + kind
                 End If
            End If
            pos2 = K - Len(A$) + 1
            If kind > 3 Then
CheckThis:
                pos1 = 1
                drop$ = aheadstatus(A$, False, pos1)
marvel:

                If pos1 > 1 Then
                    feedthis = 0
marvel2:
                    drop$ = Mid$(A$, 1, pos1 - 1)
                    ColorizeExpressions drop$, K - Len(A$) + Len(drop$), pos2, feedthis
                    If feedthis = 11 Or feedthis = 14 Or feedthis = 10 Then
                    'If kind < 5 Then lastcol = feedthis
                    feedthis = 0
                End If
newjump:
                feedback = feedthis
                If Col((UBound(Col()) \ 2) * 2) = 7 Then
                    A$ = vbNullString
                Else
                    A$ = Mid$(A$, pos1)
                End If
checkthistoo:
                If RTrim$(A$) = vbNullString Then
                    If pos2 > 0 Then
                        If feedthis <> 15 Then
                            If Col((UBound(Col()) \ 2) * 2) <> lastcol And pos2 < pos1 Then addone pos2, lastcol
                        End If
                    Else
                        makebase 14
                    End If
                    mDoc.Paracolordata(thispara) = Col()
                    If feedthis = 15 Then GoTo use15
                    Exit Sub
                End If
                If kind > 4 And onetime Then
                    If FastSymbol(A$, ",") Then
                        addone pos2, 8
                        pos2 = K - Len(A$) + 1
                        pos1 = 1
                        GoTo CheckThis
                    End If
                End If
                If FastSymbol(A$, "->", , 2) Then
                    addone pos2, 8
                    pos2 = K - Len(A$) + 1
                    If FastSymbol(A$, "{") Then
                        pos2 = K - Len(A$) + 1
                        onetime = True
                        GoTo again
                    End If
                ElseIf FastSymbol(A$, ",") Then
                    addone pos2, 8
                    pos2 = K - Len(A$) + 1
                    pos1 = 1
                End If
            ElseIf kind > 4 And onetime Then
                If FastSymbol(A$, ",") Then
                    addone pos2, 8
                    pos2 = K - Len(A$) + 1
                    pos1 = 1
                    GoTo CheckThis
                End If
            End If
            If FastSymbol(A$, ")") Then
                addone pos2, 19 ' lastcol
                pos2 = K - Len(A$) + 1
            ElseIf feedthis = 15 Then
                If FastSymbol(A$, "+") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "=") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "~") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, ">") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                    GoTo CheckThis
                ElseIf FastSymbol(A$, "<") Then
                    addone pos2, 14  ' for string
                    pos2 = K - Len(A$) + 1
                    GoTo CheckThis
                ElseIf FastSymbol(A$, "}") Then
                    If pos2 > 1 Then
                        addone pos2, 3
                    Else
                        Col(0) = 3
                    End If
                    ' for string
                    pos2 = K - Len(A$) + 1
                    GoTo CheckThis
                ElseIf MaybeIsSymbol(A$, Chr(34)) Then
                    If Not ISSTRINGA(A$, drop$) Then
                        A$ = vbNullString
                        addone pos2, 10
                    Else
                        addone pos2, 15
                    End If
                        addone pos2, 14  ' for string
                        pos2 = K - Len(A$) + 1
                        ' Chr(34))
                    Else
                        feedthis = 0
                    End If
                    '  If feedthis = 15 Then GoTo CheckThis
                    feedthis = 0
                End If
            End If
here3:
            If kind > 0 Then
                If MaybeIsSymbol(A$, ".") Then
                    If FastSymbol(A$, ".(", , 2) Then
                        addone pos2, 14
                        A$ = vbNullString
                        pos2 = K + 1
                    ElseIf FastSymbol(A$, "..", , 2) Then
                        lastcol = 10
                        lastkind = 1
                        addone pos2, lastcol
                        pos2 = K - Len(A$) + 1
                        GoTo here3
                    Else
                        pos1 = pos1 + 1
                        lastcol = 14
                        GoTo CheckThis
                    End If
                End If
            End If
            If FastSymbol(A$, "<=", , 2) Then
                addone pos2, 24
                pos2 = K - Len(A$) + 1
            Else
again4:
                If MaybeIsSymbolNoSpace(A$, "~", True) Then
                    addone pos2, 18  ' 15
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "=>", , 2) Then
                    addone pos2, 14
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "->", , 2) Then
                    addone pos2, 11
                    pos2 = K - Len(A$) + 1
                ElseIf MaybeIsSymbolLot(A$, "/*-+~^|<>", "/*-+=~^|<>") Then
                    addone pos2, 18
                    pos2 = K - Len(A$) + 1
                ElseIf FastSymbol(A$, "=") Then
                    addone pos2, 14  ' change from 15 no white =
                    pos2 = K - Len(A$) + 1
                End If
            End If
again2:
            Do
again3:
                pos1 = 1
                drop$ = aheadstatus(A$, False, pos1)
                If Len(drop$) = 0 And MaybeIsSymbol(A$, "(") Then
                    FastSymbol A$, "("
                    addone pos2, 17  ' change from 15 no white =
                    pos2 = K - Len(A$) + 1
                    GoTo again3
                ElseIf pos1 > 1 Or drop$ = "S" Then
                    If pos1 = 1 Then
                        drop$ = A$
                        pos1 = Len(A$) + 1
                        counter = 1
                    Else
                        drop$ = Mid$(A$, 1, pos1 - 1)
                        If drop$ = vbNullString Then drop$ = " "
                    End If
                    feedback = 0
                    ColorizeExpressions drop$, K - Len(A$) + Len(drop$), pos2, feedback
                    If feedback <> 15 Then
                        If feedback = 0 And lastcol = 0 Then feedback = 14
                        If drop$ = "(" Then
                            If feedback = 100 Then
                                lastcol = 14
                            Else
                                lastcol = feedback
                            End If
                        Else 'If lastcol <> 11 And lastcol <> 10 Then
                            If feedback = 100 Then
                                lastcol = 14
                            Else
                                If feedback > 0 And kind > 0 Then
                                    lastcol = feedback
                                    feedback = 0
                                    A$ = Mid$(A$, pos1)
                                    ColorizeInnerStatements A$, K, pos2, feedback, True
                                    If MaybeIsSymbol(A$, ")") Then
                                        addone pos2, 19 'lastcol
                                        FastSymbol A$, ")"
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        feedback = 0
                    End If
                    If feedback = 100 Then
                        A$ = Mid$(A$, pos1)
                        GoTo again112
                    ElseIf feedback = 15 Then
use15:
                        counter = 1
nextline:
                        feedback = 0
                        drop$ = mDoc.NextTextParagraph(thispara, feedback)
                        If feedback = 0 Then
                            A$ = Mid$(A$, pos1)
                            mDoc.Paracolordata(thispara) = Col()
                            Exit Sub
                        End If
                        pos1 = 1
                        countbrakcet drop$, pos1, counter
                        mDoc.Paracolordata(thispara) = Col()
                        A$ = drop$
                        feedthis = 15
                            'If feedthis = 15 Then
                        K = Len(A$)
                            'Else
                            'k = Len(a$) - 1
                            'End If
                        If mDoc.startfrom(thispara) = 0 Then
                            mDoc.startfrom(thispara) = thispara
                        End If
                        mDoc.startfrom(feedback) = mDoc.startfrom(thispara)
                        thispara = feedback
                        lastdata = A$
                        pos2 = 1
                        If counter > 0 Then
                            ReDim Col(3)
                            Col(0) = 15
                            Col(1) = 1
                            Col(2) = 15
                            Col(3) = K + 1
                        Else
                            ReDim Col(1)
                            Col(0) = 15
                            Col(1) = 1
                        End If
                        If counter = 0 Then
                            If UBound(Col()) > 1 Then
                                Col(3) = pos1
                                If Col(3) = Col(1) Then
                                    pos1 = pos1 + 1
                                    Col(3) = pos1
                                    Col(2) = 14
                                End If
                                '    Else
                                      '  col(1) = pos1
                            End If
                                    'If Left$(a$, 1) <> "}" Then
                            If pos1 > 0 Then
                                A$ = Mid$(A$, pos1)
                                pos2 = K - Len(A$) + 1
                            End If
                            feedback = 0 'lastcol
                        'feedthis = 0
                        Else
                            GoTo nextline
                        End If
                            'kind = 0
                           ' feedback = 0
                        pos1 = 1
                        GoTo CheckThis
                    ElseIf feedback = 0 And pos1 > 1 And kind > 0 Then
                        GoTo newjump
                    End If
                    If pos1 > 0 Then
                        If Mid$(A$, pos1 - 1, 1) = " " Then lastkind = 0
                        A$ = Mid$(A$, pos1)
                    End If
                    If MaybeIsSymbol(A$, ":{") Then
                        onetime = True
                        GoTo again
                    End If
                    If InStr(drop$, " ") = 0 Or kind >= 0 Then
                        GoTo CheckThis
                    End If
                End If
again112:
                If MaybeIsSymbol(A$, ",;", , True) Then
                    addone pos2, 8
                    pos2 = K - Len(A$) + 1
                    lastkind = 0
                    onetime = False
                ElseIf FastSymbol(A$, "}") Then
                    lastkind = 0
                    addone pos2, 14
                    pos2 = K - Len(A$) + 1
                    onetime = True
                    If MaybeIsSymbol(A$, ".;+(") Then
                        GoTo CheckThis
                    End If
                    GoTo again112
                ElseIf MaybeIsSymbol(A$, "'") Then
                    lastkind = 0
                    pos2 = K - Len(A$) + 1
                    FastSymbol A$, "'"
                    If pos2 > 0 Then
                        addone pos2, 7
                    Else
                        addone 1, 7
                    End If
                    A$ = vbNullString
                    pos2 = K
                    If pos2 > 0 Then
                        addone pos2 + 1, 14
                    End If
                    mDoc.Paracolordata(thispara) = Col()
                    Exit Sub
                Else
                    onetime = True
                    mDoc.Paracolordata(thispara) = Col()
                    Exit Do
                End If
            Loop
                            'If a$ = vbNullString Then
                            '
                            'mDoc.Paracolordata(thispara) = col()
                            'Exit Sub
                            
                            'End If
            GoTo again
        ElseIf MaybeIsSymbol(A$, "\", , True) Then
            If lastkind = 1 Then
                pos2 = K - Len(A$) + 1
                If AscW(A$ + ".") <> 32 Then
                    If MaybeIsSymbol(A$, ".", ".", True) Then
                    End If
                    If MaybeIsSymbolNoSpace(A$, "\") Then GoTo again
                    pos2 = K - Len(A$) + 1
                    lastkind = 0
                    If MaybeIsSymbol(A$, "\", , True) Then GoTo here1
                    GoTo again
                End If
            Else
here1:
                If pos2 > 0 Then
                    addone pos2, 7
                Else
                    addone 1, 7
                End If
                pos2 = K - Len(A$) + 1
                mDoc.Paracolordata(thispara) = Col()
                Exit Sub
            End If
        ElseIf MaybeIsSymbol(A$, "'") Then
            If pos2 > 1 Then
                addone pos2, 7
            Else
                addone 1, 7
            End If
            pos2 = K - Len(A$) + 1
            mDoc.Paracolordata(thispara) = Col()
            Exit Sub
        ElseIf Not onetime Then
            If FastSymbol(A$, ",") Then
                mDoc.Paracolordata(thispara) = Col()
                Exit Sub  ' stop
            ElseIf FastSymbol(A$, ":") Then
                onetime = True
            ElseIf FastSymbol(A$, ")") Then
                addone pos2, 19 'lastcol
                pos2 = K - Len(A$) + 1
                GoTo again4
            ElseIf FastSymbol(A$, ";") Then
                addone pos2, 14
                pos2 = K - Len(A$) + 1
                GoTo again2
            ElseIf FastSymbol(A$, "}") Then
                addone pos2, 14
                addone K - Len(A$), 11
                pos2 = K - Len(A$) + 1
                GoTo again3
            ElseIf FastSymbol(A$, "{") Then
                onetime = True
            ElseIf FastSymbol(A$, "?") Then
                addone pos2, 14
                addone K - Len(A$), 11
                pos2 = K - Len(A$) + 1
                GoTo again2
            Else
                addone pos2, lastcol
                mDoc.Paracolordata(thispara) = Col()
                Exit Sub
            End If
            addone pos2, 14
            pos2 = K - Len(A$) + 1
            GoTo again
        ElseIf IsSymbol(A$, "?") Then
            addone 1, 11
            pos2 = K - Len(A$) + 1
            GoTo again2
        ElseIf IsSymbol(A$, "=") Then
            If pos2 > 0 Then
                addone pos2, 14  ' 11 before
            Else
                addone 1, 14  ' 11 before
            End If
            pos2 = K - Len(A$) + 1
            onetime = False
            GoTo again2
        ElseIf FastSymbol(A$, "{") Then
            onetime = True
            If pos2 > 0 Then
                If Col((UBound(Col()) \ 2) * 2) <> 14 Then addone pos2, 14
                                    'addone pos2, 14
                pos2 = K - Len(A$) + 1
                If FastSymbol(A$, "&") Then
                    addone pos2, 10
                    pos2 = K - Len(A$) + 1
                End If
            Else
                addone 1, 14
                pos2 = K - Len(A$) + 1
            End If
            GoTo again
        ElseIf FastSymbol(A$, "}") Then
            If pos2 > 0 Then
                addone pos2, 14
            Else
                addone 1, 14
                If MaybeIsSymbol(A$, "(") Then
                    GoTo CheckThis
                ElseIf MaybeIsSymbol(A$, ",;") Then
                    A$ = Mid$(A$, 2)
                    GoTo CheckThis
                End If
            End If
                pos2 = K - Len(A$) + 1
                GoTo again
        ElseIf pos2 = 0 Then
            If IsNumberLabel(A$, drop$) Then
                addone K - Len(A$) - Len(drop$) + 1, 20
                pos2 = K - Len(A$) + 1
                GoTo again
            End If
        End If
        If pos2 > 0 Then addone pos2, 14

End Sub

Private Sub ColorizeExpressions(ByVal A$, K As Long, pos2 As Long, feedthis As Long)
Dim drop$, pos1 As Long, kind As Long, localfeedthis As Long, thispara1 As Long, lastpar As Long
Dim lastcol As Long
Dim n$
again:
If A$ = vbNullString Then
If feedthis = 0 Then
If lastpar > 0 Then
feedthis = lastpar
End If
End If
Exit Sub
End If
drop$ = "@(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then
    If IsSymbol(A$, "@") Then
        If pos2 > 0 Then
            addone pos2, 14
        Else
            addone 1, 14
        End If
        pos2 = K - Len(A$) + 1
    
    End If
getlabelonly:
    n$ = A$
    kind = IsLabelOnly(n$, drop$)
    pos2 = K - Len(A$) + 1
    If kind = 0 And n$ <> A$ Then
    feedthis = 100
    A$ = n$: Exit Sub
    End If
    A$ = n$
    
    If kind = 100 Then
    If pos2 > 0 Then
            addone pos2, 100
        Else
            addone 1, 100
        End If
         pos2 = K - Len(A$) + 1
         
    kind = 0
    GoTo again
    
    End If
End If
If kind = 0 Then
If kind = 0 Then drop$ = "$(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(A$, drop$, , 2)) * 4
If kind = 0 Then
  If FastSymbol(A$, "->", , 2) Then
   ' pos2 = k - Len(a$) - 2
    addone pos2, 11
        A$ = NLtrim$(A$)
        pos2 = K - Len(A$) + 1
        If MaybeIsSymbolNoSpace(A$, "{") Then
        pos2 = K - Len(A$) + 1
        addone pos2, 14
        A$ = Mid$(A$, 2)
         localfeedthis = 0
           pos2 = K - Len(A$) + 1
           If pos2 <= K Then
                drop$ = block(A$)
                If Len(drop$) > 0 Then
                    ColorizeInnerStatements drop$, K - Len(A$), pos2, localfeedthis
           
                    If Len(A$) > 0 Then
                    ColorizeInnerStatements A$, K, pos2, localfeedthis
                    End If
                    drop$ = vbNullString
                Else
                ColorizeInnerStatements A$, K, pos2, localfeedthis
               End If

           End If
            kind = 0
            drop$ = vbNullString
            If pos2 > K Then Exit Sub
        Else
           GoTo checklambda
            
        End If
    
    End If
End If
If kind = 0 Then drop$ = vbNullString
End If
If kind > 0 Then
    If ismine(drop$) Then
    
    lastcol = 170 + kind
    If kind > 4 Then lastpar = lastcol
    
    If pos2 > 0 Then
       addone pos2, lastcol
    Else
       addone K - Len(A$) - Len(drop$) + 1, lastcol
       End If
        '' for the end
   Else
   If kind > 4 And kind < 8 Then lastpar = 10
   If kind = 10 Then
        lastcol = 13
   ElseIf kind = 20 Then
        lastcol = 12
   Else
        lastcol = 10
   End If
   
   If pos2 > 0 Then
       addone pos2, kind + 16 * lastcol
   Else
       addone K - Len(A$) - Len(drop$) + 1, kind + 16 * lastcol
      
       End If
  End If
  pos2 = K - Len(A$) + 1
 If kind > 3 Then
checklambda:
  Do
again333:
  pos1 = 1
  drop$ = aheadstatus(A$, False, pos1)
  If Len(drop$) = 0 And MaybeIsSymbol(A$, "(") Then
    FastSymbol A$, "("
    addone pos2, 17
    pos2 = K - Len(A$) + 1
    GoTo again333
  ElseIf pos1 > 1 Then
  ColorizeExpressions Mid$(A$, 1, pos1 - 1), K - Len(A$) + pos1 - 1, pos2, feedthis
  If feedthis = 14 Or feedthis = 11 Then
  If lastcol <> 11 And lastcol <> 10 Then lastcol = feedthis
  feedthis = 0
  End If
  If feedthis > 0 And pos1 > Len(A$) Then Exit Sub
  feedthis = 0
   A$ = Mid$(A$, pos1)
   End If
   If FastSymbol(A$, ",") Then
     addone pos2, 8
    pos2 = K - Len(A$) + 1
   Else
   Exit Do
   End If
   Loop
    If MaybeIsSymbol(A$, "(") Then
addone pos2, 14
lastpar = 14
FastSymbol A$, "("
pos2 = K - Len(A$) + 1
  pos1 = 1
  drop$ = aheadstatus(A$, False, pos1)
    thispara1 = thispara
  If pos1 > 1 Then

  ColorizeExpressions Mid$(A$, 1, pos1 - 1), K - Len(A$) + pos1 - 1, pos2, feedthis
    A$ = Mid$(A$, pos1)
    End If
    If thispara1 <> thispara Then
    Exit Sub
    End If
  addone pos2, 19
FastSymbol A$, ")"
pos2 = K - Len(A$) + 1
ElseIf IsSymbol(A$, ",") Then
addone pos2, 8
pos2 = K - Len(A$) + 1
    GoTo again
ElseIf MaybeIsSymbol(A$, ")") Then
    pos2 = K - Len(A$) + 1
    If kind >= 10 Then
    addone pos2, 19
    
    ElseIf kind > 0 Then
  addone pos2, 19 ' lastcol
  Else
  addone pos2, 19
  End If
  'pos2 = k - Len(a$) + 2
  'addone pos2, 19
  FastSymbol A$, ")"
   
    pos2 = K - Len(A$) + 1
  End If
  End If

GoTo again
ElseIf MaybeIsSymbol(A$, "(", , True) Then
    addone pos2, 17
    pos2 = K - Len(A$) + 1
   GoTo again
ElseIf MaybeIsSymbol(A$, "{") Then

pos1 = 1
blockStringAhead A$, pos1
If Mid$(A$, pos1, 1) <> "}" Then
addone pos2, 15
feedthis = 15
Else
addone pos2, 16
feedthis = 0
End If
A$ = Mid$(A$, pos1)
drop$ = vbNullString
pos2 = K - Len(A$) + 1
If feedthis = 15 Then Exit Sub
addone pos2, 3
A$ = Mid$(A$, 2)
pos2 = K - Len(A$) + 1
    GoTo again

ElseIf MaybeIsSymbol(A$, Chr(34)) Then

    If Not ISSTRINGA(A$, drop$) Then
        A$ = vbNullString
        addone pos2, 10
    Else
        addone pos2, 15
        addone pos2 + Len(drop$) + 1, 3
    End If
    pos2 = K - Len(A$) + 1
    drop$ = vbNullString

    GoTo again
Else
If MaybeIsSymbol(A$, "*/+-=|<>!,") Then
If Left$(A$, 1) = "," Then
A$ = Mid$(A$, 2)
    addone pos2, 8
Else
If Left$(A$, 3) = "<=>" Then
A$ = Mid$(A$, 4)
Else
A$ = Mid$(A$, 2)

If MaybeIsSymbol(A$, "*=>") Then A$ = Mid$(A$, 2)
End If
    addone pos2, 5
    End If
    pos2 = K - Len(A$) + 1
   GoTo again
ElseIf MaybeIsSymbol(A$, ")") Then
A$ = Mid$(A$, 2)
    addone pos2, 19
    pos2 = K - Len(A$) + 1
   GoTo again
End If
If Col((UBound(Col()) \ 2) * 2) <> 14 Then addone pos2, 14
A$ = Mid$(A$, 2)
If Len(A$) > 0 Then GoTo again
End If

If pos2 > 0 Then
addone pos2, 14
End If


End Sub
Private Function blockString(s$) As String
Dim i As Long, j As Long, c As Long
Dim a1 As Boolean

If s$ = vbNullString Then Exit Function


c = Len(s$)
a1 = True
i = 1
Do
Select Case Asc(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If Asc(Mid$(s$, i, 1)) = 34 Then If Asc(Mid$(s$, i - 1, 1)) <> 92 Then Exit Do
Loop
Case 123
j = j - 1
Case 125
j = j + 1: If j = 1 Then Exit Do
End Select
i = i + 1
Loop Until i > c
If j = 1 Then
blockString = Left$(s$, i - 1)
s$ = Mid$(s$, i)
Else
blockString = s$
s$ = vbNullString
End If
End Function


Private Function ISSTRINGA(A$, r$) As Boolean
If Len(A$) < 2 Then ISSTRINGA = False: Exit Function
Dim q$, W As Long
r$ = vbNullString
q$ = Chr(34)
'W = 1
'Do
'W = W + 1
W = InStr(2, A$, q$)
'If W > 0 Then
'If Mid$(a$, W - 1, 1) <> "`" Then Exit Do
'Exit Do
'ElseIf W = 0 Then
'Exit Do
'End If
'Loop
If W = 0 Then ISSTRINGA = False: Exit Function
r$ = Mid$(A$, 2, W - 2)
A$ = Mid$(A$, W + 1)
ISSTRINGA = True
End Function
Private Sub countbrakcet(A$, K As Long, counter As Long)
Dim i As Long, j As Integer, blind As Boolean
For i = Abs(K) To Len(A$)
j = AscW(Mid$(A$, i, 1))
Select Case j
Case 34
blind = Not blind
Case 123
If Not blind Then counter = counter + 1
Case 125
If Not blind Then counter = counter - 1: If counter = 0 Then K = i: Exit For
End Select
Next i

End Sub
Function blockStringAhead(s$, pos1 As Long) As Long
Dim i As Long, j As Long, c As Long
c = Len(s$)
i = pos1
If i > c Then blockStringAhead = False: Exit Function
Do

Select Case AscW(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
Loop
Case 123
j = j - 1
Case 125
j = j + 1
End Select
i = i + 1

Loop Until i > c Or j = 0

pos1 = i - 1
blockStringAhead = True

End Function
Function IsLabelDot(where$, A$, r$, f As Boolean) As Long 'ok
' for left side...no &
Const t$ = "'\/" + vbCr

Dim rr&, one As Boolean, c$, firstdot$, gr As Boolean, i As Long
r$ = vbNullString
If A$ = vbNullString Then IsLabelDot = 0: Exit Function

A$ = NLtrim$(A$)
    Do While Len(A$) > 0
    c$ = Left$(A$, 1)
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 64  '"@"
        
            If Mid$(A$, 2, 1) <> "(" And r$ <> "" Then

                r$ = r$ + "@"
                A$ = Mid$(A$, 2)
            Else
                IsLabelDot = 0: A$ = firstdot$ + A$: Exit Function
            End If
        Case 63 '"?"
        If r$ = vbNullString And firstdot$ = vbNullString Then
        r$ = "?"
        A$ = Mid$(A$, 2)
        IsLabelDot = 1
        Exit Function
    
        ElseIf firstdot$ = vbNullString Then
        IsLabelDot = 1
        Exit Function
        Else
        IsLabelDot = 0
        Exit Function
        End If
        Case 46 '"."
            If one Then
            Exit Do
            Exit Do
            ElseIf r$ <> "" And Len(A$) > 1 Then
            If Mid$(A$, 2, 2) = ". " Or Mid$(A$, 2, 1) = " " Then Exit Do
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            rr& = 1
            Else
            firstdot$ = firstdot$ + "."
            A$ = Mid$(A$, 2)
            End If
       Case 92, 94, 123 To 126 '"\","^", "{" To "~"
        Exit Do

        Case 48 To 57, 95 '"0" To "9", "_"
           If one Then
            If firstdot$ <> "" Then A$ = firstdot$ + A$
            Exit Do
            ElseIf r$ <> "" Then
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit Do
            End If
        Case Is < 0, Is > 64 ' >=A and negative
            If one Then
            Exit Do
            Else
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        Case 36 ' "$"
        
            If one Then
            If InStr(r$, "$.") Then
            rr& = 3          ' is string variable with string pointer
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            End If
            Exit Do
            End If
            If r$ <> "" Then
            one = True
            If Mid$(A$, 2, 1) <> "." Then
            rr& = 3
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            Else
            rr& = 1       ' is string variable
            r$ = r$ & Left$(A$, 2)
            A$ = Mid$(A$, 3)
            
            End If
            
            Else
            Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            Else
            Exit Do
            End If
        Case 40 ' "("
            If r$ <> "" Then
                            If Mid$(A$, 2, 2) = ")@" Then
                                    r$ = r$ & "()."
                                  
                                 A$ = Mid$(A$, 4)
                               Else
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit Do
                                       End Select
                                       r$ = r$ & Left$(A$, 1)
                                       A$ = Mid$(A$, 2)
                                   Exit Do
                            
                          End If
               Else
                        Exit Do
            
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
            gr = True
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            rr& = 1
            End If
        End If

    Loop

    If Len(firstdot$) = 0 Then
        
        
        If f Then
        f = False
        If Left(A$, 1) = ":" Then
                i = MyTrimLi(A$, 2)
                If InStr(t$, Mid$(A$, i, 1)) > 0 Then
                    Mid$(A$, 1) = "'"
                    rr& = -500
                Else
                    r$ = myUcase(r$, gr)
                End If
                Else
                r$ = myUcase(r$, gr)
        End If
        Else
        r$ = myUcase(r$, gr)
    End If
    Else
    
       r$ = firstdot$ + myUcase(r$, gr)
       End If
    IsLabelDot = rr&

End Function

Function IsLabelOnly(A$, r$) As Long 'ok

Dim rr&, one As Boolean, c$, dot&, Extra&, num As Boolean
r$ = vbNullString
If A$ = vbNullString Then IsLabelOnly = 0: Exit Function
A$ = NLtrim$(A$)
    Do While Len(A$) > 0
     c$ = Left$(A$, 1) 'ANYCHAR HERE
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
                Case 35 ' "#"
        If r$ <> "" Then
        Exit Do
        Else
        
        r$ = "#"
           A$ = Mid$(A$, 2)
          
          rr& = 10
           End If
        Case 64  '"@"
           If r$ = vbNullString Then
              A$ = Mid$(A$, 2)
              ElseIf Mid$(A$, 2, 1) <> "(" And r$ <> "" Then
            r$ = r$ & "."
              A$ = Mid$(A$, 2)
              Else
                 IsLabelOnly = 0: Exit Function
            End If
        Case 46 '"."
            If one Then
                Exit Do
            ElseIf r$ <> "" Then
                r$ = r$ & Left$(A$, 1)
                A$ = Mid$(A$, 2)
            ElseIf Not Mid$(A$, 2, 1) Like "[0-9]" Then
                       If r$ <> "" Then
                            r$ = r$ & Left$(A$, 1)
                            rr& = 1
                        Else
                            dot& = dot& + 1
                        End If
                        If Mid$(A$, 2, 1) Like "[eEЕе]" Then
                            r$ = r$ & Mid$(A$, 2, 1)
                            A$ = Mid$(A$, 3)
                        Else
                             A$ = Mid$(A$, 2)
                        End If
            Else
                If r$ = vbNullString And dot& > 0 Then
                    r$ = String$(dot& + 1, ".")
                    A$ = Mid$(A$, 2)
                    IsLabelOnly = 1
                Else
                If MaybeIsSymbolNoSpace(Mid$(A$, 2, 1), "1234567890") Then
                    FastSymbol A$, "."
                    While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                    Wend
                    If MaybeIsSymbolNoSpace(Mid$(A$, 3, 1), "EeеЕ") Then
                        A$ = Mid$(A$, 4)
                        MaybeIsSymbolNoSpace A$, "+-", True
                        MaybeIsSymbolNoSpace A$, "123456789", True
                        While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                        Wend
                        If MaybeIsSymbolNoSpace(A$, ".") Then
                            A$ = vbNullString
                        End If
                    End If
                    If MaybeIsSymbolNoSpace(A$, ".") Then
                        IsLabelOnly = 1
                    Else
                        IsLabelOnly = 100
                    End If
                    Exit Function
                End If
                IsLabelOnly = 0
            End If
            Exit Function
        End If
     Case 92, 94, 123 To 126 '"\","^", "{" To "~"
Exit Do
        
        Case 48 To 57, 95 '"0" To "9", "_"
       
           If one Then
           
            Exit Do
            ElseIf r$ <> "" Then
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            If Extra& > 0 Then
            If Left$(A$, 1) = "%" Then
            If Len(r$) > 6 Then rr& = 1: Exit Do
            A$ = Mid$(A$, 2)
            Extra& = Extra& + 10
            rr& = 1: Exit Do
            
             ElseIf Left$(A$, 1) = "&" Then
             Extra& = Extra& + 10
            A$ = Mid$(A$, 2)
            rr& = 1: Exit Do
            End If
            If Len(r$) = 10 Then rr& = 1: Exit Do
            End If
               rr& = 1 'is an identifier or floating point variable
            ElseIf Mid$(A$ + "__", 2, 2) Like "[вВXx][0-9A-Fa-f]" Then
             r$ = r$ & Left$(A$, 3)
            A$ = Mid$(A$, 4)
               rr& = 1
            Extra& = 9
            If Left$(A$, 1) = "%" Then
                A$ = Mid$(A$, 2)
                Extra& = Extra& + 10
                rr& = 1: Exit Do
            
             ElseIf Left$(A$, 1) = "&" Then
                 Extra& = Extra& + 10
                A$ = Mid$(A$, 2)
                rr& = 1: Exit Do
            End If
            ElseIf AscW(Mid$(A$, 2, 1) + " ") = 38 Then
            A$ = Mid$(A$, 3)
            Extra& = 100
            Exit Do
            Else
           
          If dot& > 0 Then
            A$ = "." + A$: dot& = 0
            ElseIf Mid$(A$, 2, 1) = "." Then
            A$ = Mid$(A$, 3)
checkexp:
              While MaybeIsSymbolNoSpace(A$, "1234567890", True)
              Wend

              If MaybeIsSymbolNoSpace(A$, "EeеЕ", True) Then
                  MaybeIsSymbolNoSpace A$, "+-", True
                  If MaybeIsSymbolNoSpace(A$, "+-", True) Then
                  IsLabelOnly = 1
                  Exit Function
                  End If
                  MaybeIsSymbolNoSpace A$, "123456789", True
                  While MaybeIsSymbolNoSpace(A$, "1234567890", True)
                  Wend
                  If MaybeIsSymbolNoSpace(A$, ".") Then
                  IsLabelOnly = 1
                  A$ = vbNullString
                  Exit Function
                  End If
              Else
              While MaybeIsSymbolNoSpace(A$, "1234567890", True)
              Wend
               If MaybeIsSymbolNoSpace(A$, "#@&%~", True) Then IsLabelOnly = 100: Exit Function
              End If
            If MaybeIsSymbolNoSpace(A$, ".") Then
             
              IsLabelOnly = 1
              Else
              IsLabelOnly = 100
              End If
              
              Exit Function
            ElseIf MaybeIsSymbolNoSpace(A$, "1234567890", True) Then
              While MaybeIsSymbolNoSpace(A$, "1234567890", True)
              Wend
                  If MaybeIsSymbolNoSpace(A$, ".", True) Then
                                
                    GoTo checkexp
                  End If
                  MaybeIsSymbolNoSpace A$, "#@&%~", True
                  
                  
                IsLabelOnly = 100
                Exit Function
            End If
            Exit Do
            End If
Case Is < 0, Is > 64 ' >=A and negative

            If one Then
            Exit Do
            Else
  If Extra& > 0 Then
        If Not c$ Like "[a-fA-F]" Then Exit Do
        r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            
            If Left$(A$, 1) = "%" Then
                If Len(r$) > 6 Then rr& = 1: Exit Do
                Extra& = Extra& + 10
                A$ = Mid$(A$, 2)
                rr& = 1: Exit Do
            
             ElseIf Left$(A$, 1) = "&" Then
             Extra& = Extra& + 10
                A$ = Mid$(A$, 2)
                rr& = 1: Exit Do
            End If
            If Len(r$) = 10 Then rr& = 1: Exit Do
        
        
  Else
            If dot& = 0 Then
            r$ = r$ & Left$(A$, 1)
            
            A$ = Mid$(A$, 2)
       
            rr& = 1 'is an identif
            If Len(A$) > 0 Then
            If Left$(A$, 1) Like "[0-9+.\( @-\#]" Then
                If ismine66(r$, A$) Then Exit Do
            End If
            End If
            Else
            r$ = r$ & Left$(A$, 1)
            
            A$ = Mid$(A$, 2)
       
            rr& = 1 'is an identifier or floating point variable
            End If
  
  
  End If
              
        
  
            End If
      Case 38 ' "&"
            If r$ = vbNullString Then rr& = 2:    A$ = Mid$(A$, 2)
            Exit Do
       Case 36 ' "$"
       If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            Else
            Exit Do
            End If
        Case 37 ' "%"
      If one Then
            If InStr(r$, "$.") Then
            rr& = 3          ' is string variable with string pointer  ????????????
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            End If
            Exit Do
            End If
            If r$ <> "" Then
            one = True
            If Mid$(A$, 2, 1) <> "." Then
            rr& = 4
            r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            Else
            rr& = 1       ' is string variable ????
            r$ = r$ & Left$(A$, 2)
            A$ = Mid$(A$, 3)
            
            End If
            
            Else
            Exit Do
            End If
        Case 40 ' "("
            If r$ <> "" Then
                If Mid$(A$, 2, 2) = ")@" Then
                 r$ = r$ & "()."
              A$ = Mid$(A$, 4)
                Else
                Select Case rr&
                Case 1
                rr& = 5 ' float array or function
                Case 3
                rr& = 6 'string array or function
                Case 4
                rr& = 7 ' long array
                Case Else
                Exit Do
                End Select
                r$ = r$ & Left$(A$, 1)
                A$ = Mid$(A$, 2)
               Exit Do
            
            End If
           Else
           Exit Do
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
            
             If dot& = 0 Then
            r$ = r$ & Left$(A$, 1)
            
            A$ = Mid$(A$, 2)
       
            rr& = 1 'is an identif
            If Len(A$) > 0 Then
            If Left$(A$, 1) Like "[0-9+.\( @-\#]" Then
                If ismine67(r$, A$) Then Exit Do
            End If
            End If
            Else
            
             r$ = r$ & Left$(A$, 1)
            A$ = Mid$(A$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
            End If
        End If
    Loop
    If rr& > 0 Then
        If Left$(r$, 1) = "#" Then
          If Len(r$) > 1 Then
            If Mid$(r$, 2) Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                rr& = 10
            Else
                A$ = r$ + A$
                rr& = 0
            End If
         Else
        A$ = r$ + A$
        rr& = 0
        End If
        End If
    End If

    IsLabelOnly = rr& + Extra&
   
End Function
Function MaybeIsSymbolLot(A$, c$, tail$) As Boolean
Dim i As Long
If A$ = vbNullString Then Exit Function
i = 1
While Mid$(A$, i, 1) = " "
    i = i + 1
Wend
If i > Len(A$) Then Exit Function
If InStr(c$, Mid$(A$, i, 1)) > 0 Then
    If InStr(tail$, Mid$(A$, i + 1, 1)) > 0 Then
    A$ = Mid$(A$, i + 2)
    MaybeIsSymbolLot = True
    End If
End If


End Function
Function MaybeIsSymbol(A$, c$, Optional tail$ = vbNullString, Optional Remove As Boolean = False) As Boolean
Dim i As Long, ok As Boolean
If A$ = vbNullString Then Exit Function
i = 1
While Mid$(A$, i, 1) = " "
    i = i + 1
Wend
If i > Len(A$) Then Exit Function
ok = InStr(c$, Mid$(A$, i, 1)) > 0
If tail$ <> "" And ok Then
ok = Mid$(A$, i + 1, Len(tail$)) Like tail$
End If
If Remove Then If ok Then A$ = Mid$(A$, i + 1 + Len(tail$))
MaybeIsSymbol = ok
End Function
Function MaybeIsSymbolNoSpace(A$, c$, Optional Remove As Boolean = False) As Boolean
Dim ok As Boolean

If A$ = vbNullString Then Exit Function
ok = InStr(c$, Left$(A$, 1))
If Remove And ok Then A$ = Mid$(A$, 2)
MaybeIsSymbolNoSpace = ok
End Function
Function ismine60(A$) As Boolean  '  make
Select Case A$
Case "SUB", "яоутима"
ismine60 = True
End Select
End Function
Function ismine6(A$) As Boolean  '  make
Select Case A$
Case "GOTO", "пяос", "GOSUB", "диалесоу"
ismine6 = True
End Select
End Function
Function ismine66(A$, b$) As Boolean  '
If InStr("AOXaox", Left$(A$, 1)) > 0 And Len(A$) > 1 And Len(A$) < 4 Then
Select Case UCase(A$)
Case "AND", "OR", "XOR"
ismine66 = True
End Select
ElseIf InStr("DMдуdmДУ", Left$(A$, 1)) > 0 And Len(A$) = 3 Then
Select Case myUcase(A$)
Case "DIV", "MOD"
If Left$(b$, 1) = "#" Then A$ = A$ + "#": b$ = Mid$(b$, 2)
ismine66 = True
End Select
End If
End Function
Function ismine67(A$, b$) As Boolean  '
If InStr("йгаЙГАч╧", Left$(A$, 1)) > 0 And Len(A$) < 4 Then
Select Case myUcase(A$)
Case "йаи", "г", "апо"
ismine67 = True
End Select
ElseIf InStr("дДуУ", Left$(A$, 1)) > 0 And Len(A$) > 2 Then
Select Case myUcase(A$)
Case "диа", "упок", "упокоипо"
If Left$(b$, 1) = "#" Then A$ = A$ + "#": b$ = Mid$(b$, 2)
ismine67 = True
End Select
End If
End Function
Function ismine3(A$) As Boolean  ' CAN START A NEW COMMAND, PROBLEM WITH ELSE
A$ = myUcase(A$, True)
Select Case A$
Case "ELSE", "THEN", "аккиыс", "тоте"
ismine3 = True
End Select
End Function

Function ismine(A$) As Boolean
ismine = True
A$ = myUcase(A$, True)
Select Case A$
Case "@(", "$(", "~(", "?", "->", "[]"
Case "ABOUT", "ABOUT$", "ABS(", "ADD.LICENCE$(", "AFTER", "ALWAYS", "AND", "ANGLE", "APPDIR$", "APPEND", "APPEND.DOC", "APPLICATION"
Case "ARRAY", "ARRAY$(", "ARRAY(", "AS", "ASC(", "ASCENDING", "ASK$(", "ASK(", "ATN("
Case "BACK", "BACKGROUND", "BACKWARD(", "BANK(", "BASE", "BEEP", "BINARY", "BINARY.ADD(", "BINARY.AND(", "BINARY.NEG(", "BINARY.NOT("
Case "BINARY.OR(", "BINARY.ROTATE(", "BINARY.SHIFT(", "BINARY.XOR(", "BITMAPS", "BMP$(", "BOLD"
Case "BOOLEAN", "BORDER", "BREAK", "BROWSER", "BROWSER$", "BUFFER", "BUFFER(", "BYTE", "CALL", "CASE", "CAT", "CAR("
Case "CDATE(", "CDR(", "CEIL(", "CENTER", "CHANGE", "CHARSET", "CHOOSE.COLOR", "CHOOSE.FONT", "CHOOSE.OBJECT", "CHOOSE.ORGAN"
Case "CHR$(", "CHRCODE$(", "CHRCODE(", "CIRCLE", "CLASS", "CLEAR", "CLIPBOARD", "CLIPBOARD$", "CLIPBOARD.IMAGE$"
Case "CLOSE", "CLS", "CODE", "CODEPAGE", "COLLIDE(", "COLOR", "COLOR(", "COLORS"
Case "COLOUR(", "COM", "COMMAND", "COMMAND$", "COMMIT", "COMMON", "COMPARE(", "COMPRESS", "COMPUTER", "COMPUTER$", "CONCURRENT", "CONST", "CONS("
Case "CONTINUE", "CONTROL$", "COPY", "COS(", "CTIME(", "CURRENCY", "CURSOR", "CURVE"
Case "DATA", "DATE$(", "DATE(", "DATEFIELD", "DB.PROVIDER", "DB.USER", "DECIMAL", "DECLARE", "DEF", "DELETE"
Case "DESCENDING", "DESKTOP", "DIM", "DIMENSION(", "DIR", "DIR$", "DIV", "DO"
Case "DOC.LEN(", "DOC.PAR(", "DOC.UNIQUE.WORDS(", "DOC.WORDS(", "DOCUMENT", "DOS", "DOUBLE", "DOWN", "DRAW"
Case "DRAWINGS", "DRIVE$(", "DRIVE.SERIAL(", "DROP", "DRW$(", "DURATION", "EACH("
Case "EDIT", "EDIT.DOC", "ELSE", "ELSE.IF", "EMPTY", "END", "ENGLISH", "ENUM", "ENUMERATION", "ENVELOPE$(", "EOF("
Case "ERASE", "ERROR", "ERROR$", "ESCAPE", "EVAL(", "EVAL$(", "EVENT", "EVENTS", "EVERY", "EXCLUSIVE", "EXECUTE", "EXIST(", "EXIST.DIR("
Case "EXIT", "EXPORT", "EXTERN", "FALSE", "FAST", "FIELD", "FIELD$(", "FILE$("
Case "FILE.APP$(", "FILE.NAME$(", "FILE.NAME.ONLY$(", "FILE.PATH$(", "FILE.STAMP(", "FILE.TITLE$(", "FILE.TYPE$(", "FILELEN(", "FILES"
Case "FILL", "FILTER(", "FILTER$(", "FINAL", "FIND", "FKEY", "FLOODFILL", "FLOOR(", "FLUSH", "FOLD(", "FOLD$(", "FONT", "FONTNAME$", "FOR"
Case "FORM", "FORMAT$(", "FORMLABEL", "FORWARD(", "FRAC(", "FRAME", "FREQUENCY(", "FROM", "FUNCTION", "FUNCTION$(", "FUNCTION("
Case "GARBAGE", "GET", "GLOBAL", "GOSUB", "GOTO", "GRABFRAME$", "GRADIENT", "GREEK", "GROUP", "GROUP(", "GROUP$("
Case "GROUP.COUNT(", "HALT", "HEIGHT", "HELP", "HEX", "HEX$(", "HIDE", "HIDE$(", "HIGH", "HIFI", "HIGHWORD("
Case "HILOWWORD(", "HIWORD(", "HOLD", "HTML", "HWND", "ICON", "IF", "IF(", "IF$(", "IMAGE", "IMAGE(", "IMAGE.X("
Case "IMAGE.X.PIXELS(", "IMAGE.Y(", "IMAGE.Y.PIXELS(", "IN", "INFINITY", "INKEY$", "INKEY(", "INLINE", "INPUT", "INPUT$("
Case "INSERT", "INSTR(", "INT(", "INTEGER", "INTERVAL", "INVENTORY", "IS", "ISLET", "ISNUM", "ISWINE", "ITALIC"
Case "j0YPAD", "j0YPAD(", "j0YPAD.ANALOG.X(", "j0YPAD.ANALOG.Y(", "j0YPAD.DIRECTION(", "JPG$(", "KEEP", "KEY$", "KEYBOARD"
Case "KEYPRESS(", "LAMBDA", "LAMBDA(", "LAMBDA$", "LAMBDA$(", "LAN$", "LATIN", "LAYER", "LAZY$(", "LCASE$(", "LEFT$(", "LEFTPART$(", "LEGEND", "LEN"
Case "LEN(", "LEN.DISP(", "LET", "LETTER$", "LIB", "LICENCE", "LINE", "LINESPACE", "LINK", "LIST", "LN("
Case "LOAD", "LOAD.DOC", "LOCAL", "LOCALE", "LOCALE$(", "LOCALE(", "LOG(", "LONG", "LOOP"
Case "LOWORD(", "LOWWORD(", "LTRIM$(", "MAIN.TASK", "MAP(", "MARK", "MASTER", "MATCH(", "MAX(", "MAX$(", "MAX.DATA$("
Case "MAX.DATA(", "MDB(", "MEDIA", "MEDIA.COUNTER", "MEMBER$(", "MEMBER.TYPE$(", "MEMO", "MEMORY", "MENU"
Case "MENU$(", "MENU.VISIBLE", "MENUITEMS", "MERGE.DOC", "METHOD", "MID$(", "MIN(", "MIN$(", "MIN.DATA$(", "MIN.DATA("
Case "MOD", "MODE", "MODULE", "MODULE$", "MODULE(", "MODULES", "MODULE.NAME$", "MONITOR", "MONITOR.STACK", "MONITOR.STACK.SIZE", "MOTION", "MOTION.W", "MOTION.WX"
Case "MOTION.WY", "MOTION.X", "MOTION.XW", "MOTION.Y", "MOTION.YW", "MOUSE", "MOUSE.ICON", "MOUSE.KEY", "MOUSE.X"
Case "MOUSE.Y", "MOUSEA.X", "MOUSEA.Y", "MOVE", "MOVIE", "MOVIE.COUNTER", "MOVIE.DEVICE$", "MOVIE.ERROR$", "MOVIE.STATUS$"
Case "MOVIES", "MUSIC", "MUSIC.COUNTER", "NAME", "NEW", "NEXT"
Case "NORMAL", "NOT", "NOTHING", "NOW", "NUMBER", "OFF", "OLE", "ON"
Case "OPEN", "OPEN.FILE", "OPEN.IMAGE", "OPERATOR", "OPTIMIZATION", "OR", "ORDER", "ORDER(", "OSBIT", "OS$", "OUT", "OUTPUT"
Case "OVER", "OVERWRITE", "PAGE", "PARAGRAPH$(", "PARAGRAPH(", "PARAGRAPH.INDEX(", "PARAM(", "PARAM$(", "PARAMETERS$", "PART", "PARENT", "PASSWORD"
Case "PATH", "PATH$(", "PAUSE", "PEN", "PI", "PIECE$(", "PIPE", "PIPENAME$(", "PLATFORM$", "PLAY"
Case "PLAYER", "PLAYSCORE", "POINT", "POINTER", "POINTER(", "POINT(", "POLYGON", "POS", "POS(", "POS.X", "POS.Y", "PRINT"
Case "PRINTER", "PRINTERNAME$", "PRINTING", "PRIVATE", "PROFILER", "PROPERTY", "PROPERTY(", "PROPERTY$(", "PROPERTIES", "PROPERTIES$", "PROTOTYPE", "PSET", "PUBLIC", "PUSH", "PUT", "QUEUE", "QUOTE$("
Case "RANDOM", "RANDOM(", "READ", "READY(", "RECORDS(", "RECURSION.LIMIT", "REFER", "REFRESH", "RELEASE", "REM"
Case "REMOVE", "REPEAT", "REPLACE$(", "REPORT", "REPORTLINES", "RESTART", "RETRIEVE", "RETURN", "REV(", "REVISION"
Case "RIGHT", "RIGHT$(", "RIGHTPART$(", "RINSTR(", "RND", "ROUND(", "ROW", "RTRIM$(", "SAVE", "SAVE.AS", "SAVE.DOC", "SCALE.X"
Case "SCALE.Y", "SCAN", "SCORE", "SCREEN.PIXELS", "SCREEN.X", "SCREEN.Y", "SCRIPT", "SCROLL", "SEARCH"
Case "SEEK", "SEEK(", "SELECT", "SEQUENTIAL", "SET", "SETTINGS", "SGN(", "SHIFT", "SHIFTBACK", "SHORTDIR$("
Case "SHOW", "SHOW$(", "SIN(", "SINGLE", "SINT(", "SIZE", "SIZE.X(", "SIZE.Y(", "SLICE(", "SLOW", "SMOOTH"
Case "SND$(", "SORT", "SOUND", "SOUNDREC", "SOUNDS", "SPEECH", "SPEECH$(", "SPLIT", "SPRITE"
Case "SPRITE$", "SQRT(", "STACK", "STACK(", "STACK$(", "STACK.SIZE", "STACKITEM$(", "STACKITEM(", "STACKTYPE$(", "START", "STATIC"
Case "STEP", "STEREO", "STOCK", "STOP", "STR$(", "STREAM", "STRING", "STRING$(", "STRREV$(", "STRUCTURE", "SUB", "SUBDIR", "SUM(", "SUPERCLASS"
Case "SWAP", "SWEEP", "SWITCHES", "TAB", "TAB(", "TABLE", "TAN(", "TARGET"
Case "TARGETS", "TASK.MAIN", "TEMPNAME$", "TEMPORARY$", "TEST", "TEST(", "TEXT", "THEN", "THIS"
Case "THREAD", "THREAD.PLAN", "THREADS", "THREADS$", "TICK", "TIME$(", "TIME(", "TIMECOUNT", "TITLE", "TITLE$("
Case "TO", "TODAY", "TONE", "TOP", "TRIM$(", "TRUE", "TRY", "TUNE", "TWIPSX"
Case "TWIPSY", "TYPE", "TYPE$(", "UCASE$(", "UINT(", "UNARY", "UNDER", "UNICODE", "UNION.DATA$(", "UNIQUE", "UNTIL"
Case "UP", "UPDATABLE", "UPDATE", "USE", "USER", "USERS", "USER.NAME$", "USGN("
Case "VAL(", "VAL$(", "VALID(", "VALUE", "VALUE(", "VALUE$", "VERSION", "VIEW", "VOID", "VOLUME"
Case "WAIT", "WCHAR", "WEAK", "WEAK$(", "WHILE", "WIDE", "WIDTH", "WIN", "WINDOW"
Case "WITH", "WITHEVENTS", "WORDS", "WRITABLE(", "WRITE", "WRITER", "X.TWIPS", "XOR", "Y.TWIPS", "аццкийа", "адеиас", "апеияо", "апой.ая$(", "апой.де$("
Case "адеиасе", "ахя(", "ай(", "айеяаио.дуадийо(", "айеяаиос", "акт", "акгхес", "акгхгс", "аккацг", "аккацг$("
Case "аккане", "аккиыс", "аккиыс.ам", "ам", "ам(", "амап(", "амап$(", "амт(", "ам$(", "ама", "амафгтгсг", "амахеыягсг", "амайтгсг", "амакоцио"
Case "амакоцио$", "амакусг.охомгс", "амакусг.у", "амакусг.в", "амакутгс", "амаломг", "амамеысг", "амажояа", "амаье"
Case "амехесе", "амоицла.аявеиоу", "амоицла.еийомас", "амоине", "амтецяаье", "амтицяаье", "амы", "аниа", "аниа(", "аниа$", "аниа(", "апая", "апаяихлгсг", "апедысе"
Case "апкос", "апо", "апохгйеусг.ыс", "апой$(", "апойопг", "апойкеистийа", "апок(", "аяца", "аяихлос", "аяихлос.паяацяажоу("
Case "аяис$(", "аяистеяолеяос$(", "аявеиа", "аявеио", "аявеио$(", "аявеиоу.лгйос(", "аявеиоу.сталпа(", "аявг", "аукос"
Case "аукос$(", "аукоу", "ауноуса", "ауто", "ажаияесг", "ажгсе", "баке", "баке.адеиа$(", "басг"
Case "басг(", "басг.паяовос", "басг.вягстгс", "баье", "бектистопоигсг", "бгла", "богхеиа", "цецомос", "цецомота", "целисе", "целисла"
Case "цемийес", "цемийг", "цемийо", "циа", "цомийо", "цяалла", "цяалла$", "цяаллатосеияа", "цяаллатосеияа$", "цяаллесамажояас", "цяаллг"
Case "цяажг$(", "цяаье", "цягцояа", "цымиа", "дапед(", "деийтг.лояжг", "деийтгс", "деийтгс(", "деийтгс.йол", "деийтгс.у", "деийтгс.в"
Case "деийтгса.у", "деийтгса.в", "деине", "дей(", "дейаен", "дейаен$(", "дем", "дениос", "дени$(", "денилеяос$(", "дес", "дглосио"
Case "диа", "диабасе", "диацяажг", "диадовийо", "диайопг", "диайоптес", "диалесоу", "диаяхяысг", "диаяхяысг(", "диаяйеиа", "диастасг("
Case "диастиво", "диажамеиа", "диажамеиа$", "диажамо", "диажуцг", "диайопг", "диейоье", "дийтуо$", "диояхысе"
Case "дипка", "дипкос", "дойилг", "дойилг(", "долг", "дяолеас", "дуадийг.пеяистяожг(", "дуадийо", "дуадийо(", "дуадийо.айеяаио("
Case "дуадийо.амти(", "дуадийо.амтистяожо(", "дуадийо.апо(", "дуадийо.г(", "дуадийо.йаи(", "дуадийо.окисхгсг(", "дуадийо.ови(", "дуадийо.пяосхесг(", "дуадийо.пяо(", "дуолиса(", "дысе"
Case "еццяажес(", "еццяажо", "еццяажоу.кенеис(", "еццяажоу.лгйос(", "еццяажоу.ломадийес.кенеис(", "еццяажоу.пая(", "еццяаьило(", "ецйуяо(", "еий$("
Case "еийома", "еийома(", "еийома.у(", "еийома.у.сглеиа(", "еийома.в(", "еийома.в.сглеиа(", "еийомес", "еийомидио", "еимая", "еимаи", "еимця"
Case "еисацыцг", "еисацыцг$(", "еисацыцгс", "ейдосг", "ейтекесг", "ейтупысг", "ейтупысгс", "ейтупытгс", "ейтупытгс$", "ейжя(", "ейжя$("
Case "ейжяасг(", "ейжяасг$(", "екецвос", "екецвос.сыяоу", "екецвос.лецехос.сыяоу", "еккгмийа", "емаомола$", "емхесг", "емйол$", "емйол(", "емтасг", "емтокг$"
Case "емы", "емысе", "CONS(", "емысг.сеияас$(", "емысг(", "енацыцг", "енодос", "енытеяийг", "епам$(", "епамакабе", "епамекабе"
Case "епамы", "епекене", "епекене.амтийеилемо", "епекене.цяаллатосеияа", "епекене.ояцамо", "епекене.вяыла", "епицяажг", "епийаияо", "епикене", "епикене.амтийеилемо", "епикене.цяаллатосеияа"
Case "епикене.ояцамо", "епикене.вяыла", "епикоцес", "епикоцес$(", "епикоцес.жамеяес", "епикоцг", "епикоцг$(", "епикоцгс", "епипедо"
Case "епистяожг", "епижамеиа", "еполемо", "еполема(", "етийета.жоялас", "етоило(", "еуяеиа", "еуяесг", "еуяиа", "ежап(", "ежаялоцг.аявеиоу$(", "ежаялоцг.йат$", "ежаялоцг"
Case "еыс", "г", "гл(", "глеяа$(", "глеяа(", "глеяолгмиа", "гво$(", "гвоцяажгсг"
Case "гвои", "гвос", "хесе", "хесе(", "хесг", "хесг(", "хесг.у", "хесг.в", "хесгдениа(", "идиотгта", "идиотгта(", "идиотгта$("
Case "идиотгтес", "идиотгтес$", "идиытийо", "исвмг", "исвмг$(", "ивмос", "йахаяг", "йахаяо", "йахе", "йахе(", "йаи", "йакесе", "йалпукг"
Case "йаме", "йамомийа", "йат", "йат$", "йатакоцои", "йатакоцос", "йатастасг", "йатастасг.таимиас$", "йатавыягсг", "йаты"
Case "йатылисо(", "йеилемо", "йемг", "йемо", "йемтяо", "йеж$(", "йимгсг", "йимгсг.п", "йимгсг.пу"
Case "йимгсг.пв", "йимгсг.у", "йимгсг.уп", "йимгсг.в", "йимгсг.вп", "йкасг", "йкеиди", "йкеисе", "йкилан.у"
Case "йкилан.в", "йол$", "йоимос", "йоимои", "йомсока", "йяата", "йяатгсе", "йяужо$(", "йяуье", "йуйкийа", "йуйкос"
Case "йукисг", "йуяио", "йуяио.еяцо", "йыд(", "йыдийа", "йыдийосекида", "кабг", "кабг(", "кабг.амакоцийо.у("
Case "кабг.амакоцийо.в(", "кабг.йатеухумсг(", "кахос", "кахос$", "кахос.таимиас$", "калда", "калда(", "калда$", "калда$(", "катимийа", "кенеис", "киста", "коц("
Case "коцийос", "коцистийос", "коцос", "коцос$(", "кс$", "кж(", "лайяус", "ле", "лец(", "лец$(", "лецако("
Case "лецако.сеияас$(", "лецако.сеияас(", "лецецомота", "лецехос", "лецехос.сыяоу", "лецехос.у(", "лецехос.в(", "леходос", "лекос$(", "лекоус.тупос$("
Case "лекыдиа", "леяос", "леяос(", "леяос$(", "лес$(", "лета", "летахесг", "летахесг(", "левяи", "лгйос", "лгйос(", "лгйос.елж("
Case "лий(", "лий$(", "лийяо(", "лийяо.сеияас$(", "лийяо.сеияас(", "лийяос.йатакоцос$(", "лмглг", "ломадиаио", "ломадиаиос", "ломадийо", "лояжг$(", "лоусийг", "лоусийг.летягтгс", "лпип"
Case "лпяоста(", "маи", "меа", "мео", "меои", "меос", "мгла", "мглата", "мглата$"
Case "нейима", "одгциа", "одгцос$(", "охомг", "олака", "ойм$(", "олада", "олада(", "олада$(", "олада.сумоко(", "омола", "омола.аявеиоу$("
Case "омола.аявеиоу.ломо$(", "омола.тлглатос$", "омола.вягстг$", "ояио.амадяолгс", "ояисе", "ояож(", "оуяа", "ови", "паифеижымг", "паийтгс", "паине", "пай(", "пай$(", "памта"
Case "памы", "памылисо(", "паяацяажос$(", "паяацяажос(", "паяал(", "паяал$(", "паяахесг$(", "паяахуяо", "паяалетяои$", "паяе", "паяейаяе$"
Case "паяелбокг", "патглемо(", "павос", "педиа", "педио", "педио$(", "пеф$(", "пема", "пеяи"
Case "пеяи$", "пеяихыяио", "пета", "пи", "пимайас", "пимайас$(", "пимайас(", "пимайес", "писы("
Case "пкациа", "пкаисио", "пкатос", "пкатос.сглеиоу", "пкатжояла$", "пкгйтяокоцио", "покуцымо"
Case "пяос", "пяосаялоцгс", "пяосхесе.еццяажо", "пяосхгйг", "пяытотупо", "пяыто(", "пяосыяимо$", "пяовеияо", "пяовеияо$", "пяовеияо.еийома$", "яеула", "яифа("
Case "яоутима", "яоутимас", "яухлисеис", "яыта$(", "яыта(", "саяысе", "сбгсе", "се"
Case "сеияа", "сеияиайос.дисйоу(", "секида", "семаяио", "сгл", "сгл(", "сглади", "сглеио", "сглеио(", "сглеяа", "сйоупидиа", "стахеяг", "стахеяес", "статийг", "статийес"
Case "стг", "стгкг", "стгкг(", "стгм", "сто", "стой", "стовои", "стовос", "стяоцц(", "суццяажеас"
Case "суццяажг", "суцйяиме(", "суцйяоусг(", "суцвымеусе.еццяажо", "сулпиесг", "сулпкгяысг", "сум(", "сумаятгсг", "сумаятгсг$("
Case "сумаятгсг(", "сумевисе", "сумхгла", "сус", "сусйеуг.пяобокгс$", "сустгла", "сувмотгта(", "свд$(", "сведиа"
Case "сведио.мглатым", "сыяос", "сыяос(", "сыяос$(", "сыяоутупос$(", "сысе", "сысе.еццяажо", "таимиа", "таимиа.летягтгс", "таимиес"
Case "танг", "танг(", "танимолгсг", "таутисг(", "таутовяомо", "текестг", "текестгс", "текийг", "текийо", "текийос", "текос", "текос(", "тек(", "тий", "титкос.аявеиоу$(", "тилг"
Case "тилг(", "тилг$(", "тилгсыяоу$(", "тилгсыяоу(", "типота", "титкос", "титкос$(", "тлгла", "тлгла(", "тлгла$", "тлглата", "томос"
Case "тон.еж(", "топийа", "топийес", "топийг", "топийо", "топийо$(", "топийо(", "топос$(", "топос.аявеиоу$("
Case "тоте", "тяап(", "тупос", "тупос$(", "тупос.аявеиоу$(", "тупысе", "туваиос", "туваиос(", "тыяа", "у.сглеиа"
Case "упаявеи(", "упаявеи.йатакоцос(", "упеяйкасг", "уплея(", "упо", "упойатакоцос", "упок", "упокоцистг", "упокоцистгс$", "упокоипо"
Case "уполмгла", "упыяа(", "уьгкос", "уьос", "уьос.сглеиоу", "жайекос$(", "жамеяо$(", "жаядиа", "жеяе"
Case "жеяеписы", "жхимоуса", "жиктяо(", "жиктяо$(", "жомто", "жояла", "жояла$", "жоятос", "жоятысе", "жоятысе.еццяажо"
Case "жяасг", "жымг", "жыто$(", "в.сглеиа", "вая$(", "ваяайтгяес", "ваяане", "ваяйыд$("
Case "ваяйыд(", "вягсг", "вягстг", "вягстгс", "вягстым", "вяомос$(", "вяомос(", "вяыла", "вяыла(", "вяылата", "вяылатисе"
Case "выяисла", "ьеудес", "ьеудгс", "ьгжио", "ыс"
Case Else
ismine = False
End Select
End Function
Function ismine1(A$) As Boolean  '  START A BLOCK
ismine1 = True
A$ = myUcase(A$, True)
Select Case A$
Case "PART", "LIB", "PROTOTYPE"
Case "леяос", "пяытотупо"
Case Else
ismine1 = False
End Select
End Function
Function ismine2(A$) As Boolean  ' CAN START A BLOCK OR DO SOMETHING
ismine2 = True
A$ = myUcase(A$, True)
Select Case A$
Case "ABOUT", "AFTER", "BACK", "BACKGROUND", "CLASS", "COLOR", "DECLARE", "ELSE", "ENUM", "ENUMERATION", "EVENT", "EVERY", "GLOBAL", "FOR", "FKEY", "FUNCTION", "GROUP", "INVENTORY", "LAYER", "LOCAL", "MAIN.TASK", "MODULE", "OPERATOR", "PATH", "PEN", "PROPERTY", "PRINTER", "PRINTING", "REMOVE", "SET", "STACK", "START", "STRUCTURE", "TASK.MAIN", "THEN", "THREAD", "TRY", "WIDTH", "VALUE", "WHILE"
Case "аявг", "аккиыс", "аниа", "аниа(", "апая", "апаяихлгсг", "цецомос", "цемийо", "цемийг", "цемийес", "циа", "дес", "диацяажг", "долг", "ейтупытгс", "ейтупысг", "емы", "епипедо", "хесе", "хесе(", "идиотгтес", "ивмос", "идиотгта", "йахе", "йатастасг", "йкасг", "йкеиди", "йуяио.еяцо", "лета", "мгла", "олада", "ояисе", "павос", "пема", "пеяи", "пеяихыяио", "сумаятгсг", "сыяос", "текестгс", "тлгла", "топийа", "топийг", "топийес", "тоте", "вяыла"
Case "CONST", "стахеяг", "стахеяес", "SUPERCLASS", "упеяйкасг", "DO", "REPEAT", "епамекабе", "епамакабе"
Case "->"
Case Else
ismine2 = False
End Select
End Function
Function ismine22(A$) As Boolean  ' CAN START A BLOCK AFTER AN EXPRESSION, WE CAN PASS STRING BLOCK IN EXPRESSION
A$ = myUcase(A$, True)
Select Case A$
Case "FOR", "WHILE", "циа", "емы"
ismine22 = True
End Select
End Function
Function ismineEnd(A$) As Boolean
A$ = myUcase(A$, True)
Select Case A$
Case "END", "текос"
ismineEnd = True
End Select
End Function
Function ismine33(A$) As Boolean
A$ = myUcase(A$, True)
Select Case A$
Case "CASE", "ле"
ismine33 = True
End Select
End Function

Function ismine5(A$) As Boolean  '  make
A$ = myUcase(A$, True)
Select Case A$
Case "GLOBAL", "цемийо", "цемийг", "цемийес"
ismine5 = True
End Select
End Function

Public Function MyTrimRW(s$) As String
Dim i&, L As Long
Dim P2 As Long, P1 As Integer, p4 As Long, p22 As Long
L = Len(s): If L = 0 Then Exit Function

  P2 = StrPtr(s): L = L - 1
  p22 = P2
  p4 = P2 + L * 2
  For i = p4 To P2 Step -2
  GetMem2 i, P1
  Select Case P1
    Case 32, 160, 9
    Case Else
     Exit For
  End Select
  Next i
  p4 = i
   If P2 > p4 Then MyTrimRW = vbNullString Else MyTrimRW = Mid$(s$, (P2 - p22) \ 2 + 1, (p4 - P2) \ 2 + 1)
 
End Function
