Attribute VB_Name = "ServerMod"
Public Function CallEventFromSocketNow(sck As Server, a As mEvent, aString$, vrs()) As Boolean
Dim tr As Boolean, extr As Boolean, olescok As Boolean
olescok = escok
escok = False
CallEventFromSocketNow = True
extr = extreme
extreme = True
tr = trace
If Rnd * 100 > 3 Then trace = False
Dim n$, f$, F1$, bb As mStiva, oldbstack As mStiva, nowtotal As Long
Dim bstack As basetask
Set bstack = New basetask
Set bstack.Owner = Form1.DIS
Set bstack.StaticCollection = EventStaticCollection
bstack.IamAnEvent = True
Dim i As Long
If a Is Nothing Then GoTo conthere0
i = a.VarIndex
F1$ = sck.modulename
Set oldbstack = bstack.soros
Dim j As Long, k As Long, s1$, klm As Long, s2$
Dim ohere$
ohere$ = here$
here$ = "EV" + CStr(i)
If a.enabled Then
    a.ReadVar 0, n$, f$
    If f$ <> "" Then
        Set bb = New mStiva
        Set bstack.Sorosref = bb
        PushStage bstack, False
            For k = LBound(vrs()) To UBound(vrs()) - 1
                If VarType(vrs(k)) = vbString Then
                    globalvarGroup "EV" + CStr(i + k) + "$", vrs(k)
                    bb.DataStr here$ + "." + "EV" + CStr(i + k) + "$"
                Else
                    globalvarGroup "EV" + CStr(i + k), vrs(k)
                    bb.DataStr here$ + "." + "EV" + CStr(i + k)
                End If
            Next k
            bb.DataObj sck
            FastPureLabel aString$, f$, , , , , False
            n$ = Mid$(aString$, Len(f$) + 1)
            If Len(n$) > 0 Then
                n$ = Left$(n$, Len(n$) - 1)
                If n$ <> "" Then n$ = "Push " + n$ + vbCrLf
            End If
            If F1$ <> "" Then f$ = myUcase(F1$ + "." + f$ + ")", True) Else f$ = myUcase(f$ + ")", True)
            If Not GetSub(f$, klm) Then PopStage bstack: bb.Flush: CallEventFromSocketNow = False: GoTo conthere
            s1$ = sbf(klm).sb
            If Left$(s1$, 10) = "'11001EDIT" Then
                SetNextLine s1$
            End If
            If F1$ <> "" Then s1$ = n$ + "Module " + F1$ + vbCrLf + sbf(klm).sb Else s1$ = n$ + sbf(klm).sb
            Dim nn As Long
           
            If Execute(bstack, s1$, False, False) = 0 Then
                MyEr "Problem in Event " + aString$, "�������� ��� ������� " + aString$
                PopStage bstack
                bb.Flush
                GoTo conthere
            End If
            here$ = "EV" + CStr(i)
            For k = LBound(vrs()) To UBound(vrs()) - 1
                If VarType(vrs(k)) = vbString Then
                    GetlocalVar "EV" + CStr(i + k) + "$", j
                    vrs(k) = var(j)
                Else
                    GetlocalVar "EV" + CStr(i + k), j
                    vrs(k) = var(j)
                End If
            Next k
            PopStage bstack
            bb.Flush
        End If
    End If
conthere:
    Set bstack.Sorosref = oldbstack
    here$ = ohere$
conthere0:
    Set oldbstack = Nothing
    Set bb = Nothing
    If tr Then trace = tr
    extreme = extr
    escok = olescok
End Function
Public Function CallEventFromSocketOne(sck As Server, a As mEvent, aString$) As Boolean
Dim tr As Boolean, extr As Boolean, olescok As Boolean
CallEventFromSocketOne = True
olescok = escok
escok = False
tr = trace
extr = extreme
extreme = True
If Rnd * 100 > 3 Then trace = False
Dim n$, f$, F1$, bb As mStiva, uIndex As Long
Dim bstack As basetask
Set bstack = New basetask
Set bstack.Owner = Form1.DIS
Set bstack.StaticCollection = EventStaticCollection
bstack.IamAnEvent = True
Dim i As Long
If a Is Nothing Then GoTo conthere0
i = a.VarIndex
uIndex = sck.index
If uIndex >= 0 Then
bstack.soros.DataVal CDbl(uIndex)
uIndex = 1
End If
uIndex = uIndex + 1
F1$ = sck.modulename
bstack.soros.DataObj sck

Dim j As Long, k As Long, s1$, klm As Long, s2$
Dim ohere$
ohere$ = here$
here$ = "EV" + CStr(i)
If a.enabled Then
            PushStage bstack, False
            FastPureLabel aString$, f$, , , , , False
            n$ = Mid$(aString$, Len(f$) + 1)
            n$ = Left$(n$, Len(n$) - 1)
            If n$ <> "" Then
           If uIndex > 0 Then
            n$ = "Data " + n$ + " : ShiftBack Stack.Size" + Str(1 - uIndex) + "," + Str$(uIndex) + vbCrLf
            Else
            n$ = "Data " + n$ + " : ShiftBack Stack.Size" + vbCrLf
            End If
            End If
            If F1$ <> "" Then f$ = myUcase(F1$ + "." + f$ + ")", True) Else f$ = myUcase(f$ + ")", True)
            If Not GetSub(f$, klm) Then
            PopStage bstack: CallEventFromSocketOne = False: GoTo conthere
            End If
            s1$ = sbf(klm).sb
            If Left$(s1$, 10) = "'11001EDIT" Then
            SetNextLine s1$
            End If
            If F1$ <> "" Then s1$ = n$ + "Module " + F1$ + vbCrLf + sbf(klm).sb Else s1$ = n$ + sbf(klm).sb
            If Execute(bstack, s1$, False, False) = 0 Then
            If F1$ = vbNullString Then
            MyEr "Problem in Event " + aString$, "�������� ��� ������� " + aString$
            Else
            MyEr "Problem in Event " + aString$ + " in module " + F1$, "�������� ��� ������� " + aString$ + " ��� ����� " + F1$
            End If
            bstack.soros.Flush
                PopStage bstack
                GoTo conthere
            End If
            PopStage bstack
End If
conthere:
Set bstack = Nothing
here$ = ohere$
conthere0:
If tr Then trace = tr
extreme = extr
escok = olescok
End Function
