VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsyncCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Event Complete(x As Long)

Public Function CallEventFromGui(gui As Object, a As mEvent, aString$) As Boolean
Dim tr As Boolean, extr As Boolean, olescok As Boolean
olescok = escok
escok = False
extr = extreme
extreme = True
tr = trace
If Rnd * 100 > 3 Then trace = False
On Error Resume Next

CallEventFromGui = True
Dim n$, f$, bb As mStiva, oldbstack As mStiva, nowtotal As Long
Dim bstack As basetask
Set bstack = basestack1
bstack.IamAnEvent = True
Dim i As Long
If a Is Nothing Then GoTo conthere1
i = a.VarIndex
bstack.soros.DataStr aString$
If gui.index >= 0 Then
bstack.soros.DataVal gui.index
End If
bstack.soros.DataObj gui

Set oldbstack = bstack.soros
Dim j As Long, s1$, klm As Long
Dim ohere$
ohere$ = here$
For j = 0 To a.Count - 1
here$ = "EV" + CStr(i) + "." + CStr(j)
If a.enabled Then
a.ReadVar j, n$, f$
If f$ <> "" Then
Set bb = New mStiva
Set bstack.Sorosref = bb
            bb.Copy2TopNItems2FromStiva a.params, oldbstack
            PushStage bstack, False
            s1$ = Mid$(f$, 2, rinstr(f$, "}") - 2)
            klm = GlobalSub("A_()", s1$, Trim$(Mid$(f$, Len(s1$) + 3)))
            
            If Not ProcModuleEntry(bstack, "A_()", klm, "") Then
                PopStage bstack
                bb.Flush
                GoTo conthere
            End If
            PopStage bstack

bb.Flush
End If
End If
Next j
conthere:
Set bstack.Sorosref = oldbstack
Set oldbstack = Nothing
bstack.soros.drop a.params
Set bb = Nothing
here$ = ohere$
conthere1:
extreme = extr
If tr Then
'If STEXIT Then trace = tr
trace = tr
End If
escok = olescok
RaiseEvent Complete(3)
End Function
Public Function CallEventFromGuiOne(gui As Object, a As mEvent, aString$) As Boolean
Dim tr As Boolean, extr As Boolean, olescok As Boolean
CallEventFromGuiOne = True
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
bstack.IamAnEvent = True
Dim i As Long
If a Is Nothing Then GoTo conthere0
i = a.VarIndex
uIndex = gui.index
If uIndex >= 0 Then
bstack.soros.DataVal CDbl(uIndex)
uIndex = 1
End If
uIndex = uIndex + 1
F1$ = gui.modulename$
bstack.soros.DataObj gui

Dim j As Long, k As Long, s1$, klm As Long, s2$
Dim ohere$
ohere$ = here$
here$ = "EV" + CStr(i)
If a.enabled Then
            PushStage bstack, False
            IsLabelOnly (aString$), f$
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
            If Not GetSub(f$, klm) Then PopStage bstack: CallEventFromGuiOne = False: GoTo conthere
            s1$ = sbf(klm).sb
            If Left$(s1$, 10) = "'11001EDIT" Then
            SetNextLine s1$
            End If
            If F1$ <> "" Then s1$ = n$ + "Module " + F1$ + vbCrLf + sbf(klm).sb Else s1$ = n$ + sbf(klm).sb
            If Execute(bstack, s1$, False, False) <> 1 Then
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
If tr Then
'If STEXIT Then
trace = tr
End If
extreme = extr
escok = olescok
RaiseEvent Complete(2)
End Function
Public Function CallEventFromGuiNow(gui As Object, a As mEvent, aString$, vrs()) As Boolean
Dim tr As Boolean, extr As Boolean, olescok As Boolean
olescok = escok
escok = False
CallEventFromGuiNow = True
extr = extreme
extreme = True
tr = trace
If Rnd * 100 > 3 Then trace = False
Dim n$, f$, F1$, bb As mStiva, oldbstack As mStiva, nowtotal As Long
Dim bstack As basetask
Set bstack = New basetask
Set bstack.Owner = Form1.DIS
bstack.IamAnEvent = True
Dim i As Long
If a Is Nothing Then GoTo conthere0
i = a.VarIndex
F1$ = gui.modulename$
Set oldbstack = bstack.soros
Dim j As Long, k As Long, s1$, klm As Long, s2$
Dim ohere$
ohere$ = here$
'For j = 0 To a.Count - 1

here$ = "EV" + CStr(i)

If a.enabled Then
a.ReadVar 0, n$, f$
If f$ <> "" Then
Set bb = New mStiva
Set bstack.Sorosref = bb
            PushStage bstack, False
            For k = LBound(vrs()) To UBound(vrs()) - 1
            If VarType(vrs(k)) = vbString Then
            GlobalVar "EV" + CStr(i + k) + "$", vrs(k)
            bb.DataStr here$ + "." + "EV" + CStr(i + k) + "$"
            Else
            GlobalVar "EV" + CStr(i + k), vrs(k)
            bb.DataStr here$ + "." + "EV" + CStr(i + k)
            End If
            
            
            Next k
            bb.DataObj gui
             
            IsLabelOnly (aString$), f$
            n$ = Mid$(aString$, Len(f$) + 1)
            If Len(n$) > 0 Then
            
            n$ = Left$(n$, Len(n$) - 1)
            If n$ <> "" Then n$ = "Push " + n$ + vbCrLf
         End If
            If F1$ <> "" Then f$ = myUcase(F1$ + "." + f$ + ")", True) Else f$ = myUcase(f$ + ")", True)
  
            If Not GetSub(f$, klm) Then PopStage bstack: bb.Flush: CallEventFromGuiNow = False: GoTo conthere
            '' look for '11001EDIT
            s1$ = sbf(klm).sb
            If Left$(s1$, 10) = "'11001EDIT" Then
            SetNextLine s1$
            End If
            If F1$ <> "" Then s1$ = n$ + "Module " + F1$ + vbCrLf + sbf(klm).sb Else s1$ = n$ + sbf(klm).sb
            
          
            
            If Execute(bstack, s1$, False, False) <> 1 Then
            
            
            
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

If tr Then
'If STEXIT Then trace = tr
trace = tr
End If
extreme = extr
escok = olescok
RaiseEvent Complete(1)
End Function
