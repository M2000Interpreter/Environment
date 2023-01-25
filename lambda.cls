VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' lambda functions in M2000
' in running state copy nonlocal variables to current list
Public Name$
Public Code$
Private Type Vars
    Value As Variant
    there As Long
    isvariant As Boolean
End Type
Private Total As Long, fixed As Long, used As Long
Private myVar() As Vars, myHash As New Hash, refcount As Long
Public subs As FastCollection
Public OriginalCode As Long, lastlen As Long
Sub CopyToVar(bstack As basetask, glb As Boolean, v() As Variant)
Dim i As Long, nm$, where As Long, dummy As Variant, x1 As Long
bstack.IamLambda = True
If used > 0 Then
    For i = 0 To used - 1
        myHash.ReadVar i, nm$, where
        If refcount = 0 Then
            If IsObject(myVar(where).Value) Then
                dummy = 0
                myVar(where).there = globalvar(nm$, dummy, , glb, , Not myVar(where).isvariant)
                If Typename(myVar(where).Value) = "lambda" Then
                    If glb Then
                        GlobalSub nm$ + "()", "", , , myVar(where).there
                    Else
                        GlobalSub here$ + "." + bstack.GroupName + nm$ + "()", "", , , myVar(where).there
                    End If
                    Set v(myVar(where).there) = myVar(where).Value
                ElseIf Typename(myVar(where).Value) = "Group" Then
                    
                    Dim aa As Object, g As Group
                    Set g = myVar(where).Value
                    If g.IamApointer Then
                    Set v(myVar(where).there) = g
                    Else
                    Set aa = g
                    UnFloatGroup bstack, nm$, myVar(where).there, aa
                    Set aa = Nothing
                    End If
                    Set g = Nothing
                Else
                    Set v(myVar(where).there) = myVar(where).Value
                End If
            Else
                myVar(where).there = globalvar(nm$, myVar(where).Value, , glb, , Not myVar(where).isvariant)
            End If
        Else
            Dim there As Long
            there = myVar(where).there
            If IsObject(myVar(where).Value) Then
                If Typename(myVar(where).Value) = "lambda" Then
                    globalvar nm$, there, True, glb, , Not myVar(where).isvariant
                    If glb Then
                        GlobalSub nm$ + "()", "", , , there
                    Else
                        GlobalSub here$ + "." + bstack.GroupName + nm$ + "()", "", , , there
                    End If
                ElseIf Typename(myVar(where).Value) = "Group" Then
                    If myVar(where).Value.IamApointer Then
                        globalvar nm$, there, True, glb, , Not myVar(where).isvariant
                    Else
                        LinkGroup bstack, nm$, v(there)
                    End If
                Else
                    globalvar nm$, there, True, glb, , Not myVar(where).isvariant
                End If
            Else
                globalvar nm$, there, True, glb, , Not myVar(where).isvariant
            End If
        End If
    Next i
End If
refcount = refcount + 1
End Sub

Sub CopyFromVar(bstack As basetask, v() As Variant)
Dim i As Long, aa As Group
bstack.IamLambda = False
refcount = refcount - 1
' maybe not needed
Set bstack.lastobj = Nothing
If refcount = 0 Then
    For i = 0 To used - 1
    ' check if is better with MyIsObject
        If myVarType(v(myVar(i).there), vbObject) Then
            If Typename(myVar(i).Value) = "Group" Then
                If Typename(v(myVar(i).there)) = "Group" Then
                    Set myVar(i).Value = CopyGroupObj(v(myVar(i).there))
                Else
                    Set myVar(i).Value = v(myVar(i).there)
                End If
                ' why not  Set v(myVar(i).there) = Nothing
            Else
                Set myVar(i).Value = v(myVar(i).there)
                Set v(myVar(i).there) = Nothing
            End If
        Else
            myVar(i).Value = v(myVar(i).there)
        End If
        myVar(i).there = 0
    Next i
End If

End Sub

Sub FeedNonLocal(aName$, val As Variant, isvariant As Boolean, v() As Variant)
Dim where As Long
If Not myHash.Find(aName$, where) Then
where = GetFree()
myHash.ItemCreator aName$, where ', , , True
End If
If IsObject(val) Then
'' maybe here we can copy
'' check first the type
If TypeOf val Is mHandler Then
If val.indirect > -1 Then
    Set myVar(where).Value = v(val.indirect)
    myVar(where).isvariant = False
    Exit Sub
End If
End If
    Set myVar(where).Value = val
    myVar(where).isvariant = isvariant
Else
    myVar(where).Value = val
    myVar(where).isvariant = isvariant
End If
End Sub

Private Sub Class_Initialize()
Total = 10
ReDim myVar(Total) As Vars

End Sub
Private Function GetFree() As Long
    If used >= Total Then
        Total = Total * 2
        ReDim Preserve myVar(Total) As Vars
    End If
    GetFree = used
    used = used + 1
End Function

Private Sub Class_Terminate()
Erase myVar()
Set subs = Nothing
End Sub
Public Sub CopyTo(a As lambda, v() As Variant)
' only for new
Dim b As lambda, VR As Variant, where As Long, nm$
Set a = New lambda
a.Code$ = Me.Code$
a.OriginalCode = Me.OriginalCode
a.lastlen = Me.lastlen
Dim i As Long
If used > 0 Then
For i = 0 To used - 1
myHash.ReadVar i, nm$, where
If myVar(where).there <> 0 Then
If Typename(v(myVar(where).there)) = "lambda" Then
    With v(myVar(where).there)
    .CopyTo b, v()
    End With
    Set VR = b
    a.FeedNonLocal nm$, VR, myVar(where).isvariant, v()
    Set VR = Nothing
    Set b = Nothing
Else
    a.FeedNonLocal nm$, v(myVar(where).there), myVar(where).isvariant, v()
End If
Else
If Typename(myVar(where).Value) = "lambda" Then
On Error Resume Next
    myVar(where).Value.CopyTo b, v()
    Set VR = b
    a.FeedNonLocal nm$, VR, myVar(where).isvariant, v()
    Set VR = Nothing
    Set b = Nothing
Else
    a.FeedNonLocal nm$, myVar(where).Value, myVar(where).isvariant, v()
End If
End If
Next i
End If
End Sub
