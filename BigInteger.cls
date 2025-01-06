VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BigInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private val As String
Private cReminder As String
Private outbase As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Friend Property Get Value() As String
    Value = val
End Property
Property Get Reminder()
    Reminder = CVar(cReminder)
End Property
Property Get Reminder16()
    Reminder16 = CVar(StrConv(cReminder, vbUnicode))
End Property
Friend Property Get Unique() As Boolean
    Dim obj As IUnknown, GetRefCount As Long
    Set obj = Me
    CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
    If GetRefCount > 4 Then Exit Property
    Unique = True
End Property
Property Get Length() As Long
Length = LenB(val)
End Property
Property Get Pack() As String
If LenB(val) < 15 Then
    Pack = "*[BigInteger] = " + StrConv(val, vbUnicode)
Else
    Pack = "*[BigInteger:" & LenB(val) & "] = " + StrConv(MidB$(val, 1, 6), vbUnicode) + "..." + StrConv(RightB$(val, 6), vbUnicode)
End If
End Property

Public Sub negate()
If LeftB$(val, 1) <> ChrB$(45) Then
val = ChrB$(45) + val
Else
val = MidB$(val, 2)
End If
End Sub
Property Get ToString()
    If outbase = 0 Or outbase = 10 Then
        ToString = CVar(StrConv(Module13.TrimZero(val), vbUnicode))
    Else
        ToString = CVar(StrConv(Module13.TrimZero(BigIntToString(Value, outbase)), vbUnicode))
    End If
End Property
Friend Sub Load(v As String, b As Integer)
    SwapStrings val, v
    outbase = b
End Sub
Friend Sub Load2(A, Optional o, Optional re)
    Dim aa As BigInteger, bb As BigInteger
    If VarType(A) = vbString Then
        SwapString2Variant val, A
        If IsMissing(o) Then
            outbase = 10
        Else
            Set bb = o
            outbase = o.OutputBase
            
        End If
            If Not IsMissing(re) Then
                SwapString2Variant cReminder, re
            End If
    Else
        Set aa = A
        SwapStrings val, aa.Value
        If IsMissing(re) Then
        SwapString2Variant cReminder, aa.Reminder
        Else
        SwapString2Variant cReminder, re
        End If
        If IsMissing(o) Then
            outbase = aa.OutputBase
        Else
            Set bb = o
            outbase = o.OutputBase
    End If
    End If
    
    
End Sub
Sub AnyBaseInput(v As String, Optional ByVal basem = 10)

    Dim b As Integer
    b = CInt(Abs(basem))
    If b < 2 Or b > 36 Then
        MyEr "base out of limit", "η βάση είναι εκτός ορίων"
        Exit Sub
    End If
    val = Module13.BigIntFromString(StrConv(v, vbFromUnicode), b)
End Sub
Property Get OutputBase()
    OutputBase = outbase
End Property
Property Let OutputBase(b)
If b > 1 And b < 37 Then
    outbase = b
End If
End Property
Function AnyBaseOuput(Optional b = 10) As String
    If b > 1 And b < 37 Then
    AnyBaseOuput = StrConv(BigIntToString(Value, CInt(b)), vbUnicode)
    End If
End Function

Function Add(v As BigInteger) As BigInteger
    Set Add = New BigInteger
    Add.Load Module13.Add(Value, v.Value), outbase
End Function
Function subtract(v As BigInteger) As BigInteger
    Set subtract = New BigInteger
    subtract.Load Module13.subtract(Value, v.Value), outbase
End Function
Function divide(v As BigInteger) As BigInteger
    Set divide = New BigInteger
    divide.Load2 Module13.divide(Value, v.Value), CVar(Me), CVar(Module13.LastRemainder)
End Function
Function divideE(v As BigInteger) As BigInteger
    Dim s$, p$, M$, k As Integer
 ' ? Module13.CreateBigInteger("-25").ModulusE(Module13.CreateBigInteger("-7")).ToString
    p$ = v.Value
    If LeftB$(p$, 1) = ChrB$(45) Then MidB$(p$, 1, 1) = ChrB$(48)
    If bSgn() < 0 Then
        If Module13.compare(Value, (p$), 0) = -1 Then ' value<v.value
            s$ = Module13.subtract(Value, Module13.multiply((p$), ChrB$(45) + ChrB$(49)))
        Else
            s$ = Module13.subtract(Value, Module13.multiply((p$), Module13.subtract(Module13.divide(Value, (p$)), ChrB$(49))))
        End If
    Else
        s$ = Module13.subtract(Value, Module13.multiply((p$), Module13.divide(Value, (p$))))
    End If
    If LeftB$(s$, 1) = ChrB$(45) Then MidB$(s$, 1, 1) = ChrB$(48)
    If Module13.compare((s$), p$) > -1 Then s$ = ChrB$(48)
    M$ = s$
    s$ = Module13.divide(Module13.subtract(Value, s$), v.Value)
    Set divideE = New BigInteger
    divideE.Load2 CVar(s$), CVar(Me), CVar(M$)
End Function
Function multiply(v As BigInteger) As BigInteger
    Set multiply = New BigInteger
    multiply.Load Module13.multiply(Value, v.Value), outbase
End Function
Function Modulus(v As BigInteger) As BigInteger
    Set Modulus = New BigInteger
    Modulus.Load Module13.Modulus(Value, v.Value), outbase
End Function
Function ModulusE(v As BigInteger) As BigInteger
    Dim s$, p$, k As Integer
 ' ? Module13.CreateBigInteger("-25").ModulusE(Module13.CreateBigInteger("-7")).ToString
    p$ = v.Value
    If LeftB$(p$, 1) = ChrB$(45) Then MidB$(p$, 1, 1) = ChrB$(48)
    If bSgn() < 0 Then
        If Module13.compare(Value, (p$), 0) = -1 Then ' value<v.value
            s$ = Module13.subtract(Value, Module13.multiply((p$), ChrB$(45) + ChrB$(49)))
        Else
            s$ = Module13.subtract(Value, Module13.multiply((p$), Module13.subtract(Module13.divide(Value, (p$)), ChrB$(49))))
        End If
    Else
        s$ = Module13.subtract(Value, Module13.multiply((p$), Module13.divide(Value, (p$))))
    End If
    If LeftB$(s$, 1) = ChrB$(45) Then MidB$(s$, 1, 1) = ChrB$(48)
    If Module13.compare((s$), p$) > -1 Then s$ = ChrB$(48)
    Set ModulusE = New BigInteger
    ModulusE.Load2 CVar(s$), CVar(Me), CVar(ChrB$(0))
End Function
Function intpower(v As BigInteger) As BigInteger
    Set intpower = New BigInteger
    intpower.Load Module13.IntStrByExp(Value, v.Value), outbase
End Function
Function modpow(v As BigInteger, M As BigInteger) As BigInteger
    Set modpow = New BigInteger
    modpow.Load Module13.modpow(Value, v.Value, M.Value), outbase
End Function
Friend Function modpow1(vv, mm) As BigInteger
    Dim v As BigInteger, M As BigInteger
    Set v = vv
    Set M = mm
    Set modpow1 = New BigInteger
    modpow1.Load Module13.modpow(Value, v.Value, M.Value), outbase
End Function
Function IntSqr() As BigInteger
    Set IntSqr = New BigInteger
    IntSqr.Load Module13.IntSqr(Value), outbase
End Function

Function IsProbablyPrime(k As Integer) As Boolean
    IsProbablyPrime = Module13.IsProbablyPrime(Value, k)
End Function
Function bAbs() As BigInteger
    Set bAbs = New BigInteger
    If LeftB$(val, 1) = ChrB$(45) Then
    bAbs.Load MidB$(val, 2), OutputBase
    Else
    bAbs.Load val, OutputBase
    End If
End Function
Function bSgn() As Integer
   
    If LeftB$(val, 1) = ChrB$(45) Then
        bSgn = -1
    ElseIf val = ChrB$(48) Then
        bSgn = 0
    Else
        bSgn = 1
    End If
End Function
Function compare(v As BigInteger) As Integer
    compare = Module13.compare(Value, v.Value)
End Function

Private Sub Class_Initialize()
val = ChrB$(48)
cReminder = ChrB$(48)
outbase = 10
End Sub
