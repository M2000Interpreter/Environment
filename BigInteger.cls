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
Friend Property Get Value() As String
    Value = val
End Property
Property Get Reminder()
    Reminder = CVar(cReminder)
End Property

Property Get ToString()
    If outbase = 0 Or outbase = 10 Then
        ToString = CVar(Module13.TrimZero(val))
    Else
        ToString = CVar(Module13.TrimZero(BigIntToString(Value, outbase)))
    End If
End Property
Friend Sub Load(v As String, b As Integer)
    SwapStrings val, v
    outbase = b
End Sub
Sub AnyBaseInput(v As String, Optional ByVal basem = 10)

    Dim b As Integer

    b = CInt(Abs(basem))
    If b < 2 Or b > 36 Then
        MyEr "base out of limit", "η βάση είναι εκτός ορίων"
        Exit Sub
    End If
    val = Module13.BigIntFromString(v, b)
End Sub
Property Let OutputBase(b)
If b > 1 And b < 37 Then
    outbase = b
End If
End Property
Function AnyBaseOuput(Optional b = 10) As String
    If b > 1 And b < 37 Then
    AnyBaseOuput = BigIntToString(Value, CInt(b))
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
    divide.Load Module13.divide(Value, v.Value), outbase
    cReminder = Module13.TrimZero(Module13.LastModulus())
End Function
Function multiply(v As BigInteger) As BigInteger
    Set multiply = New BigInteger
    multiply.Load Module13.multiply(Value, v.Value), outbase
End Function
Function Modulus(v As BigInteger) As BigInteger
    Set Modulus = New BigInteger
    Modulus.Load Module13.Modulus(Value, v.Value), outbase
End Function
Function intpower(v As BigInteger) As BigInteger
    Set intpower = New BigInteger
    intpower.Load Module13.IntStrByExp(Value, v.Value), outbase
End Function
Function modpow(v As BigInteger, m As BigInteger) As BigInteger
    Set modpow = New BigInteger
    modpow.Load Module13.modpow(Value, v.Value, m.Value), outbase
End Function
Function IntSqr() As BigInteger
    Set IntSqr = New BigInteger
    IntSqr.Load Module13.IntSqr(Value), outbase
End Function

Function IsProbablyPrime(k As Integer) As Boolean
    IsProbablyPrime = Module13.IsProbablyPrime(Value, k)
End Function
Function compare(v As BigInteger) As Integer
    compare = Module13.compare(Value, v.Value)
End Function

Private Sub Class_Initialize()
val = "0"
cReminder = "0"
End Sub
