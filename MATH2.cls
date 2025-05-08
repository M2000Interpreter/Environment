VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MATH2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' FROM THE TRICK VBFORUMS
Private Declare Function GetMem2 Lib "msvbvm60" (pSrc As Any, pDst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (pSrc As Any, pDst As Any) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (arr() As Any) As Long
'Public Type cxTest
'    v   As Double
'    S   As String
'    K() As String
'End Type

Public Type Complex
    r   As Double
    i   As Double
End Type

Public Type Matrix
    Col As Long                 ' Number of columns
    Row As Long                 ' Number of rows
    D() As Double
End Type

Private Const Pi = 3.14159265358979
Private Const e = 2.71828182845905

Private Const PI2 = Pi / 2

'Public Function TstCreate(v As Double, S As String, mSize As Long) As cxTest
'    TstCreate.v = v
'    TstCreate.S = S
'    ReDim TstCreate.K(mSize - 1)
'End Function

'+=====================================================================================================================================+
'|                                                            Real numbers                                                             |
'+=====================================================================================================================================+

' // From degree to radians
Public Function Deg(ByVal Value As Double) As Double
    Deg = 1.74532925199433E-02 * Value
End Function

' // The logarithm to the base of a real number X
Public Function LogX(ByVal Value As Double, ByVal Base As Double) As Double
    LogX = Log(Value) / Log(Base)
End Function

' // The decimal logarithm of a real number
Public Function Log10(ByVal Value As Double) As Double
    Log10 = Log(Value) / 2.30258509299405
End Function

' // The binary logarithm of a real number
Public Function Log2(ByVal Value As Double) As Double
    Log2 = Log(Value) / 0.693147180559945
End Function

' // Rounding up
Public Function Ceil(ByVal Value As Double) As Double
    Ceil = -Int(-Value)
End Function

' // Rounding down (Int)
Public Function Floor(ByVal Value As Double) As Double
    Floor = Int(Value)
End Function

' // Secant of a real number
Public Function Sec(ByVal Value As Double) As Double
    Sec = 1 / Cos(Value)
End Function

' // Cosecant of a real number
Public Function Csc(ByVal Value As Double) As Double
    Csc = 1 / Sin(Value)
End Function

' // Cotangent of a real number
Public Function Ctg(ByVal Value As Double) As Double
    Ctg = 1 / Tan(Value)
End Function

' // Arcsine of a real number
Public Function Asin(ByVal Value As Double) As Double
    If Value = -1 Then Asin = -PI2: Exit Function
    If Value = 1 Then Asin = PI2: Exit Function
    Asin = Atn(Value / Sqr(-Value * Value + 1))
End Function

' // Arccosine of a real number
Public Function Acos(ByVal Value As Double) As Double
    If CSng(Value) = -1# Then Acos = Pi: Exit Function
    If CSng(Value) = 1# Then Acos = 0: Exit Function
    Acos = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
End Function

' // Arcsecant of a real number
Public Function Asec(ByVal Value As Double) As Double
    Asec = 1.5707963267949 - Atn(Sgn(Value) / Sqr(Value * Value - 1))
End Function

' // Arccosecant of a real number
Public Function Acsc(ByVal Value As Double) As Double
    Acsc = Atn(Sgn(Value) / Sqr(Value * Value - 1))
End Function

' // Returns the angle whose tangent is the ratio of the two numbers
Public Function Atan2(ByVal Y As Double, ByVal X As Double) As Double
    If Y > 0 Then
        If X >= Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= -Y Then
            Atan2 = Atn(Y / X) + Pi
        Else
            Atan2 = Pi / 2 - Atn(X / Y)
        End If
    Else
        If X >= -Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= Y Then
            Atan2 = Atn(Y / X) - Pi
        Else
            Atan2 = -Atn(X / Y) - Pi / 2
        End If
    End If
End Function

' // Arccotangent of a real number
Public Function Actg(ByVal Value As Double) As Double
    Actg = 1.5707963267949 - Atn(Value)
End Function

' // Hyperbolic sine of a real number
Public Function Sinh(ByVal Value As Double) As Double
    Sinh = (EXP(Value) - EXP(-Value)) / 2
End Function

' // Hyperbolic cosine of a real number
Public Function Cosh(ByVal Value As Double) As Double
    Cosh = (EXP(Value) + EXP(-Value)) / 2
End Function

' // Hyperbolic tangent of a real number
Public Function Tanh(ByVal Value As Double) As Double
    Tanh = (EXP(2 * Value) - 1) / (EXP(2 * Value) + 1)
End Function

' // Hyperbolic cotangent of a real number
Public Function Ctgh(ByVal Value As Double) As Double
    Ctgh = 1 / (EXP(2 * Value) + 1) / (EXP(2 * Value) - 1)
End Function

' // Hyperbolic secant of a real number
Public Function Sech(ByVal Value As Double) As Double
    Sech = 2 / (EXP(Value) + EXP(-Value))
End Function

' // Hyperbolic cosecant of a real number
Public Function Csch(ByVal Value As Double) As Double
    Csch = 2 / (EXP(Value) - EXP(-Value))
End Function

' // Hyperbolic arcsine of a real number
Public Function Asinh(ByVal Value As Double) As Double
    Asinh = Log(Value + Sqr(Value * Value + 1))
End Function

' // Hyperbolic arcosine of a real number
Public Function Acosh(ByVal Value As Double) As Double
    Acosh = Log(Value + Sqr(Value * Value - 1))
End Function

' // Hyperbolic arctangent of a real number
Public Function Atanh(ByVal Value As Double) As Double
    Atanh = Log((1 + Value) / (1 - Value)) / 2
End Function

' // Hyperbolic arccotangent of a real number
Public Function Actan(ByVal Value As Double) As Double
    Actan = Log((Value + 1) / (Value - 1)) / 2
End Function

' // Hyperbolic arcsecant of a real number
Public Function Asech(ByVal Value As Double) As Double
    Asech = Log((Sqr(-Value * Value + 1) + 1) / Value)
End Function

' // Hyperbolic arccosecant of a real number
Public Function Acsch(ByVal Value As Double) As Double
    Acsch = Log((Sgn(Value) * Sqr(Value * Value + 1) + 1) / Value)
End Function

' // Return maximum of two numbers
Public Function Max(ByVal Op1 As Double, ByVal Op2 As Double) As Double
    Max = IIf(Op1 > Op2, Op1, Op2)
End Function

' // Return maximum of three numbers
Public Function Max3(ByVal Op1 As Double, ByVal Op2 As Double, ByVal Op3 As Double) As Double
    Max3 = IIf(Op1 > Op2, IIf(Op1 > Op3, Op1, Op3), IIf(Op2 > Op3, Op2, Op3))
End Function

' // Return maximum of four numbers
Public Function Max4(ByVal Op1 As Double, ByVal Op2 As Double, ByVal Op3 As Double, ByVal Op4 As Double) As Double
    Max4 = Max(Max3(Op1, Op2, Op3), Op4)
End Function

' // Return minimum of two numbers
Public Function Min(ByVal Op1 As Double, ByVal Op2 As Double) As Double
    Min = IIf(Op1 < Op2, Op1, Op2)
End Function

' // Return minimum of three numbers
Public Function Min3(ByVal Op1 As Double, ByVal Op2 As Double, ByVal Op3 As Double) As Double
    Min3 = IIf(Op1 < Op2, IIf(Op1 < Op3, Op1, Op3), IIf(Op2 < Op3, Op2, Op3))
End Function

' // Return minimum of four numbers
Public Function Min4(ByVal Op1 As Double, ByVal Op2 As Double, ByVal Op3 As Double, ByVal Op4 As Double) As Double
    Min4 = Min(Min3(Op1, Op2, Op3), Op4)
End Function

' // Returns the remainder of dividing one specified number by another specified number.
Public Function IEEERemainder(ByVal Op1 As Double, ByVal Op2 As Double) As Double
    IEEERemainder = Op1 - (Op2 * Round(Op1 / Op2))
End Function

' // Returns the remainder of dividing one specified number by another specified number.
Public Function rMod(ByVal Op1 As Double, ByVal Op2 As Double) As Double
    rMod = (Abs(Op1) - (Abs(Op2) * (Int(Abs(Op1) / Abs(Op2))))) * Sgn(Op1)
End Function
'+=====================================================================================================================================+
'|                                                          Complex numbers                                                            |
'+=====================================================================================================================================+

' // R = 1, I = 0
Public Function cxOne() As Complex
    cxOne.r = 1
End Function

' // R = 0, I = 1
Public Function cxImgOne() As Complex
    cxOne.i = 1
End Function

' // R = 0, I = 0
Public Function cxZero() As Complex
End Function

' // Creating a new complex number
Public Function cxNew(ByVal Real As Double, ByVal Imaginary As Double) As Complex
    cxNew.r = Real: cxNew.i = Imaginary
End Function

' // Creating a new complex number by polar coordinates
Public Function cxPolar(ByVal Magnitude As Double, ByVal Phase As Double) As Complex
    cxPolar.r = Magnitude * Cos(Phase): cxPolar.i = Magnitude * Sin(Phase)
End Function

' // Return the additive inverse of a specified complex number
Public Function cxNeg(Op As Complex) As Complex
    cxNeg.r = -Op.r: cxNeg.i = -Op.i
End Function

' // Return the inverse value of a complex number
Public Function cxInv(Op As Complex) As Complex
    Dim Ab2 As Double
    Ab2 = Op.r * Op.r + Op.i * Op.i
    cxInv.r = Op.r / Ab2: cxInv.i = -Op.i / Ab2
End Function

' // Addition of two complex numbers
Public Function cxAddVar(Op1, Op2) As Complex
    cxAddVar.r = Op1.r + Op2.r
    cxAddVar.i = Op1.i + Op2.i
End Function

Public Function cxAddComVar(Op1 As Complex, Op2) As Complex
    cxAddComVar.r = Op1.r + Op2.r
    cxAddComVar.i = Op1.i + Op2.i
End Function
Public Function cxAdd(Op1 As Complex, Op2 As Complex) As Complex
    cxAdd.r = Op1.r + Op2.r
    cxAdd.i = Op1.i + Op2.i
End Function

' // Subtraction of two complex numbers
Public Function cxSub(Op1 As Complex, Op2 As Complex) As Complex
    cxSub.r = Op1.r - Op2.r
    cxSub.i = Op1.i - Op2.i
End Function

' // Multiplication of two complex numbers
Public Function cxMul(Op1 As Complex, Op2 As Complex) As Complex
    cxMul.r = Op1.r * Op2.r - Op1.i * Op2.i
    cxMul.i = Op1.r * Op2.i + Op1.i * Op2.r
End Function
Public Function cxMulVar(Op1, Op2) As Complex
    cxMulVar.r = Op1.r * Op2.r - Op1.i * Op2.i
    cxMulVar.i = Op1.r * Op2.i + Op1.i * Op2.r
End Function
' // Division of two complex numbers
Public Function cxDiv(Op1 As Complex, Op2 As Complex) As Complex
    Dim R2 As Double, i2 As Double
    On Error Resume Next
    R2 = Op2.r * Op2.r: i2 = Op2.i * Op2.i
    cxDiv.r = (Op1.r * Op2.r + Op1.i * Op2.i) / (R2 + i2)
    cxDiv.i = (Op1.i * Op2.r - Op1.r * Op2.i) / (R2 + i2)
    If Err.Number > 0 Then Err.Clear: DevZero
End Function
Public Function cxDivVar(Op1, Op2) As Complex
    Dim R2 As Double, i2 As Double
    On Error Resume Next
    R2 = Op2.r * Op2.r: i2 = Op2.i * Op2.i
    cxDivVar.r = (Op1.r * Op2.r + Op1.i * Op2.i) / (R2 + i2)
    cxDivVar.i = (Op1.i * Op2.r - Op1.r * Op2.i) / (R2 + i2)
    If Err.Number > 0 Then Err.Clear: DevZero
End Function
' // Exponentiation of a complex number
Public Function cxDgr(Op As Complex, ByVal Degree As Long) As Complex
    Dim Md As Double, ar As Double
    Md = cxMod(Op): ar = cxArg(Op): Md = Md ^ Degree: ar = ar * Degree
    cxDgr.r = Md * Cos(ar): cxDgr.i = Md * Sin(ar)
End Function

' // The square root of a complex number
Public Function cxSqr(Op As Complex) As Complex
On Error Resume Next
    Dim M As Double, a As Double
    M = Sqr(cxMod(Op)): a = cxArg(Op) / 2
    cxSqr.r = M * Cos(a): cxSqr.i = M * Sin(a)
End Function

' // Module of a complex number
Public Function cxMod(Op As Complex) As Double
    Dim R2 As Double, i2 As Double
    On Error GoTo 100
    R2 = Op.r * Op.r: i2 = Op.i * Op.i
    cxMod = Sqr(R2 + i2)
    Exit Function
100 Overflow
    Err.Clear
End Function
Public Function cxModVar(Op) As Double
    Dim R2 As Double, i2 As Double
    On Error GoTo 100
    R2 = Op.r * Op.r: i2 = Op.i * Op.i
    cxModVar = Sqr(R2 + i2)
    Exit Function
100 Overflow
    Err.Clear
End Function
' // Phase of a complex number
Public Function cxPhase(Op As Complex) As Double
    cxPhase = Atan2(Op.i, Op.r)
End Function

' // Argument of a complex number (equal phase)
Public Function cxArg(Op As Complex) As Double
    If Op.i = 0 Then
        If Op.r >= 0 Then cxArg = 0 Else cxArg = Pi
    ElseIf Op.r = 0 Then
        If Op.i >= 0 Then cxArg = PI2 Else cxArg = -PI2
    Else
        If Op.r > 0 Then
            cxArg = Atn(Op.i / Op.r)
        ElseIf Op.r < 0 And Op.i > 0 Then
            cxArg = Pi + Atn(Op.i / Op.r)
        ElseIf Op.r < 0 And Op.i < 0 Then
            cxArg = -Pi + Atn(Op.i / Op.r)
        End If
    End If
End Function
Public Function cxArgVar(Op) As Double
    If Op.i = 0 Then
        If Op.r >= 0 Then cxArgVar = 0 Else cxArgVar = Pi
    ElseIf Op.r = 0 Then
        If Op.i >= 0 Then cxArgVar = PI2 Else cxArgVar = -PI2
    Else
        If Op.r > 0 Then
            cxArgVar = Atn(Op.i / Op.r)
        ElseIf Op.r < 0 And Op.i > 0 Then
            cxArgVar = Pi + Atn(Op.i / Op.r)
        ElseIf Op.r < 0 And Op.i < 0 Then
            cxArgVar = -Pi + Atn(Op.i / Op.r)
        End If
    End If
End Function
Public Function cxPow(Op1 As Complex, Op2 As Complex, Optional NumDigitsAfterDecimal = 20) As Complex
If Op1.i = 0 And Op2.i = 0 Then
    cxPow = cxNew(Op1.r ^ Op2.r, 0)
Else
    cxPow = cxExp(cxMul(cxlog(Op1), Op2))
    End If
    If NumDigitsAfterDecimal < 20 Then
    cxPow.r = Round(cxPow.r, NumDigitsAfterDecimal)
   cxPow.i = Round(cxPow.i, NumDigitsAfterDecimal)
   End If
End Function
Public Function cxPowVar(Op1, Op2) As Complex
    If Op1.i = 0 And Op2.i = 0 Then
        cxPowVar = cxNew(Op1.r ^ Op2.r, 0)
    Else
        cxPowVar = cxExp(cxMulVar(cxlogVar(Op1), Op2))
    End If

End Function
Public Function cxPowRealVar(Op1, ByVal Op2 As Double) As Complex
On Error GoTo 100
    If Op1.i = 0 Then
        cxPowRealVar = cxNew(Op1.r ^ Op2, 0)
    Else
        cxPowRealVar = cxExp(cxMulRealVar(cxlogVar(Op1), Op2))
    End If
100
End Function
Public Function cxPowReal(Op1 As Complex, Op2 As Double, Optional NumDigitsAfterDecimal = 20) As Complex
On Error GoTo 100
If Op1.i = 0 Then
    cxPowReal = cxNew(Op1.r ^ Op2, 0)
Else
    cxPowReal = cxExp(cxMulReal(cxlog(Op1), Op2))
    End If
    If NumDigitsAfterDecimal < 20 Then
    cxPowReal.r = Round(cxPowReal.r, NumDigitsAfterDecimal)
    cxPowReal.i = Round(cxPowReal.i, NumDigitsAfterDecimal)
    End If
100
End Function
Public Function cxPowRealComp(Op1 As Double, Op2 As Complex, Optional NumDigitsAfterDecimal = 20) As Complex
    If Round(Op2.i, NumDigitsAfterDecimal) = 0 Then
        cxPowRealComp = cxNew(Op1 ^ Op2.r, 0)
    Else
        cxPowRealComp = cxExp(cxMulReal(Op2, Log(Op1)))
    End If
   If NumDigitsAfterDecimal < 20 Then
   cxPowRealComp.r = Round(cxPowRealComp.r, NumDigitsAfterDecimal)
   cxPowRealComp.i = Round(cxPowRealComp.i, NumDigitsAfterDecimal)
   End If
End Function
Public Function cxPowRealCompVar(ByVal Op1 As Double, Op2) As Complex
    If Round(Op2.i, NumDigitsAfterDecimal) = 0 Then
        cxPowRealCompVar = cxNew(Op1 ^ Op2.r, 0)
    Else
        cxPowRealCompVar = cxExp(cxMulRealVar(Op2, Log(Op1)))
    End If
End Function
' // Returns the number e, raised to power by complex number
Public Function cxExp(Op As Complex) As Complex
    On Error Resume Next
    cxExp.r = EXP(Op.r) * Cos(Op.i): cxExp.i = EXP(Op.r) * Sin(Op.i)
    If Err.Number = 6 Then Overflow
End Function
' // Addition real number and complex number
Public Function cxAddRealVar(Op1, ByVal Op2 As Double) As Complex
    cxAddRealVar.r = Op1.r + Op2
    cxAddRealVar.i = Op1.i
End Function
Public Function cxAddReal(Op1 As Complex, ByVal Op2 As Double) As Complex
    cxAddReal.r = Op1.r + Op2
    cxAddReal.i = Op1.i
End Function

' // Subtraction from complex number a real number
Public Function cxSubReal(Op1 As Complex, ByVal Op2 As Double) As Complex
    cxSubReal.r = Op1.r - Op2
    cxSubReal.i = Op1.i
End Function

' // Subtraction from real number a complex number
Public Function cxRealSub(ByVal Op1 As Double, Op2 As Complex) As Complex
    cxRealSub.r = Op1 - Op2.r
    cxRealSub.i = -Op2.i
End Function

' // Multiplication complex number on a real number
Public Function cxMulReal(Op1 As Complex, ByVal Op2 As Double) As Complex
    cxMulReal.r = Op1.r * Op2
    cxMulReal.i = Op1.i * Op2
End Function
Public Function cxMulRealVar(Op1, ByVal Op2 As Double) As Complex
    cxMulRealVar.r = Op1.r * Op2
    cxMulRealVar.i = Op1.i * Op2
End Function
' // Division a complex number on a real number
Public Function cxDivReal(Op1 As Complex, ByVal Op2 As Double) As Complex
    Dim R2 As Double
    R2 = Op2 * Op2
    cxDivReal.r = (Op1.r * Op2) / R2
    cxDivReal.i = (Op1.i * Op2) / R2
End Function
Public Function cxDivRealVar(Op1, ByVal Op2 As Double) As Complex
    Dim R2 As Double
    R2 = Op2 * Op2
    cxDivRealVar.r = (Op1.r * Op2) / R2
    cxDivRealVar.i = (Op1.i * Op2) / R2
End Function

' // Division a real number on a complex number
Public Function cxRealDiv(ByVal Op1 As Double, Op2 As Complex) As Complex
    Dim R2 As Double, i2 As Double
    R2 = Op2.r * Op2.r: i2 = Op2.i * Op2.i
    cxRealDiv.r = (Op1 * Op2.r) / (R2 + i2)
    cxRealDiv.i = (-Op1 * Op2.i) / (R2 + i2)
End Function
Public Function cxRealDivVar(ByVal Op1 As Double, Op2) As Complex
    Dim R2 As Double, i2 As Double
    R2 = Op2.r * Op2.r: i2 = Op2.i * Op2.i
    cxRealDivVar.r = (Op1 * Op2.r) / (R2 + i2)
    cxRealDivVar.i = (-Op1 * Op2.i) / (R2 + i2)
End Function
' // Addition of a complex number and imaginary part
Public Function cxAddImg(Op1 As Complex, ByVal Op2 As Double) As Complex
    cxAddImg.r = Op1.r
    cxAddImg.i = Op1.i + Op2
End Function

' // Subtraction from a complex number a imaginary part
Public Function cxSubImg(Op1 As Complex, Op2 As Double) As Complex
    cxSubImg.r = Op1.r
    cxSubImg.i = Op1.i - Op2
End Function

' // Subtraction from imaginary part a complex number
Public Function cxImgSub(ByVal Op1 As Double, Op2 As Complex) As Complex
    cxImgSub.r = -Op2.r
    cxImgSub.i = Op1 - Op2.i
End Function

' // Multiplication complex number on a imaginary part
Public Function cxMulImg(Op1 As Complex, ByVal Op2 As Double) As Complex
    cxMulImg.r = -Op1.i * Op2
    cxMulImg.i = Op1.r * Op2
End Function

' // Division a complex number on a imaginary part
Public Function cxDivImg(Op1 As Complex, ByVal Op2 As Double) As Complex
    Dim i2 As Double
    i2 = Op2 * Op2
    cxDivImg.r = (Op1.i * Op2) / i2
    cxDivImg.i = (-Op1.r * Op2) / i2
End Function

' // Division imaginary part on a complex number
Public Function cxImgDiv(ByVal Op1 As Double, Op2 As Complex) As Complex
    Dim R2 As Double, i2 As Double
    R2 = Op2.r * Op2.r: i2 = Op2.i * Op2.i
    cxImgDiv.r = (Op1 * Op2.i) / (R2 + i2)
    cxImgDiv.i = (Op1 * Op2.r) / (R2 + i2)
End Function

' // Return true if complex number is equal
Public Function cxEq(Op1 As Complex, Op2 As Complex, _
                Optional NumDigitsAfterDecimal As Long = -1) As Boolean
        Dim z As Double
    If NumDigitsAfterDecimal = -1 Then
        z = Op1.r + Op2.r
        If z > 10000000000# Then
            If Abs((Op1.r - Op2.r) / (Op1.r + Op2.r)) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEq = True
        Else
        If Abs(Op1.r - Op2.r) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEq = True
        End If
    Else
        z = Op1.r + Op2.r
        If z > 10000000000# Then
            If Abs((Op1.r - Op2.r) / (Op1.r + Op2.r)) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEq = True
        Else
        If Round(Op1.r, NumDigitsAfterDecimal) = Round(Op2.r, NumDigitsAfterDecimal) And _
           Round(Op1.i, NumDigitsAfterDecimal) = Round(Op2.i, NumDigitsAfterDecimal) Then cxEq = True
        End If
    End If
End Function
Public Function cxEqVar(Op1, Op2) As Boolean
        Dim z As Double
        z = Op1.r + Op2.r
        If z > 10000000000# Then
            If Abs((Op1.r - Op2.r) / (Op1.r + Op2.r)) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEqVar = True
        Else
            If Abs(Op1.r - Op2.r) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEqVar = True
        End If
End Function
Public Function cxEqVar3(Op1, Op2, ByVal NumDigitsAfterDecimal As Long) As Boolean
        Dim z As Double
        z = Op1.r + Op2.r
        z = Op1.r + Op2.r
        If z > 10000000000# Then
            If Abs((Op1.r - Op2.r) / (Op1.r + Op2.r)) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEqVar3 = True
        Else
        If Round(Op1.r, NumDigitsAfterDecimal) = Round(Op2.r, NumDigitsAfterDecimal) And _
           Round(Op1.i, NumDigitsAfterDecimal) = Round(Op2.i, NumDigitsAfterDecimal) Then cxEqVar3 = True
        End If
End Function
Public Function cxEqComVar3(Op1 As Complex, Op2, ByVal NumDigitsAfterDecimal As Long) As Boolean
        Dim z As Double
        z = Op1.r + Op2.r
        z = Op1.r + Op2.r
        If z > 10000000000# Then
            If Abs((Op1.r - Op2.r) / (Op1.r + Op2.r)) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEqComVar3 = True
        Else
        If Round(Op1.r, NumDigitsAfterDecimal) = Round(Op2.r, NumDigitsAfterDecimal) And _
           Round(Op1.i, NumDigitsAfterDecimal) = Round(Op2.i, NumDigitsAfterDecimal) Then cxEqComVar3 = True
        End If
End Function
Public Function cxEqComVar(Op1 As Complex, Op2) As Boolean
        Dim z As Double
        On Error GoTo 1000
        z = Op1.r + Op2.r
        If z > 10000000000# Then
            If Abs((Op1.r - Op2.r) / (Op1.r + Op2.r)) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEqComVar = True
        Else
            If Abs(Op1.r - Op2.r) < 0.00000000000001 And Abs(Op1.i - Op2.i) < 0.00000000000001 Then cxEqComVar = True
        End If
        Exit Function
1000         Overflow
End Function
' // Return absolute value of a complex number
' // correction: https://learn.microsoft.com/en-us/dotnet/api/system.numerics.complex.abs?view=net-9.0
Public Function cxAbs(Op As Complex) As Double
    On Error GoTo 1000
    If Op.i = 0 Then
        cxAbs = Op.r
    ElseIf Op.r > Op.i Then
        cxAbs = Op.r * Sqr(1 + (Op.i * Op.i) / (Op.r * Op.r))
    ElseIf Op.r <= Op.i Then
        cxAbs = Op.i * Sqr(1 + (Op.r * Op.r) / (Op.i * Op.i))
    End If
    Exit Function
1000
    cxAbs = Infinity
End Function

' // Return complex conjugate of complex number
Public Function cxConj(Op As Complex) As Complex
    cxConj.r = Op.r
    cxConj.i = -Op.i
End Function

' // The natural logarithm of a complex number
Public Function cxlog(Op As Complex) As Complex
    Dim M As Double, a As Double
    M = cxMod(Op): a = cxArg(Op)
    cxlog.r = Log(M): cxlog.i = a
End Function
Public Function cxlogVar(Op) As Complex
    Dim M As Double, a As Double
    M = cxModVar(Op): a = cxArgVar(Op)
    cxlogVar.r = Log(M): cxlogVar.i = a
End Function
' // The logarithm of a complex number by base X
Public Function cxLogX(Op As Complex, ByVal Base As Double) As Complex
    Dim M As Double, a As Double, Nc As Complex
    M = cxMod(Op): a = cxArg(Op): Nc.r = Log(Base)
    cxLogX.r = Log(M): cxLogX.i = a
    cxLogX = cxDiv(cxLogX, Nc)
End Function

' // Sine of a complex number
Public Function cxSin(Op As Complex) As Complex
    cxSin.r = Sin(Op.r) * Cosh(Op.i): cxSin.i = Cos(Op.r) * Sinh(Op.i)
End Function
' // Cosine of a complex number
Public Function cxCos(Op As Complex) As Complex
    cxCos.r = Cos(Op.r) * Cosh(Op.i): cxCos.i = -Sin(Op.r) * Sinh(Op.i)
End Function
' // Tangent of a complex number
Public Function cxTan(Op As Complex) As Complex
    Dim c2 As Double, S2 As Double
    c2 = Cos(Op.r): c2 = c2 * c2: S2 = Sinh(Op.i): S2 = S2 * S2
    cxTan.r = (Sin(Op.r) * Cos(Op.r)) / (c2 + S2)
    cxTan.i = (Sinh(Op.i) * Cosh(Op.i)) / (c2 + S2)
End Function
' // Cotangent of a complex number
Public Function cxCtg(Op As Complex) As Complex
    Dim c2 As Double, S2 As Double
    c2 = Sin(Op.r): c2 = c2 * c2: S2 = Sinh(Op.i): S2 = S2 * S2
    cxCtg.r = (Sin(Op.r) * Cos(Op.r)) / (c2 + S2)
    cxCtg.i = -(Sinh(Op.i) * Cosh(Op.i)) / (c2 + S2)
End Function

' // Secant of a complex number
Public Function cxSec(Op As Complex) As Complex
    Dim c2 As Double, S2 As Double
    c2 = Cos(Op.r): c2 = c2 * c2: S2 = Sinh(Op.i): S2 = S2 * S2
    cxSec.r = (Cos(Op.r) * Cosh(Op.i)) / (c2 + S2)
    cxSec.i = -(Sin(Op.r) * Sinh(Op.i)) / (c2 + S2)
End Function

' // Cosecant of a complex number
Public Function cxCsc(Op As Complex) As Complex
    Dim c2 As Double, S2 As Double
    c2 = Sin(Op.r): c2 = c2 * c2: S2 = Sinh(Op.i): S2 = S2 * S2
    cxCsc.r = (Sin(Op.r) * Cosh(Op.i)) / (c2 + S2)
    cxCsc.i = (Cos(Op.r) * Sinh(Op.i)) / (c2 + S2)
End Function

' // Arcsine of a complex number
Public Function cxAsin(Op As Complex) As Complex
    cxAsin = cxMulImg(cxlog(cxAdd(cxMulImg(Op, 1), cxSqr(cxRealSub(1, cxMul(Op, Op))))), -1)
End Function

' // Arccosine of a complex number
Public Function cxAcos(Op As Complex) As Complex
    cxAcos = cxAddReal(cxMulImg(cxlog(cxAdd(cxMulImg(Op, 1), cxSqr(cxRealSub(1, cxMul(Op, Op))))), 1), PI2)
End Function

' // Arctangent of a complex number
Public Function cxAtan(Op As Complex) As Complex
    Dim Iz As Complex
    Iz = cxMulImg(Op, 1)
    cxAtan = cxMulImg(cxSub(cxlog(cxRealSub(1, Iz)), cxlog(cxAddReal(Iz, 1))), 0.5)
End Function

' // Arccotangent of a complex number
Public Function cxActg(Op As Complex) As Complex
    cxActg = cxMulImg(cxSub(cxlog(cxDiv(cxSubImg(Op, 1), Op)), cxlog(cxDiv(cxAddImg(Op, 1), Op))), 0.5)
End Function

' // Arcsecant of a complex number
Public Function cxAsec(Op As Complex) As Complex
    cxAsec = cxAcos(cxDgr(Op, -1))
End Function

' // Arccosecant of a complex number
Public Function cxAcsc(Op As Complex) As Complex
    cxAcsc = cxAsin(cxDgr(Op, -1))
End Function

' // Hyperbolic sine of a complex number
Public Function cxSinh(Op As Complex) As Complex
    cxSinh = cxMulImg(cxSin(cxMulImg(Op, 1)), -1)
End Function

' // Hyperbolic cosine of a complex number
Public Function cxCosh(Op As Complex) As Complex
    cxCosh = cxCos(cxMulImg(Op, 1))
End Function

' // Hyperbolic tangent of a complex number
Public Function cxTanh(Op As Complex) As Complex
    cxTanh = cxMulImg(cxTan(cxMulImg(Op, 1)), -1)
End Function

' // Hyperbolic cotangent of a complex number
Public Function cxCtgh(Op As Complex) As Complex
    cxCtgh = cxRealDiv(1, cxTanh(Op))
End Function

' // Hyperbolic secant of a complex number
Public Function cxSech(Op As Complex) As Complex
    cxSech = cxRealDiv(1, cxCosh(Op))
End Function

' // Hyperbolic cosecant of a complex number
Public Function cxCsch(Op As Complex) As Complex
    cxCsch = cxRealDiv(1, cxSinh(Op))
End Function

' // Hyperbolic arcsine of a complex number
Public Function cxAsinh(Op As Complex) As Complex
    cxAsinh = cxlog(cxAdd(Op, cxSqr(cxAddReal(cxMul(Op, Op), 1))))
End Function

' // Hyperbolic arccosine of a complex number
Public Function cxAcosh(Op As Complex) As Complex
    cxAcosh = cxlog(cxAdd(Op, cxMul(cxSqr(cxAddReal(Op, 1)), cxSqr(cxSubReal(Op, 1)))))
End Function

' // Hyperbolic arctangent of a complex number
Public Function cxAtanh(Op As Complex) As Complex
    cxAtanh = cxMulReal(cxlog(cxDiv(cxAddReal(Op, 1), cxRealSub(1, Op))), 0.5)
End Function

' // Hyperbolic arccotangent of a complex number
Public Function cxActgh(Op As Complex) As Complex
    cxActgh = cxMulReal(cxlog(cxDiv(cxAddReal(Op, 1), cxSubReal(Op, 1))), 0.5)
End Function

' // Hyperbolic arcsecant of a complex number
Public Function cxAsech(Op As Complex) As Complex
    Dim z As Complex
    z = cxRealDiv(1, Op)
    cxAsech = cxlog(cxAdd(z, cxSqr(cxAddReal(cxMul(z, z), 1))))
End Function

' // Hyperbolic arccosecant of a complex number
Public Function cxAcsch(Op As Complex) As Complex
    Dim z As Complex
    z = cxRealDiv(1, Op)
    cxAcsch = cxlog(cxAdd(z, cxMul(cxSqr(cxAddReal(z, 1)), cxSqr(cxSubReal(z, 1)))))
End Function
' // Print matrix to immediate window
Public Function PrintMtrx(Op As Matrix)
    Dim Ts As String, i As Long, j As Long
    Debug.Print vbNewLine
    Debug.Print "Col= " & Op.Col & " ; Row= " & Op.Row
    For i = 0 To Op.Row - 1: For j = 0 To Op.Col - 1
        Ts = space(10)
        LSet Ts = str(Round(Op.D(i * Op.Col + j), 3))
        Debug.Print Ts;
    Next: Debug.Print vbNewLine;: Next
End Function

' // Creating a matrix
Public Function mxCreate(ByVal Row As Long, ByVal Col As Long, ParamArray Y()) As Matrix
    Dim p As Variant, c As Long
    If Row <= 0 Or Col <= 0 Then Exit Function
    If Row * Col < UBound(Y) + 1 Then Exit Function
    mxCreate.Row = Row: mxCreate.Col = Col
    ReDim mxCreate.D(Row * Col - 1): c = 0
    For Each p In Y
        mxCreate.D(c) = p: c = c + 1
    Next
End Function

' // Creating the null-matrix
Public Function mxNull(ByVal Row As Long, ByVal Col As Long) As Matrix
    If Row <= 0 Or Col <= 0 Then Exit Function
    ReDim mxNull.D(Row * Col - 1): mxNull.Col = Col: mxNull.Row = Row
End Function

' // Creating the identity matrix
Public Function mxIdt(ByVal Size As Long) As Matrix
    Dim ij As Long
    If Size <= 0 Then Exit Function
    ReDim mxIdt.D(Size * Size - 1): mxIdt.Col = Size: mxIdt.Row = Size
    For ij = 0 To Size - 1: mxIdt.D(ij + (ij * Size)) = 1: Next
End Function

' // Transpose matrix
Public Function mxTrans(Op As Matrix) As Matrix
    Dim i As Long, j As Long, p As Long
    GetMem4 ByVal ArrPtr(Op.D), p: If p = 0 Then Exit Function
    mxTrans.Row = Op.Col: mxTrans.Col = Op.Row: ReDim mxTrans.D(UBound(Op.D))
    For j = 0 To mxTrans.Col - 1: For i = 0 To mxTrans.Row - 1
        mxTrans.D(i + j * mxTrans.Row) = Op.D(j + i * Op.Row)
    Next: Next
End Function

' // Multiplication matrix on a real number
Public Function mxMulReal(Op As Matrix, Op2 As Double) As Matrix
    Dim p As Long, ij As Long
    GetMem4 ByVal ArrPtr(Op.D), p: If p = 0 Then Exit Function
    ReDim mxMulReal.D(UBound(Op.D)): mxMulReal.Col = Op.Col: mxMulReal.Row = Op.Row
    For ij = 0 To UBound(Op.D): mxMulReal.D(ij) = Op.D(ij) * Op2: Next
End Function

' // Addition of a two matrix
Public Function mxAdd(Op1 As Matrix, Op2 As Matrix) As Matrix
    Dim p As Long, ij As Long
    GetMem4 ByVal ArrPtr(Op1.D), p: If p = 0 Then Exit Function
    GetMem4 ByVal ArrPtr(Op2.D), p: If p = 0 Then Exit Function
    If Op1.Col <> Op2.Col Or Op1.Row <> Op2.Row Then Exit Function
    ReDim mxAdd.D(UBound(Op1.D)): mxAdd.Col = Op1.Col: mxAdd.Row = Op1.Row
    For ij = 0 To UBound(Op1.D): mxAdd.D(ij) = Op1.D(ij) + Op2.D(ij): Next
End Function

' // Subtaction of a two matrix
Public Function mxSub(Op1 As Matrix, Op2 As Matrix) As Matrix
    Dim p As Long, ij As Long
    GetMem4 ByVal ArrPtr(Op1.D), p: If p = 0 Then Exit Function
    GetMem4 ByVal ArrPtr(Op2.D), p: If p = 0 Then Exit Function
    If Op1.Col <> Op2.Col Or Op1.Row <> Op2.Row Then Exit Function
    ReDim mxSub.D(UBound(Op1.D)): mxSub.Col = Op1.Col: mxSub.Row = Op1.Row
    For ij = 0 To UBound(Op1.D): mxSub.D(ij) = Op1.D(ij) - Op2.D(ij): Next
End Function

' // Multiplication of a two matrix
Public Function mxMul(Op1 As Matrix, Op2 As Matrix) As Matrix
    Dim p As Long, i As Long, j As Long, K As Long, iM As Long, i1 As Long, i2 As Long
    GetMem4 ByVal ArrPtr(Op1.D), p: If p = 0 Then Exit Function
    GetMem4 ByVal ArrPtr(Op2.D), p: If p = 0 Then Exit Function
    If Op1.Col <> Op2.Row Then Exit Function
    ReDim mxMul.D(Op1.Row * Op2.Col - 1): mxMul.Col = Op2.Col: mxMul.Row = Op1.Row
    For i = 0 To Op1.Row - 1
        For j = 0 To Op2.Col - 1
        i2 = j
        For K = 0 To Op1.Col - 1
            mxMul.D(iM) = mxMul.D(iM) + Op1.D(i1 + K) * Op2.D(i2)
            i2 = i2 + Op2.Col
        Next
        iM = iM + 1
        Next
    i1 = i1 + Op1.Col
    Next
End Function

' // Determinant of a square matrix
Public Function mxDtm(Op As Matrix) As Double
    Dim p1 As Long, p2 As Long, ij1 As Long, ij2 As Long, Ct As Long, l As Long, t1 As Double, t2 As Double
    GetMem4 ByVal ArrPtr(Op.D), p1: If p1 = 0 Then Exit Function
    If Op.Col <> Op.Row Then Exit Function
    p2 = Op.Col - 1: ij1 = 0: ij2 = p2: Ct = Op.Col * Op.Row: p1 = Op.Col + 1
    For K = 0 To Op.Col - 1
        t1 = Op.D(ij1): t2 = Op.D(ij2)
        For l = 1 To Op.Col - 1
            ij1 = (ij1 + p1) Mod Ct: ij2 = (ij2 + p2) Mod Ct
            t1 = t1 * Op.D(ij1): t2 = t2 * Op.D(ij2)
        Next
        mxDtm = mxDtm + t1 - t2: ij1 = (ij1 + p1) Mod Ct: ij2 = (ij2 + p2) Mod Ct
    Next
End Function

