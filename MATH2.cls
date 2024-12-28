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

Public Type cxComplex
    r   As Double
    I   As Double
End Type

Public Type Matrix
    Col As Long                 ' Number of columns
    Row As Long                 ' Number of rows
    D() As Double
End Type

Private Const PI = 3.14159265358979
Private Const E = 2.71828182845905

Private Const PI2 = PI / 2

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
    If CSng(Value) = -1# Then Acos = PI: Exit Function
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
            Atan2 = Atn(Y / X) + PI
        Else
            Atan2 = PI / 2 - Atn(X / Y)
        End If
    Else
        If X >= -Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= Y Then
            Atan2 = Atn(Y / X) - PI
        Else
            Atan2 = -Atn(X / Y) - PI / 2
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
Public Function cxOne() As cxComplex
    cxOne.r = 1
End Function

' // R = 0, I = 1
Public Function cxImgOne() As cxComplex
    cxOne.I = 1
End Function

' // R = 0, I = 0
Public Function cxZero() As cxComplex
End Function

' // Creating a new complex number
Public Function cxNew(ByVal Real As Double, ByVal Imaginary As Double) As cxComplex
    cxNew.r = Real: cxNew.I = Imaginary
End Function

' // Creating a new complex number by polar coordinates
Public Function cxPolar(ByVal Magnitude As Double, ByVal Phase As Double) As cxComplex
    cxPolar.r = Magnitude * Cos(Phase): cxPolar.I = Magnitude * Sin(Phase)
End Function

' // Return the additive inverse of a specified complex number
Public Function cxNeg(Op As cxComplex) As cxComplex
    cxNeg.r = -Op.r: cxNeg.I = -Op.I
End Function

' // Return the inverse value of a complex number
Public Function cxInv(Op As cxComplex) As cxComplex
    Dim Ab2 As Double
    Ab2 = Op.r * Op.r + Op.I * Op.I
    cxInv.r = Op.r / Ab2: cxInv.I = -Op.I / Ab2
End Function

' // Addition of two complex numbers
Public Function cxAdd(Op1 As cxComplex, Op2 As cxComplex) As cxComplex
    cxAdd.r = Op1.r + Op2.r
    cxAdd.I = Op1.I + Op2.I
End Function

' // Subtraction of two complex numbers
Public Function cxSub(Op1 As cxComplex, Op2 As cxComplex) As cxComplex
    cxSub.r = Op1.r - Op2.r
    cxSub.I = Op1.I - Op2.I
End Function

' // Multiplication of two complex numbers
Public Function cxMul(Op1 As cxComplex, Op2 As cxComplex) As cxComplex
    cxMul.r = Op1.r * Op2.r - Op1.I * Op2.I
    cxMul.I = Op1.r * Op2.I + Op1.I * Op2.r
End Function

' // Division of two complex numbers
Public Function cxDiv(Op1 As cxComplex, Op2 As cxComplex) As cxComplex
    Dim R2 As Double, i2 As Double
    R2 = Op2.r * Op2.r: i2 = Op2.I * Op2.I
    cxDiv.r = (Op1.r * Op2.r + Op1.I * Op2.I) / (R2 + i2)
    cxDiv.I = (Op1.I * Op2.r - Op1.r * Op2.I) / (R2 + i2)
End Function

' // Exponentiation of a complex number
Public Function cxDgr(Op As cxComplex, ByVal Degree As Long) As cxComplex
    Dim Md As Double, Ar As Double
    Md = cxMod(Op): Ar = cxArg(Op): Md = Md ^ Degree: Ar = Ar * Degree
    cxDgr.r = Md * Cos(Ar): cxDgr.I = Md * Sin(Ar)
End Function

' // The square root of a complex number
Public Function cxSqr(Op As cxComplex) As cxComplex
    Dim M As Double, A As Double
    M = Sqr(cxMod(Op)): A = cxArg(Op) / 2
    cxSqr.r = M * Cos(A): cxSqr.I = M * Sin(A)
End Function

' // Module of a complex number
Public Function cxMod(Op As cxComplex) As Double
    Dim R2 As Double, i2 As Double
    R2 = Op.r * Op.r: i2 = Op.I * Op.I
    cxMod = Sqr(R2 + i2)
End Function

' // Phase of a complex number
Public Function cxPhase(Op As cxComplex) As Double
    cxPhase = Atan2(Op.I, Op.r)
End Function

' // Argument of a complex number (equal phase)
Public Function cxArg(Op As cxComplex) As Double
    If Op.I = 0 Then
        If Op.r >= 0 Then cxArg = 0 Else cxArg = PI
    ElseIf Op.r = 0 Then
        If Op.I >= 0 Then cxArg = PI2 Else cxArg = -PI2
    Else
        If Op.r > 0 Then
            cxArg = Atn(Op.I / Op.r)
        ElseIf Op.r < 0 And Op.I > 0 Then
            cxArg = PI + Atn(Op.I / Op.r)
        ElseIf Op.r < 0 And Op.I < 0 Then
            cxArg = -PI + Atn(Op.I / Op.r)
        End If
    End If
End Function

' // Returns the number e, raised to power by complex number
Public Function cxExp(Op As cxComplex) As cxComplex
    cxExp.r = EXP(Op.r) * Cos(Op.I): cxExp.I = EXP(Op.r) * Sin(Op.I)
End Function

' // Addition real number and complex number
Public Function cxAddReal(Op1 As cxComplex, ByVal Op2 As Double) As cxComplex
    cxAddReal.r = Op1.r + Op2
    cxAddReal.I = Op1.I
End Function

' // Subtraction from complex number a real number
Public Function cxSubReal(Op1 As cxComplex, ByVal Op2 As Double) As cxComplex
    cxSubReal.r = Op1.r - Op2
    cxSubReal.I = Op1.I
End Function

' // Subtraction from real number a complex number
Public Function cxRealSub(ByVal Op1 As Double, Op2 As cxComplex) As cxComplex
    cxRealSub.r = Op1 - Op2.r
    cxRealSub.I = -Op2.I
End Function

' // Multiplication complex number on a real number
Public Function cxMulReal(Op1 As cxComplex, ByVal Op2 As Double) As cxComplex
    cxMulReal.r = Op1.r * Op2
    cxMulReal.I = Op1.I * Op2
End Function

' // Division a complex number on a real number
Public Function cxDivReal(Op1 As cxComplex, ByVal Op2 As Double) As cxComplex
    Dim R2 As Double
    R2 = Op2 * Op2
    cxDivReal.r = (Op1.r * Op2) / R2
    cxDivReal.I = (Op1.I * Op2) / R2
End Function

' // Division a real number on a complex number
Public Function cxRealDiv(ByVal Op1 As Double, Op2 As cxComplex) As cxComplex
    Dim R2 As Double, i2 As Double
    R2 = Op2.r * Op2.r: i2 = Op2.I * Op2.I
    cxRealDiv.r = (Op1 * Op2.r) / (R2 + i2)
    cxRealDiv.I = (-Op1 * Op2.I) / (R2 + i2)
End Function

' // Addition of a complex number and imaginary part
Public Function cxAddImg(Op1 As cxComplex, ByVal Op2 As Double) As cxComplex
    cxAddImg.r = Op1.r
    cxAddImg.I = Op1.I + Op2
End Function

' // Subtraction from a complex number a imaginary part
Public Function cxSubImg(Op1 As cxComplex, Op2 As Double) As cxComplex
    cxSubImg.r = Op1.r
    cxSubImg.I = Op1.I - Op2
End Function

' // Subtraction from imaginary part a complex number
Public Function cxImgSub(ByVal Op1 As Double, Op2 As cxComplex) As cxComplex
    cxImgSub.r = -Op2.r
    cxImgSub.I = Op1 - Op2.I
End Function

' // Multiplication complex number on a imaginary part
Public Function cxMulImg(Op1 As cxComplex, ByVal Op2 As Double) As cxComplex
    cxMulImg.r = -Op1.I * Op2
    cxMulImg.I = Op1.r * Op2
End Function

' // Division a complex number on a imaginary part
Public Function cxDivImg(Op1 As cxComplex, ByVal Op2 As Double) As cxComplex
    Dim i2 As Double
    i2 = Op2 * Op2
    cxDivImg.r = (Op1.I * Op2) / i2
    cxDivImg.I = (-Op1.r * Op2) / i2
End Function

' // Division imaginary part on a complex number
Public Function cxImgDiv(ByVal Op1 As Double, Op2 As cxComplex) As cxComplex
    Dim R2 As Double, i2 As Double
    R2 = Op2.r * Op2.r: i2 = Op2.I * Op2.I
    cxImgDiv.r = (Op1 * Op2.I) / (R2 + i2)
    cxImgDiv.I = (Op1 * Op2.r) / (R2 + i2)
End Function

' // Return true if complex number is equal
Public Function cxEq(Op1 As cxComplex, Op2 As cxComplex, _
                Optional NumDigitsAfterDecimal As Long = -1) As Boolean
    If NumDigitsAfterDecimal = -1 Then
        If Op1.r = Op2.r And Op1.I = Op2.I Then cxEq = True
    Else
        If Round(Op1.r, NumDigitsAfterDecimal) = Round(Op2.r, NumDigitsAfterDecimal) And _
           Round(Op1.I, NumDigitsAfterDecimal) = Round(Op2.I, NumDigitsAfterDecimal) Then cxEq = True
    End If
End Function

' // Return absolute value of a complex number
Public Function cxAbs(Op As cxComplex) As Double
    If Op.I = 0 Then
        cxAbs = 0
    ElseIf Op.r > Op.I Then
        cxAbs = Sqr(1 + (Op.I * Op.I) / (Op.r * Op.r))
    ElseIf Op.r <= Op.I Then
        cxAbs = Sqr(1 + (Op.r * Op.r) / (Op.I * Op.I))
    End If
End Function

' // Return complex conjugate of complex number
Public Function cxConj(Op As cxComplex) As cxComplex
    cxConj.r = Op.r
    cxConj.I = -Op.I
End Function

' // The natural logarithm of a complex number
Public Function cxLog(Op As cxComplex) As cxComplex
    Dim M As Double, A As Double
    M = cxMod(Op): A = cxArg(Op)
    cxLog.r = Log(M): cxLog.I = A
End Function

' // The logarithm of a complex number by base X
Public Function cxLogX(Op As cxComplex, ByVal Base As Double) As cxComplex
    Dim M As Double, A As Double, Nc As cxComplex
    M = cxMod(Op): A = cxArg(Op): Nc.r = Log(Base)
    cxLogX.r = Log(M): cxLogX.I = A
    cxLogX = cxDiv(cxLogX, Nc)
End Function

' // Sine of a complex number
Public Function cxSin(Op As cxComplex) As cxComplex
    cxSin.r = Sin(Op.r) * Cosh(Op.I): cxSin.I = Cos(Op.r) * Sinh(Op.I)
End Function
 
' // Cosine of a complex number
Public Function cxCos(Op As cxComplex) As cxComplex
    cxCos.r = Cos(Op.r) * Cosh(Op.I): cxCos.I = -Sin(Op.r) * Sinh(Op.I)
End Function

' // Tangent of a complex number
Public Function cxTan(Op As cxComplex) As cxComplex
    Dim C2 As Double, S2 As Double
    C2 = Cos(Op.r): C2 = C2 * C2: S2 = Sinh(Op.I): S2 = S2 * S2
    cxTan.r = (Sin(Op.r) * Cos(Op.r)) / (C2 + S2)
    cxTan.I = (Sinh(Op.I) * Cosh(Op.I)) / (C2 + S2)
End Function

' // Cotangent of a complex number
Public Function cxCtg(Op As cxComplex) As cxComplex
    Dim C2 As Double, S2 As Double
    C2 = Sin(Op.r): C2 = C2 * C2: S2 = Sinh(Op.I): S2 = S2 * S2
    cxCtg.r = (Sin(Op.r) * Cos(Op.r)) / (C2 + S2)
    cxCtg.I = -(Sinh(Op.I) * Cosh(Op.I)) / (C2 + S2)
End Function

' // Secant of a complex number
Public Function cxSec(Op As cxComplex) As cxComplex
    Dim C2 As Double, S2 As Double
    C2 = Cos(Op.r): C2 = C2 * C2: S2 = Sinh(Op.I): S2 = S2 * S2
    cxSec.r = (Cos(Op.r) * Cosh(Op.I)) / (C2 + S2)
    cxSec.I = -(Sin(Op.r) * Sinh(Op.I)) / (C2 + S2)
End Function

' // Cosecant of a complex number
Public Function cxCsc(Op As cxComplex) As cxComplex
    Dim C2 As Double, S2 As Double
    C2 = Sin(Op.r): C2 = C2 * C2: S2 = Sinh(Op.I): S2 = S2 * S2
    cxCsc.r = (Sin(Op.r) * Cosh(Op.I)) / (C2 + S2)
    cxCsc.I = (Cos(Op.r) * Sinh(Op.I)) / (C2 + S2)
End Function

' // Arcsine of a complex number
Public Function cxAsin(Op As cxComplex) As cxComplex
    cxAsin = cxMulImg(cxLog(cxAdd(cxMulImg(Op, 1), cxSqr(cxRealSub(1, cxMul(Op, Op))))), -1)
End Function

' // Arccosine of a complex number
Public Function cxAcos(Op As cxComplex) As cxComplex
    cxAcos = cxAddReal(cxMulImg(cxLog(cxAdd(cxMulImg(Op, 1), cxSqr(cxRealSub(1, cxMul(Op, Op))))), 1), PI2)
End Function

' // Arctangent of a complex number
Public Function cxAtan(Op As cxComplex) As cxComplex
    Dim Iz As cxComplex
    Iz = cxMulImg(Op, 1)
    cxAtan = cxMulImg(cxSub(cxLog(cxRealSub(1, Iz)), cxLog(cxAddReal(Iz, 1))), 0.5)
End Function

' // Arccotangent of a complex number
Public Function cxActg(Op As cxComplex) As cxComplex
    cxActg = cxMulImg(cxSub(cxLog(cxDiv(cxSubImg(Op, 1), Op)), cxLog(cxDiv(cxAddImg(Op, 1), Op))), 0.5)
End Function

' // Arcsecant of a complex number
Public Function cxAsec(Op As cxComplex) As cxComplex
    cxAsec = cxAcos(cxDgr(Op, -1))
End Function

' // Arccosecant of a complex number
Public Function cxAcsc(Op As cxComplex) As cxComplex
    cxAcsc = cxAsin(cxDgr(Op, -1))
End Function

' // Hyperbolic sine of a complex number
Public Function cxSinh(Op As cxComplex) As cxComplex
    cxSinh = cxMulImg(cxSin(cxMulImg(Op, 1)), -1)
End Function

' // Hyperbolic cosine of a complex number
Public Function cxCosh(Op As cxComplex) As cxComplex
    cxCosh = cxCos(cxMulImg(Op, 1))
End Function

' // Hyperbolic tangent of a complex number
Public Function cxTanh(Op As cxComplex) As cxComplex
    cxTanh = cxMulImg(cxTan(cxMulImg(Op, 1)), -1)
End Function

' // Hyperbolic cotangent of a complex number
Public Function cxCtgh(Op As cxComplex) As cxComplex
    cxCtgh = cxRealDiv(1, cxTanh(Op))
End Function

' // Hyperbolic secant of a complex number
Public Function cxSech(Op As cxComplex) As cxComplex
    cxSech = cxRealDiv(1, cxCosh(Op))
End Function

' // Hyperbolic cosecant of a complex number
Public Function cxCsch(Op As cxComplex) As cxComplex
    cxCsch = cxRealDiv(1, cxSinh(Op))
End Function

' // Hyperbolic arcsine of a complex number
Public Function cxAsinh(Op As cxComplex) As cxComplex
    cxAsinh = cxLog(cxAdd(Op, cxSqr(cxAddReal(cxMul(Op, Op), 1))))
End Function

' // Hyperbolic arccosine of a complex number
Public Function cxAcosh(Op As cxComplex) As cxComplex
    cxAcosh = cxLog(cxAdd(Op, cxMul(cxSqr(cxAddReal(Op, 1)), cxSqr(cxSubReal(Op, 1)))))
End Function

' // Hyperbolic arctangent of a complex number
Public Function cxAtanh(Op As cxComplex) As cxComplex
    cxAtanh = cxMulReal(cxLog(cxDiv(cxAddReal(Op, 1), cxRealSub(1, Op))), 0.5)
End Function

' // Hyperbolic arccotangent of a complex number
Public Function cxActgh(Op As cxComplex) As cxComplex
    cxActgh = cxMulReal(cxLog(cxDiv(cxAddReal(Op, 1), cxSubReal(Op, 1))), 0.5)
End Function

' // Hyperbolic arcsecant of a complex number
Public Function cxAsech(Op As cxComplex) As cxComplex
    Dim Z As cxComplex
    Z = cxRealDiv(1, Op)
    cxAsech = cxLog(cxAdd(Z, cxSqr(cxAddReal(cxMul(Z, Z), 1))))
End Function

' // Hyperbolic arccosecant of a complex number
Public Function cxAcsch(Op As cxComplex) As cxComplex
    Dim Z As cxComplex
    Z = cxRealDiv(1, Op)
    cxAcsch = cxLog(cxAdd(Z, cxMul(cxSqr(cxAddReal(Z, 1)), cxSqr(cxSubReal(Z, 1)))))
End Function
' // Print matrix to immediate window
Public Function PrintMtrx(Op As Matrix)
    Dim Ts As String, I As Long, j As Long
    Debug.Print vbNewLine
    Debug.Print "Col= " & Op.Col & " ; Row= " & Op.Row
    For I = 0 To Op.Row - 1: For j = 0 To Op.Col - 1
        Ts = space(10)
        LSet Ts = str(Round(Op.D(I * Op.Col + j), 3))
        Debug.Print Ts;
    Next: Debug.Print vbNewLine;: Next
End Function

' // Creating a matrix
Public Function mxCreate(ByVal Row As Long, ByVal Col As Long, ParamArray Y()) As Matrix
    Dim p As Variant, C As Long
    If Row <= 0 Or Col <= 0 Then Exit Function
    If Row * Col < UBound(Y) + 1 Then Exit Function
    mxCreate.Row = Row: mxCreate.Col = Col
    ReDim mxCreate.D(Row * Col - 1): C = 0
    For Each p In Y
        mxCreate.D(C) = p: C = C + 1
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
    Dim I As Long, j As Long, p As Long
    GetMem4 ByVal ArrPtr(Op.D), p: If p = 0 Then Exit Function
    mxTrans.Row = Op.Col: mxTrans.Col = Op.Row: ReDim mxTrans.D(UBound(Op.D))
    For j = 0 To mxTrans.Col - 1: For I = 0 To mxTrans.Row - 1
        mxTrans.D(I + j * mxTrans.Row) = Op.D(j + I * Op.Row)
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
    Dim p As Long, I As Long, j As Long, K As Long, iM As Long, i1 As Long, i2 As Long
    GetMem4 ByVal ArrPtr(Op1.D), p: If p = 0 Then Exit Function
    GetMem4 ByVal ArrPtr(Op2.D), p: If p = 0 Then Exit Function
    If Op1.Col <> Op2.Row Then Exit Function
    ReDim mxMul.D(Op1.Row * Op2.Col - 1): mxMul.Col = Op2.Col: mxMul.Row = Op1.Row
    For I = 0 To Op1.Row - 1
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
    Dim P1 As Long, P2 As Long, ij1 As Long, ij2 As Long, Ct As Long, L As Long, T1 As Double, T2 As Double
    GetMem4 ByVal ArrPtr(Op.D), P1: If P1 = 0 Then Exit Function
    If Op.Col <> Op.Row Then Exit Function
    P2 = Op.Col - 1: ij1 = 0: ij2 = P2: Ct = Op.Col * Op.Row: P1 = Op.Col + 1
    For K = 0 To Op.Col - 1
        T1 = Op.D(ij1): T2 = Op.D(ij2)
        For L = 1 To Op.Col - 1
            ij1 = (ij1 + P1) Mod Ct: ij2 = (ij2 + P2) Mod Ct
            T1 = T1 * Op.D(ij1): T2 = T2 * Op.D(ij2)
        Next
        mxDtm = mxDtm + T1 - T2: ij1 = (ij1 + P1) Mod Ct: ij2 = (ij2 + P2) Mod Ct
    Next
End Function
