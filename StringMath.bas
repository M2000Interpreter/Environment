Attribute VB_Name = "Module13"
Option Explicit
'Option Compare Text 'Database for Access
'--------------------------------------------------------------------------------------------------------------
'https://cosxoverx.livejournal.com/47220.html
'Credit to Rebecca Gabriella's String Math Module (Big Integer Library) for VBA (Visual Basic for Applications)
' Minor edits made with comments and other.
' Additions from George Karras
'--------------------------------------------------------------------------------------------------------------

Private Type PartialDivideInfo
    Quotient As Long
    Subtrahend As String
    Remainder As String
End Type

Private sLastRemainder As String
' Alphabet moved to Module1 as variable, now has an Ansi format(1 byte per letter/digit)
'Private Const Alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Property Get LastRemainder()
    LastRemainder = CVar(sLastRemainder)
End Property

Public Function compare(sA As String, sb As String, Optional absfirst) As Integer
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns an integer that represents one of three states
    'sA > sB returns 1, sA < sB returns -1, and sA = sB returns 0
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim I As Long, iA As Long, iB As Long
    
    'handle any early exits on basis of signs
    
    bAN = (LeftB$(sA, 1) = ChrB$(45))
    If Not IsMissing(absfirst) Then
        If bAN Then
            MidB$(sA, 1, 1) = ChrB$(48)
            bAN = False
        End If
    End If
    bBN = (LeftB$(sb, 1) = ChrB$(45))
    If bAN Then MidB$(sA, 1, 1) = ChrB$(48)
    If bBN Then MidB$(sb, 1, 1) = ChrB$(48)
    If bAN And bBN Then
        bRN = True
    ElseIf bBN Then
        compare = 1
        Exit Function
    ElseIf bAN Then
        compare = -1
        Exit Function
    Else
        bRN = False
    End If
    
    'remove any leading zeros
    Dim CROP As Long, lim As Long
    
    CROP = 1
    lim = LenB(sA)
    Do While CROP <= lim
        If MidB$(sA, CROP, 1) <> ChrB$(48) Then Exit Do
        CROP = CROP + 1
    Loop
    sA = MidB$(sA, CROP)
    CROP = 1
    lim = LenB(sb)
    Do While CROP <= lim
       If MidB$(sb, CROP, 1) <> ChrB$(48) Then Exit Do
       CROP = CROP + 1
    Loop
    sb = MidB$(sb, CROP)
    
    'then decide size first on basis of length
    If LenB(sA) < LenB(sb) Then
        compare = -1
    ElseIf LenB(sA) > LenB(sb) Then
        compare = 1
    Else 'unless they are the same length
        compare = 0
        'then check each digit by digit
        For I = 1 To LenB(sA)
            iA = AscB(MidB$(sA, I, 1))
            iB = AscB(MidB$(sb, I, 1))
            If iA < iB Then
                compare = -1
                Exit For
            ElseIf iA > iB Then
                compare = 1
                Exit For
            Else 'defaults zero
            End If
        Next I
    End If
    
    'decide about any negative signs
    If bRN Then
        compare = -compare
    End If

End Function
'ByVal sA As String, ByVal sB As String

Public Function Add(sA As String, sb As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sum of sA and sB as string integer in Add()
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim iA As Long, iB As Long, iCarry As Integer
       
    'test for empty parameters
    If LenB(sA) = 0 Or LenB(sb) = 0 Then
        MyEr "Empty parameter in Add", "Κενοί Παράμετροι στην Add"
        Exit Function
    End If
        
    'handle some negative values with Subtract()
    bAN = (LeftB$(sA, 1) = ChrB$(45))
    bBN = (LeftB$(sb, 1) = ChrB$(45))
    If bAN Then sA = MidB$(sA, 2)
    If bBN Then sb = MidB$(sb, 2)
    If bAN And bBN Then 'both negative
        bRN = True      'set output reminder
    ElseIf bBN Then     'use subtraction
        Add = subtract(sA, sb)
        Exit Function
    ElseIf bAN Then     'use subtraction
        Add = subtract(sb, sA)
        Exit Function
    Else
        bRN = False
    End If
    Dim cur As Long
    
    'add column by column
    iA = LenB(sA)
    iB = LenB(sb)
    iCarry = 0
    If iA > iB Then
    cur = iA + 1
    Add = SpaceB(cur)
    Else
    cur = iB + 1
    Add = SpaceB(cur)
    End If
    Do While iA > 0 And iB > 0
        iCarry = iCarry + AscB(MidB$(sA, iA, 1)) + AscB(MidB$(sb, iB, 1)) - 96
        MidB$(Add, cur, 1) = ChrB$(iCarry Mod 10 + 48)
        iCarry = iCarry \ 10
        iA = iA - 1
        iB = iB - 1
        cur = cur - 1
    Loop
    
    'Assuming param sA is longer
    Do While iA > 0
        iCarry = iCarry + AscB(MidB$(sA, iA, 1)) - 48
        MidB$(Add, cur, 1) = ChrB$(iCarry Mod 10 + 48)
        iCarry = iCarry \ 10
        iA = iA - 1
        cur = cur - 1
    Loop
    'Assuming param sB is longer
    Do While iB > 0
        iCarry = iCarry + AscB(MidB$(sb, iB, 1)) - 48
        MidB$(Add, cur, 1) = ChrB$(iCarry Mod 10 + 48)
        iCarry = iCarry \ 10
        iB = iB - 1
        cur = cur - 1
    Loop
    MidB$(Add, cur, 1) = ChrB$(iCarry + 48)
    cur = cur - 1
    Do While cur > 0
        MidB$(Add, cur, 1) = ChrB$(48)
        cur = cur - 1
    Loop
    'remove any leading zeros
    'Do While LenB(Add) > 1 And LeftB$(Add, 1) = ChrB$(48)
    '    Add = MidB$(Add, 2)
    'Loop
  '  If cur + 1 = 1 Then
    Add = TrimZero(Add)
 '   Else
 '   Add = TrimZero(MidB$(Add, cur + 2))
  '  End If
    'decide about any negative signs
    If Add <> ChrB$(48) And bRN Then
        Add = ChrB$(45) + Add
    End If

End Function

Private Function RealMod(ByVal iA As Long, ByVal iB As Long) As Long
    'Returns iA mod iB in RealMod() as an integer. Good for small values.
    'Normally Mod takes on the sign of iA but here
    'negative values are increased by iB until result is positive.
    'Credit to Rebecca Gabriella's String Math Module with added edits.
    'https://cosxoverx.livejournal.com/47220.html
        
    If iB = 0 Then
        MyEr "Divide by zero", "Διαίρεση με το μηδέν"
        Exit Function
    End If
    
    If iA Mod iB = 0 Then
        RealMod = 0
    ElseIf iA < 0 Then
        RealMod = iB + iA Mod iB 'increase till pos
    Else
        RealMod = iA Mod iB
    End If

End Function

Private Function RealDiv(ByVal iA As Long, ByVal iB As Long) As Long
    'Returns integer division iA divided by iB in RealDiv().Good for small values.
    'Credit to Rebecca Gabriella's String Math Module with added edits.
    'https://cosxoverx.livejournal.com/47220.html
    
    If iB = 0 Then
        MyEr "Divide by zero", "Διαίρεση με το μηδέν"
        Exit Function
    End If
    
    If iA Mod iB = 0 Then
        RealDiv = iA \ iB
    ElseIf iA < 0 Then
        RealDiv = iA \ iB - 1 'round down
    Else
        RealDiv = iA \ iB
    End If

End Function
' ByVal sA As String, ByVal sB As String
Public Function subtract(sA As String, sb As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sA minus sB as string integer in Subtract()
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim iA As Long, iB As Long, iComp As Integer
    
    'test for empty parameters
    If LenB(sA) = 0 Or LenB(sb) = 0 Then
        MyEr "Empty parameter in Subtract", "Κενοί Παράμετροι στην Subtract"
        Exit Function
    End If
        
    'handle some negative values with Add()
    bAN = (LeftB$(sA, 1) = ChrB$(45))
    bBN = (LeftB$(sb, 1) = ChrB$(45))
    If bAN Then sA = MidB$(sA, 2)
    If bBN Then sb = MidB$(sb, 2)
    If bAN And bBN Then
        bRN = True
    ElseIf bBN Then
        subtract = Add(sA, sb)
        Exit Function
    ElseIf bAN Then
        subtract = ChrB$(45) + Add(sA, sb)
        Exit Function
    Else
        bRN = False
    End If
    
    'get biggest value into variable sA
    iComp = compare(sA, sb)
    If iComp = 0 Then     'parameters equal in size
        subtract = ChrB$(48)
        Exit Function
    ElseIf iComp < 0 Then 'sA < sB
        SwapStrings sA, sb
        'subtract = sA     'so swop sA and sB
        'sA = sb           'to ensure sA >= sB
        'sb = subtract
        bRN = Not bRN     'and reverse output sign
    End If
    iA = LenB(sA)          'recheck lengths
    iB = LenB(sb)
    iComp = 0
    subtract = ""
    subtract = SpaceB(iA)
    'subtract column by column
    Do While iA > 0 And iB > 0
        iComp = iComp + AscB(MidB$(sA, iA, 1)) - AscB(MidB$(sb, iB, 1))
        MidB$(subtract, iA, 1) = ChrB$(RealMod(iComp, 10) + 48) '+ subtract
        iComp = RealDiv(iComp, 10)
        iA = iA - 1
        iB = iB - 1
    Loop
    'then assuming param sA is longer
    Do While iA > 0
        iComp = iComp + AscB(MidB$(sA, iA, 1)) - 48
        MidB$(subtract, iA, 1) = ChrB$(RealMod(iComp, 10) + 48)
        iComp = RealDiv(iComp, 10)
        iA = iA - 1
    Loop
    
    'remove any leading zeros from result
    'Do While LenB(subtract) > 1 And LeftB$(subtract, 1) = ChrB$(48)
    '    subtract = MidB$(subtract, 2)
    'Loop
    subtract = TrimZero(subtract)
    'decide about any negative signs
    If subtract <> ChrB$(48) And bRN Then
        subtract = ChrB$(45) + subtract
    End If

End Function

Public Function multiply(sA As String, sb As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sA times sB as string integer in Multiply()
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim M() As Long, iCarry As Long
    Dim iAL As Long, iBL As Long, iA As Long, iB As Long
        
    'test for empty parameters
    If LenB(sA) = 0 Or LenB(sb) = 0 Then
        MyEr "Empty parameter in Multiply", "Κενοί Παράμετροι στην Multiply"
        Exit Function
    End If
        
    'handle any negative signs
    bAN = (LeftB$(sA, 1) = ChrB$(45))
    bBN = (LeftB$(sb, 1) = ChrB$(45))
    If bAN Then sA = MidB$(sA, 2)
    If bBN Then sb = MidB$(sb, 2)
    bRN = (bAN <> bBN)
    iAL = LenB(sA)
    iBL = LenB(sb)
    
    'perform long multiplication without carry in notional columns
    ReDim M(1 To (iAL + iBL - 1)) 'expected length of product
    For iA = 1 To iAL
        For iB = 1 To iBL
            M(iA + iB - 1) = M(iA + iB - 1) + CLng(AscB(MidB$(sA, iAL - iA + 1, 1)) - 48) * CLng(AscB(MidB$(sb, iBL - iB + 1, 1)) - 48)
        Next iB
    Next iA
    iCarry = 0
    multiply = ""
    
    'add up column results with carry
    For iA = 1 To iAL + iBL - 1
        iCarry = iCarry + M(iA)
        multiply = ChrB$(iCarry Mod 10 + 48) + multiply
        iCarry = iCarry \ 10
    Next iA
    multiply = ChrB$(iCarry + 48) + multiply
    
    'remove any leading zeros
   ' Do While LenB(multiply) > 1 And LeftB$(multiply, 1) = ChrB$(48)
   '     multiply = MidB$(multiply, 2)
   ' Loop
    multiply = TrimZero(multiply)
    'decide about any negative signs
    If multiply <> ChrB$(48) And bRN Then
        multiply = ChrB$(45) + multiply
    End If

End Function

Private Function PartialDivide(sA As String, sb As String) As PartialDivideInfo
    'Called only by Divide() to assist in fitting trials for long division
    'All of Quotient, Subtrahend, and Remainder are returned as elements of type PartialDivideInfo
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
        
    For PartialDivide.Quotient = 9 To 1 Step -1                                'propose a divisor to fit
        SwapStrings PartialDivide.Subtrahend, multiply((sb), ChrB$(PartialDivide.Quotient + 48)) 'test by multiplying it out
        If compare(PartialDivide.Subtrahend, (sA)) <= 0 Then                      'best fit found
            SwapStrings PartialDivide.Remainder, subtract((sA), (PartialDivide.Subtrahend))   'get remainder
            Exit Function                                                      'exit with best fit details
        End If
    Next PartialDivide.Quotient
    
    'no fit found, divisor too big
    PartialDivide.Quotient = 0
    PartialDivide.Subtrahend = ChrB$(48)
    PartialDivide.Remainder = sA

End Function

Public Function divide(sA As String, sb As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sA divided by sB as string integer in Divide()
    'The remainder is available as sLastRemainder at Module level
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN  As Boolean, bBN As Boolean, bRN As Boolean
    Dim iC As Long
    Dim s As String
    Dim D As PartialDivideInfo
    
    'test for empty parameters
    If LenB(sA) = 0 Or LenB(sb) = 0 Then
        MyEr "Empty parameter in Divide", "Κενοί Παράμετροι στην Divide"
        Exit Function
    End If
    
    bAN = (LeftB$(sA, 1) = ChrB$(45)) 'true for neg
    bBN = (LeftB$(sb, 1) = ChrB$(45))
    If bAN Then sA = MidB$(sA, 2) 'take two charas if neg
    If bBN Then sb = MidB$(sb, 2)
    bRN = (bAN <> bBN)
    If compare(sb, ChrB$(48)) = 0 Then
        DevZero
        Exit Function
    ElseIf compare(sA, ChrB$(48)) = 0 Then
        divide = ChrB$(48)
        sLastRemainder = ChrB$(48)
        Exit Function
    End If
    iC = compare(sA, sb)
    If iC < 0 Then
        divide = ChrB$(48)
        sLastRemainder = sA
        Exit Function
    ElseIf iC = 0 Then
        If bRN Then
            divide = ChrB$(45) + ChrB$(49)
        Else
            divide = ChrB$(49)
        End If
        sLastRemainder = ChrB$(48)
        Exit Function
    End If
    divide = SpaceB(LenB(sA) + 1)
    s = ""
    
    'Long division method
    For iC = 1 To LenB(sA)
        'take increasing number of digits
        s = s + MidB$(sA, iC, 1)
        D = PartialDivide(s, sb)   'find best fit
        MidB$(divide, iC, 1) = ChrB$(D.Quotient + 48)
        s = D.Remainder
    Next iC
    
    'remove any leading zeros
    'Do While LenB(divide) > 1 And LeftB$(divide, 1) = ChrB$(48)
    '    divide = MidB$(divide, 2)
    'Loop
    divide = TrimZero(RtrimB(divide))
    'decide about the signs
    If divide <> ChrB$(48) And bRN Then
        divide = ChrB$(45) + divide
    End If
    If bAN Then
        sLastRemainder = ChrB$(45) + s
    Else
        sLastRemainder = s 'string integer remainder
    End If
End Function

Public Function LastModulus() As String
    LastModulus = sLastRemainder
End Function

Public Function Modulus(sA As String, sb As String) As String
    divide sA, sb
    Modulus = sLastRemainder
End Function

Public Function BigIntFromString(sIn As String, iBaseIn As Integer) As String
    'Returns base10 integer string from sIn of different base (iBaseIn).
    'Example for sIn = "1A" and iBaseIn = 16, returns the base10 result 26.
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
            
    Dim bRN As Boolean
    Dim sBS As String
    Dim iP As Long, iV As Long
    
    'test for empty parameters
    If LenB(sIn) = 0 Or iBaseIn = 0 Then
        MyEr "Bad parameter in BigIntFromString", "Προβληματικοί παράμετροι στη BigIntFromString"
        Exit Function
    End If
        
    'handle negative signs
    If LeftB$(sIn, 1) = ChrB$(45) Then
        bRN = True
        sIn = MidB$(sIn, 2)
    Else
        bRN = False
    End If
    sBS = StrConv(CStr(iBaseIn), vbFromUnicode)
    
    BigIntFromString = ChrB$(48)
    For iP = 1 To LenB(sIn)
        'use constant list position and base for conversion
        iV = InStrB(Alphabet, MidB$(sIn, iP, 1))
        If iV > 0 Then 'accumulate
            BigIntFromString = multiply(BigIntFromString, sBS)
            BigIntFromString = Add(BigIntFromString, StrConv(CStr(iV - 1), vbFromUnicode))
        End If
    Next iP
    
    'decide on any negative signs
    If bRN Then
        BigIntFromString = ChrB$(45) + BigIntFromString
    End If

End Function

Public Function BigIntToString(sIn As String, iBaseOut As Integer) As String
    'Returns integer string of specified iBaseOut (iBaseOut) from base10 (sIn) integer string.
    'Example for sIn = "26" and iBaseOut = 16, returns the output "1A".
    'Credit to Rebecca Gabriella'sIn String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
            
    Dim bRN As Boolean
    Dim sb As String
    Dim iV As Long
    
    'test for empty parameters
    If LenB(sIn) = 0 Or iBaseOut = 0 Then
        MyEr "Bad parameter in BigIntToString", "Προβληματικοί παράμετροι στη BigIntToString"
        Exit Function
    End If
    
    'handle negative signs
    If LeftB$(sIn, 1) = ChrB$(45) Then
        bRN = True
        sIn = MidB$(sIn, 2)
    Else
        bRN = False
    End If
    sb = StrConv(CStr(iBaseOut), vbFromUnicode)
    
    BigIntToString = ""
    On Error GoTo 100
    Dim ivs As String
    Do While compare((sIn), ChrB$(48)) > 0
        sIn = divide(sIn, sb)
        ivs = LastModulus()
        iV = AscB(RightB$(ivs, 1)) - 48
        If LenB(ivs) > 1 Then
            iV = iV + (AscB(MidB$(ivs, LenB(ivs) - 1, 1)) - 48) * 10
        End If
        
        'locates appropriate alphabet character
        BigIntToString = MidB$(Alphabet, iV + 1, 1) + BigIntToString
    Loop
    
    'decide on any negative signs
    If BigIntToString = "" Then
        BigIntToString = ChrB$(48)
    ElseIf BigIntToString <> ChrB$(48) And bRN Then
        BigIntToString = ChrB$(45) + BigIntToString
    End If
    Exit Function
100
    MyEr "BigInteger Error", "Πρόβλημα Μεγάλου Ακέραιου"
End Function

Function IntStrByExp(sA As String, sExp As String) As String
    'Returns integer string raised to exponent iExp As Long string
    'Assumes posiive exponent, and pos or neg string integer
    
    Dim ba As Boolean, br As Boolean
    
    'check parameter
    If LeftB$(sExp, 1) = ChrB$(45) Then
        MyEr "Negative power in IntPower", "Αρνητική δύναμη στην IntPower"
        Exit Function
    End If
    
    Dim VarZ
    
    
    'handle any negative signs
    ba = (LeftB$(sA, 1) = ChrB$(45))
    If ba Then sA = MidB$(sA, 2) 'Else sA = midb$(sA, 1)
    If LenB(sExp) > 4 Then
        MyEr "too big exponent", "υπερβολικά μεγάλη δύναμη"
        IntStrByExp = ChrB$(49)
        Exit Function
    End If
    If ba And val(RightB$(sExp, 1)) Mod 2 <> 0 Then br = True
    VarZ = CDec(StrConv(sExp, vbUnicode))
    'run multiplication loop
    IntStrByExp = ChrB$(49)
    Do Until VarZ = 0
        IntStrByExp = multiply(IntStrByExp, sA)
        VarZ = VarZ - 1
    Loop

    'remove any leading zeros
    IntStrByExp = TrimZero(IntStrByExp)
    
    'decide on any signs
    If IntStrByExp <> ChrB$(48) And br Then
       IntStrByExp = ChrB$(45) & IntStrByExp
    End If

End Function
Function IsProbablyPrime(sA As String, K As Integer) As Boolean
    If LenB(sA) < 1 Then Exit Function
    If LeftB$(sA, 1) = ChrB$(45) Then
        MyEr "Negative Prime not exist", "Αρνητικός πρώτος δεν υπάρχει"
        Exit Function
    End If
    If sA = ChrB$(50) Then IsProbablyPrime = True: Exit Function
    If val(RightB$(sA, 1)) Mod 2 = 0 Then Exit Function
    If sA = ChrB$(49) Then Exit Function
    
    Dim nn As String, D As String, s As Long, Z As Long
    nn = subtract(sA, ChrB$(49))
    D = nn
    While compare(Modulus((D), ChrB$(50)), ChrB$(48)) = 0
        s = s + 1
        D = divide(D, ChrB$(50))
    Wend
    Z = LenB(sA)
    Dim A As String, X As String, I As Long, j As Long
    
    IsProbablyPrime = True
    For I = 1 To K
        
        Do
            A = SpaceB(LenB(sA))
            For j = 1 To LenB(sA)
                MidB$(A, j, 1) = ChrB$(47 + Int(10 * RndM(rndbase) + 1))
            Next
        Loop Until compare(nn, A) = 1 And compare(A, ChrB$(49)) > -1
        X = modpow(A, (D), (sA))
        If compare(X, ChrB$(49)) <> 0 Then ' continue
            If compare(X, nn) <> 0 Then ' continue
                For j = 1 To s
                    X = modpow(X, ChrB$(50), (sA))
                    If compare(X, ChrB$(49)) = 0 Then IsProbablyPrime = False: Exit Function
                    If compare(X, nn) = 0 Then Exit For
                Next
            End If
            If compare(X, nn) <> 0 Then IsProbablyPrime = False: Exit Function
        End If
    Next
End Function

Public Function modpow(sBase As String, sExp As String, sMod As String) As String
    If LeftB$(sExp, 1) = ChrB$(45) Then
        MyEr "Negative power in modpow", "Αρνητική δύναμη στην modpow"
        Exit Function
    End If
    If LeftB$(sMod, 1) = ChrB$(45) Or sMod = ChrB$(48) Then
        MyEr "Zero or negative Modules in modpow", "Μηδενικό ή Αρνητικό Μέτρο στην modpow"
        Exit Function
    End If
    Dim br As Boolean, ba As Boolean
    ba = (LeftB$(sBase, 1) = ChrB$(45))
    If ba Then sBase = MidB$(sBase, 2)
    If ba And AscB(RightB$(sExp, 1)) Mod 2 <> 0 Then br = True
    modpow = ChrB$(49)
    Do While sExp <> ChrB$(48)
        If AscB(RightB$(sExp, 1)) Mod 2 = 1 Then
            modpow = Module13.Modulus(Module13.multiply(modpow, (sBase)), (sMod))
        End If
        sExp = divide(sExp, ChrB$(50))
        sBase = Modulus(multiply(sBase, (sBase)), sMod)
    Loop
    
    If modpow <> ChrB$(48) And br Then
       modpow = ChrB$(45) & modpow
    End If
End Function
Public Function IntSqr(sA As String) As String
    If LeftB$(sA, 1) = ChrB$(45) Or sA = ChrB$(48) Then
        MyEr "Zero or negative paramter for integer Square Root", "Μηδενική ή Αρνητική παράμετρος για ακέραια τετραγωνική ρίζα"
        Exit Function
    End If
    Dim q As String, r As String, t As String, Z As String, minusone As String
    minusone = ChrB$(45) + ChrB$(49)
    Z = sA
    r = ChrB$(48)
    q = ChrB$(49)
    Do
    q = multiply(q, ChrB$(52))
    Loop Until compare((q), (sA)) = 1
    Do
        If compare((q), ChrB$(49)) < 1 Then Exit Do
        q = divide(q, ChrB$(52))
        t = subtract(subtract((Z), (r)), (q))
        r = divide(r, ChrB$(50))
        If compare((t), (minusone)) > -1 Then
            SwapStrings Z, t
            r = Add(r, (q))
        End If
    Loop
    IntSqr = r
End Function
Public Function IsPrime(sA As String) As Boolean
    ' works but not used - use IsProbablyPrime()
    If LenB(sA) < 1 Then Exit Function
    If LeftB$(sA, 1) = ChrB$(45) Then
        MyEr "Negative Prime not exist", "Αρνητικός πρώτος δεν υπάρχει"
        Exit Function
    End If
    Dim D As String
    
    If sA = ChrB$(50) Then IsPrime = True: Exit Function
    If val(RightB$(sA, 1)) Mod 2 = 0 Then Exit Function
    If sA = ChrB$(49) Then Exit Function
    If sA = ChrB$(51) Then IsPrime = True: Exit Function
    If sA = ChrB$(53) Then IsPrime = True: Exit Function
    If compare(Modulus((sA), ChrB$(51)), ChrB$(48)) = 0 Then Exit Function
    Dim x1 As String
    x1 = IntSqr(sA)
    D = ChrB$(53)
    Do
        If compare(Modulus((sA), (D)), ChrB$(48)) = 0 Then Exit Do
        D = Add(ChrB$(50), D)
        If compare((D), (x1)) = 1 Then IsPrime = True: Exit Function
        If compare(Modulus((sA), (D)), ChrB$(48)) = 0 Then Exit Do
        D = Add(ChrB$(52), D)
        If compare((D), (x1)) = 1 Then IsPrime = True: Exit Function
    Loop
End Function
' GET UNICODE AND CHANGE IT TO ANSI
Public Function CreateBigInteger(s$, Optional basenum) As BigInteger
    Set CreateBigInteger = New BigInteger
    s$ = TrimZeroU(s$)
    If IsMissing(basenum) Then
        If TestNumber(s$) Then
            CreateBigInteger.Load StrConv(s$, vbFromUnicode), 10
        Else
            MyEr "not in base 10 (invalid chars)", "Δεν είναι στη βάση 10 (αντικανονικοί χαρακτήρες)"
        End If
    ElseIf basenum > 1 And basenum < 37 Then
        If TestNumberOnBase(s$, CInt(basenum)) Then
        
            CreateBigInteger.AnyBaseInput s$, CInt(basenum)
        Else
            MyEr "not in base " & basenum & " (invalid chars)", "Δεν είναι στη βάση " & basenum & " (αντικανονικοί χαρακτήρες)"
        End If
    Else
        MyEr "base out of limits", "η βάση είναι εκτός ορίων"
        
    End If
End Function
Public Function TestNumberOnBase(s$, b As Integer) As Boolean
    Dim I As Long, lim As Long, ss As String
    Const Alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    ss = Mid$(Alphabet, 1, b)
    lim = Len(s$)
    If lim = 0 Then Exit Function
    I = 1
    If Left$(s$, 1) = Chr$(45) Then I = I + 1: If lim = 1 Then Exit Function
    Do While I <= lim
    If InStr(ss, Mid$(s$, I, 1)) = 0 Then Exit Function
    I = I + 1
    Loop
    TestNumberOnBase = True
End Function
Public Function TestNumber(s$) As Boolean

    Dim I As Long, lim As Long, W As Long
    lim = Len(s$)
    If lim = 0 Then Exit Function
    I = 1
    If Left$(s$, 1) = Chr$(45) Then I = I + 1: If lim = 1 Then Exit Function
    Do While I <= lim
    
    If InStr("0123456789", Mid$(s$, I, 1)) = 0 Then Exit Function
    I = I + 1
    Loop
    TestNumber = True
End Function
Public Function TrimZero(s$) As String
    Dim I As Long, j As Long, lim As Long
    lim = LenB(s$)
    If lim = 0 Then Exit Function
    j = 1
    TrimZero = SpaceB(LenB(s$))
    If LeftB$(s$, j) = ChrB$(45) Then MidB$(TrimZero, j, 1) = ChrB$(45): j = j + 1
    I = j
    lim = lim - 1
    Do While I <= lim
        If MidB$(s$, I, 1) <> ChrB$(48) Then Exit Do
        I = I + 1
    Loop
    MidB$(TrimZero, j, lim - I + 2) = MidB$(s$, I)
    TrimZero = RtrimB(TrimZero)
End Function
Public Function RtrimB(s$) As String
    Dim I As Long, j As Long
    j = LenB(s$)
    If j = 0 Then Exit Function
    For I = j To 1 Step -1
        If MidB$(s$, I, 1) <> ChrB$(32) Then Exit For
    Next
    If I = j Then RtrimB = s$: Exit Function
    If I < 1 Then RtrimB = "": Exit Function
    RtrimB = LeftB$(s$, I)
End Function
Public Function SpaceB(n As Long) As String
If n = 1 Then
    SpaceB = ChrB$(32)
Else
    SpaceB = StrConv(space(n), vbFromUnicode)
End If
End Function

Public Function TrimZeroU(s$) As String
    Dim I As Long, j As Long, lim As Long
    lim = LenB(s$)
    If lim = 0 Then Exit Function
    j = 1
    TrimZeroU = space(Len(s$))
    If Left$(s$, j) = "-" Then Mid$(TrimZeroU, j, 1) = "-": j = j + 1
    I = j
    lim = lim - 1
    Do While I <= lim
    If Mid$(s$, I, 1) <> ChrB$(48) Then Exit Do
    I = I + 1
    Loop
    Mid$(TrimZeroU, j, lim - I + 2) = Mid$(s$, I)
    TrimZeroU = RTrim$(TrimZeroU)
End Function
