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
Private Const Alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Property Get LastRemainder()
    LastRemainder = CVar(sLastRemainder)
End Property
Sub TestMultAndDiv()
    'Run this to test multiplication and division with integer strings
    'Open immediate window in View or with ctrl-g to see results
    
    Dim sP1 As String, sP2 As String, sRes1 As String, sRes2 As String
    
    sP1 = "100"
    sP2 = "2"             '33 digits and also prime


    sRes1 = Modulus(sP1, sP2)
    
    Debug.Print "Modulus : " & sRes1

End Sub

Public Function compare(sA As String, sb As String) As Integer
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns an integer that represents one of three states
    'sA > sB returns 1, sA < sB returns -1, and sA = sB returns 0
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim i As Long, iA As Long, iB As Long
    
    'handle any early exits on basis of signs
    bAN = (Left(sA, 1) = "-")
    bBN = (Left(sb, 1) = "-")
    If bAN Then sA = Mid(sA, 2)
    If bBN Then sb = Mid(sb, 2)
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
    lim = Len(sA)
    Do While CROP <= lim
        If Mid$(sA, CROP, 1) <> "0" Then Exit Do
        CROP = CROP + 1
    Loop
    sA = Mid$(sA, CROP)
    CROP = 1
    lim = Len(sb)
    Do While CROP <= lim
       If Mid$(sb, CROP, 1) <> "0" Then Exit Do
       CROP = CROP + 1
    Loop
    sb = Mid$(sb, CROP)
    
    'then decide size first on basis of length
    If Len(sA) < Len(sb) Then
        compare = -1
    ElseIf Len(sA) > Len(sb) Then
        compare = 1
    Else 'unless they are the same length
        compare = 0
        'then check each digit by digit
        For i = 1 To Len(sA)
            iA = CInt(Mid(sA, i, 1))
            iB = CInt(Mid(sb, i, 1))
            If iA < iB Then
                compare = -1
                Exit For
            ElseIf iA > iB Then
                compare = 1
                Exit For
            Else 'defaults zero
            End If
        Next i
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
    Dim iA As Integer, iB As Integer, iCarry As Integer
       
    'test for empty parameters
    If Len(sA) = 0 Or Len(sb) = 0 Then
        MyEr "Empty parameter in Add", "Κενοί Παράμετροι στην Add"
        Exit Function
    End If
        
    'handle some negative values with Subtract()
    bAN = (Left(sA, 1) = "-")
    bBN = (Left(sb, 1) = "-")
    If bAN Then sA = Mid(sA, 2)
    If bBN Then sb = Mid(sb, 2)
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
    
    'add column by column
    iA = Len(sA)
    iB = Len(sb)
    iCarry = 0
    Add = ""
    Do While iA > 0 And iB > 0
        iCarry = iCarry + CInt(Mid(sA, iA, 1)) + CInt(Mid(sb, iB, 1))
        Add = CStr(iCarry Mod 10) + Add
        iCarry = iCarry \ 10
        iA = iA - 1
        iB = iB - 1
    Loop
    
    'Assuming param sA is longer
    Do While iA > 0
        iCarry = iCarry + CInt(Mid(sA, iA, 1))
        Add = CStr(iCarry Mod 10) + Add
        iCarry = iCarry \ 10
        iA = iA - 1
    Loop
    'Assuming param sB is longer
    Do While iB > 0
        iCarry = iCarry + CInt(Mid(sb, iB, 1))
        Add = CStr(iCarry Mod 10) + Add
        iCarry = iCarry \ 10
        iB = iB - 1
    Loop
    Add = CStr(iCarry) + Add
    
    'remove any leading zeros
    Do While Len(Add) > 1 And Left(Add, 1) = "0"
        Add = Mid(Add, 2)
    Loop
    
    'decide about any negative signs
    If Add <> "0" And bRN Then
        Add = "-" + Add
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
    If Len(sA) = 0 Or Len(sb) = 0 Then
        MyEr "Empty parameter in Subtract", "Κενοί Παράμετροι στην Subtract"
        Exit Function
    End If
        
    'handle some negative values with Add()
    bAN = (Left(sA, 1) = "-")
    bBN = (Left(sb, 1) = "-")
    If bAN Then sA = Mid(sA, 2)
    If bBN Then sb = Mid(sb, 2)
    If bAN And bBN Then
        bRN = True
    ElseIf bBN Then
        subtract = Add(sA, sb)
        Exit Function
    ElseIf bAN Then
        subtract = "-" + Add(sA, sb)
        Exit Function
    Else
        bRN = False
    End If
    
    'get biggest value into variable sA
    iComp = compare(sA, sb)
    If iComp = 0 Then     'parameters equal in size
        subtract = "0"
        Exit Function
    ElseIf iComp < 0 Then 'sA < sB
        subtract = sA     'so swop sA and sB
        sA = sb           'to ensure sA >= sB
        sb = subtract
        bRN = Not bRN     'and reverse output sign
    End If
    iA = Len(sA)          'recheck lengths
    iB = Len(sb)
    iComp = 0
    subtract = ""
        
    'subtract column by column
    Do While iA > 0 And iB > 0
        iComp = iComp + CInt(Mid(sA, iA, 1)) - CInt(Mid(sb, iB, 1))
        subtract = CStr(RealMod(iComp, 10)) + subtract
        iComp = RealDiv(iComp, 10)
        iA = iA - 1
        iB = iB - 1
    Loop
    'then assuming param sA is longer
    Do While iA > 0
        iComp = iComp + CInt(Mid(sA, iA, 1))
        subtract = CStr(RealMod(iComp, 10)) + subtract
        iComp = RealDiv(iComp, 10)
        iA = iA - 1
    Loop
    
    'remove any leading zeros from result
    Do While Len(subtract) > 1 And Left(subtract, 1) = "0"
        subtract = Mid(subtract, 2)
    Loop
    
    'decide about any negative signs
    If subtract <> "0" And bRN Then
        subtract = "-" + subtract
    End If

End Function

Public Function multiply(sA As String, sb As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sA times sB as string integer in Multiply()
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim m() As Long, iCarry As Long
    Dim iAL As Long, iBL As Long, iA As Long, iB As Long
        
    'test for empty parameters
    If Len(sA) = 0 Or Len(sb) = 0 Then
        MyEr "Empty parameter in Multiply", "Κενοί Παράμετροι στην Multiply"
        Exit Function
    End If
        
    'handle any negative signs
    bAN = (Left(sA, 1) = "-")
    bBN = (Left(sb, 1) = "-")
    If bAN Then sA = Mid(sA, 2)
    If bBN Then sb = Mid(sb, 2)
    bRN = (bAN <> bBN)
    iAL = Len(sA)
    iBL = Len(sb)
    
    'perform long multiplication without carry in notional columns
    ReDim m(1 To (iAL + iBL - 1)) 'expected length of product
    For iA = 1 To iAL
        For iB = 1 To iBL
            m(iA + iB - 1) = m(iA + iB - 1) + CLng(Mid(sA, iAL - iA + 1, 1)) * CLng(Mid(sb, iBL - iB + 1, 1))
        Next iB
    Next iA
    iCarry = 0
    multiply = ""
    
    'add up column results with carry
    For iA = 1 To iAL + iBL - 1
        iCarry = iCarry + m(iA)
        multiply = CStr(iCarry Mod 10) + multiply
        iCarry = iCarry \ 10
    Next iA
    multiply = CStr(iCarry) + multiply
    
    'remove any leading zeros
    Do While Len(multiply) > 1 And Left(multiply, 1) = "0"
        multiply = Mid(multiply, 2)
    Loop
    
    'decide about any negative signs
    If multiply <> "0" And bRN Then
        multiply = "-" + multiply
    End If

End Function

Private Function PartialDivide(sA As String, sb As String) As PartialDivideInfo
    'Called only by Divide() to assist in fitting trials for long division
    'All of Quotient, Subtrahend, and Remainder are returned as elements of type PartialDivideInfo
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
        
    For PartialDivide.Quotient = 9 To 1 Step -1                                'propose a divisor to fit
        PartialDivide.Subtrahend = multiply((sb), CStr(PartialDivide.Quotient))   'test by multiplying it out
        If compare(PartialDivide.Subtrahend, (sA)) <= 0 Then                      'best fit found
            PartialDivide.Remainder = subtract((sA), (PartialDivide.Subtrahend))   'get remainder
            Exit Function                                                      'exit with best fit details
        End If
    Next PartialDivide.Quotient
    
    'no fit found, divisor too big
    PartialDivide.Quotient = 0
    PartialDivide.Subtrahend = "0"
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
    Dim d As PartialDivideInfo
    
    'test for empty parameters
    If Len(sA) = 0 Or Len(sb) = 0 Then
        MyEr "Empty parameter in Divide", "Κενοί Παράμετροι στην Divide"
        Exit Function
    End If
    
    bAN = (Left(sA, 1) = "-") 'true for neg
    bBN = (Left(sb, 1) = "-")
    If bAN Then sA = Mid(sA, 2) 'take two charas if neg
    If bBN Then sb = Mid(sb, 2)
    bRN = (bAN <> bBN)
    If compare(sb, "0") = 0 Then
        Err.Raise 11
        Exit Function
    ElseIf compare(sA, "0") = 0 Then
        divide = "0"
        sLastRemainder = "0"
        Exit Function
    End If
    iC = compare(sA, sb)
    If iC < 0 Then
        divide = "0"
        sLastRemainder = sA
        Exit Function
    ElseIf iC = 0 Then
        If bRN Then
            divide = "-1"
        Else
            divide = "1"
        End If
        sLastRemainder = "0"
        Exit Function
    End If
    divide = ""
    s = ""
    
    'Long division method
    For iC = 1 To Len(sA)
        'take increasing number of digits
        s = s + Mid(sA, iC, 1)
        d = PartialDivide(s, sb)   'find best fit
        divide = divide + CStr(d.Quotient)
        s = d.Remainder
    Next iC
    
    'remove any leading zeros
    Do While Len(divide) > 1 And Left(divide, 1) = "0"
        divide = Mid(divide, 2)
    Loop
    
    'decide about the signs
    If divide <> "0" And bRN Then
        divide = "-" + divide
    End If
    
    sLastRemainder = s 'string integer remainder

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
    If Len(sIn) = 0 Or iBaseIn = 0 Then
        MyEr "Bad parameter in BigIntFromString", "Προβληματικοί παράμετροι στη BigIntFromString"
        Exit Function
    End If
        
    'handle negative signs
    If Left(sIn, 1) = "-" Then
        bRN = True
        sIn = Mid(sIn, 2)
    Else
        bRN = False
    End If
    sBS = CStr(iBaseIn)
    
    BigIntFromString = "0"
    For iP = 1 To Len(sIn)
        'use constant list position and base for conversion
        iV = InStr(Alphabet, UCase(Mid(sIn, iP, 1)))
        If iV > 0 Then 'accumulate
            BigIntFromString = multiply(BigIntFromString, sBS)
            BigIntFromString = Add(BigIntFromString, CStr(iV - 1))
        End If
    Next iP
    
    'decide on any negative signs
    If bRN Then
        BigIntFromString = "-" + BigIntFromString
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
    If Len(sIn) = 0 Or iBaseOut = 0 Then
        MyEr "Bad parameter in BigIntToString", "Προβληματικοί παράμετροι στη BigIntToString"
        Exit Function
    End If
    
    'handle negative signs
    If Left(sIn, 1) = "-" Then
        bRN = True
        sIn = Mid(sIn, 2)
    Else
        bRN = False
    End If
    sb = CStr(iBaseOut)
    
    BigIntToString = ""
    On Error GoTo 100
    Do While compare((sIn), "0") > 0
        sIn = divide(sIn, sb)
        iV = CInt(LastModulus())
        'locates appropriate alphabet character
        BigIntToString = Mid(Alphabet, iV + 1, 1) + BigIntToString
    Loop
    
    'decide on any negative signs
    If BigIntToString = "" Then
        BigIntToString = "0"
    ElseIf BigIntToString <> "0" And bRN Then
        BigIntToString = "-" + BigIntToString
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
    If Left$(sExp, 1) = "-" Then
        MyEr "Negative power in IntPower", "Αρνητική δύναμη στην IntPower"
        Exit Function
    End If
    
    Dim VarZ
    
    
    'handle any negative signs
    ba = (Left(sA, 1) = "-")
    If ba Then sA = Mid(sA, 2) 'Else sA = Mid(sA, 1)
    If Len(sExp) > 4 Then
        MyEr "too big exponent", "υπερβολικά μεγάλη δύναμη"
        IntStrByExp = "1"
        Exit Function
    End If
    If ba And val(Right$(sExp, 1)) Mod 2 <> 0 Then br = True
    VarZ = CDec(sExp)
    'run multiplication loop
    IntStrByExp = "1"
    Do Until VarZ = 0
        IntStrByExp = multiply(IntStrByExp, sA)
        VarZ = VarZ - 1
    Loop

    'remove any leading zeros
    IntStrByExp = TrimZero(IntStrByExp)
    
    'decide on any signs
    If IntStrByExp <> "0" And br Then
       IntStrByExp = "-" & IntStrByExp
    End If

End Function
Function IsProbablyPrime(sA As String, k As Integer) As Boolean
    If Len(sA) < 1 Then Exit Function
    If Left$(sA, 1) = "-" Then
        MyEr "Negative Prime not exist", "Αρνητικός πρώτος δεν υπάρχει"
        Exit Function
    End If
    If sA = "2" Then IsProbablyPrime = True: Exit Function
    If val(Right$(sA, 1)) Mod 2 = 0 Then Exit Function
    If sA = "1" Then Exit Function
    
    Dim nn As String, d As String, s As Long, z As Long
    nn = subtract(sA, "1")
    d = nn
    While compare(Modulus((d), "2"), "0") = 0
        s = s + 1
        d = divide(d, "2")
    Wend
    z = Len(sA)
    Dim a As String, x As String, i As Integer, j As Integer
    
    IsProbablyPrime = True
    For i = 1 To k
        Do
            a = ""
            For j = 1 To Len(sA)
                a = a + Chr$(47 + Int(10 * RndM(rndbase) + 1))
            Next
        Loop Until compare(nn, a) = 1 And compare(a, "1") > -1
        x = modpow(a, (d), (sA))
        If compare(x, "1") <> 0 Then ' continue
            If compare(x, nn) <> 0 Then ' continue
                For j = 1 To s
                    x = modpow(x, "2", (sA))
                    If compare(x, "1") = 0 Then IsProbablyPrime = False: Exit Function
                    If compare(x, nn) = 0 Then Exit For
                Next
            End If
            If compare(x, nn) <> 0 Then IsProbablyPrime = False: Exit Function
        End If
    Next
End Function

Public Function modpow(sBase As String, sExp As String, sMod As String) As String
    If Left$(sExp, 1) = "-" Then
        MyEr "Negative power in modpow", "Αρνητική δύναμη στην modpow"
        Exit Function
    End If
    If Left$(sMod, 1) = "-" Or sMod = "0" Then
        MyEr "Zero or negative Modules in modpow", "Μηδενικό ή Αρνητικό Μέτρο στην modpow"
        Exit Function
    End If
    Dim br As Boolean, ba As Boolean
    ba = (Left(sBase, 1) = "-")
    If ba Then sBase = Mid(sBase, 2)
    If ba And AscW(Right$(sExp, 1)) Mod 2 <> 0 Then br = True
    modpow = "1"
    Do While sExp <> "0"
        If AscW(Right$(sExp, 1)) Mod 2 = 1 Then
            modpow = Module13.Modulus(Module13.multiply(modpow, (sBase)), (sMod))
        End If
        sExp = divide(sExp, "2")
        sBase = Modulus(multiply(sBase, (sBase)), sMod)
    Loop
    
    If modpow <> "0" And br Then
       modpow = "-" & modpow
    End If
End Function
Public Function IntSqr(sA As String) As String
    If Left$(sA, 1) = "-" Or sA = "0" Then
        MyEr "Zero or negative paramter for integer Square Root", "Μηδενική ή Αρνητική παράμετρος για ακέραια τετραγωνική ρίζα"
        Exit Function
    End If
    Dim q As String, r As String, t As String, z As String
    z = sA
    r = "0"
    q = "1"
    Do
    q = multiply(q, "4")
    Loop Until compare((q), (sA)) = 1
    Do
        If compare((q), "1") < 1 Then Exit Do
        q = divide(q, "4")
        t = subtract(subtract((z), (r)), (q))
        r = divide(r, "2")
        If compare((t), "-1") > -1 Then
            SwapStrings z, t
            r = Add(r, (q))
        End If
    Loop
    IntSqr = r
End Function
Private Function IsPrime(sA As String) As Boolean
    ' works but not used - use IsProbablyPrime()
    If Len(sA) < 1 Then Exit Function
    If Left$(sA, 1) = "-" Then
        MyEr "Negative Prime not exist", "Αρνητικός πρώτος δεν υπάρχει"
        Exit Function
    End If
    Dim d As String
    
    If sA = "2" Then IsPrime = True: Exit Function
    If val(Right$(sA, 1)) Mod 2 = 0 Then Exit Function
    If sA = "1" Then Exit Function
    If sA = "3" Then IsPrime = True: Exit Function
    If sA = "5" Then IsPrime = True: Exit Function
    If compare(Modulus((sA), "3"), "0") = 0 Then Exit Function
    Dim x1 As String
    x1 = IntSqr(sA)
    d = "5"
    Do
        If compare(Modulus((sA), (d)), "0") = 0 Then Exit Do
        d = Add("2", d)
        If compare((d), (x1)) = 1 Then IsPrime = True: Exit Function
        If compare(Modulus((sA), (d)), "0") = 0 Then Exit Do
        d = Add("4", d)
        If compare((d), (x1)) = 1 Then IsPrime = True: Exit Function
    Loop
End Function
Public Function CreateBigInteger(s$, Optional basenum) As BigInteger
    Set CreateBigInteger = New BigInteger
    s$ = TrimZero(s$)
    If IsMissing(basenum) Then
        If TestNumber(s$) Then
            CreateBigInteger.Load s$, 10
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
    Dim i As Long, lim As Long, ss As String
    
    ss = Mid$(Alphabet, 1, b)
    lim = Len(s$)
    If lim = 0 Then Exit Function
    i = 1
    If Left$(s$, 1) = "-" Then i = i + 1: If lim = 1 Then Exit Function
    Do While i <= lim
    If InStr(ss, Mid$(s$, i, 1)) = 0 Then Exit Function
    i = i + 1
    Loop
    TestNumberOnBase = True
End Function
Public Function TestNumber(s$) As Boolean
    Dim i As Long, lim As Long
    lim = Len(s$)
    If lim = 0 Then Exit Function
    i = 1
    If Left$(s$, 1) = "-" Then i = i + 1: If lim = 1 Then Exit Function
    Do While i <= lim
    If InStr("0123456789", Mid$(s$, i, 1)) = 0 Then Exit Function
    i = i + 1
    Loop
    TestNumber = True
End Function
Public Function TrimZero(s$) As String
    Dim i As Long, j As Long, lim As Long
    lim = Len(s$)
    If lim = 0 Then Exit Function
    j = 1
    TrimZero = space(Len(s$))
    If Left$(s$, j) = "-" Then Mid$(TrimZero, j, 1) = "-": j = j + 1
    i = j
    lim = lim - 1
    Do While i <= lim
    If Mid$(s$, i, 1) <> "0" Then Exit Do
    i = i + 1
    Loop
    Mid$(TrimZero, j, lim - i + 2) = Mid$(s$, i)
    TrimZero = RTrim$(TrimZero)
End Function
