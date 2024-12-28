Attribute VB_Name = "Module13"
Option Explicit
'Option Compare Text 'Database for Access
'--------------------------------------------------------------------------------------------------------------
'https://cosxoverx.livejournal.com/47220.html
'Credit to Rebecca Gabriella's String Math Module (Big Integer Library) for VBA (Visual Basic for Applications)
' Minor edits made with comments and other.
'--------------------------------------------------------------------------------------------------------------

Private Type PartialDivideInfo
    Quotient As Integer
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

Public Function compare(sA As String, sB As String) As Integer
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns an integer that represents one of three states
    'sA > sB returns 1, sA < sB returns -1, and sA = sB returns 0
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim i As Integer, iA As Integer, iB As Integer
    
    'handle any early exits on basis of signs
    bAN = (Left(sA, 1) = "-")
    bBN = (Left(sB, 1) = "-")
    If bAN Then sA = Mid(sA, 2)
    If bBN Then sB = Mid(sB, 2)
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
    Dim crop As Long, lim As Long
    
    crop = 1
    lim = Len(sA)
    Do While crop <= lim
        If Mid$(sA, crop, 1) <> "0" Then Exit Do
        crop = crop + 1
    Loop
    sA = Mid$(sA, crop)
    crop = 1
    lim = Len(sB)
    Do While crop <= lim
       If Mid$(sB, crop, 1) <> "0" Then Exit Do
       crop = crop + 1
    Loop
    sB = Mid$(sB, crop)
    
    'then decide size first on basis of length
    If Len(sA) < Len(sB) Then
        compare = -1
    ElseIf Len(sA) > Len(sB) Then
        compare = 1
    Else 'unless they are the same length
        compare = 0
        'then check each digit by digit
        For i = 1 To Len(sA)
            iA = CInt(Mid(sA, i, 1))
            iB = CInt(Mid(sB, i, 1))
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

Public Function Add(sA As String, sB As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sum of sA and sB as string integer in Add()
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim iA As Integer, iB As Integer, iCarry As Integer
       
    'test for empty parameters
    If Len(sA) = 0 Or Len(sB) = 0 Then
        MyEr "Empty parameter in Add()", "Κενοί Παράμετροι στην Add()"
        Exit Function
    End If
        
    'handle some negative values with Subtract()
    bAN = (Left(sA, 1) = "-")
    bBN = (Left(sB, 1) = "-")
    If bAN Then sA = Mid(sA, 2)
    If bBN Then sB = Mid(sB, 2)
    If bAN And bBN Then 'both negative
        bRN = True      'set output reminder
    ElseIf bBN Then     'use subtraction
        Add = subtract(sA, sB)
        Exit Function
    ElseIf bAN Then     'use subtraction
        Add = subtract(sB, sA)
        Exit Function
    Else
        bRN = False
    End If
    
    'add column by column
    iA = Len(sA)
    iB = Len(sB)
    iCarry = 0
    Add = ""
    Do While iA > 0 And iB > 0
        iCarry = iCarry + CInt(Mid(sA, iA, 1)) + CInt(Mid(sB, iB, 1))
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
        iCarry = iCarry + CInt(Mid(sB, iB, 1))
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

Private Function RealMod(ByVal iA As Integer, ByVal iB As Integer) As Integer
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

Private Function RealDiv(ByVal iA As Integer, ByVal iB As Integer) As Integer
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
Public Function subtract(sA As String, sB As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sA minus sB as string integer in Subtract()
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim iA As Integer, iB As Integer, iComp As Integer
    
    'test for empty parameters
    If Len(sA) = 0 Or Len(sB) = 0 Then
        MyEr "Empty parameter in Subtract()", "Κενοί Παράμετροι στην Subtract()"
        Exit Function
    End If
        
    'handle some negative values with Add()
    bAN = (Left(sA, 1) = "-")
    bBN = (Left(sB, 1) = "-")
    If bAN Then sA = Mid(sA, 2)
    If bBN Then sB = Mid(sB, 2)
    If bAN And bBN Then
        bRN = True
    ElseIf bBN Then
        subtract = Add(sA, sB)
        Exit Function
    ElseIf bAN Then
        subtract = "-" + Add(sA, sB)
        Exit Function
    Else
        bRN = False
    End If
    
    'get biggest value into variable sA
    iComp = compare(sA, sB)
    If iComp = 0 Then     'parameters equal in size
        subtract = "0"
        Exit Function
    ElseIf iComp < 0 Then 'sA < sB
        subtract = sA     'so swop sA and sB
        sA = sB           'to ensure sA >= sB
        sB = subtract
        bRN = Not bRN     'and reverse output sign
    End If
    iA = Len(sA)          'recheck lengths
    iB = Len(sB)
    iComp = 0
    subtract = ""
        
    'subtract column by column
    Do While iA > 0 And iB > 0
        iComp = iComp + CInt(Mid(sA, iA, 1)) - CInt(Mid(sB, iB, 1))
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

Public Function multiply(sA As String, sB As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sA times sB as string integer in Multiply()
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN As Boolean, bBN As Boolean, bRN As Boolean
    Dim m() As Long, iCarry As Long
    Dim iAL As Integer, iBL As Integer, iA As Integer, iB As Integer
        
    'test for empty parameters
    If Len(sA) = 0 Or Len(sB) = 0 Then
        MyEr "Empty parameter in Multiply()", "Κενοί Παράμετροι στην Multiply()"
        Exit Function
    End If
        
    'handle any negative signs
    bAN = (Left(sA, 1) = "-")
    bBN = (Left(sB, 1) = "-")
    If bAN Then sA = Mid(sA, 2)
    If bBN Then sB = Mid(sB, 2)
    bRN = (bAN <> bBN)
    iAL = Len(sA)
    iBL = Len(sB)
    
    'perform long multiplication without carry in notional columns
    ReDim m(1 To (iAL + iBL - 1)) 'expected length of product
    For iA = 1 To iAL
        For iB = 1 To iBL
            m(iA + iB - 1) = m(iA + iB - 1) + CLng(Mid(sA, iAL - iA + 1, 1)) * CLng(Mid(sB, iBL - iB + 1, 1))
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

Private Function PartialDivide(sA As String, sB As String) As PartialDivideInfo
    'Called only by Divide() to assist in fitting trials for long division
    'All of Quotient, Subtrahend, and Remainder are returned as elements of type PartialDivideInfo
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
        
    For PartialDivide.Quotient = 9 To 1 Step -1                                'propose a divisor to fit
        PartialDivide.Subtrahend = multiply((sB), CStr(PartialDivide.Quotient))   'test by multiplying it out
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

Public Function divide(sA As String, sB As String) As String
    'Parameters are string integers of any length, for example "-345...", "973..."
    'Returns sA divided by sB as string integer in Divide()
    'The remainder is available as sLastRemainder at Module level
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
    
    Dim bAN  As Boolean, bBN As Boolean, bRN As Boolean
    Dim iC As Integer
    Dim s As String
    Dim d As PartialDivideInfo
    
    'test for empty parameters
    If Len(sA) = 0 Or Len(sB) = 0 Then
        MyEr "Empty parameter in Divide()", "Κενοί Παράμετροι στην Divide()"
        Exit Function
    End If
    
    bAN = (Left(sA, 1) = "-") 'true for neg
    bBN = (Left(sB, 1) = "-")
    If bAN Then sA = Mid(sA, 2) 'take two charas if neg
    If bBN Then sB = Mid(sB, 2)
    bRN = (bAN <> bBN)
    If compare(sB, "0") = 0 Then
        Err.Raise 11
        Exit Function
    ElseIf compare(sA, "0") = 0 Then
        divide = "0"
        sLastRemainder = "0"
        Exit Function
    End If
    iC = compare(sA, sB)
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
        d = PartialDivide(s, sB)   'find best fit
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

Public Function Modulus(sA As String, sB As String) As String
    divide sA, sB
    Modulus = sLastRemainder
End Function

Public Function BigIntFromString(sIn As String, iBaseIn As Integer) As String
    'Returns base10 integer string from sIn of different base (iBaseIn).
    'Example for sIn = "1A" and iBaseIn = 16, returns the base10 result 26.
    'Credit to Rebecca Gabriella's String Math Module with added edits
    'https://cosxoverx.livejournal.com/47220.html
            
    Dim bRN As Boolean
    Dim sBS As String
    Dim iP As Integer, iV As Integer
    
    'test for empty parameters
    If Len(sIn) = 0 Or iBaseIn = 0 Then
        MyEr "Bad parameter in BigIntFromString()", "Προβληματικοί παράμετροι στη BigIntFromString()"
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
    Dim sB As String
    Dim iV As Integer
    
    'test for empty parameters
    If Len(sIn) = 0 Or iBaseOut = 0 Then
        MyEr "Bad parameter in BigIntToString()", "Προβληματικοί παράμετροι στη BigIntToString()"
        Exit Function
    End If
    
    'handle negative signs
    If Left(sIn, 1) = "-" Then
        bRN = True
        sIn = Mid(sIn, 2)
    Else
        bRN = False
    End If
    sB = CStr(iBaseOut)
    
    BigIntToString = ""
    On Error GoTo 100
    Do While compare((sIn), "0") > 0
        sIn = divide(sIn, sB)
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

Function IntStrByExp(sA As String, iExp As Integer) As String
    'Returns integer string raised to exponent iExp as integer string
    'Assumes posiive exponent, and pos or neg string integer
    
    Dim bA As Boolean, bR As Boolean
    
    'check parameter
    If iExp < 0 Then
        MsgBox "Cannot handle negative powers yet"
        MyEr "Negative power in IntStrByExp()", "Αρνητική δύναμη στην IntStrByExp()"
        Exit Function
    End If
    
    'handle any negative signs
    bA = (Left(sA, 1) = "-")
    If bA Then sA = Mid(sA, 2) Else sA = Mid(sA, 1)
    If bA And RealMod(iExp, 2) <> 0 Then bR = True
    
    'run multiplication loop
    IntStrByExp = "1"
    Do Until iExp <= 0
        DoEvents 'permits break key use
        IntStrByExp = multiply(IntStrByExp, sA)
        iExp = iExp - 1
    Loop

    'remove any leading zeros
    Do While Len(IntStrByExp) > 1 And Left(IntStrByExp, 1) = "0"
        IntStrByExp = Mid(IntStrByExp, 2)
    Loop
    
    'decide on any signs
    If IntStrByExp <> "0" And bR Then
       IntStrByExp = "-" & IntStrByExp
    End If

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
