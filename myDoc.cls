VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event MayQuit(Yes)
Public Event BreakLine(data$, datanext$)
Public Event Lastline()
Public Event ColorDataEvent(thispara As Long, data$, feedback As Long)
Private lastlinelastchar As Long
Private lastcolordatacur As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameW" (ByVal lpszLongPath As Long, _
ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar& Lib "kernel32" (ByVal CodePage&, ByVal dwFlags&, MultiBytes As Any, ByVal cBytes&, ByVal pWideChars&, ByVal cWideChars&)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Final$
Public MaxParaLength As Long
'Dim mPosId As Long
'Dim mCurrentParagraph As Long
' each paragraph has a number
Private Type Paragraph
  content() As String
  colordata() As Long
  startfrom As Long
lines As Long  ' computed by the width of ruler...
End Type
Public LfLead As Boolean
Public CrTail As Boolean
Dim DocParaBack() As Long, DocParaNext() As Long, para() As Paragraph
Dim neopara As Long, paramax As Long, docmax As Long, docpara As Long
Dim delDocPara As Long
Public MarkParagraphID As Long  ' this is the direct paranum non indirect
Dim mKeystart As String
Public BackMove As Boolean
Public LastSelStart As Long
Public ListLoadedType As Long
Private Declare Function CompareString Lib "kernel32" Alias "CompareStringW" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long) As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, retval As Integer)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Integer)
Private words$, unique As Long
Private p3 As Long
Dim mlcid As Long
Dim sortmlcid As Long
Private pos() As Long, slen() As Long, hit() As Long
Dim Buffer() As Byte, bufOffset As Long, bufsize As Long, SeekBuf As Long
Public Busy As Boolean
Private mSortstyle As VbCompareMethod
Public ColorEvent As Boolean, ForEditBox As Boolean
Public Switchon As Boolean
Public NoIgnoreCase As Boolean
Function IsLast(ByVal ParaNum As Long) As Boolean
' for MarkParagraphID
IsLast = (DocParaNext(ParaNum) = 0)
End Function
Function IsEmpty() As Boolean

If DocParaNext(0) = 0 And DocParaBack(0) = 0 Then IsEmpty = True
End Function
Function FindNewPara() As Long
If delDocPara > 0 Then
FindNewPara = delDocPara
delDocPara = DocParaNext(delDocPara)
Else
docpara = docpara + 1
If docpara > docmax Then
docmax = docmax * 2
ReDim Preserve DocParaBack(0 To docmax)
ReDim Preserve DocParaNext(0 To docmax)
ReDim Preserve para(1 To docmax)
End If
FindNewPara = docpara
If Not para(docpara).startfrom = 0 Then Err.Raise 50501, , "internal problem in document"
End If
End Function
Sub InsertDocPara(ByVal ParaNum As Long, ByVal Previus As Long, ByVal NextPara As Long)
 'ok
    DocParaNext(ParaNum) = NextPara
    DocParaBack(ParaNum) = Previus
    DocParaNext(Previus) = ParaNum
    DocParaBack(NextPara) = ParaNum
End Sub
Sub UnlinkDocPara(ParaNum As Long)
   DocParaNext(DocParaBack(ParaNum)) = DocParaNext(ParaNum)
    DocParaBack(DocParaNext(ParaNum)) = DocParaBack(ParaNum)
    
    para(ParaNum).startfrom = 0
End Sub
Sub LinkDocParaAfter(ParaNum As Long, thisparanum As Long)
    DocParaBack(DocParaNext(thisparanum)) = ParaNum
    DocParaNext(ParaNum) = DocParaNext(thisparanum)
    DocParaBack(ParaNum) = thisparanum
    DocParaNext(thisparanum) = ParaNum
    If thisparanum > 0 Then para(ParaNum).startfrom = para(thisparanum).startfrom
End Sub

Sub RemoveDocParaIndex(ParaNum As Long)
    DocParaNext(DocParaBack(ParaNum)) = DocParaNext(ParaNum)
    DocParaBack(DocParaNext(ParaNum)) = DocParaBack(ParaNum)
    para(ParaNum).lines = 0   ' clear' also can be a flag to show that this is deleted
    ReDim para(ParaNum).content(1 To 1)
    para(ParaNum).content(1) = vbNullString
    Dim emp() As Long
    para(ParaNum).colordata() = emp()
    DocParaNext(ParaNum) = delDocPara
    para(ParaNum).startfrom = 0
    delDocPara = ParaNum
End Sub
Function InvalidPara(ParaNum) As Boolean
If ParaNum < 1 Or ParaNum > docpara Then
InvalidPara = True
Else
InvalidPara = para(ParaNum).lines = 0
End If
End Function
Sub WritePara(ParaNum As Long, ByVal data$)
    Dim NeoData$, neolines As Long, thatmax As Long, lookthis As Long
    Do
    neolines = para(ParaNum).lines + 1
    If neolines = 1 Then
    thatmax = 10
    ReDim Preserve para(ParaNum).content(1 To thatmax)
        ReDim para(ParaNum).colordata(1)
          
    End If
    thatmax = UBound(para(ParaNum).content())
    If thatmax < neolines Then
    thatmax = thatmax * 2
    ReDim Preserve para(ParaNum).content(1 To thatmax)
    ReDim para(ParaNum).colordata(1)
    
    End If
    If False Then
    
   End If
    
    RaiseEvent BreakLine(data$, NeoData$)  ' no method inside this class
    If NeoData$ = vbNullString Then NeoData$ = data$: data$ = vbNullString
    para(ParaNum).content(neolines) = NeoData$
    para(ParaNum).lines = neolines
    NeoData$ = vbNullString
    Loop Until data$ = vbNullString
End Sub
Sub WriteParaOneline(ParaNum As Long, ByVal NeoData$)
    
    ReDim para(ParaNum).content(1 To 1)
    para(ParaNum).content(1) = NeoData$
    para(ParaNum).lines = 1
    ReDim para(ParaNum).colordata(1)
        para(ParaNum).startfrom = 0
        
End Sub
Public Sub ResetColor(ParaNum)
On Error Resume Next
ReDim para(ParaNum).colordata(2)
End Sub
Public Sub ColorThis(mNum As Long, Optional Count As Long = 0, Optional LookDown As Boolean = True)
If ForEditBox Then
    ColorThis3 mNum, Count, LookDown
Else
    ColorThis2 mNum, Count, LookDown
End If
End Sub
Private Sub ColorThis3(mNum As Long, Optional Count As Long = 0, Optional LookDown As Boolean = True)
Dim docback As Long, lookthis As Long, data$, oldstart As Long, lenfinal As Long, closed As Long, lastnum As Long
lenfinal = Len(Final$)
If lenfinal = 0 Then Final$ = ";": lenfinal = 1
    If mNum = 0 Then Exit Sub
    data$ = TextParagraph(mNum)
    If Trim$(data$) = vbNullString And Switchon Then
        lookthis = DocParaBack(mNum)
        If lookthis > 0 Then
                    If para(lookthis).startfrom > 0 And UBound(para(lookthis).colordata()) = 2 Then
                                    ReDim para(mNum).colordata(2)
                                    para(mNum).colordata(0) = 15
                                    para(mNum).colordata(1) = 1
                                    If para(mNum).startfrom <> mNum Then
                                            para(mNum).startfrom = para(lookthis).startfrom
                                    End If
                    End If
        End If
        GoTo finish
    End If
    docback = DocParaBack(mNum)
    If docback = 0 Then
            docback = mNum
again:
            oldstart = para(mNum).startfrom
            data$ = TextParagraph(mNum)
            
again55:
            RaiseEvent ColorDataEvent(mNum, data$, lookthis)
            If UBound(para(mNum).colordata()) = 0 Then
                  ReDim para(mNum).colordata(1)
                            para(mNum).colordata(0) = 14
                            para(mNum).colordata(1) = 1
            ElseIf para(mNum).colordata(1) = 0 Then
                             para(mNum).colordata(0) = 14
                            para(mNum).colordata(1) = 1
            End If
            If Final$ = "}" Then
                        closed = -1
                       Count = countbrakcet(data$, closed, Count)
                       End If
            If lookthis = 15 Then
                    
                        para(mNum).startfrom = mNum
                        docback = mNum
            Else
                     para(mNum).startfrom = 0
                        If oldstart <> 0 Then
                                 lookthis = DocParaNext(mNum)
                                Do While lookthis > 0
                                        If para(lookthis).startfrom <> oldstart Then mNum = lookthis: GoTo finish
                                        para(lookthis).startfrom = 0
                                        ColorThis1 lookthis, 0
                                        
                                Loop
                        Else
                        lookthis = mNum
                        End If
                        mNum = lookthis
                        GoTo finish
            End If
again2:
            lookthis = DocParaNext(mNum)
            Do While lookthis > 0
                        data$ = TextParagraph(lookthis)
                        If Final$ = "}" Then
                        closed = -1
                        Count = countbrakcet(data$, closed, Count)
                        If closed > -1 Then
                        

                                       

                                        If closed < Len(data$) Then
                                           oldstart = mNum
                                         ReDim para(lookthis).colordata(1)
                                        para(lookthis).colordata(0) = 15
                                        para(lookthis).colordata(1) = 1
                                         mNum = lookthis
                                            lookthis = closed
                                            GoTo again55
                                        End If
                                         ReDim para(lookthis).colordata(3)
                                        para(lookthis).colordata(0) = 15
                                        para(lookthis).colordata(1) = 1
                                        para(lookthis).colordata(2) = 14
                                        para(lookthis).colordata(3) = closed + 1
                                      ' para(lookthis).startfrom = mNum
                                         mNum = lookthis
                                        GoTo finish
                        Else
                        GoTo conthere
                        End If
                        ElseIf Left$(Trim$(data$), lenfinal) = Final$ Or Right$(RTrim(data$), lenfinal) = Final$ Then
                                        ReDim para(lookthis).colordata(4)
                                        para(lookthis).colordata(0) = 15
                                        para(lookthis).colordata(1) = 1
                                        para(lookthis).colordata(2) = 14
                                        para(lookthis).colordata(3) = rinstr(data$, Final$) + lenfinal
                                        para(lookthis).startfrom = para(docback).startfrom
                                        mNum = lookthis
                                        lookthis = DocParaNext(mNum)
                                        Do While lookthis > 0
                                            If para(lookthis).startfrom <> para(docback).startfrom Then GoTo finish
                                            para(lookthis).startfrom = 0
                                            ColorThis1 lookthis, Count
                                        Loop
                                        mNum = lookthis
                                        GoTo finish
                    Else
conthere:
                    If data$ <> "" Then
                            ReDim para(lookthis).colordata(3)
                            para(lookthis).colordata(0) = 15
                            para(lookthis).colordata(1) = 1
                            
                               para(lookthis).colordata(2) = 15
                            para(lookthis).colordata(3) = Len(data$) + 1
                            Else
                            ReDim para(lookthis).colordata(1)
                            para(lookthis).colordata(0) = 15
                            para(lookthis).colordata(1) = 1
                            End If
                            para(lookthis).startfrom = para(docback).startfrom
                    End If
                    lookthis = DocParaNext(lookthis)
            Loop
                    mNum = lookthis
                    GoTo finish
    ElseIf para(docback).startfrom = 0 Then
    lookthis = 0
      GoTo again
    Else   ' no is  a part
    If UBound(para(docback).colordata()) = 4 Then
        lookthis = 0
    GoTo again
    Else
    lastnum = mNum
        mNum = docback
        If para(docback).startfrom = docback Then
        lookthis = 0: GoTo again
        End If
        If Final = "}" Then
        If para(docback).startfrom > 0 Then mNum = para(docback).startfrom: Count = 0
        lookthis = 0: GoTo again
        End If
         GoTo again2
    End If
    End If
finish:
If mNum = 0 Then Exit Sub
    mNum = DocParaNext(mNum)
End Sub

Private Sub ColorThis2(mNum As Long, Optional Count As Long = 0, Optional LookDown As Boolean = True)
Dim mNum1 As Long, laststart As Long
If mNum = 0 Then Exit Sub
If Count = 0 Then laststart = para(mNum).startfrom
If ColorEvent Then
mNum1 = mNum
ColorThis1 mNum, Count, True

If Not mNum1 = mNum And mNum <> 0 Then

If laststart <> para(mNum).startfrom Or para(mNum1).colordata(UBound(para(mNum1).colordata) - 1) <> 15 Then
If Count = 0 And laststart <> 0 Then

laststart = para(mNum1).startfrom
mNum1 = DocParaNext(mNum1)
Do While mNum1 <> 0
If para(mNum1).startfrom = laststart Then
para(mNum1).startfrom = 0
ColorThis1 mNum1, (0), True
mNum1 = DocParaNext(mNum1)
Else
Exit Do
End If
Loop

End If



End If
End If
End If
mNum = DocParaNext(mNum)

End Sub
Private Sub ColorThis1(mNum As Long, Count As Long, Optional ByPass As Boolean = False)
Dim lookthis As Long, data$, mNum1 As Long
If para(mNum).startfrom <> 0 Then

    mNum1 = para(mNum).startfrom
    
    mNum = mNum1
data$ = TextParagraph(mNum1)

RaiseEvent ColorDataEvent(mNum1, data$, lookthis)
If mNum1 = mNum Then
If lookthis = 0 Then
If para(mNum).startfrom = 0 Then
lookthis = mNum
mNum = DocParaNext(mNum1)
Else
lookthis = para(mNum).startfrom
End If

mNum1 = mNum
Do While mNum1 <> 0
    If para(mNum1).startfrom = lookthis Then
    ReDim para(mNum1).colordata(1)
    Else
    If Switchon Then Exit Do   ' maybe is wrong
    End If
    
     para(mNum1).startfrom = 0
    mNum1 = DocParaNext(mNum1)
    
Loop
mNum = mNum1
End If

Count = -1
Else

If lookthis = 0 And (Switchon Or ByPass) Then
lookthis = para(mNum1).startfrom
mNum1 = DocParaNext(mNum1)
Do While mNum1 <> 0
    If para(mNum1).startfrom = lookthis Then ReDim para(mNum1).colordata(1)
     para(mNum1).startfrom = 0
    mNum1 = DocParaNext(mNum1)
    
Loop
Count = -1
End If


End If

Else

data$ = TextParagraph(mNum)
RaiseEvent ColorDataEvent(mNum, data$, lookthis)
Count = 0
End If
End Sub
 

Sub PlaceColorData(ParaNum As Long, thatcolordata() As Long)
    para(ParaNum).colordata = thatcolordata()
End Sub
Sub ReWritePara(ParaNum As Long, ByVal data$, Optional skipcolor As Boolean = False)
    Dim NeoData$, neolines As Long, thatmax As Long, lookthis As Long
    If ParaNum = 0 Then Exit Sub
    para(ParaNum).lines = 0  ' clear
    ReDim para(ParaNum).content(1 To 1)
    Do
    neolines = para(ParaNum).lines + 1
    If neolines = 1 Then
    thatmax = 10
    ReDim Preserve para(ParaNum).content(1 To thatmax)
    End If
    thatmax = UBound(para(ParaNum).content())
    If thatmax < neolines Then
    thatmax = thatmax * 2
    ReDim Preserve para(ParaNum).content(1 To thatmax)
    End If
    RaiseEvent BreakLine(data$, NeoData$)  ' no method inside this class
    If NeoData$ = vbNullString Then NeoData$ = data$: data$ = vbNullString
    If Len(NeoData$) > 1000 And LenB(data$) = 0 Then
    data$ = Mid$(NeoData$, 1001)
    NeoData$ = Left$(NeoData$, 1000)
    End If
    
    para(ParaNum).content(neolines) = NeoData$
    para(ParaNum).lines = neolines
    NeoData$ = vbNullString
    Loop Until data$ = vbNullString
    If ColorEvent And Not skipcolor Then
     ColorThis (ParaNum)
    End If
End Sub
Sub AppendParagraphOneLine(myparagraph$)
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    WriteParaOneline getAparaNum, myparagraph$
    If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
End Sub
Sub InsertParagraphOneLine(myparagraph$)
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    WriteParaOneline getAparaNum, myparagraph$
    If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        InsertDocPara getAparaNum, MarkParagraphID, DocParaNext(MarkParagraphID)   ' DocParaBack(0) link to last
    End If
    MarkParagraphID = getAparaNum
    End Sub

Function AppendParagraph(myparagraph$) As Long
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    
    ' getAparaNum is paragraph absolute address
    If docpara = 0 Then
    WritePara getAparaNum, myparagraph$
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
    WritePara getAparaNum, myparagraph$
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    AppendParagraph = getAparaNum
End Function
Function AppendParagraphPlusColor(myparagraph$, colordata() As Long, Optional Offset As Long = 0) As Long
    Dim getAparaNum As Long, mColor() As Long
    mColor() = colordata()
    getAparaNum = FindNewPara
    ' getAparaNum is paragraph absolute address
    If docpara = 0 Then
    WritePara getAparaNum, myparagraph$
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
    WritePara getAparaNum, myparagraph$
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    If Offset > 0 Then mColor() = Relocate(mColor(), Offset)
    para(getAparaNum).colordata() = mColor()
    AppendParagraphPlusColor = getAparaNum
End Function

Function InsertAfterParagraph(ParaNum As Long, myparagraph$) As Long
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    WritePara getAparaNum, myparagraph$
    
    If ParaNum = -1 Then
     If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    Else
    para(getAparaNum).startfrom = para(ParaNum).startfrom
    InsertDocPara getAparaNum, ParaNum, DocParaNext(ParaNum)
    
    End If
    InsertAfterParagraph = getAparaNum
End Function
Function InsertAfterParagraphPlusColor(ParaNum As Long, myparagraph$, colordata() As Long, Optional Offset As Long = 0) As Long
    Dim getAparaNum As Long, mColor() As Long
    mColor() = colordata()
    getAparaNum = FindNewPara
    WritePara getAparaNum, myparagraph$
    If ParaNum = -1 Then
     If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    Else
    InsertDocPara getAparaNum, ParaNum, DocParaNext(ParaNum)
    End If
        If Offset > 0 Then mColor() = Relocate(mColor(), Offset)
    para(getAparaNum).colordata() = mColor()
    
    InsertAfterParagraphPlusColor = getAparaNum
End Function

Function InsertAfterParagraph2(ParaNum As Long, myparagraph$) As Long
    Dim getAparaNum As Long
    getAparaNum = FindNewPara

    WritePara getAparaNum, myparagraph$

    If ParaNum = -1 Then
     If docpara = 0 Then
        InsertDocPara getAparaNum, 0, 0  ' back is 0 and next is zero.
    Else
        InsertDocPara getAparaNum, DocParaBack(0), 0  ' DocParaBack(0) link to last
    End If
    Else
    InsertDocPara getAparaNum, ParaNum, DocParaNext(ParaNum)
    End If
    
  If ParaNum > 0 Then para(DocParaNext(ParaNum)).startfrom = para(ParaNum).startfrom
  'ColorThis DocParaNext(Paranum)
    InsertAfterParagraph2 = getAparaNum
End Function
 
Function InsertBeforeParagraph(ParaNum As Long, myparagraph$) As Long
    Dim getAparaNum As Long
    getAparaNum = FindNewPara
    WritePara getAparaNum, myparagraph$
       InsertDocPara getAparaNum, DocParaBack(ParaNum), ParaNum
       para(getAparaNum).startfrom = para(ParaNum).startfrom
    InsertBeforeParagraph = getAparaNum
End Function
Property Let textDocFast(norawtext As String)
    Dim where As Long ', before As Long, t_text As String
    Dim a() As String
    where = 1
    a() = Split(norawtext, vbCrLf)
        
        If UBound(a()) = LBound(a()) Then
        AppendParagraphOneLine norawtext
        Else
      For where = LBound(a()) To UBound(a())
      AppendParagraphOneLine a(where)
      Next where
    End If
End Property
Property Let textDoc(norawtext As String)
    Busy = True
    If Len(norawtext) = 0 Then
    AppendParagraph norawtext
    Else
    Dim where As Long, ok
    Dim a() As String
    
    If InStr(norawtext, vbLf) = 0 Then
    a() = Split(norawtext, vbCr)
    Else
    a() = Split(norawtext, vbCrLf)
    End If
        norawtext = Join(a(), vbLf)
             If InStr(norawtext, vbCr) Then
              norawtext = Join(Split(norawtext, vbCr), vbLf)
     End If
        If InStr(norawtext, vbLf) = 0 Then
        AppendParagraph norawtext
        Else
      a() = Split(norawtext, vbLf)
      For where = LBound(a()) To UBound(a())
      RaiseEvent MayQuit(ok)
      If ok Then Busy = False: Exit Property
      AppendParagraph a(where)
      Next where
    End If
    
    End If
    ' colorize at a second stage...(new)
MarkParagraphID = 0
WrapAgain
Busy = False
End Property
Property Let textDoc2(norawtext As String)
    Busy = True

    
    If Len(norawtext) = 0 Then
    AppendParagraph norawtext
    Else
    Dim where As Long, ok
    Dim a() As String
      a() = Split(norawtext, vbCrLf)
      For where = LBound(a()) To UBound(a())
      RaiseEvent MayQuit(ok)
      If ok Then Busy = False: Exit Property
      AppendParagraph a(where)
      
      Next where
    End If
    ' colorize at a second stage...(new)
MarkParagraphID = 0
WrapAgain
Busy = False
End Property

Property Let FasttextDoc(norawtext As String)
    EmptyDoc
    Busy = True

    If Len(norawtext) = 0 Then
    AppendParagraph norawtext
    Else
        Dim where As Long
        Dim a() As String
        If InStr(norawtext, vbLf) = 0 Then
        a() = Split(norawtext, vbCr)
        Else
        a() = Split(norawtext, vbCrLf)
        End If
        norawtext = Join(a(), vbLf)
          If InStr(norawtext, vbCr) Then
              norawtext = Join(Split(norawtext, vbCr), vbLf)
         End If
            If InStr(norawtext, vbLf) = 0 Then
            AppendParagraph norawtext
            Else
          a() = Split(norawtext, vbLf)
          For where = LBound(a()) To UBound(a())
          AppendParagraph a(where)
          Next where
        End If
    End If
    MarkParagraphID = 0
    Busy = False
End Property
Property Let textDocColor(norawtext As String)

    'where = 1
   ' If InStr(norawtext, vbLf) = 0 Then
    'a() = Split(norawtext, vbCr)
    'Else
    'a() = Split(norawtext, vbCrLf)
    'End If
     '   norawtext = Join(a(), vbLf)
        
        If InStr(norawtext, vbCrLf) = 0 Then
            ColorThis AppendParagraph(norawtext)
        Else
    Dim where As Long
    Dim a() As String
            a() = Split(norawtext, vbCrLf)
            For where = LBound(a()) To UBound(a())
                ColorThis AppendParagraph(a(where))
            Next where
        End If
    ' colorize at a second stage...(new)
WrapAgain
End Property

Sub textParagraphPlusColor(norawtext As String, colordata() As Long, Optional Offset As Long = 0)
Dim where As Long, copyColor() As Long
where = AppendParagraph(norawtext)
copyColor() = colordata()
If Offset > 0 Then copyColor() = Relocate(copyColor(), Offset)
para(where).colordata = copyColor()
End Sub
Public Function textFormat(ByVal linebreak2use$) As String
'
Dim mlines As Long, NeoData As String
 Dim mycur As Long, thiscur, l As Long
mycur = DocParaNext(0)
textFormat = space$(500)
Do While mycur <> 0
    If thiscur <> 0 Then NeoData = linebreak2use$ Else thiscur = 1
    For mlines = 1 To para(mycur).lines
        NeoData = NeoData + para(mycur).content(mlines)
    Next mlines
l = Len(NeoData)
If Len(textFormat) < thiscur + l Then textFormat = textFormat + space$((thiscur + l) + 100)
Mid$(textFormat, thiscur, l) = NeoData
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
textFormat = Left$(textFormat, thiscur - 1)
End Function
Private Function textArray(ByVal linebreak2use$, mycur As Long, minimumbytes2exit As Long) As Byte()
'
Dim mlines As Long, NeoData As String
 Dim thiscur, l As Long, textbuf As String
If minimumbytes2exit < 500 Then minimumbytes2exit = 500
textbuf = space$(500)
Do While mycur <> 0 And Not Len(textbuf) > minimumbytes2exit
    If thiscur <> 0 Then NeoData = linebreak2use$ Else thiscur = 1
    For mlines = 1 To para(mycur).lines
        NeoData = NeoData + para(mycur).content(mlines)
    Next mlines
     
l = Len(NeoData)
If Len(textbuf) < thiscur + l Then textbuf = textbuf + space$((thiscur + l) + 100)
Mid$(textbuf, thiscur, l) = NeoData
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
If mycur <> 0 Then
textArray = Left$(textbuf, thiscur - 1) + linebreak2use$
Else
textArray = Left$(textbuf, thiscur - 1)
End If

End Function
Private Function textString(ByVal linebreak2use$, mycur As Long, minimumbytes2exit As Long) As String
'
Dim mlines As Long, NeoData As String
 Dim thiscur, l As Long, textbuf As String
If minimumbytes2exit < 500 Then minimumbytes2exit = 500
textbuf = space$(500)
Do While mycur <> 0 And Not Len(textbuf) > minimumbytes2exit
    If thiscur <> 0 Then NeoData = linebreak2use$ Else thiscur = 1
    For mlines = 1 To para(mycur).lines
        NeoData = NeoData + para(mycur).content(mlines)
    Next mlines
l = Len(NeoData)
If Len(textbuf) < thiscur + l Then textbuf = textbuf + space$((thiscur + l) + 100)
Mid$(textbuf, thiscur, l) = NeoData
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
If mycur <> 0 Then
textString = Left$(textbuf, thiscur - 1) + linebreak2use$
Else
textString = Left$(textbuf, thiscur - 1)
End If
End Function
Property Get textDoc() As String
Attribute textDoc.VB_UserMemId = 0
Dim mlines As Long, NeoData As String
 Dim mycur As Long, thiscur, l As Long
mycur = DocParaNext(0)
textDoc = space$(500)
Do While mycur <> 0
    If thiscur <> 0 Then NeoData = vbCrLf Else thiscur = 1
    For mlines = 1 To para(mycur).lines
        NeoData = NeoData + para(mycur).content(mlines)
    Next mlines
l = Len(NeoData)
If Len(textDoc) < thiscur + l Then textDoc = textDoc + space$((thiscur + l) + 100)
Mid$(textDoc, thiscur, l) = NeoData
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
If Typename(thiscur) = "Empty" Then
Else
textDoc = Left$(textDoc, thiscur - 1)
End If
End Sub
Function textDocFormated(Optional mode2saveV) As String
Dim mlines As Long, NeoData As String, a() As Byte, bb As Byte, otherlcid As Long, lf$
Dim textDoc As String, mult As Long, bom$, mode2save As Long
If IsMissing(mode2saveV) Then mode2save = ListLoadedType Else mode2save = CLng(mode2saveV)

    If mode2save >= 500 Then
        otherlcid = mode2save
        mode2save = 3
        If GetCodePage(otherlcid) = 0 Then
        NoValidLocale
        Exit Function
        End If
    Else
    otherlcid = Module1.Clid
    End If
lf$ = choose(Abs(ListLoadedType \ 10) + 1, vbCrLf, vbLf, vbCr)
If Abs(mode2save \ 10) > 0 Then
lf$ = choose(Abs(mode2save \ 10) Mod 3 + 1, vbCrLf, vbLf, vbCr)
End If
mode2save = mode2save Mod 10
If Abs(mode2save) = 1 Or Abs(mode2save) = 0 Or Abs(mode2save) = 7 Then
mult = 2
Else
mult = 1
End If
If Abs(mode2save) = 7 Or mode2save = 0 Then
    bom$ = ChrW(&HFEFF)
 ElseIf mode2save = 1 Then
 
    bom$ = ChrW(&HFFFE)
ElseIf mode2save = 2 Then
ReDim a(0 To 2)
 a(0) = CByte(&HEF)
 a(1) = CByte(&HBB)
 a(2) = CByte(&HBF)
 bom$ = a()
End If
 Dim mycur As Long, thiscur, l As Long
mycur = DocParaNext(0)
Dim flag As Boolean
flag = False
textDoc = bom$ + space$(500)
Do While mycur <> 0
    If flag Then
        NeoData = lf$
    Else
        thiscur = LenB(bom$) + 1
        flag = True
    End If
    For mlines = 1 To para(mycur).lines
        NeoData = NeoData + para(mycur).content(mlines)
    Next mlines
      If Abs(mode2save) = 2 Then
        NeoData = utf8encode(NeoData)
     ElseIf Abs(mode2save) = 3 Then
    
        NeoData = StrConv(NeoData, vbFromUnicode, otherlcid)
     ElseIf Abs(mode2save) = 1 Then
        a() = NeoData
        Dim i As Long
        For i = 0 To UBound(a()) - 1 Step 2
                bb = a(i): a(i) = a(i + 1): a(i + 1) = bb
        Next i
        NeoData = a()
      End If
l = LenB(NeoData)
If LenB(textDoc) < thiscur + l Then textDoc = textDoc + space$((thiscur + l) + 100)
If l > 0 Then MidB$(textDoc, thiscur, l) = NeoData
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
If Typename(thiscur) = "Empty" Then
    textDocFormated = bom$
Else
    textDocFormated = LeftB$(textDoc, thiscur - 1)
End If
End Function
Function textHtmlDoc(colortable() As String) As String
Dim mlines As Long, NeoData As String, a77$
 Dim mycur As Long, thiscur, l As Long, a$, mColor() As Long, qu$
 qu$ = """"
 a77$ = space$(EditTabWidth)
 a77$ = Replace(a77$, " ", "&nbsp;")
mycur = DocParaNext(0)
textHtmlDoc = space$(500)
Do While mycur <> 0
    NeoData = vbNullString
    For mlines = 1 To para(mycur).lines
        NeoData = NeoData + para(mycur).content(mlines)
    Next mlines
If NeoData = vbNullString Then
If thiscur <> 0 Then NeoData = "<p>" Else thiscur = 1
Else

a$ = Replace$(NeoData, "<", ChrW$(3))
a$ = Replace$(a$, ">", ChrW$(4))
a$ = Replace$(a$, "  ", ChrW$(7) + ChrW$(7))
a$ = Replace$(a$, ChrW$(7) + " ", ChrW$(7) + ChrW$(7))
If UBound(para(mycur).colordata()) > 0 Then
mColor() = para(mycur).colordata()
NeoData = vbNullString
l = 1
If mColor(l) > 1 Then mColor(l) = 1
Do
If l > 3 Then
If mColor(l - 3) <> 15 And mColor(l - 3) <> 7 Then
If mColor(l) - mColor(l - 2) < 1 Then
If l > 1 Then NeoData = NeoData + Replace(Mid$(a$, mColor(l - 2)), ChrW$(7), " ") + "</span>"
Else
If l > 1 Then NeoData = NeoData + Replace(Mid$(a$, mColor(l - 2), mColor(l) - mColor(l - 2)), ChrW$(7), " ") + "</span>"
End If
Else
If l > 1 Then
    If mColor(l) - mColor(l - 2) < 1 Then
    NeoData = NeoData + Mid$(a$, mColor(l - 2)) + "</span>"
    Else
NeoData = NeoData + Mid$(a$, mColor(l - 2), mColor(l) - mColor(l - 2)) + "</span>"
End If
End If
End If
Else
If l > 1 Then
    If mColor(l) - mColor(l - 2) < 1 Then
    NeoData = NeoData + Mid$(a$, mColor(l - 2)) + "</span>"
    Else
    NeoData = NeoData + Mid$(a$, mColor(l - 2), mColor(l) - mColor(l - 2)) + "</span>"
    End If
End If
End If
NeoData = NeoData + "<span style=" + qu$ + "color: #" + colortable(mColor(l - 1)) + ";" + qu$ + ">"

  l = l + 2
Loop Until l > UBound(mColor())
If mColor(l - 2) > 0 Then
If mColor(l - 3) <> 15 And mColor(l - 3) <> 7 Then
a$ = NeoData + Replace(Mid$(a$, mColor(l - 2)), ChrW$(7), " ") + "</span>"
Else
a$ = NeoData + Mid$(a$, mColor(l - 2)) + "</span>"
End If
Else
a$ = NeoData + "</span>"
End If

End If

a$ = Replace$(a$, ChrW$(3), "&lt;")
a$ = Replace$(a$, ChrW$(4), "&gt;")
a$ = Replace$(a$, ChrW$(7), "&nbsp;")
If thiscur <> 0 Then NeoData = "<br />" + a$ Else thiscur = 1: NeoData = a$
End If
If Nonbsp Then
NeoData = Replace$(NeoData, vbTab, a77$)
End If
l = Len(NeoData)
If Len(textHtmlDoc) < thiscur + l Then textHtmlDoc = textHtmlDoc + space$((thiscur + l) + 100)
Mid$(textHtmlDoc, thiscur, l) = NeoData
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
If Typename(thiscur) = "Empty" Then

Else
If InStr(textHtmlDoc, vbTab) > 0 Then
    textHtmlDoc = "<pre >" + Left$(textHtmlDoc, thiscur - 1) + vbCrLf + "</pre>"
Else
    textHtmlDoc = Left$(textHtmlDoc, thiscur - 1) + vbCrLf
End If
End If
End Function
Property Get textDoc1LINE() As String  ' MORE SPEED IF WE KNOW THAT WE DON'T USE LINE BREAK
Dim mlines As Long, NeoData As String
 Dim mycur As Long, thiscur, l As Long
mycur = DocParaNext(0)
textDoc1LINE = space$(1000)
Do While mycur <> 0
    If thiscur <> 0 Then NeoData = vbCrLf Else thiscur = 1
        NeoData = NeoData + para(mycur).content(1)
   
l = Len(NeoData)
If Len(textDoc1LINE) < thiscur + l Then textDoc1LINE = textDoc1LINE + space$((thiscur + l) + 100)
Mid$(textDoc1LINE, thiscur, l) = NeoData
thiscur = thiscur + l
mycur = DocParaNext(mycur)
Loop
textDoc1LINE = Left$(textDoc1LINE, thiscur - 1)
End Property
Private Sub InsertLine(b$, ParaNum, cur)
    Dim mlines As Long, NeoData As String
    If ParaNum = 0 Then Exit Sub
    If cur <> 0 Then NeoData = vbCrLf Else cur = 1
    For mlines = 1 To para(ParaNum).lines
        NeoData = NeoData + para(ParaNum).content(mlines)
    Next mlines

If Len(b$) < cur + Len(NeoData) Then b$ = b$ + space$(Len(b$) * 2)
Mid$(b$, cur, Len(NeoData)) = NeoData
cur = cur + Len(NeoData)
End Sub

Public Sub InsertDoc(ByVal ParaNum As Long, fromPos As Long, norawtext As String, Optional TabWidth As Long = 0, Optional gotonextfree As Boolean = False, Optional removenbsp As Boolean = False)
    Dim t_text As String, once As Boolean
    Dim a() As String, t() As String
    Dim b$
   ' If para(ParaNum).startfrom = ParaNum Then
   ' para(ParaNum).startfrom = 0
   ' End If
    once = True
      a() = Split(norawtext, vbCrLf)
        norawtext = Join(a(), vbLf)
     If InStr(norawtext, vbCr) Then
      a() = Split(norawtext, vbCr)
      norawtext = Join(a(), vbLf)
     End If
     a() = Split(norawtext + vbLf, vbLf)
     
Dim i As Long, k As Long, j As Long, ii As Long
    i = LBound(a())
    Do
    If i > UBound(a()) Then Exit Do
    If Left$(a(i), 1) = Chr$(160) Then

    Dim JJ As Long
    If removenbsp Then
    For JJ = 1 To Len(a(i))
    If Mid$(a(i), JJ, 1) <> Chr$(160) Then Mid$(a(i), 1, JJ - 1) = space$(JJ - 1): Exit For
    Next JJ
   ' If Left$(a(i), 1) = Chr$(160) Then a(i) = space$(Len(a(i)))
    End If
    End If
    If TabWidth > 0 Then
    t() = Split(a(i), Chr(9))
    k = LBound(t())
    j = UBound(t())
        If k = j Then
        t_text = t(k)
        Else
        t_text = vbNullString
        For ii = k To j
        If Len(t(ii)) = 0 Then
        t_text = t_text + space(TabWidth)
        ElseIf Len(t_text) Mod TabWidth = 0 Then
        t_text = t_text + t(ii)
        Else
            t_text = t_text + space(TabWidth - (Len(t_text) Mod TabWidth)) + t(ii)
            End If
        Next ii
      
        End If
    Else
    t_text = a(i)
    End If
    
    i = i + 1
    If t_text = vbNullString And i > UBound(a()) Then
    
    If b$ <> "" Then
        ReWritePara ParaNum, TextParagraph(ParaNum) + b$, True
        b$ = vbNullString
    End If
    Exit Do
    End If
    If once Then
    once = False
        b$ = TextParagraph(ParaNum)
        If fromPos = 1 Then
          ReWritePara ParaNum, t_text ', True
          fromPos = Len(t_text) + 1
       
        ElseIf fromPos > Len(b$) Then
           
           ReWritePara ParaNum, b$ & t_text, True
            fromPos = Len(b$ & t_text) + 1
            
            b$ = vbNullString
        Else
        If fromPos <= 0 Then fromPos = 1
        
        ReWritePara ParaNum, Left$(b$, fromPos - 1) + t_text, True
        b$ = Mid$(b$, fromPos)
        fromPos = fromPos + Len(t_text)
        
        End If
    Else
    'ColorThis (Paranum)
    ResetColor ParaNum
    ParaNum = InsertAfterParagraph(ParaNum, t_text)
         fromPos = Len(t_text) + 1
       
    End If
    Loop
    If gotonextfree Then
    MarkParagraphID = DocParaNext(ParaNum)
    Else
    MarkParagraphID = ParaNum
    End If
    ResetColor ParaNum
   ' ColorThis (Paranum)
End Sub
Property Get DocLines() As Long
    Dim mycur As Long, mlines As Long
    mycur = DocParaNext(0)
    While mycur <> 0
        mlines = para(mycur).lines + mlines
        mycur = DocParaNext(mycur)
    Wend
    DocLines = mlines
End Property
Property Get DocParagraphs() As Long
    Dim mycur As Long, dPara As Long
    mycur = DocParaNext(0)
    While mycur <> 0
        dPara = dPara + 1
        mycur = DocParaNext(mycur)
    Wend
    DocParagraphs = dPara
End Property

Property Get FirstParagraphLine(ParaNum As Long) As Long
    Dim mycur As Long, mlines As Long, mlinesold As Long, valid As Boolean
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        If mycur = ParaNum Then valid = True: Exit Do
        mlines = para(mycur).lines + mlines
        mycur = DocParaNext(mycur)
    Loop
    If valid Then FirstParagraphLine = mlinesold
End Property

'
Public Sub FindPos(ByVal fromLineNo As Long, ByVal fromPos As Long, CharNumber As Long, toLineNo As Long, toPos As Long, toParagraph As Long, toParaPos As Long)
Dim probelen As Long, i As Long, target As Long, sum As Long
Dim mycur As Long, mlines As Long, mlinesold As Long, sumtohere As Long
toLineNo = fromLineNo
toPos = fromPos
toParagraph = DocParaNext(0)
toParaPos = 1
If fromLineNo < 1 Then Exit Sub
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        If fromLineNo <= mlines Then Exit Do
         sumtohere = sumtohere + TextParagraphLen(mycur)
        mycur = DocParaNext(mycur)
        If Not mycur = 0 Then sumtohere = sumtohere + 2         ' now we put the vbcrlf length (these are not bytes)
    Loop
    toParagraph = mycur
    If CharNumber = 0 Then Exit Sub
    If mycur <> 0 Then
        target = fromLineNo - mlinesold
        probelen = 0
        For i = 1 To target - 1
             toPos = Len(para(mycur).content(i))
             probelen = probelen + toPos
             sumtohere = sumtohere + toPos
        Next i
        toLineNo = fromLineNo
        toPos = Len(para(mycur).content(target))
        If CharNumber < 0 Then
            If fromPos = 0 Then
                CharNumber = sumtohere
            Else
                CharNumber = sumtohere + fromPos - 1
            End If
            Exit Sub
        Else
           sum = toPos - fromPos
        End If
        If target = para(mycur).lines Then
           
                        If (CharNumber - sum) <= 2 Then
                            If (CharNumber - sum) = 2 Then
                                 If DocParaNext(mycur) = 0 Then
                              toPos = CharNumber + fromPos
                              toParaPos = toPos + probelen - 2
                              
                              Else
                                toPos = 0
                                toParagraph = DocParaNext(mycur)
                                toLineNo = toLineNo + 1
                                toParaPos = 0
                                End If
                            Else
                                toPos = CharNumber + fromPos
                                toParaPos = toPos + probelen
                            End If
                            Exit Sub
                        End If
               
        Else
                If CharNumber <= sum Then
                    toPos = CharNumber + fromPos
                    toParaPos = toPos + probelen
                    Exit Sub
                End If
        End If
        sum = toPos - fromPos
        probelen = probelen + toPos
        Do
                For i = target + 1 To para(mycur).lines - 1
                    toLineNo = i + mlinesold  ' set the right lineNo
                    toPos = Len(para(mycur).content(i))
                    probelen = probelen + toPos
                    sum = sum + toPos
                    If sum >= CharNumber Then
                        probelen = probelen - toPos  ' rewind
                        toParaPos = probelen + CharNumber - (sum - toPos)
                        toPos = CharNumber - (sum - toPos)
                        Exit Sub
                    End If
                Next i
    
        If para(mycur).lines < i Then
            toPos = 0
        Else
                toPos = Len(para(mycur).content(i))
        End If
             toLineNo = i + mlinesold
               

                
                If DocParaNext(mycur) = 0 Then
                        If (CharNumber - sum) <= toPos Then
                            toPos = CharNumber - sum - 1
                        End If
                        toParaPos = toPos + probelen
                        Exit Do
                Else
                        If (CharNumber - sum - toPos) <= 2 Then
                               If (CharNumber - sum - toPos) = 2 Then
                                    toPos = 0
                                    toParagraph = DocParaNext(mycur)
                                    toLineNo = toLineNo + 1
                                    toParaPos = 0
                                Else
                                    toPos = CharNumber - sum
                                    toParaPos = toPos + probelen
                                End If
                                Exit Sub
                        End If
                End If
   
                probelen = 0
                sum = sum + toPos + 2
                mycur = DocParaNext(mycur)
                toParagraph = mycur
                target = 0
                mlinesold = mlines
                mlines = para(mycur).lines + mlines
        Loop
    End If
    toParaPos = toParaPos + 1
    toPos = toPos + 1
End Sub
Public Function RemoveEmptyParagraph(ParaNum As Long) As Boolean
' but not last, only erased
Dim oldone As Long
If TextParagraphLen(ParaNum) = 0 Then
RemoveEmptyParagraph = True
oldone = DocParaNext(ParaNum)
If oldone = 0 Then oldone = DocParaBack(ParaNum)
If oldone = 0 Then
    para(ParaNum).lines = 1  ' clear
    ReDim para(ParaNum).content(1 To 1)
    ReDim para(ParaNum).colordata(1)
    para(ParaNum).startfrom = 0
Else
RemoveDocParaIndex ParaNum
ParaNum = oldone
End If
End If
End Function

Property Get LastCurLineChar() As Long
LastCurLineChar = lastlinelastchar
End Property
Property Get lastcolordata() As Long()
If lastcolordatacur <> 0 Then
lastcolordata = Relocate(para(lastcolordatacur).colordata(), lastlinelastchar)
Else
Dim dummy(1) As Long
lastcolordata = dummy()
End If
End Property
Function Relocate(aa() As Long, ByVal Offset As Long) As Long()
Dim i As Long, bbb() As Long, j As Long
If Offset = 0 Or UBound(aa()) <= 1 Then
Relocate = aa()
Exit Function
End If
For i = 1 To UBound(aa()) - 1 Step 2
If aa(i) > Offset Then Exit For
Next i
If i > UBound(aa()) Then GoTo there:
    If aa(i) <= Offset Then
there:
        ReDim bbb(2)
        bbb(0) = aa(UBound(aa()) - 1)
        bbb(1) = 1
    Else
    If aa(i) - Offset <> 1 Then
        ReDim bbb(UBound(aa()) - i + 4)
        bbb(0) = aa(i - 3) ' preserve color
        bbb(1) = 1
        j = 2
    Else
        ReDim bbb(UBound(aa()) - i + 2)
    End If
        For i = i - 1 To UBound(aa()) - 1 Step 2
        bbb(j) = aa(i) 'color
        bbb(j + 1) = aa(i + 1) - Offset
        j = j + 2
        Next i
    End If

Relocate = bbb()
End Function
Property Get TextLine(lineno As Long) As String
If lineno < 1 Then Exit Property
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
    lastcolordatacur = mycur
    lastlinelastchar = 0
    For mlines = 1 To lineno - mlinesold - 1
    lastlinelastchar = lastlinelastchar + Len(para(mycur).content(mlines))
    Next mlines
        TextLine = para(mycur).content(lineno - mlinesold)
    End If
End Property
Property Get TextLine2(lineno As Long, showspaces As Boolean, Optional skipcolor As Boolean = False) As String
If lineno < 1 Then Exit Property
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
    
    If ColorEvent And Not skipcolor Then
    If para(mycur).colordata(1) = 0 Then
    ColorThis (mycur)
        End If
End If
    lastcolordatacur = mycur
    lastlinelastchar = 0
    For mlines = 1 To lineno - mlinesold - 1
    lastlinelastchar = lastlinelastchar + Len(para(mycur).content(mlines))
    Next mlines
    If showspaces Then
        TextLine2 = Replace(Replace(para(mycur).content(lineno - mlinesold), " ", ChrW(&HB7)), ChrW(&H2007), ChrW(&H2423))
    TextLine2 = Replace(TextLine2, vbTab, ChrW(&H21E5))
    Else
    TextLine2 = para(mycur).content(lineno - mlinesold)
        End If
    End If
End Property
Property Let TextLine(lineno As Long, that As String)
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
        para(mycur).content(lineno - mlinesold) = that
    End If
End Property
Property Get ParagraphFromOrder(ByVal order As Long) As Long
    ParagraphFromOrder = -1
    Dim mycur As Long
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order - 1
        If order = 0 Then ParagraphFromOrder = mycur: Exit Do
        mycur = DocParaNext(mycur)
    Loop
End Property
Property Let startfrom(ParaNum As Long, there As Long)
para(ParaNum).startfrom = there
End Property
Property Get startfrom(ParaNum As Long) As Long
startfrom = para(ParaNum).startfrom
End Property
Property Let Paracolordata(ParaNum As Long, col() As Long)
If ParaNum <> 0 Then para(ParaNum).colordata() = col()
End Property
Property Get Paracolordata(ParaNum As Long) As Long()
' it is valid
On Error Resume Next
Paracolordata = para(ParaNum).colordata()

End Property
Property Get ParagraphOrder(ParaNum As Long) As Long
    ParagraphOrder = 0
    Dim mycur As Long, order As Long
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order + 1
        If mycur = ParaNum Then ParagraphOrder = order: Exit Do
        mycur = DocParaNext(mycur)
    Loop
End Property
Property Get LastParagraph() As Long
LastParagraph = DocParaBack(0)
End Property
Property Get SizeCRLF() As Long
Dim mycur As Long, suma As Long, k As Long
If IsEmpty Then
Else
 mycur = DocParaNext(0)
  k = -1
    Do While mycur > 0
    k = k + 1
        suma = suma + TextParagraphLen(mycur)
        mycur = DocParaNext(mycur)
    Loop
    End If
    SizeCRLF = suma + k * 2
End Property
Property Get TextParagraphOrder(ByVal order As Long) As String
    ' 1 is the first
    ' -1 is the last
    Dim mycur As Long
    If order = -1 Then
        mycur = DocParaBack(0)
        TextParagraphOrder = TextParagraph(mycur)
        Exit Property
    ElseIf order = 0 Then
        order = 1
    End If
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order - 1
        If order = 0 Then
            TextParagraphOrder = TextParagraph(mycur)
            Exit Property
        End If
        mycur = DocParaNext(mycur)
    Loop
End Property
Property Get TextParagraphOrderRemove(ByVal order As Long) As String
    ' 1 is the first
    ' -1 is the last
    Dim mycur As Long
    If order = -1 Then
        mycur = DocParaBack(0)
        TextParagraphOrderRemove = TextParagraph(mycur)
         RemoveDocParaIndex mycur
        Exit Property
    ElseIf order = 0 Then
        order = 1
    End If
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order - 1
        If order = 0 Then
            TextParagraphOrderRemove = TextParagraph(mycur)
            RemoveDocParaIndex mycur
            Exit Property
        End If
        mycur = DocParaNext(mycur)
    Loop
End Property
Public Function FindStrUp(s$, FirstPara As Long, FirstPosPara As Long) As Boolean
Dim a As Long, b As Long, mC As Long
a = FirstPara
mC = FirstPosPara - 1
If FirstPosPara <= 0 Then
mC = -1
a = DocParaBack(a)
If a = 0 Then Exit Function
Else
mC = mC - Len(s$)
End If
If mC > 0 Then b = InstrRevTextCompareAnyLocale(Me.TextParagraph(a), s$, mC, mlcid) 'using default locale????
If b = 0 Then
a = DocParaBack(a)
Do While b = 0 And a > 0
b = InstrRevTextCompareAnyLocale(Me.TextParagraph(a), s$, , mlcid)   'using default locale????
If b Then Exit Do
a = DocParaBack(a)
Loop
End If
If b > 0 Then
FirstPara = a
FirstPosPara = b
FindStrUp = True
End If
End Function
Public Function FindStrDown(s$, FirstPara As Long, FirstPosPara As Long) As Boolean
'return true if found
Dim a As Long, b As Long, mC As Long
If FirstPosPara <= 0 Then mC = 1 Else mC = FirstPosPara + 1
a = FirstPara
b = InstrTextCompareAnyLocale(mC, Me.TextParagraph(a), s$, mlcid)
If b = 0 Then
a = DocParaNext(a)
Do While b = 0 And a > 0
b = InstrTextCompareAnyLocale(1, TextParagraph(a), s$, mlcid)
If b Then Exit Do
a = DocParaNext(a)
Loop
End If
If b > 0 Then
FirstPara = a
FirstPosPara = b
FindStrDown = True
End If
End Function
Public Function FindStr(s$, ByVal st&, FindParaOrder As Long, FindPosPara As Long) As Long
Dim mycur As Long, ac&, ac1&, acp&, mlines As Long, wh&, order As Long
mycur = DocParaNext(0)
order = 1
    Do While mycur > 0
 order = order + 1
    ac& = acp&
    If st& <= 0 Then st& = 1
    For mlines = 1 To para(mycur).lines
        ac1& = ac& + Len(para(mycur).content(mlines))
        If ac1& > st& Then
        If st& - ac& <= 0 Then st& = ac& + 1
        wh& = InstrTextCompareAnyLocale(st& - ac&, para(mycur).content(mlines), s$, mlcid)
        If wh& > 0 Then
            FindParaOrder = order - 1
            FindPosPara = wh& + ac& - acp&
            FindStr = wh& + ac&
        Exit Function
        
        End If
        Exit Do
        
        End If
    Next mlines
      acp& = ac1& + 2  'for linefeed
      mycur = DocParaNext(mycur)
      
    Loop
    If mycur > 0 Then
        If mlines < para(mycur).lines Then
            For mlines = mlines To para(mycur).lines
            ac1& = ac& + Len(para(mycur).content(mlines))
            wh& = InstrTextCompareAnyLocale(1, para(mycur).content(mlines), s$, mlcid)
                If wh& > 0 Then
                        FindParaOrder = order
                        FindPosPara = wh& + ac& - acp&
                        FindStr = wh& + ac&
                        Exit Function
                End If
        Next mlines
        Else
            order = order - 1
        End If
          acp& = ac1& + 2   'for linefeed
          mycur = DocParaNext(mycur)
       
    End If
    
        Do While mycur > 0
        order = order + 1
        ac& = acp&
    For mlines = 1 To para(mycur).lines
        ac1& = ac& + Len(para(mycur).content(mlines))
        wh& = InstrTextCompareAnyLocale(1, para(mycur).content(mlines), s$, mlcid)
        If wh& > 0 Then
                 FindParaOrder = order
                 FindPosPara = wh& + ac& - acp&
                 FindStr = wh& + ac&
                 Exit Function
                 Exit Do
         End If
     Next mlines
      acp& = ac1& + 2  'for linefeed
      mycur = DocParaNext(mycur)
    Loop
    

    
End Function
Property Let TextParagraphOrder(ByVal order As Long, data$)
    Dim mycur As Long
    If order = -1 Then
        mycur = DocParaBack(0)
        ReWritePara mycur, data$
        Exit Property
    ElseIf order = 0 Then
        order = 1
    End If
    mycur = DocParaNext(0)
    Do While mycur > 0
        order = order - 1
        If order = 0 Then
            ReWritePara mycur, data$
            Exit Property
        End If
        mycur = DocParaNext(mycur)
    Loop
End Property
Property Get TextParagraph(ByVal ParaNum As Long) As String
    If InvalidPara(ParaNum) Then Exit Property
    If ParaNum = 0 Then TextParagraph = vbNullString: Exit Property
    TextParagraph = Join(para(ParaNum).content, "")
     
End Property
Property Get TextParagraphKey(ByVal ParaNum As Long) As String
    If ParaNum = 0 Then TextParagraphKey = vbNullString: Exit Property
    TextParagraphKey = Mid$(Join(para(ParaNum).content, ""), mKeystart)
End Property

Property Let TextParagraph(ByVal ParaNum As Long, data$)
    If ParaNum = 0 Then Exit Property
    para(ParaNum).lines = 1   ' clear' also can be a flag to show that this is deleted
    ReDim para(ParaNum).content(1 To 1)
    ReDim para(ParaNum).colordata(1)   ' erase old color  --
    para(ParaNum).content(1) = data$
End Property

Property Get TextParagraphLen(ByVal ParaNum As Long) As Long
' without addition of 2 chars vbcrlf
    Dim mlines As Long, neodatalen As Long
    If ParaNum = 0 Then TextParagraphLen = 0: Exit Property
    For mlines = 1 To para(ParaNum).lines
        neodatalen = neodatalen + Len(para(ParaNum).content(mlines))
    Next mlines
    TextParagraphLen = neodatalen
End Property
Function TextNextParagraph(ParaNum As Long) As String
Dim mycur As Long
mycur = DocParaNext(ParaNum)
If mycur > 0 Then
TextNextParagraph = TextParagraph(mycur)
ParaNum = mycur
End If

End Function
Function DeleteNextParagraph(ByVal ParaNum As Long) As String
    Dim mynextcur As Long
    mynextcur = DocParaNext(ParaNum)
    If mynextcur <> 0 Then
        DeleteNextParagraph = TextParagraph(mynextcur)
        RemoveDocParaIndex mynextcur
    End If
End Function
Function DeletePreviusParagraph(ByVal ParaNum As Long) As String
    Dim mybackcur As Long
    mybackcur = DocParaBack(ParaNum)
    If mybackcur <> 0 Then
        DeletePreviusParagraph = TextParagraph(mybackcur)
        RemoveDocParaIndex mybackcur
    End If
End Function

Property Get TextParagraph2Line(ByVal ParaNum As Long, ByVal Charpos As Long, newselstart, Optional skipcolor As Boolean = False) As Long
Dim mlines As Long, counter As Long
    If ParaNum <= 0 Then Exit Property
    If para(ParaNum).lines = 0 Then Exit Property
    
    If ColorEvent And Not skipcolor Then
If para(ParaNum).colordata(1) = 0 Then
    ColorThis (ParaNum)
    If para(ParaNum).colordata(1) = 0 Then para(ParaNum).startfrom = 0
End If
End If
     For mlines = 1 To para(ParaNum).lines
        counter = counter + Len(para(ParaNum).content(mlines))
        If Charpos <= counter Then
            newselstart = Charpos - (counter - Len(para(ParaNum).content(mlines))) '  - 1 to write from right to left
            TextParagraph2Line = mlines - 1
            Exit Property
        End If
    Next mlines
    If Charpos > counter Then newselstart = Len(para(ParaNum).content(para(ParaNum).lines)) + 1
    TextParagraph2Line = mlines - 2
End Property
Function TextParagraphLine2(ByVal lineno As Long, thistext As String, ParaNum As Long, showpar As Boolean) As Long
    ' for lineno get line and paragraph id..
    ' we want to mark the current paragraph
    Dim mycur As Long, mlines As Long, mlinesold As Long, i As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
    If lineno = mlines Then RaiseEvent Lastline
      If (lineno - mlinesold) > 0 Then
      If showpar Then
      thistext = Replace(para(mycur).content(lineno - mlinesold), " ", ChrW(&HB7))
      thistext = Replace(thistext, ChrW(&H2007), ChrW(&H2423))
      thistext = Replace(thistext, vbTab, ChrW(&H21E5))
      Else
      thistext = para(mycur).content(lineno - mlinesold)
      End If
      End If
        ParaNum = mycur
        If (lineno - mlinesold) = 1 Then
            TextParagraphLine2 = 0
        Else
            For i = 1 To lineno - mlinesold - 1
                TextParagraphLine2 = TextParagraphLine2 + Len(para(mycur).content(i))
            Next i
        End If
        
    Else
        thistext = vbNullString
        ParaNum = 0
    End If
End Function
Function TextParagraphLine(ByVal lineno As Long, thistext As String, ParaNum As Long) As Long
    ' for lineno get line and paragraph id..
    ' we want to mark the current paragraph
    Dim mycur As Long, mlines As Long, mlinesold As Long, i As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        mlines = para(mycur).lines + mlines
        If lineno <= mlines Then Exit Do
        mycur = DocParaNext(mycur)
    Loop
    If lineno <= mlines Then
    If lineno = mlines Then RaiseEvent Lastline
      If (lineno - mlinesold) > 0 Then thistext = para(mycur).content(lineno - mlinesold)
        ParaNum = mycur
        If (lineno - mlinesold) = 1 Then
            TextParagraphLine = 0
        Else
            For i = 1 To lineno - mlinesold - 1
                TextParagraphLine = TextParagraphLine + Len(para(mycur).content(i))
            Next i
        End If
        
    Else
        thistext = vbNullString
        ParaNum = 0
    End If
End Function

Sub GoTopLinePreviusParagraph(lineno As Long, thistext As String, ParaNum As Long)
' we give Paranum and we get lineno and thistext as first line, and new paranum
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    If mycur > 0 Then
        If mycur = ParaNum Then
            thistext = para(ParaNum).content(1)  ' first line
            lineno = 1
            Exit Sub
        End If
    End If
    Do While mycur <> 0
        mlinesold = mlines
        If DocParaNext(mycur) = ParaNum Then Exit Do
        mlines = para(mycur).lines + mlines
        mycur = DocParaNext(mycur)
    Loop
    If DocParaNext(mycur) = ParaNum Then
        ParaNum = mycur
        thistext = para(ParaNum).content(1)  ' first line
        lineno = mlinesold + 1
    Else
        ParaNum = DocParaNext(0): mlinesold = 0
        thistext = para(ParaNum).content(1)  ' first line
        lineno = mlinesold + 1
    End If
End Sub
Sub GoTopLineNextParagraph(lineno As Long, thistext As String, ParaNum As Long)
' we give Paranum and we get lineno and thistext as first line at the next Paranum..
    Dim mycur As Long, mlines As Long, mlinesold As Long
    mlinesold = 0
    mycur = DocParaNext(0)
    Do While mycur <> 0
        mlinesold = mlines
        If mycur = ParaNum Then Exit Do  'lineno=mlinesold+1
        mlines = para(mycur).lines + mlines
        mycur = DocParaNext(mycur)
    Loop
    If mycur = ParaNum Then
        If DocParaNext(mycur) <> 0 Then
            mlines = para(mycur).lines + mlines
            mycur = DocParaNext(mycur)
        End If
        ParaNum = mycur ' new paragraph id
        thistext = para(ParaNum).content(1)  ' first line
        lineno = mlinesold + 1
    Else
        mycur = DocParaNext(0): mlinesold = 0  ' go to start again if not found..
        ParaNum = mycur ' new paragraph id
        thistext = para(ParaNum).content(1)  ' first line
        lineno = mlinesold + 1
    End If
End Sub
Public Sub WrapAgain()
Busy = True
Dim mycur As Long, ok
mycur = DocParaNext(0)
Do While mycur <> 0
 ReWritePara mycur, TextParagraph(mycur), True
       RaiseEvent MayQuit(ok)
      If ok Then Busy = False: Exit Sub
mycur = DocParaNext(mycur)
Loop
Busy = False
End Sub
Public Sub WrapAgainBlock(startblock As Long, endblock As Long)
Busy = True
Dim mycur As Long, ok
mycur = startblock
Do While mycur <> 0
 ReWritePara mycur, TextParagraph(mycur), True
       RaiseEvent MayQuit(ok)
      If ok Then Busy = False: Exit Sub
 If mycur = endblock Then Exit Do
 mycur = DocParaNext(mycur)
Loop
Busy = False
End Sub

Public Sub WrapAgainColor2()
Busy = False = True
Dim mycur As Long, ok
mycur = DocParaNext(0)

Do While mycur <> 0
      RaiseEvent MayQuit(ok)
      If ok Then Busy = False: Exit Sub
If para(mycur).colordata(1) = 0 Then
ColorThis mycur
Else
mycur = DocParaNext(mycur)
End If

Loop
Busy = False
End Sub

Public Sub WrapAgainColorBlock(startblock As Long, endblock As Long)
Busy = True
Dim mycur As Long, ok
mycur = startblock
If startblock = endblock Then
ColorThis mycur
Else
Do While mycur <> 0
      RaiseEvent MayQuit(ok)
      If ok Then Busy = False: Exit Sub
ColorThis mycur

If mycur = endblock Then Exit Do
Loop
If mycur = endblock Then ColorThis mycur
End If
Busy = False
End Sub
Public Sub WrapAgainColorBlock2(startblock As Long, endblock As Long)
Busy = True
Dim mycur As Long, ok
mycur = startblock
If startblock = endblock Then
ColorThis mycur
Else
Do While mycur <> 0
      RaiseEvent MayQuit(ok)
      If ok Then Exit Sub
If para(mycur).colordata(1) = 0 Then
ColorThis (mycur)
mycur = DocParaNext(mycur)
Else
mycur = DocParaNext(mycur)
End If
If mycur = endblock Then Exit Do
Loop

If mycur = endblock Then ColorThis mycur
End If
Busy = False
End Sub
Private Sub Class_Initialize()
' from this version paramax= docmax
'paramax = 20 ' 20 parapraphs ...content
Switchon = True  ' check this please
mSortstyle = vbTextCompare
docmax = 20  ' 20 paragraphs..links only
ReDim para(1 To docmax)
ReDim DocParaBack(0 To docmax)
ReDim DocParaNext(0 To docmax)
ReDim Buffer(1 To 1) As Byte
bufOffset = 0
bufsize = 0
SeekBuf = 1
mlcid = Module1.Clid
Final$ = "}"
End Sub

Sub BackSpaceNchars(fromPARA As Long, fromPos As Long, ByVal Nchars As Long)
' we return the current frompara and from pos
Dim ww As Long, ThatPara
ww = TextParagraphLen(fromPARA)
'If fromPos > ww Then Nchars = Nchars - (fromPos - ww): fromPos = ww

Do While Nchars > 0
If fromPos = 0 Then
fromPARA = DocParaBack(fromPARA)
fromPos = TextParagraphLen(fromPARA) + 1
TextParagraph(fromPARA) = TextParagraph(fromPARA) + Chr$(10) + DeleteNextParagraph(fromPARA)
Nchars = Nchars - 1
ElseIf fromPos = 1 Then
If Nchars > 2 Then
Nchars = Nchars - 2
fromPARA = DocParaBack(fromPARA)
fromPos = TextParagraphLen(fromPARA) + 1
TextParagraph(fromPARA) = TextParagraph(fromPARA) + DeleteNextParagraph(fromPARA)

' CONTINUE
Else
Nchars = Nchars - 1
fromPARA = DocParaBack(fromPARA)
fromPos = TextParagraphLen(fromPARA) + 2
TextParagraph(fromPARA) = TextParagraph(fromPARA) + Chr$(13) + DeleteNextParagraph(fromPARA)
End If
ElseIf fromPos > Nchars Then

TextParagraph(fromPARA) = Left(TextParagraph(fromPARA), fromPos - Nchars - 1) + Mid$(TextParagraph(fromPARA), fromPos)
fromPos = fromPos - Nchars
Exit Do
Else
Nchars = Nchars - fromPos + 1
      ThatPara = DocParaBack(fromPARA)
      
        If fromPos = TextParagraphLen(fromPARA) Then
        RemoveDocParaIndex fromPARA
        
        Else
        TextParagraph(fromPARA) = Mid$(TextParagraph(fromPARA), fromPos)
        End If
        
        If ThatPara = 0 Then ThatPara = DocParaNext(0)
        fromPARA = ThatPara
        If fromPARA = 0 Then
        ' make only one empty paragraph
        Exit Sub
        End If
      fromPos = TextParagraphLen(fromPARA) + 1
      If Nchars >= 2 Then
      Nchars = Nchars - 2
      
      TextParagraph(fromPARA) = TextParagraph(fromPARA) + DeleteNextParagraph(fromPARA)
      ElseIf Nchars = 1 Then
      TextParagraph(fromPARA) = TextParagraph(fromPARA) + Chr$(13) + DeleteNextParagraph(fromPARA)
      fromPos = fromPos + 1
      Nchars = 0
      Else
      
      Exit Sub
      End If
     
    
    
End If

Loop


End Sub
Public Sub SetBinaryCompare()
sortmlcid = 0
mSortstyle = vbBinaryCompare
End Sub
Public Sub SetLocaleCompare(Optional RHS As Variant)
If IsMissing(RHS) Then
sortmlcid = mlcid
Else
    If CLng(RHS) = 0 Then
        sortmlcid = OsInfo.LangNonUnicodeCode
    ElseIf GetCodePage(CLng(RHS)) <> 0 Then
        sortmlcid = CLng(RHS)
    Else
        NoValidLocale
    End If
End If
If sortmlcid = 0 Then sortmlcid = OsInfo.LangNonUnicodeCode
mSortstyle = vbBinaryCompare
End Sub

Public Sub SetTextCompare()
mSortstyle = vbTextCompare
End Sub
Public Sub SetDataBaseCompare()
mSortstyle = vbDatabaseCompare
End Sub
Private Sub QuickSortExtended(Arr() As Long, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
Dim Piv As String, tmp As String '<- adjust types here, when switching to something different than Long
     If UB - LB = 1 Then
     M1 = LB
      If CompareStr2(TextParagraphKey(Arr(M1)), TextParagraphKey(Arr(UB)), mSortstyle) = 1 Then SwapPara M1, UB, Arr
      Exit Sub
     Else
       M1 = (LB + UB) \ 2 '+ 1
             If CompareStr2(TextParagraphKey(Arr(M1)), TextParagraphKey(Arr(LB)), mSortstyle) = 0 Then
                M2 = UB - 1
                M1 = LB
                Piv = TextParagraphKey(Arr(LB))
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If CompareStr2(TextParagraphKey(Arr(UB)), Piv, mSortstyle) = -1 Then SwapPara LB, UB, Arr
                        Exit Sub
                    End If
                Loop Until CompareStr2(TextParagraphKey(Arr(M1)), Piv, mSortstyle) <> 0
                Piv = TextParagraphKey(Arr(M1))
                If M1 > LB Then If CompareStr2(TextParagraphKey(Arr(LB)), Piv, mSortstyle) = 1 Then SwapPara M1, LB, Arr: Piv = TextParagraphKey(Arr(M1))
            Else
                Piv = TextParagraphKey(Arr(M1))
                M1 = LB
                Do While CompareStr2(TextParagraphKey(Arr(M1)), Piv, mSortstyle) = -1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While CompareStr2(TextParagraphKey(Arr(M2)), Piv, mSortstyle) = 1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
       If M1 <> M2 Then SwapPara M1, M2, Arr
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While CompareStr2(TextParagraphKey(Arr(M1)), Piv, mSortstyle) = -1: M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortExtended Arr, LB, M2
    If M1 < UB Then QuickSortExtended Arr, M1, UB
    Exit Sub
abc1:
    
    
End Sub
Private Sub QuickSortExtendedDes(Arr() As Long, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
Dim Piv As String, tmp As String '<- adjust types here, when switching to something different than Long
     If UB - LB = 1 Then
     M1 = LB
      If CompareStr2(TextParagraphKey(Arr(M1)), TextParagraphKey(Arr(UB)), mSortstyle) = -1 Then SwapPara M1, UB, Arr
      Exit Sub
     Else
       M1 = (LB + UB) \ 2 ' + 1
             If CompareStr2(TextParagraphKey(Arr(M1)), TextParagraphKey(Arr(LB)), mSortstyle) = 0 Then
                M2 = UB - 1
                M1 = LB
                Piv = TextParagraphKey(Arr(LB))
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If CompareStr2(TextParagraphKey(Arr(UB)), Piv, mSortstyle) = 1 Then SwapPara LB, UB, Arr
                        Exit Sub
                    End If
                Loop Until CompareStr2(TextParagraphKey(Arr(M1)), Piv, mSortstyle) <> 0
                Piv = TextParagraphKey(Arr(M1))
                If M1 > LB Then If CompareStr2(TextParagraphKey(Arr(LB)), Piv, mSortstyle) = -1 Then SwapPara M1, LB, Arr: Piv = TextParagraphKey(Arr(M1))
            Else
                Piv = TextParagraphKey(Arr(M1))
                M1 = LB
                Do While CompareStr2(TextParagraphKey(Arr(M1)), Piv, mSortstyle) = 1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While CompareStr2(TextParagraphKey(Arr(M2)), Piv, mSortstyle) = -1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
       If M1 <> M2 Then SwapPara M1, M2, Arr
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While CompareStr2(TextParagraphKey(Arr(M1)), Piv, mSortstyle) = 1: M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortExtendedDes Arr, LB, M2
    If M1 < UB Then QuickSortExtendedDes Arr, M1, UB
    Exit Sub
abc1:
    
    
End Sub
Sub SwapPara(ByVal order1 As Long, ByVal order2 As Long, Arr() As Long)
Dim paranum1 As Long, paranum2 As Long
If order1 = order2 Then Exit Sub
''If order1 > order2 Then paranum1 = order1: order1 = order2: order2 = paranum1
paranum1 = Arr(order1)
paranum2 = Arr(order2)
Arr(order2) = paranum1  ' now we fix the order in the temporary array.
Arr(order1) = paranum2



UnlinkDocPara paranum2
LinkDocParaAfter paranum2, Arr(order1 - 1)
UnlinkDocPara paranum1
LinkDocParaAfter paranum1, Arr(order2 - 1)



End Sub


Sub SortDoc(KeyStart As Long, order1 As Long, order2 As Long)
Dim ndx() As Long, i As Long
ReDim ndx(DocParagraphs) As Long
If KeyStart < 0 Then
mKeystart = 1
Else
mKeystart = KeyStart
End If
ndx(0) = 0
For i = 1 To UBound(ndx())
ndx(i) = DocParaNext(ndx(i - 1))
Next i
If order1 <= 0 Then order1 = 1
If order2 <= 0 Then order2 = i - 1
QuickSortExtended ndx, order1, order2

End Sub
Sub SortDocDes(KeyStart As Long, order1 As Long, order2 As Long)
Dim ndx() As Long, i As Long
ReDim ndx(DocParagraphs) As Long
If KeyStart < 0 Then
mKeystart = 1
Else
mKeystart = KeyStart
End If
ndx(0) = 0
For i = 1 To UBound(ndx())
ndx(i) = DocParaNext(ndx(i - 1))
Next i
If order1 <= 0 Then order1 = 1
If order2 <= 0 Then order2 = i - 1
QuickSortExtendedDes ndx, order1, order2

End Sub
Public Function Advance(cursor As Long)
If cursor < 0 Then cursor = 0
If cursor > docpara Then cursor = 0: Exit Function
cursor = DocParaNext(cursor)
Advance = cursor
End Function
Public Function BackStep(cursor As Long)
If cursor < 0 Then cursor = 0
If cursor > docpara Then cursor = 0: Exit Function
cursor = DocParaBack(cursor)
BackStep = cursor
End Function

Private Function GetDosPath(LongPath As String) As String

Dim s As String
Dim i As Long
Dim PathLength As Long

        i = Len(LongPath) * 2 + 2

        s = String(1024, 0)

        PathLength = GetShortPathName(StrPtr(LongPath), StrPtr(s), i)

        GetDosPath = Left$(s, PathLength)

End Function
Private Function NeoUnicodeFile(Filename$) As Boolean
Dim hFile
Dim F$
Const FILE_ATTRIBUTE_NORMAL = &H80
Const GENERIC_WRITE = &H40000000
If Left$(Filename$, 2) <> "\\" Then

F$ = "\\?\" + Filename$
Else
F$ = Filename$
End If
On Error Resume Next
hFile = CreateFile(StrPtr(F$), GENERIC_WRITE, ByVal 0, ByVal 0, 2, FILE_ATTRIBUTE_NORMAL, ByVal 0)
''Sleep 1
CloseHandle hFile
NeoUnicodeFile = (GetDosPath(F$) <> "")
End Function
Public Function SaveUnicodeOrAnsi(ByVal Filename As String, Optional mode2save As Long = 0, Optional Append As Boolean = False) As Boolean
' using doc as extension you can read it from word...with automatic conversion to unicode
' OVERWRITE ALWAYS
Dim w As Long, a() As Byte, F$, i As Long, bb As Byte, yesswap As Boolean, lf$
Const Utf8CodePage As Long = 65001
On Error GoTo t12345
'' 0 vbcrlf, 1 vblf, 2 vbcr
lf$ = choose(Abs(ListLoadedType \ 10) + 1, vbCrLf, vbLf, vbCr)
If Abs(mode2save \ 10) > 0 Then
' overide loaded status
        lf$ = choose(Abs(mode2save \ 10) Mod 3 + 1, vbCrLf, vbLf, vbCr)
End If
mode2save = mode2save Mod 10
If Not Append Then
    If Not NeoUnicodeFile(Filename) Then Exit Function
End If
F$ = GetDosPath(Filename)
If Err.Number > 0 Or F$ = vbNullString Then Exit Function
w = FreeFile
Open F$ For Binary As w
' mode2save
' 0 is utf-le
If Append Then Seek #w, LOF(w) + 1
mode2save = mode2save Mod 10
Dim maxmw As Long, iPos As Long
 Dim bbBuf() As Byte, iLen As Long



If Abs(mode2save) = 7 Or mode2save = 0 Then
 a() = ChrW(&HFEFF)
If Not Append And mode2save >= 0 Then Put #w, , a()

ElseIf mode2save = 1 Then
a() = ChrW(&HFFFE) ' big endian...need swap
If Not Append And mode2save > 0 Then Put #w, , a()
yesswap = True
ElseIf Abs(mode2save) = 2 Then  'utf8
If mode2save > 0 And Not Append Then
        Put #w, , CByte(&HEF)
        Put #w, , CByte(&HBB)
        Put #w, , CByte(&HBF)
        End If
     ''   Put #W, , Utf16toUtf8(buf)
            iPos = DocParaNext(0)
            If iPos = 0 Then
            ElseIf DocParaNext(0) = DocParaBack(0) And Me.TextParagraphLen(iPos) = 0 Then
            Else
            maxmw = 327680
        While iPos <> 0
            a() = textArray(lf$, iPos, maxmw)
            iLen = WideCharToMultiByte(Utf8CodePage, 0, VarPtr(a(0)), CLng(UBound(a) - LBound(a) + 1) \ 2, 0, 0, 0, 0)
            If iLen > 0 Then
            ReDim bbBuf(0 To iLen - 1)
            iLen = WideCharToMultiByte(Utf8CodePage, 0, VarPtr(a(0)), CLng(UBound(a) - LBound(a) + 1) \ 2, VarPtr(bbBuf(0)), iLen, 0, 0)
            Put #w, , bbBuf()
            End If
        Wend
     End If
     
        Close w
    SaveUnicodeOrAnsi = True
        Exit Function
ElseIf Abs(mode2save) = 3 Then ' ascii
            iPos = DocParaNext(0)
            maxmw = 327680
If mlcid < 4 Then mlcid = Module1.Clid
            While iPos <> 0
            
                Put #w, , Convert2Ansi(textString(lf$, iPos, maxmw), mlcid)
            Wend
            Close w
            SaveUnicodeOrAnsi = True
            Exit Function
End If


iPos = DocParaNext(0)
maxmw = 327680
If yesswap Then
While iPos <> 0
a() = textArray(lf$, iPos, maxmw)
For i = 0 To UBound(a()) - 1 Step 2
bb = a(i): a(i) = a(i + 1): a(i + 1) = bb
Next i
Put #w, , a()
Wend
Else


While iPos <> 0
a() = textArray(lf$, iPos, maxmw)
Put #w, , a()
Wend
End If
Close w
SaveUnicodeOrAnsi = True
t12345:
End Function
Public Function ContainsUTF8(ByRef Source() As Byte) As Boolean
  Dim i As Long, lUBound As Long, lUBound2 As Long, lUBound3 As Long
  Dim CurByte As Byte
    lUBound = UBound(Source)
    lUBound2 = lUBound - 2
    lUBound3 = lUBound - 3
    If lUBound > 2 Then
    
    For i = 0 To lUBound - 1
      CurByte = Source(i)
        If (CurByte And &HE0) = &HC0 Then
        If (Source(i + 1) And &HC0) = &H80 Then
            ContainsUTF8 = ContainsUTF8 Or True
             i = i + 1
             Else
                ContainsUTF8 = False
                Exit For
            End If
        

        ElseIf (CurByte And &HF0) = &HE0 Then
        ' 2 bytes
        If (Source(i + 1) And &HC0) = &H80 Then
            i = i + 1
            If i < lUBound2 Then
            If (Source(i + 1) And &HC0) = &H80 Then
                ContainsUTF8 = ContainsUTF8 Or True
                i = i + 1
            Else
                ContainsUTF8 = False
                Exit For
            End If
                Else
                ContainsUTF8 = False
                Exit For
            End If
        Else
            ContainsUTF8 = False
            Exit For
        End If
        ElseIf (CurByte And &HF8) = &HF0 Then
        ' 2 bytes
        If (Source(i + 1) And &HC0) = &H80 Then
            i = i + 1
            If i < lUBound2 Then
               If (Source(i + 1) And &HC0) = &H80 Then
                    ContainsUTF8 = ContainsUTF8 Or True
                    i = i + 1
                    If i < lUBound3 Then
                       If (Source(i + 1) And &HC0) = &H80 Then
                            ContainsUTF8 = ContainsUTF8 Or True
                            i = i + 1
                        Else
                            ContainsUTF8 = False
                            Exit For
                        End If
                        
                    Else
                        ContainsUTF8 = False
                        Exit For
                    End If
                Else
                    ContainsUTF8 = False
                    Exit For
                End If
                
            Else
                ContainsUTF8 = False
                Exit For
            End If
        Else
            ContainsUTF8 = False
            Exit For
        End If
        
        
        End If
        
    Next i
    End If
    

End Function
Public Function ContainsUTF16(ByRef Source() As Byte) As Long
  Dim i As Long, lUBound As Long, lUBound2 As Long, lUBound3 As Long
  Dim CurByte As Byte, CurByte1 As Byte
  Dim CurBytes As Long, CurBytes1 As Long
    lUBound = UBound(Source)
    If lUBound > 4 Then
    CurByte = Source(0)
    CurByte1 = Source(1)
    For i = 2 To lUBound - 1 Step 2
        If CurByte1 = 0 And CurByte < 31 Then CurBytes1 = CurBytes1 + 1
        If CurByte = 0 And CurByte1 < 31 Then CurBytes = CurBytes + 1
        If Source(i) = CurByte Then
            CurBytes = CurBytes + 1
        Else
            CurByte = Source(i)
        End If
        If Source(i + 1) = CurByte1 Then
            CurBytes1 = CurBytes1 + 1
        Else
            CurByte1 = Source(i + 1)
        End If
        
    Next i
    End If
    If CurBytes1 = CurBytes And CurBytes1 * 3 >= lUBound Then
    ContainsUTF16 = 0
    Else
    If CurBytes1 * 3 >= lUBound Then
    ContainsUTF16 = 1
    ElseIf CurBytes * 3 >= lUBound Then
    ContainsUTF16 = 2
    Else
    ContainsUTF16 = 0
    End If
    End If
End Function
Public Sub ReadUnicodeOrANSI(Filename As String, Optional Merge As Boolean = False, Optional feedback As Long, Optional insert2index As Boolean = False)
Busy = True
Dim i&, FNr&, BLen&, WChars&, bom As Integer, BTmp As Byte, b() As Byte, lbt As Long
Dim mLof As Long, ok, nobom As Long
nobom = 1
' code from Schmidt, member of vbforums
' this code changed for Document Class
' Change to read line by line, so no two copies needed
If Filename = vbNullString Then Exit Sub
If Not Merge Then EmptyDoc
ReDim Buffer(1 To 1) As Byte
bufOffset = 0
bufsize = 0
SeekBuf = 1
On Error GoTo ErrHandler
  BLen = FileLen(GetDosPath(Filename))
    If BLen = 0 Then
  If DocLines = 0 Then
  textDoc = vbNullString
  End If
  Exit Sub
  End If
  
  FNr = FreeFile
  Open GetDosPath(Filename) For Binary Access Read As FNr
  Get FNr, , bom
    Select Case bom
      Case &HFEFF, &HFFFE 'one of the two possible 16 Bit BOMs
        If BLen >= 3 Then
         NewSeek FNr, 3
          ' now we can read line by line
utf16:
          
           mLof = LOF(FNr)
           While Not (mLof < SeekBuf)
           If insert2index Then
           InsertParagraphOneLine getUniStringlINE(FNr, mLof, bom = &HFEFF, lbt)
           Else
           AppendParagraphOneLine getUniStringlINE(FNr, mLof, bom = &HFEFF, lbt)
           End If
            RaiseEvent MayQuit(ok)
            If ok Then GoTo ErrHandler
          Wend
          
          feedback = nobom * (7 + lbt)
          
        End If
      Case &HBBEF 'the start of a potential UTF8-BOM
        Get FNr, , BTmp
        If BTmp = &HBF Then 'it's indeed the UTF8-BOM
        
          If BLen >= 4 Then
           NewSeek FNr, 4
          ' now we can read line by line
           mLof = LOF(FNr)
utf8cont:
           
           While Not (mLof < SeekBuf)
           If insert2index Then
           InsertParagraphOneLine getUniStringLineUtF8(FNr, mLof, lbt)
           Else
          AppendParagraphOneLine getUniStringLineUtF8(FNr, mLof, lbt)
          End If
            RaiseEvent MayQuit(ok)
            If ok Then GoTo ErrHandler
          Wend
          End If
         feedback = nobom * (2 + lbt)
        Else
       NewSeek FNr, 1
       mLof = LOF(FNr)
        While Not (mLof < SeekBuf)
            If insert2index Then
            InsertParagraphOneLine getAnsiStringlINE(FNr, mLof, lbt)
            Else
           AppendParagraphOneLine getAnsiStringlINE(FNr, mLof, lbt)
           End If
            RaiseEvent MayQuit(ok)
            If ok Then GoTo ErrHandler
          Wend
          feedback = 3 + lbt
        End If
        
      Case Else
      
       NewSeek FNr, 1
       
       mLof = LOF(FNr)
       Dim buf() As Byte
       If mLof > 1000 Then
       ReDim buf(1000)
       Else
       ReDim buf(mLof)
       End If
       Get FNr, 1, buf()
       NewSeek FNr, 1
      If ContainsUTF8(buf()) Then nobom = -1: GoTo utf8cont
' check if it is utf-16
    Select Case ContainsUTF16(buf())
    Case 1
        nobom = -1
        bom = &HFEFF
        NewSeek FNr, 1
        GoTo utf16
    Case 2
        nobom = -1
        bom = &HFFFE
        NewSeek FNr, 1
        GoTo utf16
    End Select
    Dim cc$
    cc$ = space$(30)
    LSet cc$ = buf()
lbt = 0
If mlcid < 4 Then mlcid = Module1.Clid

        While Not (mLof < SeekBuf)
         If insert2index Then
         InsertParagraphOneLine getAnsiStringlINE(FNr, mLof, lbt)
         Else
           AppendParagraphOneLine getAnsiStringlINE(FNr, mLof, lbt)
           End If
            RaiseEvent MayQuit(ok)
            If ok Then GoTo ErrHandler
          Wend
          feedback = 3 + lbt
    End Select

    
ErrHandler:
Busy = False
If FNr Then Close FNr
ReDim Buffer(1 To 1) As Byte
bufOffset = 0
bufsize = 0
SeekBuf = 1
ListLoadedType = lbt \ 10
If Err Then MyEr Err.Description, Err.Description
End Sub
Public Property Get WordCount() As Long
 Dim mycur As Long
mycur = DocParaNext(0)
Do While mycur <> 0
WordCount = WordCount + mWordCount(TextParagraph(mycur))
mycur = DocParaNext(mycur)
Loop
End Property
Private Function mWordCount(ByVal s$) As Long
Dim i&, l As Long
Static p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then Exit Function
  p2 = StrPtr(s): l = l - 1
  p4 = p2 - 2
  For i = p2 To p2 + l * 2 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 65 To 90, 97 To 122, Is > 127
    If i - p4 = 2 Then mWordCount = mWordCount + 1
    Case 39
    If i - p4 = 2 Then p4 = i
    Case Else
    p4 = i
  End Select
  Next i
End Function
Private Function mWordMap(s$) As String
Dim i&, l As Long
Static p2 As Long, p1 As Integer, p3 As Long, p4 As Long
Dim ss$, pos As Long
pos = -2
  l = Len(s): If l = 0 Then mWordMap = " ": Exit Function

  p2 = StrPtr(s$): l = l - 1
  p3 = p2
  p4 = p2 - 2
  For i = p2 To p2 + l * 2 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 65 To 90, 97 To 122, Is > 127
    If i - p4 = 2 Then p2 = i: pos = pos + 2
    Case 39
    If i - p4 = 2 Then p4 = i
    Case Else
    If i - p4 > 2 Then
            CopyMemory ByVal p3 + pos, ByVal p2, i - p2
            pos = pos + i - p2
            PutMem2 p3 + pos, 32   ' place a space...
            
    End If
    p4 = i
  End Select
  Next i
  If i - p4 > 2 Then
        CopyMemory ByVal p3 + pos, ByVal p2, i - p2
        pos = pos + i - p2
  End If
  pos = pos / 2
  If pos <= 0 Then
  mWordMap = " "
  Else
  mWordMap = Left$(s$, pos)
  End If

End Function

Public Sub CompactLines(Optional addtimes As Boolean = True)
Busy = True
Dim mycur As Long, nextCur As Long, s$, many As Long, ok
mycur = DocParaNext(0)
Do While mycur <> 0
      RaiseEvent MayQuit(ok)
      If ok Then Busy = False: Exit Sub
many = 1
s$ = TextParagraph(mycur)
nextCur = DocParaNext(mycur)
Do While nextCur <> 0
If StrComp(s$, TextParagraph(nextCur), vbTextCompare) = 0 Then
many = many + 1
RemoveDocParaIndex nextCur
nextCur = DocParaNext(mycur)
Else
Exit Do
End If
Loop
If addtimes Then TextParagraph(mycur) = TextParagraph(mycur) & " - " & CStr(many)
mycur = nextCur
Loop
Busy = False
End Sub
Public Function UniqueWords(Optional uLCID As Long = -1, Optional nowrap As Boolean = True) As Long
If uLCID = -1 Then uLCID = sortmlcid
 ResetWords
 If nowrap Then
 words$ = textDoc1LINE
 Else
words$ = textDoc
End If
If uLCID = 0 Then uLCID = &H800
unique = CountWords2(words$, uLCID)
UniqueWords = unique
End Function
Private Function CountWords2(a$, uLCID As Long) As Long
Dim l As Long, NORM_IGNORECASE As Long
If Not NoIgnoreCase Then NORM_IGNORECASE = &H1001 Else NORM_IGNORECASE = &H1000
Dim top As Long
l = Len(a$)
If l = 1 Then
CountWords2 = -(Trim(a$) <> "")
top = CountWords2
If top > 0 Then
ReDim pos(1 To 1) As Long, slen(1 To 1) As Long, hit(1 To 1) As Long
pos(1) = StrPtr(a$)
slen(1) = 1
End If
Else
Dim p5 As Long
p5 = 512 '' L \ 32 + 1
ReDim pos(1 To p5) As Long, slen(1 To p5) As Long, hit(1 To p5) As Long
Dim i As Long, k As Long
Dim p2 As Long, p4 As Long, p6 As Long, bb As Integer
p2 = StrPtr(a$)
p3 = p2
Dim md As Long, mda As Long, mtop As Long
l = p2 + 2 * l
p4 = p2
Do
         If p4 >= l Then
         If p4 <> p2 Then GoTo there1
        CountWords2 = top: Exit Function
        End If
            GetMem2 p4, bb
            Select Case bb
            Case 65 To 90, 97 To 122, Is > 127
            Case 39
            If p4 = p2 Then p2 = p4 + 2
            Case Else
            If p4 <> p2 Then
there1:
                k = (p4 - p2) \ 2
                mtop = top
                mda = 0
                md = (mtop - mda + 1) \ 2
                p6 = mda + md
                Do While md > 0
                    Select Case CompareString(uLCID, NORM_IGNORECASE, p2, k, pos(p6), slen(p6))
                        Case 1
                            mtop = mtop - md
                            md = (mtop - mda + 1) \ 2
                        Case 3
                            mda = mda + md
                            md = (mtop - mda + 1) \ 2
                        Case 2
                            hit(p6) = hit(p6) + 1
                            Exit Do
                    End Select
                    p6 = mda + md
                Loop
                If md = 0 Then
                If top = p5 Then
                p5 = p5 * 2
                ReDim Preserve pos(1 To p5) As Long, slen(1 To p5) As Long, hit(1 To p5) As Long
                End If
                    top = top + 1
                    
                    If mda = top - 1 Then
                        pos(top) = p2
                        slen(top) = k
                    Else
                        CopyMemory ByVal VarPtr(pos(1)) + mda * 4 + 4, ByVal VarPtr(pos(1)) + mda * 4, (top - mda - 1) * 4
                        CopyMemory ByVal VarPtr(slen(1)) + mda * 4 + 4, ByVal VarPtr(slen(1)) + mda * 4, (top - mda - 1) * 4
                        CopyMemory ByVal VarPtr(hit(1)) + mda * 4 + 4, ByVal VarPtr(hit(1)) + mda * 4, (top - mda - 1) * 4
                        pos(mda + 1) = p2
                        slen(mda + 1) = k
                        hit(mda + 1) = 0
                    End If
                End If
            End If
            p2 = p4 + 2
            End Select
            p4 = p4 + 2
  Loop
CountWords2 = top
End If
End Function
Private Function MakeString(copythat As Long, ByVal lSize As Long) As String
MakeString = space$(lSize)
CopyMemory ByVal StrPtr(MakeString), ByVal copythat, lSize * 2
End Function
Function ExportWords(Optional usenumbers As Boolean = True) As String()
Dim a() As String
If unique = 0 Then
ExportWords = Split(mWordMap(textDoc))
Exit Function
Else
ReDim a(1 To unique)
Dim i As Long
For i = 1 To unique
a(i) = MakeString(pos(i), slen(i)) & " (" & CStr(hit(i) + 1) & ")"
Next i
ExportWords = a()
End If
End Function
Sub ExportWords2ObjectAddItem(ThisObject1 As Object, Optional usenumbers As Boolean = True)
Dim a() As String, i As Long, THISOBJECT As mStiva
Set THISOBJECT = ThisObject1
If unique = 0 Then
a() = Split(mWordMap(textDoc))
QuickSortString a(), LBound(a()), UBound(a())
For i = LBound(a()) To UBound(a())
THISOBJECT.additem a(i)
Next i
Exit Sub
Else
If usenumbers Then
For i = 1 To unique
THISOBJECT.additem MakeString(pos(i), slen(i)) & " (" & CStr(hit(i) + 1) & ")"
Next i
Else
For i = 1 To unique
THISOBJECT.additem MakeString(pos(i), slen(i))
Next i

End If

End If
End Sub
Public Sub ResetWords()
words$ = vbNullString
p3 = 0
ReDim pos(1 To 1) As Long, slen(1 To 1) As Long, hit(1 To 1) As Long
unique = 0
End Sub
Public Sub QuickSortString(Arr() As String, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
Dim Piv As String, tmp As String '<- adjust types here, when switching to something different than Long
  If UB - LB = 1 Then
     M1 = LB

     If StrComp(Arr(M1), Arr(UB), vbTextCompare) = 1 Then tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = tmp
     Exit Sub
  Else
    M1 = (LB + UB) \ 2
 
    If StrComp(Arr(M1), Arr(LB), vbTextCompare) = 0 Then
      M2 = UB - 1
      M1 = LB
      Do
        M1 = M1 + 1
        If M1 > M2 Then
          If StrComp(Arr(UB), Arr(LB), vbTextCompare) = -1 Then tmp = Arr(LB): Arr(LB) = Arr(UB): Arr(UB) = tmp
          Exit Sub
        End If

      Loop Until StrComp(Arr(M1), Arr(LB), vbTextCompare) <> 0
      Piv = Arr(M1)
      If M1 > LB Then
        If StrComp(Arr(LB), Piv, vbTextCompare) = 1 Then Arr(M1) = Arr(LB): Arr(LB) = Piv: Piv = Arr(M1)
      End If
    Else
      Piv = Arr(M1)
      M1 = LB
      Do While StrComp(Arr(M1), Piv, vbTextCompare) = -1: M1 = M1 + 1: Loop
    End If
  End If
  
  M2 = UB
  Do
    Do While StrComp(Arr(M2), Piv, vbTextCompare) = 1: M2 = M2 - 1: Loop

    If M1 <= M2 Then
      tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = tmp 'swap
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While StrComp(Arr(M1), Piv, vbTextCompare) = -1: M1 = M1 + 1: Loop

  Loop
  If LB < M2 Then QuickSortString Arr, LB, M2
  If M1 < UB Then QuickSortString Arr, M1, UB
End Sub
Private Function getAnsiStringlINE(F As Long, thislof As Long, Optional linebreaktype As Long) As String
Dim a() As Byte, s1 As String, ss As Long, lbreak As Byte, check As Byte
Dim se As Long, st1 As Long
On Error GoTo a11
Do While Not (thislof < SeekBuf)
    GetByte F, SeekBuf, check
    If check <> 10 And check <> 13 Then

                  getAnsiStringlINE = getAnsiStringlINE + ChrW(AscW(StrConv(ChrW(check), vbUnicode, mlcid)))
                
            

    Else
        If Not (thislof < SeekBuf) Then
            ss = SeekBuf
            lbreak = check
            GetByte F, SeekBuf, check
            If check <> 10 And check <> 13 Or lbreak = check Then
                NewSeek F, ss  ' restore it
            If check = 10 Then linebreaktype = 10 Else linebreaktype = 20
            End If
            
        End If
        Exit Do
    End If
Loop
Exit Function
a11:
On Error GoTo 0
SeekBuf = thislof + 1
End Function
Private Function getUniStringlINE(F As Long, thislof As Long, Optional LEndian As Boolean = True, Optional linebreaktype As Long) As String
Dim a() As Byte, s1 As String, ss As Long, lbreak As String, tmp As Byte
a = " "
Dim p As Integer
On Error GoTo a12
Do While Not (thislof < SeekBuf)
    ''Get #f, , a()
    GetWord F, SeekBuf, p
    s1 = ChrW(p)
    If Not LEndian Then
    a() = s1
    tmp = a(0): a(0) = a(1): a(1) = tmp
    s1 = a()
    End If
    If MaxParaLength > 0 Then If Len(getUniStringlINE) > MaxParaLength Then GoTo a12
    If s1 <> vbCr And s1 <> vbLf Then
        getUniStringlINE = getUniStringlINE + s1
     
    Else
    ' linebreak mybe is a CR or LF or CRLF (or LFCR???)
        If Not (thislof < SeekBuf) Then
            ss = SeekBuf
            lbreak = s1
                GetWord F, SeekBuf, p
                s1 = ChrW(p)
                If Not LEndian Then
                a() = s1
                tmp = a(0): a(0) = a(1): a(1) = tmp
                s1 = a()
                End If
            If s1 <> vbCr And s1 <> vbLf Or lbreak = s1 Then
                  NewSeek F, ss
            If s1 = vbLf Then linebreaktype = 10 Else linebreaktype = 20
            End If
        End If
        Exit Do
    End If
Loop
Exit Function
a12:
On Error GoTo 0
SeekBuf = thislof + 1
End Function
Private Function getUniStringLineUtF8(F As Long, thislof As Long, Optional linebreaktype As Long) As String

Dim ss As Long, lbreak As Byte, check As Byte
Dim st As Long, WChars&
Dim shortbuf() As Byte
ReDim shortbuf(1 To 200) As Byte
On Error GoTo a13
Do While Not (thislof < SeekBuf)
tryagain:
    If thislof < SeekBuf Then Exit Do
    GetByte F, SeekBuf, check

    If (check And &H80) <> 0 Then
        st = st + 1
    If st > UBound(shortbuf) Then ReDim Preserve shortbuf(1 To UBound(shortbuf) * 2) As Byte
    shortbuf(st) = check
    GoTo tryagain
    End If
    If check <> 10 And check <> 13 Then
        st = st + 1
    If st > UBound(shortbuf) Then ReDim Preserve shortbuf(1 To UBound(shortbuf) * 2) As Byte
    shortbuf(st) = check
    Else
        If Not (thislof < SeekBuf) Then
            ss = SeekBuf
            lbreak = check
              GetByte F, SeekBuf, check
            If (check And &H80) <> 0 Then GoTo notryother
            
            If (check <> 10) And (check <> 13) Or (lbreak = check) Then
notryother:

                NewSeek F, ss  ' restore it
                If check = 10 Then linebreaktype = 10 Else linebreaktype = 20
     
            End If
            
        End If
        Exit Do
    End If
Loop
If st > 0 Then
            WChars = MultiByteToWideChar(65001, 0, shortbuf(1), st, 0, 0)
            getUniStringLineUtF8 = space$(WChars)
            MultiByteToWideChar 65001, 0, shortbuf(1), st, StrPtr(getUniStringLineUtF8), WChars

End If
Exit Function
a13:
On Error GoTo 0
SeekBuf = thislof + 1
End Function
Sub EmptyDoc()
delDocPara = 0
docpara = 0
docmax = 20  '
ReDim para(1 To docmax)
ReDim DocParaBack(0 To docmax)
ReDim DocParaNext(0 To docmax)
End Sub
Sub FillBuffer(F As Long)
Dim c As Long
If bufsize = 0 Then
bufsize = LOF(F) Mod 327680
If bufsize = 0 Then bufsize = LOF(F)
End If
c = LOF(F) - SeekBuf + 1
If UBound(Buffer()) > 1 Then bufOffset = bufOffset + UBound(Buffer())
If c > 327680 Then
        ReDim Buffer(1 To 327680) As Byte
        Get F, , Buffer
ElseIf c > 0 Then
        ReDim Buffer(1 To c) As Byte
        Get F, , Buffer
End If
End Sub
Sub NewSeek(F As Long, ByVal location As Long)
If bufsize = 0 Then
bufsize = LOF(F) Mod 327680
If bufsize = 0 Then bufsize = LOF(F) Else bufsize = 327680
GoTo fillmeplease
End If
If location - bufOffset > 0 And location <= bufsize + bufOffset Then
Else
fillmeplease:
Seek #F, (location \ bufsize) * bufsize + 1
FillBuffer F
End If
SeekBuf = location
End Sub
Sub GetByte(F As Long, ByVal location As Long, b As Byte)
If location > UBound(Buffer) + bufOffset Then
Seek #F, (location \ bufsize) * bufsize + 1
FillBuffer F
End If
b = Buffer(location - bufOffset)
SeekBuf = location + 1
End Sub
Sub GetWord(F As Long, ByVal location As Long, b As Integer)
If location > UBound(Buffer) + bufOffset Then
Seek #F, (location \ bufsize) * bufsize + 1
FillBuffer F
End If
GetMem2 VarPtr(Buffer(location - bufOffset)), b
SeekBuf = location + 2
End Sub
Function InstrTextCompareAnyLocale(Position As Long, from$, this$, Optional uLCID As Long = 1033) As Long
Dim NORM_IGNORECASE As Long
If Not NoIgnoreCase Then NORM_IGNORECASE = &H1
Dim p1 As Long, p2 As Long, l1 As Long, l2 As Long, cu As Long
l1 = (Len(from$) - Len(this$)) * 2 - (Position - 1) * 2
If l1 < 0 Then Exit Function
p1 = StrPtr(from$) + (Position - 1) * 2
p2 = StrPtr(this$)
l2 = Len(this$)
For cu = p1 To p1 + l1 Step 2
If CompareString(uLCID, NORM_IGNORECASE, cu, l2, p2, l2) = 2 Then
InstrTextCompareAnyLocale = (cu - StrPtr(from$)) \ 2 + 1
Exit Function
End If
Next cu
End Function
Function InstrRevTextCompareAnyLocale(from$, this$, Optional Position As Long = -1, Optional uLCID As Long = 1033) As Long
Dim NORM_IGNORECASE As Long
If Not NoIgnoreCase Then NORM_IGNORECASE = &H1
Dim p1 As Long, p2 As Long, l1 As Long, l2 As Long, cu As Long
If Len(from$) < Position Then Position = -1
If Position > 0 Then
l1 = (Len(from$) - Len(this$)) * 2 - (Len(from$) - Position) * 2
Else
l1 = (Len(from$) - Len(this$)) * 2
End If
If l1 < 0 Then Exit Function
p1 = StrPtr(from$)
p2 = StrPtr(this$)
l2 = Len(this$)
For cu = p1 + l1 To p1 Step -2
If CompareString(uLCID, NORM_IGNORECASE, cu, l2, p2, l2) = 2 Then
InstrRevTextCompareAnyLocale = (cu - p1) \ 2 + 1
Exit Function
End If
Next cu
End Function

Public Property Get lcid() As Long
lcid = mlcid
End Property

Public Property Let lcid(ByVal RHS As Long)
If CLng(RHS) = 0 Then
    mlcid = OsInfo.LangNonUnicodeCode
ElseIf GetCodePage(CLng(RHS)) <> 0 Then
    mlcid = RHS
Else
    NoValidLocale
End If
End Property

Function RangeUnsignInteger(mValue As Integer, ByVal mFrom As Long, ByVal mTo As Long) As Boolean
Dim acc As Long
PutMem2 VarPtr(acc), mValue
RangeUnsignInteger = acc >= Abs(mFrom) And acc <= Abs(mTo)
End Function
Function FindWord(Find$, direction As Boolean, there As Long, cur As Long) As Boolean
Dim oncemore As Boolean, Result As Boolean, res$, code As Integer
again:
If direction Then
 Result = FindStrDown(Find$, there, cur)
 Else
 Result = FindStrUp(Find$, there, cur)
 End If
If Result Then
res$ = TextParagraph(there)
If cur > 1 Then
    code = AscW(Mid$(res$, cur - 1, 1))
    If Not RangeUnsignInteger(code, &H2000&, &H206F&) Then
        Select Case code
        Case 39, 65 To 90, 97 To 122, Is > 127
        If direction Then cur = cur + 1 Else cur = cur - 1
        GoTo again
        Case Else
        End Select
    End If
End If
If (Len(res$) - cur + 1) > Len(Find$) Then
    code = AscW(Mid$(res$, cur + Len(Find$), 1))
    If Not RangeUnsignInteger(code, &H2000&, &H206F&) Then
    Select Case code
    Case 39, 65 To 90, 97 To 122, Is > 127
    
    If direction Then
    cur = cur + 1
    Else
        If cur <= 0 Then
            there = DocParaBack(there)
                If there > 0 Then
                    cur = TextParagraphLen(there)
                Else
                    Result = False
                Exit Function
            End If
        Else
        cur = cur + Len(Find$)
        End If
    End If
    GoTo again
    Case Else
    End Select
    End If
End If
''
FindWord = True
Exit Function
Else
If oncemore Then Exit Function
If direction Then
there = 1
cur = 1
Else
there = DocParagraphs
cur = TextParagraphLen(ParagraphFromOrder(there))
End If
oncemore = True
End If
End Function

Function FindIdentifier(Find$, direction As Boolean, there As Long, cur As Long) As Boolean
' there is paragraph number not paragraph order
Dim oncemore As Boolean, Result As Boolean, res$
If Len(Find$) > 1 And Left$(Find$, 1) = "&" Then Find$ = Mid$(Find$, 2): cur = cur + 1
If InStr(Find$, Chr$(13)) > 0 Then Exit Function
If InStr(Find$, Chr$(10)) > 0 Then Exit Function
again:
If direction Then
 Result = FindStrDown(Find$, there, cur)
 Else
 Result = FindStrUp(Find$, there, cur)
 End If
If Result Then
res$ = TextParagraph(there)
If cur > 1 Then
If AscW(Mid$(res$, cur, 1)) = 35 Then
Else
    Select Case AscW(Mid$(res$, cur - 1, 1))
    Case 160
    Case 35, 46, 65 To 90, 95, 97 To 122, Is > 127
    If direction Then cur = cur + 1 Else cur = cur - 1
    GoTo again
    Case Else
    End Select
    End If
End If
If Find$ = "" Then Exit Function
Select Case AscW(Right$(Find$, 1))
Case 36, 37, 40
Case Else
If (Len(res$) - cur + 1) > Len(Find$) Then
    Select Case AscW(Mid$(res$, cur + Len(Find$), 1))
    Case 40, 46, 48 To 57, 65 To 90, 95, 97 To 122, Is > 127
    If direction Then
    cur = cur + 1
    Else
        If cur <= Len(Find$) Then
            there = DocParaBack(there)
                If there > 0 Then
                    cur = TextParagraphLen(there)
                Else
                    Result = False
                Exit Function
            End If
        Else
        cur = cur + Len(Find$)
        End If
    End If
    GoTo again
    Case Else
    End Select
End If
End Select
''
FindIdentifier = True
Exit Function
Else
If oncemore Then Exit Function
If direction Then
there = 1
cur = 1
Else
there = DocParagraphs
If there <> 0 Then
cur = TextParagraphLen(ParagraphFromOrder(there))
End If
End If
oncemore = True
End If
End Function
Public Function NextTextParagraph(ByVal ParaNum As Long, nextparanum As Long) As String
If Not IsLast(ParaNum) Then
nextparanum = DocParaNext(ParaNum)
NextTextParagraph = TextParagraph(nextparanum)
End If
End Function

Function CompareStr2(a$, b$, what As Long) As Long
On Error GoTo comperr
Select Case what
Case 1 To 2
CompareStr2 = StrComp(a$, b$, what)
Case Else
If sortmlcid = 0 Or UserCodePage = DefCodePage Then
CompareStr2 = StrComp(a$, b$)
Else
CompareStr2 = -2 + CompareString(sortmlcid, &H1000, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))
End If
End Select
Exit Function
comperr:
CompareStr2 = -2 + CompareString(0, 0, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))

End Function

Private Function countbrakcet(a$, k As Long, counter As Long) As Long
Dim i As Long, j As Integer, blind As Boolean
countbrakcet = counter
For i = Abs(k) To Len(a$)
j = AscW(Mid$(a$, i, 1))
Select Case j
Case 34
blind = Not blind
Case 123
If Not blind Then countbrakcet = countbrakcet + 1
Case 125
If Not blind Then countbrakcet = countbrakcet - 1: If countbrakcet = 0 Then k = i: Exit For
End Select
Next i

End Function
Property Get IsPart(ParaNum As Long) As Boolean
If Not InvalidPara(ParaNum) Then
IsPart = para(ParaNum).startfrom > 0
End If
End Property
