Attribute VB_Name = "mdQRCodegen"
'=========================================================================
'
' QR Code generator library (VB6/VBA)
'
' Copyright (c) Project Nayuki. (MIT License)
' https://www.nayuki.io/page/qr-code-generator-library
'
' Copyright (c) wqweto@gmail.com (MIT License)
'
'=========================================================================
Option Explicit
DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0) Or (TWINBASIC <> 0)

'=========================================================================
' Public enums
'=========================================================================

Public Enum QRCodegenEcc
    QRCodegenEcc_LOW = 0  ' The QR Code can tolerate about  7% erroneous codewords
    QRCodegenEcc_MEDIUM   ' The QR Code can tolerate about 15% erroneous codewords
    QRCodegenEcc_QUARTILE ' The QR Code can tolerate about 25% erroneous codewords
    QRCodegenEcc_HIGH     ' The QR Code can tolerate about 30% erroneous codewords
End Enum

Public Enum QRCodegenMask
    QRCodegenMask_AUTO = -1
    QRCodegenMask_0 = 0
    QRCodegenMask_1
    QRCodegenMask_2
    QRCodegenMask_3
    QRCodegenMask_4
    QRCodegenMask_5
    QRCodegenMask_6
    QRCodegenMask_7
End Enum

Public Enum QRCodegenMode
    QRCodegenMode_NUMERIC = &H1
    QRCodegenMode_ALPHANUMERIC = &H2
    QRCodegenMode_BYTE = &H4
    QRCodegenMode_KANJI = &H8
    QRCodegenMode_ECI = &H7
End Enum

Public Type QRCodegenSegment
    Mode            As QRCodegenMode
    NumChars        As Long
    Data()          As Byte
    BitLength       As Long
End Type

'=========================================================================
' API
'=========================================================================

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileW" (ByVal hdcRef As LongPtr, ByVal lpFileName As LongPtr, ByVal lpRect As LongPtr, ByVal lpDescription As LongPtr) As Longptr
Private Declare PtrSafe Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As LongPtr, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileW" (ByVal hdcRef As LongPtr, ByVal lpFileName As LongPtr, ByVal lpRect As LongPtr, ByVal lpDescription As LongPtr) As LongPtr
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As LongPtr, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#End If

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type PICTDESC
    Size                As Long
    Type                As Long
    hBmpOrIcon          As LongPtr
    hPal                As LongPtr
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const VERSION_MIN               As Long = 1
Private Const VERSION_MAX               As Long = 40
Private Const PENALTY_N1                As Long = 3
Private Const PENALTY_N2                As Long = 3
Private Const PENALTY_N3                As Long = 40
Private Const PENALTY_N4                As Long = 10
Private Const INT16_MAX                 As Long = 32767
Private Const LONG_MAX                  As Long = 2147483647
Private Const ALPHANUMERIC_CHARSET      As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:"

Private LNG_POW2(0 To 31)                            As Long
Private ECC_CODEWORDS_PER_BLOCK(0 To 3, 0 To 40)     As Long
Private NUM_ERROR_CORRECTION_BLOCKS(0 To 3, 0 To 40) As Long

'=========================================================================
' Functions
'=========================================================================

Public Function QRCodegenBarcode(TextOrByteArray As Variant, _
            Optional ByVal ForeColor As OLE_COLOR = vbBlack, _
            Optional ByVal ModuleSize As Long = 10, _
            Optional ByVal Ecl As QRCodegenEcc = QRCodegenEcc_LOW, _
            Optional ByVal MinVersion As Long = VERSION_MIN, _
            Optional ByVal MaxVersion As Long = VERSION_MAX, _
            Optional ByVal Mask As QRCodegenMask = QRCodegenMask_AUTO, _
            Optional ByVal BoostEcl As Boolean = True) As StdPicture
    Dim baQrCode()      As Byte
    
    If QRCodegenEncode(TextOrByteArray, baQrCode, Ecl, MinVersion, MaxVersion, Mask, BoostEcl) Then
        Set QRCodegenBarcode = QRCodegenConvertToPicture(baQrCode, ForeColor, ModuleSize)
    End If
End Function

Public Function QRCodegenEncode(TextOrByteArray As Variant, baQrCode() As Byte, _
            Optional ByVal Ecl As QRCodegenEcc = QRCodegenEcc_LOW, _
            Optional ByVal MinVersion As Long = VERSION_MIN, _
            Optional ByVal MaxVersion As Long = VERSION_MAX, _
            Optional ByVal Mask As QRCodegenMask = QRCodegenMask_AUTO, _
            Optional ByVal BoostEcl As Boolean = True) As Boolean
    Dim baData()        As Byte
    Dim lDataLen        As Long
    Dim sText           As String
    Dim lBufLen         As Long
    Dim uSegments()     As QRCodegenSegment
    
    pvInit
    If IsArray(TextOrByteArray) Then
        baData = TextOrByteArray
        lDataLen = UBound(baData) + 1
    Else
        sText = TextOrByteArray
        lDataLen = Len(sText)
    End If
    If lDataLen = 0 Then
        ReDim uSegments(-1 To -1) As QRCodegenSegment
    Else
        ReDim uSegments(0 To 0) As QRCodegenSegment
        lBufLen = pvGetBufferLenForVersion(MaxVersion)
        If IsArray(TextOrByteArray) Then
            If QRCodegenCalcSegmentBufferSize(QRCodegenMode_BYTE, lDataLen) > lBufLen Then
                GoTo QH
            End If
            uSegments(0) = QRCodegenMakeBytes(baData)
        ElseIf QRCodegenIsNumeric(sText) Then
            If QRCodegenCalcSegmentBufferSize(QRCodegenMode_NUMERIC, lDataLen) > lBufLen Then
                GoTo QH
            End If
            uSegments(0) = QRCodegenMakeNumeric(sText)
        ElseIf QRCodegenIsAlphanumeric(sText) Then
            If QRCodegenCalcSegmentBufferSize(QRCodegenMode_ALPHANUMERIC, lDataLen) > lBufLen Then
                GoTo QH
            End If
            uSegments(0) = QRCodegenMakeAlphanumeric(sText)
        Else
            baData = pvToUtf8Array(sText)
            lDataLen = UBound(baData) + 1
            If QRCodegenCalcSegmentBufferSize(QRCodegenMode_BYTE, lDataLen) > lBufLen Then
                GoTo QH
            End If
            uSegments(0) = QRCodegenMakeBytes(baData)
        End If
    End If
    QRCodegenEncode = QRCodegenEncodeSegments(uSegments, baQrCode, Ecl, MinVersion, MaxVersion, Mask, BoostEcl)
    Exit Function
QH:
    ReDim baQrCode(0 To 0) As Byte
End Function

Public Function QRCodegenEncodeSegments(uSegments() As QRCodegenSegment, baQrCode() As Byte, _
            Optional ByVal Ecl As QRCodegenEcc = QRCodegenEcc_LOW, _
            Optional ByVal MinVersion As Long = VERSION_MIN, _
            Optional ByVal MaxVersion As Long = VERSION_MAX, _
            Optional ByVal Mask As QRCodegenMask = QRCodegenMask_AUTO, _
            Optional ByVal BoostEcl As Boolean = True) As Boolean
    Dim lVersion        As Long
    Dim lDataUsedBits   As Long
    Dim lDataCapacityBits As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lBitLen         As Long
    Dim lBit            As Long
    Dim lTerminatorBits As Long
    Dim lPadByte        As Long
    Dim baTemp()        As Byte
    Dim lMinPenalty     As Long
    Dim lPenalty        As Long
    
    pvInit
    '--- Find the minimal version number to use
    For lVersion = MinVersion To MaxVersion
        lDataCapacityBits = pvGetNumDataCodewords(lVersion, Ecl) * 8
        lDataUsedBits = pvGetTotalBits(uSegments, lVersion)
        If lDataUsedBits <> -1 And lDataUsedBits <= lDataCapacityBits Then
            Exit For
        End If
        If lVersion >= MaxVersion Then
            baQrCode = vbNullString
            GoTo QH
        End If
    Next
    Debug.Assert lDataUsedBits <> -1
    '--- Increase the error correction level while the data still fits in the current version number
    If BoostEcl Then
        For lIdx = QRCodegenEcc_MEDIUM To QRCodegenEcc_HIGH
            If lDataUsedBits <= pvGetNumDataCodewords(lVersion, lIdx) * 8 Then
                Ecl = lIdx
            End If
        Next
    End If
    '--- Concatenate all segments to create the data bit string
    ReDim baQrCode(0 To pvGetBufferLenForVersion(lVersion) - 1) As Byte
    For lIdx = 0 To UBound(uSegments)
        With uSegments(lIdx)
            pvAppendBitsToBuffer .Mode, 4, baQrCode, lBitLen
            pvAppendBitsToBuffer .NumChars, pvNumCharCountBits(.Mode, lVersion), baQrCode, lBitLen
            For lJdx = 0 To .BitLength - 1
                lBit = -((.Data(lJdx \ 8) And LNG_POW2(7 - (lJdx And 7))) <> 0)
                pvAppendBitsToBuffer lBit, 1, baQrCode, lBitLen
            Next
        End With
    Next
    Debug.Assert lBitLen = lDataUsedBits
    '--- Add terminator and pad up to a byte if applicable
    lDataCapacityBits = pvGetNumDataCodewords(lVersion, Ecl) * 8
    Debug.Assert lBitLen <= lDataCapacityBits
    lTerminatorBits = lDataCapacityBits - lBitLen
    If lTerminatorBits > 4 Then
        lTerminatorBits = 4
    End If
    pvAppendBitsToBuffer 0, lTerminatorBits, baQrCode, lBitLen
    pvAppendBitsToBuffer 0, (8 - (lBitLen And 7)) And 7, baQrCode, lBitLen
    Debug.Assert lBitLen Mod 8 = 0
    '--- Pad with alternating bytes until data capacity is reached
    lPadByte = &HEC
    Do While lBitLen < lDataCapacityBits
        pvAppendBitsToBuffer lPadByte, 8, baQrCode, lBitLen
        lPadByte = lPadByte Xor (&HEC Xor &H11)
    Loop
    '--- Compute ECC, draw modules
    pvAddEccAndInterleave baQrCode, lVersion, Ecl, baTemp
    pvInitializeFunctionModules lVersion, baQrCode
    Debug.Assert UBound(baTemp) + 1 = pvGetNumRawDataModules(lVersion) \ 8
    pvDrawCodewords baTemp, baQrCode
    pvDrawLightFunctionModules lVersion, baQrCode
    pvInitializeFunctionModules lVersion, baTemp
    '--- Do masking
    If Mask = QRCodegenMask_AUTO Then
        lMinPenalty = LONG_MAX
        For lIdx = QRCodegenMask_0 To QRCodegenMask_7
            pvApplyMask baTemp, baQrCode, lIdx
            pvDrawFormatBits Ecl, lIdx, baQrCode
            lPenalty = pvGetPenaltyScore(baQrCode)
            If lPenalty < lMinPenalty Then
                Mask = lIdx
                lMinPenalty = lPenalty
            End If
            pvApplyMask baTemp, baQrCode, lIdx '--- Undoes the mask due to XOR
        Next
    End If
    Debug.Assert QRCodegenMask_0 <= Mask And Mask <= QRCodegenMask_7
    pvApplyMask baTemp, baQrCode, Mask
    pvDrawFormatBits Ecl, Mask, baQrCode
    '--- success
    QRCodegenEncodeSegments = True
QH:
End Function

Public Function QRCodegenConvertToPicture(baQrCode() As Byte, _
            Optional ByVal ForeColor As OLE_COLOR = vbBlack, _
            Optional ByVal ModuleSize As Long = 10) As StdPicture
    Const WHITE_BRUSH   As Long = 0
    Const vbPicTypeEMetafile As Long = 4
    Dim lQrSize         As Long
    Dim lY              As Long
    Dim lX              As Long
    Dim hDC             As LongPtr
    Dim uRect           As RECT
    Dim hBrush          As LongPtr
    Dim uDesc           As PICTDESC
    Dim aGUID(0 To 3)   As Long
    Dim vErr            As Variant
    
    On Error GoTo EH
    lQrSize = QRCodegenGetSize(baQrCode)
    hDC = CreateEnhMetaFile(0, 0, 0, 0)
    uRect.Right = lQrSize * ModuleSize
    uRect.Bottom = lQrSize * ModuleSize
    Call FillRect(hDC, uRect, GetStockObject(WHITE_BRUSH))
    hBrush = CreateSolidBrush(ForeColor)
    For lY = 0 To lQrSize - 1
        For lX = 0 To lQrSize - 1
            If QRCodegenGetModule(baQrCode, lX, lY) Then
                uRect.Left = lX * ModuleSize
                uRect.Right = uRect.Left + ModuleSize
                uRect.Top = lY * ModuleSize
                uRect.Bottom = uRect.Top + ModuleSize
                Call FillRect(hDC, uRect, hBrush)
            End If
        Next
    Next
    '--- fill struct
    With uDesc
        .Size = LenB(uDesc)
        .Type = vbPicTypeEMetafile
        .hBmpOrIcon = CloseEnhMetaFile(hDC)
    End With
    hDC = 0
    '--- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    '--- Create picture from bitmap handle
    Call OleCreatePictureIndirect(uDesc, aGUID(0), True, QRCodegenConvertToPicture)
QH:
    If hBrush <> 0 Then
        Call DeleteObject(hBrush)
        hBrush = 0
    End If
    If IsArray(vErr) Then
        On Error GoTo 0
        Err.Raise vErr(0), vErr(1), vErr(2)
    End If
    Exit Function
EH:
    vErr = Array(Err.Number, Err.Source, Err.Description)
    Resume QH
End Function

Public Function QRCodegenDebugDump(baQrCode() As Byte) As String
    Dim lQrSize         As Long
    Dim aRows()         As String
    Dim lY              As Long
    Dim lX              As Long
    
    lQrSize = QRCodegenGetSize(baQrCode)
    ReDim aRows(0 To lQrSize - 1) As String
    For lY = 0 To lQrSize - 1
        For lX = 0 To lQrSize - 1
            aRows(lY) = aRows(lY) & IIf(QRCodegenGetModule(baQrCode, lX, lY), "##", "  ")
        Next
        aRows(lY) = RTrim$(aRows(lY))
    Next
    QRCodegenDebugDump = Join(aRows, vbCrLf)
End Function

Public Function QRCodegenGetSize(baQrCode() As Byte) As Long
    QRCodegenGetSize = baQrCode(0)
    Debug.Assert VERSION_MIN * 4 + 17 <= QRCodegenGetSize And QRCodegenGetSize <= VERSION_MAX * 4 + 17
End Function

Public Function QRCodegenGetModule(baQrCode() As Byte, ByVal lX As Long, ByVal lY As Long) As Long
    Dim lQrSize         As Long
    
    lQrSize = baQrCode(0)
    If 0 <= lX And lX < lQrSize And 0 <= lY And lY < lQrSize Then
        QRCodegenGetModule = pvGetModuleBounded(baQrCode, lX, lY)
    End If
End Function

Public Function QRCodegenIsNumeric(sText As String) As Boolean
    If LenB(sText) <> 0 Then
        QRCodegenIsNumeric = Not (sText Like "*[!0-9]*")
    End If
End Function

Public Function QRCodegenIsAlphanumeric(sText As String) As Boolean
    Dim lIdx            As Long
    
    If LenB(sText) <> 0 Then
        For lIdx = 1 To Len(sText)
            If InStr(Mid$(sText, lIdx, 1), ALPHANUMERIC_CHARSET) = 0 Then
                Exit Function
            End If
        Next
    End If
    QRCodegenIsAlphanumeric = True
End Function

Public Function QRCodegenCalcSegmentBufferSize(ByVal eMode As QRCodegenMode, ByVal lNumChars As Long) As Long
    Dim lSize           As Long
    
    lSize = pvCalcSegmentBitLength(eMode, lNumChars)
    If lSize = -1 Then
        QRCodegenCalcSegmentBufferSize = LONG_MAX
    Else
        Debug.Assert 0 <= lSize And lSize < INT16_MAX
        QRCodegenCalcSegmentBufferSize = (lSize + 7) \ 8
    End If
End Function

Public Function QRCodegenMakeBytes(baData() As Byte) As QRCodegenSegment
    With QRCodegenMakeBytes
        .Mode = QRCodegenMode_BYTE
        .BitLength = pvCalcSegmentBitLength(.Mode, UBound(baData) + 1)
        Debug.Assert .BitLength <> -1
        .NumChars = UBound(baData) + 1
        .Data = baData
    End With
End Function

Public Function QRCodegenMakeNumeric(sDigits As String) As QRCodegenSegment
    Dim lLen            As Long
    Dim lBitLen         As Long
    Dim lAccumData      As Long
    Dim lAccumCount     As Long
    Dim lIdx            As Long
    Dim lDigit          As Long
    
    With QRCodegenMakeNumeric
        lLen = Len(sDigits)
        .Mode = QRCodegenMode_NUMERIC
        lBitLen = pvCalcSegmentBitLength(.Mode, lLen)
        Debug.Assert lBitLen <> -1
        ReDim .Data(0 To (lBitLen + 7) \ 8 - 1) As Byte
        .NumChars = lLen
        For lIdx = 1 To lLen
            lDigit = Asc(Mid$(sDigits, lIdx, 1)) - 48    '--- Asc("0") = 48
            Debug.Assert 0 <= lDigit And lDigit <= 9
            lAccumData = lAccumData * 10 + lDigit
            lAccumCount = lAccumCount + 1
            If lAccumCount = 3 Then
                pvAppendBitsToBuffer lAccumData, 10, .Data, .BitLength
                lAccumData = 0
                lAccumCount = 0
            End If
        Next
        If lAccumCount > 0 Then
            pvAppendBitsToBuffer lAccumData, lAccumCount * 3 + 1, .Data, .BitLength
        End If
        Debug.Assert lBitLen = .BitLength
    End With
End Function

Public Function QRCodegenMakeAlphanumeric(sText As String) As QRCodegenSegment
    Dim lLen            As Long
    Dim lBitLen         As Long
    Dim lAccumData      As Long
    Dim lAccumCount     As Long
    Dim lIdx            As Long
    Dim lChar          As Long
    
    With QRCodegenMakeAlphanumeric
        lLen = Len(sText)
        .Mode = QRCodegenMode_ALPHANUMERIC
        lBitLen = pvCalcSegmentBitLength(.Mode, lLen)
        Debug.Assert lBitLen <> -1
        ReDim .Data(0 To (lBitLen + 7) \ 8 - 1) As Byte
        .NumChars = lLen
        For lIdx = 1 To lLen
            lChar = InStr(Mid$(sText, lIdx, 1), ALPHANUMERIC_CHARSET) - 1
            Debug.Assert 0 <= lChar
            lAccumData = lAccumData * 45 + lChar
            lAccumCount = lAccumCount + 1
            If lAccumCount = 2 Then
                pvAppendBitsToBuffer lAccumData, 11, .Data, .BitLength
                lAccumData = 0
                lAccumCount = 0
            End If
        Next
        If lAccumCount > 0 Then
            pvAppendBitsToBuffer lAccumData, 6, .Data, .BitLength
        End If
        Debug.Assert lBitLen = .BitLength
    End With
End Function

Public Function QRCodegenMakeEci(ByVal lAssignVal As Long) As QRCodegenSegment
    With QRCodegenMakeEci
        .Mode = QRCodegenMode_ECI
        ReDim .Data(0 To 2) As Byte
        If lAssignVal < 0 Then
            Debug.Assert False
        ElseIf lAssignVal < LNG_POW2(7) Then
            pvAppendBitsToBuffer lAssignVal, 8, .Data, .BitLength
        ElseIf lAssignVal < LNG_POW2(14) Then
            pvAppendBitsToBuffer 2, 2, .Data, .BitLength
            pvAppendBitsToBuffer lAssignVal, 14, .Data, .BitLength
        ElseIf lAssignVal < 1000000 Then
            pvAppendBitsToBuffer 6, 3, .Data, .BitLength
            pvAppendBitsToBuffer lAssignVal \ LNG_POW2(10), 11, .Data, .BitLength
            pvAppendBitsToBuffer lAssignVal And &H3FF, 10, .Data, .BitLength
        Else
            Debug.Assert False
        End If
    End With
End Function

'= private ===============================================================

Private Sub pvInit()
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
    If ECC_CODEWORDS_PER_BLOCK(0, 0) <> 0 Then
        Exit Sub
    End If
    LNG_POW2(0) = 1
    For lIdx = 1 To UBound(LNG_POW2) - 1
        LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
    Next
    LNG_POW2(31) = &H80000000
    vSplit = Split("-1| 7|10|15|20|26|18|20|24|30|18|20|24|26|30|22|24|28|30|28|28|28|28|30|30|26|28|30|30|30|30|30|30|30|30|30|30|30|30|30|30|" & _
                   "-1|10|16|26|18|24|16|18|22|22|26|30|22|22|24|24|28|28|26|26|26|26|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|28|" & _
                   "-1|13|22|18|26|18|24|18|22|20|24|28|26|24|20|30|24|28|28|26|30|28|30|30|30|30|28|30|30|30|30|30|30|30|30|30|30|30|30|30|30|" & _
                   "-1|17|28|22|16|22|28|26|26|24|28|24|28|22|24|24|30|28|28|26|28|30|24|30|30|30|30|30|30|30|30|30|30|30|30|30|30|30|30|30|30", "|")
    For lIdx = 0 To UBound(vSplit)
        ECC_CODEWORDS_PER_BLOCK(lIdx \ 41, lIdx Mod 41) = vSplit(lIdx)
    Next
    vSplit = Split("-1|1|1|1|1|1|2|2|2|2|4| 4| 4| 4| 4| 6| 6| 6| 6| 7| 8| 8| 9| 9|10|12|12|12|13|14|15|16|17|18|19|19|20|21|22|24|25|" & _
                   "-1|1|1|1|2|2|4|4|4|5|5| 5| 8| 9| 9|10|10|11|13|14|16|17|17|18|20|21|23|25|26|28|29|31|33|35|37|38|40|43|45|47|49|" & _
                   "-1|1|1|2|2|4|4|6|6|8|8| 8|10|12|16|12|17|16|18|21|20|23|23|25|27|29|34|34|35|38|40|43|45|48|51|53|56|59|62|65|68|" & _
                   "-1|1|1|2|4|4|4|5|6|8|8|11|11|16|16|18|16|19|21|25|25|25|34|30|32|35|37|40|42|45|48|51|54|57|60|63|66|70|74|77|81", "|")
    For lIdx = 0 To UBound(vSplit)
        NUM_ERROR_CORRECTION_BLOCKS(lIdx \ 41, lIdx Mod 41) = vSplit(lIdx)
    Next
End Sub

Private Function pvGetNumDataCodewords(ByVal lVersion As Long, ByVal eEcl As QRCodegenEcc) As Long
    Debug.Assert QRCodegenEcc_LOW <= eEcl And eEcl <= QRCodegenEcc_HIGH
    pvGetNumDataCodewords = pvGetNumRawDataModules(lVersion) \ 8 - ECC_CODEWORDS_PER_BLOCK(eEcl, lVersion) * NUM_ERROR_CORRECTION_BLOCKS(eEcl, lVersion)
End Function
 
Private Function pvGetNumRawDataModules(ByVal lVersion As Long) As Long
    Dim lNumAlign       As Long
    
    Debug.Assert VERSION_MIN <= lVersion And lVersion <= VERSION_MAX
    pvGetNumRawDataModules = (16 * lVersion + 128) * lVersion + 64
    If lVersion >= 2 Then
        lNumAlign = lVersion \ 7 + 2
        pvGetNumRawDataModules = pvGetNumRawDataModules - (25 * lNumAlign - 10) * lNumAlign + 55
        If lVersion >= 7 Then
            pvGetNumRawDataModules = pvGetNumRawDataModules - 36
        End If
    End If
    Debug.Assert 208 <= pvGetNumRawDataModules And pvGetNumRawDataModules <= 29648
End Function
 
Private Function pvGetTotalBits(uSegments() As QRCodegenSegment, ByVal lVersion As Long) As Long
    Dim lIdx            As Long
    Dim lCcBits         As Long
    
    For lIdx = 0 To UBound(uSegments)
        lCcBits = pvNumCharCountBits(uSegments(lIdx).Mode, lVersion)
        Debug.Assert 0 <= lCcBits And lCcBits <= 16
        If uSegments(lIdx).NumChars >= LNG_POW2(lCcBits) Then
            pvGetTotalBits = -1
            GoTo QH
        End If
        pvGetTotalBits = pvGetTotalBits + 4 + lCcBits + uSegments(lIdx).BitLength
        If pvGetTotalBits > INT16_MAX Then
            pvGetTotalBits = -1
            GoTo QH
        End If
    Next
    Debug.Assert 0 <= pvGetTotalBits And pvGetTotalBits <= INT16_MAX
QH:
End Function

Private Function pvNumCharCountBits(ByVal eMode As QRCodegenMode, ByVal lVersion As Long) As Long
    Dim lIdx            As Long
    
    lIdx = (lVersion + 7) \ 17 + 1
    Select Case eMode
    Case QRCodegenMode_NUMERIC
        pvNumCharCountBits = Choose(lIdx, 10, 12, 14)
    Case QRCodegenMode_ALPHANUMERIC
        pvNumCharCountBits = Choose(lIdx, 9, 11, 13)
    Case QRCodegenMode_BYTE
        pvNumCharCountBits = Choose(lIdx, 8, 16, 16)
    Case QRCodegenMode_KANJI
        pvNumCharCountBits = Choose(lIdx, 0, 10, 12)
    Case QRCodegenMode_ECI
        pvNumCharCountBits = 0
    Case Else
        Debug.Assert False
    End Select
End Function

Private Sub pvAddEccAndInterleave(baData() As Byte, ByVal lVersion As Long, ByVal eEcl As QRCodegenEcc, baResult() As Byte)
    Dim lNumBlocks      As Long
    Dim lBlockEccLen    As Long
    Dim lRawCodewords   As Long
    Dim lDataLen        As Long
    Dim lNumShortBlocks As Long
    Dim lShortBlockDataLen  As Long
    Dim baDiv()         As Byte
    Dim lIdx            As Long
    Dim lBlockPos       As Long
    Dim lBlockLen       As Long
    Dim baEcc()         As Byte
    Dim lJdx            As Long
    Dim lKdx            As Long
    
    Debug.Assert VERSION_MIN <= lVersion And lVersion <= VERSION_MAX
    Debug.Assert QRCodegenEcc_LOW <= eEcl And eEcl <= QRCodegenEcc_HIGH
    lNumBlocks = NUM_ERROR_CORRECTION_BLOCKS(eEcl, lVersion)
    lBlockEccLen = ECC_CODEWORDS_PER_BLOCK(eEcl, lVersion)
    lRawCodewords = pvGetNumRawDataModules(lVersion) \ 8
    lDataLen = pvGetNumDataCodewords(lVersion, eEcl)
    lNumShortBlocks = lNumBlocks - lRawCodewords Mod lNumBlocks
    lShortBlockDataLen = lRawCodewords \ lNumBlocks - lBlockEccLen
    ReDim baResult(0 To lDataLen + lBlockEccLen * lNumBlocks - 1) As Byte
    pvReedSolomonComputeDivisor lBlockEccLen, baDiv
    For lIdx = 0 To lNumBlocks - 1
        lBlockLen = lShortBlockDataLen + IIf(lIdx < lNumShortBlocks, 0, 1)
        pvReedSolomonComputeRemainder baData, lBlockPos, lBlockLen, baDiv, lBlockEccLen, baEcc
        lKdx = lIdx
        For lJdx = 0 To lBlockLen - 1
            If lJdx = lShortBlockDataLen Then
                lKdx = lKdx - lNumShortBlocks
            End If
            baResult(lKdx) = baData(lBlockPos + lJdx)
            lKdx = lKdx + lNumBlocks
        Next
        lKdx = lDataLen + lIdx
        For lJdx = 0 To lBlockEccLen - 1
            baResult(lKdx) = baEcc(lJdx)
            lKdx = lKdx + lNumBlocks
        Next
        lBlockPos = lBlockPos + lBlockLen
    Next
End Sub

Private Sub pvReedSolomonComputeDivisor(ByVal lDegree As Long, baResult() As Byte)
    Dim lRoot           As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    ReDim baResult(0 To lDegree - 1) As Byte
    baResult(lDegree - 1) = 1
    lRoot = 1
    For lIdx = 0 To lDegree - 1
        For lJdx = 0 To lDegree - 1
            baResult(lJdx) = pvReedSolomonMultiply(baResult(lJdx), lRoot)
            If lJdx + 1 < lDegree Then
                baResult(lJdx) = baResult(lJdx) Xor baResult(lJdx + 1)
            End If
        Next
        lRoot = pvReedSolomonMultiply(lRoot, 2)
    Next
End Sub

Private Sub pvReedSolomonComputeRemainder( _
            baData() As Byte, ByVal lDataPos As Long, ByVal lDataSize As Long, _
            baGen() As Byte, ByVal lDegree As Long, baResult() As Byte)
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim bFactor         As Byte
    
    ReDim baResult(0 To lDegree - 1) As Byte
    For lIdx = lDataPos To lDataPos + lDataSize - 1
        bFactor = baData(lIdx) Xor baResult(0)
        Call CopyMemory(baResult(0), baResult(1), lDegree - 1)
        baResult(lDegree - 1) = 0
        For lJdx = 0 To lDegree - 1
            baResult(lJdx) = baResult(lJdx) Xor pvReedSolomonMultiply(baGen(lJdx), bFactor)
        Next
    Next
End Sub

Private Function pvReedSolomonMultiply(ByVal bX As Byte, ByVal bY As Byte) As Byte
    Dim lIdx            As Long
    Dim lTemp           As Long
    
    For lIdx = 7 To 0 Step -1
        If (pvReedSolomonMultiply And &H80) <> 0 Then
            lTemp = &H11D
        Else
            lTemp = 0
        End If
        pvReedSolomonMultiply = ((pvReedSolomonMultiply * 2) Xor lTemp) And &HFF
        If (bY And LNG_POW2(lIdx)) <> 0 Then
            pvReedSolomonMultiply = pvReedSolomonMultiply Xor bX
        End If
    Next
End Function

Private Sub pvInitializeFunctionModules(ByVal lVersion As Long, baQrCode() As Byte)
    Dim lQrSize         As Long
    Dim lNumAlign       As Long
    Dim baAlignPatPos(0 To 6) As Byte
    Dim lIdx            As Long
    Dim lJdx            As Long
    
    lQrSize = lVersion * 4 + 17
    ReDim baQrCode(0 To (lQrSize * lQrSize + 7) \ 8) As Byte
    baQrCode(0) = lQrSize
    '--- Fill horizontal and vertical timing patterns
    pvFillRectangle 6, 0, 1, lQrSize, baQrCode
    pvFillRectangle 0, 6, lQrSize, 1, baQrCode
    '--- Fill 3 finder patterns (all corners except bottom right) and format bits
    pvFillRectangle 0, 0, 9, 9, baQrCode
    pvFillRectangle lQrSize - 8, 0, 8, 9, baQrCode
    pvFillRectangle 0, lQrSize - 8, 9, 8, baQrCode
    '--- Fill numerous alignment patterns
    lNumAlign = pvGetAlignmentPatternPositions(lVersion, baAlignPatPos)
    For lIdx = 0 To lNumAlign - 1
        For lJdx = 0 To lNumAlign - 1
            If (lIdx = 0 And lJdx = 0) Or (lIdx = 0 And lJdx = lNumAlign - 1) Or (lIdx = lNumAlign - 1 And lJdx = 0) Then
                '--- Don't draw on the three finder corners
            Else
                pvFillRectangle baAlignPatPos(lIdx) - 2, baAlignPatPos(lJdx) - 2, 5, 5, baQrCode
            End If
        Next
    Next
    '--- Fill version blocks
    If lVersion >= 7 Then
        pvFillRectangle lQrSize - 11, 0, 3, 6, baQrCode
        pvFillRectangle 0, lQrSize - 11, 6, 3, baQrCode
    End If
End Sub

Private Sub pvDrawLightFunctionModules(ByVal lVersion As Long, baQrCode() As Byte)
    Dim lQrSize         As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lKdx            As Long
    Dim lDy             As Long
    Dim lDx             As Long
    Dim lDist           As Long
    Dim lNumAlign       As Long
    Dim baAlignPatPos(0 To 6) As Byte
    Dim bIsDark         As Boolean
    Dim lRem            As Long
    Dim lBits           As Long
    
    lQrSize = baQrCode(0)
    '--- Draw horizontal and vertical timing patterns
    For lIdx = 7 To lQrSize - 7 Step 2
        pvSetModuleBounded baQrCode, 6, lIdx, False
        pvSetModuleBounded baQrCode, lIdx, 6, False
    Next
    '--- Draw 3 finder patterns (all corners except bottom right; overwrites some timing modules)
    For lDy = -4 To 4
        For lDx = -4 To 4
            lDist = Abs(lDx)
            If Abs(lDy) > lDist Then
                lDist = Abs(lDy)
            End If
            If lDist = 2 Or lDist = 4 Then
                pvSetModuleUnbounded baQrCode, 3 + lDx, 3 + lDy, False
                pvSetModuleUnbounded baQrCode, lQrSize - 4 + lDx, 3 + lDy, False
                pvSetModuleUnbounded baQrCode, 3 + lDx, lQrSize - 4 + lDy, False
            End If
        Next
    Next
    '--- Draw numerous alignment patterns
    lNumAlign = pvGetAlignmentPatternPositions(lVersion, baAlignPatPos)
    For lIdx = 0 To lNumAlign - 1
        For lJdx = 0 To lNumAlign - 1
            If (lIdx = 0 And lJdx = 0) Or (lIdx = 0 And lJdx = lNumAlign - 1) Or (lIdx = lNumAlign - 1 And lJdx = 0) Then
                '--- Don't draw on the three finder corners
            Else
                For lDy = -1 To 1
                    For lDx = -1 To 1
                        bIsDark = (lDx = 0 And lDy = 0)
                        pvSetModuleBounded baQrCode, baAlignPatPos(lIdx) + lDx, baAlignPatPos(lJdx) + lDy, bIsDark
                    Next
                Next
            End If
        Next
    Next
    '--- Draw version blocks
    If lVersion >= 7 Then
        '--- Calculate error correction code and pack bits
        lRem = lVersion
        For lIdx = 0 To 11
            lRem = (lRem * 2) Xor ((lRem \ LNG_POW2(11)) * &H1F25)
        Next
        lBits = lVersion * LNG_POW2(12) Or lRem
        Debug.Assert lBits < LNG_POW2(18)
        '--- Draw two copies
        For lIdx = 0 To 5
            For lJdx = 0 To 2
                lKdx = lQrSize - 11 + lJdx
                bIsDark = ((lBits And 1) <> 0)
                pvSetModuleBounded baQrCode, lKdx, lIdx, bIsDark
                pvSetModuleBounded baQrCode, lIdx, lKdx, bIsDark
                lBits = lBits \ 2
            Next
        Next
    End If
End Sub

Private Sub pvDrawFormatBits(ByVal eEcl As QRCodegenEcc, ByVal eMask As QRCodegenMask, baQrCode() As Byte)
    Dim lData           As Long
    Dim lRem            As Long
    Dim lBits           As Long
    Dim lIdx            As Long
    Dim lQrSize         As Long
    
    '--- Calculate error correction code and pack bits
    lData = Choose(eEcl + 1, 1, 0, 3, 2) * 8 Or eMask
    lRem = lData
    For lIdx = 0 To 9
        lRem = (lRem * 2) Xor ((lRem \ LNG_POW2(9)) * &H537)
    Next
    lBits = (lData * LNG_POW2(10) Or lRem) Xor &H5412
    '--- Draw first copy
    For lIdx = 0 To 5
        pvSetModuleBounded baQrCode, 8, lIdx, pvGetBit(lBits, lIdx)
    Next
    pvSetModuleBounded baQrCode, 8, 7, pvGetBit(lBits, 6)
    pvSetModuleBounded baQrCode, 8, 8, pvGetBit(lBits, 7)
    pvSetModuleBounded baQrCode, 7, 8, pvGetBit(lBits, 8)
    For lIdx = 9 To 14
        pvSetModuleBounded baQrCode, 14 - lIdx, 8, pvGetBit(lBits, lIdx)
    Next
    '--- Draw second copy
    lQrSize = baQrCode(0)
    For lIdx = 0 To 7
        pvSetModuleBounded baQrCode, lQrSize - 1 - lIdx, 8, pvGetBit(lBits, lIdx)
    Next
    For lIdx = 8 To 14
        pvSetModuleBounded baQrCode, 8, lQrSize - 15 + lIdx, pvGetBit(lBits, lIdx)
    Next
    pvSetModuleBounded baQrCode, 8, lQrSize - 8, True
End Sub

Private Function pvGetAlignmentPatternPositions(ByVal lVersion As Long, baResult() As Byte) As Long
    Dim lNumAlign       As Long
    Dim lStep           As Long
    Dim lIdx            As Long
    Dim lPos            As Long
    
    If lVersion > 1 Then
        lNumAlign = lVersion \ 7 + 2
        lStep = IIf(lVersion = 32, 26, ((lVersion * 4 + lNumAlign * 2 + 1) \ (lNumAlign * 2 - 2)) * 2)
        lPos = lVersion * 4 + 10
        For lIdx = lNumAlign - 1 To 1 Step -1
            baResult(lIdx) = lPos
            lPos = lPos - lStep
        Next
        baResult(0) = 6
        pvGetAlignmentPatternPositions = lNumAlign
    End If
End Function

Private Sub pvFillRectangle(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, baQrCode() As Byte)
    Dim lDy             As Long
    Dim lDx             As Long
    
    For lDy = 0 To lHeight - 1
        For lDx = 0 To lWidth - 1
            pvSetModuleBounded baQrCode, lLeft + lDx, lTop + lDy, True
        Next
    Next
End Sub

Private Sub pvDrawCodewords(baData() As Byte, baQrCode() As Byte)
    Dim lQrSize         As Long
    Dim lBitLen         As Long
    Dim lIdx            As Long
    Dim lRight          As Long
    Dim lVert           As Long
    Dim lJdx            As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim bIsDark         As Boolean
    
    lQrSize = baQrCode(0)
    lBitLen = (UBound(baData) + 1) * 8
    For lRight = lQrSize - 1 To 1 Step -2
        If lRight = 6 Then
            lRight = 5
        End If
        For lVert = 0 To lQrSize - 1
            For lJdx = 0 To 1
                lX = lRight - lJdx
                If ((lRight + 1) And 2) = 0 Then
                    lY = lQrSize - 1 - lVert
                Else
                    lY = lVert
                End If
                If Not pvGetModuleBounded(baQrCode, lX, lY) And lIdx < lBitLen Then
                    bIsDark = pvGetBit(baData(lIdx \ 8), 7 - (lIdx And 7))
                    pvSetModuleBounded baQrCode, lX, lY, bIsDark
                    lIdx = lIdx + 1
                End If
            Next
        Next
    Next
    Debug.Assert lIdx = lBitLen
End Sub

Private Sub pvApplyMask(baFunctionModules() As Byte, baQrCode() As Byte, ByVal eMask As QRCodegenMask)
    Dim lQrSize         As Long
    Dim lY              As Long
    Dim lX              As Long
    Dim bInvert         As Boolean
    Dim bVal            As Boolean
    
    Debug.Assert QRCodegenMask_0 <= eMask And eMask <= QRCodegenMask_7
    lQrSize = baQrCode(0)
    For lY = 0 To lQrSize - 1
        For lX = 0 To lQrSize - 1
            If Not pvGetModuleBounded(baFunctionModules, lX, lY) Then
                Select Case eMask
                Case QRCodegenMask_0
                    bInvert = (lX + lY) Mod 2 = 0
                Case QRCodegenMask_1
                    bInvert = lY Mod 2 = 0
                Case QRCodegenMask_2
                    bInvert = lX Mod 3 = 0
                Case QRCodegenMask_3
                    bInvert = (lX + lY) Mod 3 = 0
                Case QRCodegenMask_4
                    bInvert = (lX \ 3 + lY \ 2) Mod 2 = 0
                Case QRCodegenMask_5
                    bInvert = (lX * lY Mod 2 + lX * lY Mod 3) = 0
                Case QRCodegenMask_6
                    bInvert = (lX * lY Mod 2 + lX * lY Mod 3) Mod 2 = 0
                Case QRCodegenMask_7
                    bInvert = ((lX + lY) Mod 2 + lX * lY Mod 3) Mod 2 = 0
                End Select
                bVal = pvGetModuleBounded(baQrCode, lX, lY)
                pvSetModuleBounded baQrCode, lX, lY, (bVal Xor bInvert)
            End If
        Next
    Next
End Sub

Private Function pvGetPenaltyScore(baQrCode() As Byte) As Long
    Dim lQrSize         As Long
    Dim lX              As Long
    Dim lY              As Long
    Dim bRunColor       As Boolean
    Dim lRunX           As Long
    Dim lRunY           As Long
    Dim aRunHistory()   As Long
    Dim lDark           As Long
    Dim lTotal          As Long
    Dim lKdx            As Long
    
    lQrSize = baQrCode(0)
    '--- Adjacent modules in row having same color, and finder-like patterns
    For lY = 0 To lQrSize - 1
        bRunColor = False
        lRunX = 0
        ReDim aRunHistory(0 To 6) As Long
        For lX = 0 To lQrSize - 1
            If pvGetModuleBounded(baQrCode, lX, lY) = bRunColor Then
                lRunX = lRunX + 1
                If lRunX = 5 Then
                    pvGetPenaltyScore = pvGetPenaltyScore + PENALTY_N1
                ElseIf lRunX > 5 Then
                    pvGetPenaltyScore = pvGetPenaltyScore + 1
                End If
            Else
                pvFinderPenaltyAddHistory lRunX, aRunHistory, lQrSize
                If Not bRunColor Then
                    pvGetPenaltyScore = pvGetPenaltyScore + pvFinderPenaltyCountPatterns(aRunHistory, lQrSize) * PENALTY_N3
                End If
                bRunColor = pvGetModuleBounded(baQrCode, lX, lY)
                lRunX = 1
            End If
        Next
        pvGetPenaltyScore = pvGetPenaltyScore + pvFinderPenaltyTerminateAndCount(bRunColor, lRunX, aRunHistory, lQrSize) * PENALTY_N3
    Next
    '--- Adjacent modules in column having same color, and finder-like patterns
    For lX = 0 To lQrSize - 1
        bRunColor = False
        lRunY = 0
        ReDim aRunHistory(0 To 6) As Long
        For lY = 0 To lQrSize - 1
            If pvGetModuleBounded(baQrCode, lX, lY) = bRunColor Then
                lRunY = lRunY + 1
                If lRunY = 5 Then
                    pvGetPenaltyScore = pvGetPenaltyScore + PENALTY_N1
                ElseIf lRunY > 5 Then
                    pvGetPenaltyScore = pvGetPenaltyScore + 1
                End If
            Else
                pvFinderPenaltyAddHistory lRunY, aRunHistory, lQrSize
                If Not bRunColor Then
                    pvGetPenaltyScore = pvGetPenaltyScore + pvFinderPenaltyCountPatterns(aRunHistory, lQrSize) * PENALTY_N3
                End If
                bRunColor = pvGetModuleBounded(baQrCode, lX, lY)
                lRunY = 1
            End If
        Next
        pvGetPenaltyScore = pvGetPenaltyScore + pvFinderPenaltyTerminateAndCount(bRunColor, lRunY, aRunHistory, lQrSize) * PENALTY_N3
    Next
    '--- 2*2 blocks of modules having same color
    For lY = 0 To lQrSize - 2
        For lX = 0 To lQrSize - 2
            bRunColor = pvGetModuleBounded(baQrCode, lX, lY)
            If bRunColor = pvGetModuleBounded(baQrCode, lX + 1, lY) And _
                    bRunColor = pvGetModuleBounded(baQrCode, lX, lY + 1) And _
                    bRunColor = pvGetModuleBounded(baQrCode, lX + 1, lY + 1) Then
                pvGetPenaltyScore = pvGetPenaltyScore + PENALTY_N2
            End If
        Next
    Next
    '--- Balance of dark and light modules
    For lY = 0 To lQrSize - 1
        For lX = 0 To lQrSize - 1
            If pvGetModuleBounded(baQrCode, lX, lY) Then
               lDark = lDark + 1
            End If
        Next
    Next
    lTotal = lQrSize * lQrSize
    '--- Compute the smallest integer k >= 0 such that (45-5k)% <= dark/total <= (55+5k)%
    lKdx = ((Abs(lDark * 20 - lTotal * 10) + lTotal - 1) \ lTotal) - 1
    Debug.Assert 0 <= lKdx And lKdx <= 9
    pvGetPenaltyScore = pvGetPenaltyScore + lKdx * PENALTY_N4
    Debug.Assert 0 <= pvGetPenaltyScore And pvGetPenaltyScore <= 2568888
End Function

Private Function pvFinderPenaltyCountPatterns(aRunHistory() As Long, ByVal lQrSize As Long) As Long
    Dim lN              As Long
    Dim bCore           As Boolean
    
    lN = aRunHistory(1)
    Debug.Assert lN <= lQrSize * 3
    bCore = (lN > 0 And aRunHistory(2) = lN And aRunHistory(3) = lN * 3 And aRunHistory(4) = lN And aRunHistory(5) = lN)
    '-- The maximum QR Code size is 177, hence the dark run length n <= 177.
    pvFinderPenaltyCountPatterns = IIf(bCore And aRunHistory(0) >= lN * 4 And aRunHistory(6) >= lN, 1, 0) _
                                 + IIf(bCore And aRunHistory(6) >= lN * 4 And aRunHistory(0) >= lN, 1, 0)
End Function

Private Function pvFinderPenaltyTerminateAndCount(ByVal bCurrentRunColor As Boolean, ByVal lCurrentRunLength As Long, aRunHistory() As Long, ByVal lQrSize As Long) As Long
    If bCurrentRunColor Then
        pvFinderPenaltyAddHistory lCurrentRunLength, aRunHistory, lQrSize
        lCurrentRunLength = 0
    End If
    lCurrentRunLength = lCurrentRunLength + lQrSize
    pvFinderPenaltyAddHistory lCurrentRunLength, aRunHistory, lQrSize
    pvFinderPenaltyTerminateAndCount = pvFinderPenaltyCountPatterns(aRunHistory, lQrSize)
End Function

Private Function pvFinderPenaltyAddHistory(ByVal lCurrentRunLength As Long, aRunHistory() As Long, ByVal lQrSize As Long) As Long
    If aRunHistory(0) = 0 Then
        lCurrentRunLength = lCurrentRunLength + lQrSize
    End If
    Debug.Assert UBound(aRunHistory) + 1 = 7
    Call CopyMemory(aRunHistory(1), aRunHistory(0), 6 * LenB(aRunHistory(0)))
    aRunHistory(0) = lCurrentRunLength
End Function

Private Function pvGetModuleBounded(baQrCode() As Byte, ByVal lX As Long, ByVal lY As Long) As Long
    Dim lQrSize         As Long
    Dim lIndex          As Long
    
    lQrSize = baQrCode(0)
    Debug.Assert 21 <= lQrSize And lQrSize <= 177 And 0 <= lX And lX < lQrSize And 0 <= lY And lY < lQrSize
    lIndex = lY * lQrSize + lX
    pvGetModuleBounded = pvGetBit(baQrCode(lIndex \ 8 + 1), lIndex And 7)
End Function

Private Function pvSetModuleBounded(baQrCode() As Byte, ByVal lX As Long, ByVal lY As Long, ByVal bIsDark As Boolean) As Long
    Dim lQrSize         As Long
    Dim lIndex          As Long
    Dim lByteIndex      As Long
    
    lQrSize = baQrCode(0)
    Debug.Assert 21 <= lQrSize And lQrSize <= 177 And 0 <= lX And lX < lQrSize And 0 <= lY And lY < lQrSize
    lIndex = lY * lQrSize + lX
    lByteIndex = lIndex \ 8 + 1
    If bIsDark Then
        baQrCode(lByteIndex) = baQrCode(lByteIndex) Or LNG_POW2(lIndex And 7)
    Else
        baQrCode(lByteIndex) = baQrCode(lByteIndex) And Not LNG_POW2(lIndex And 7)
    End If
End Function

Private Function pvSetModuleUnbounded(baQrCode() As Byte, ByVal lX As Long, ByVal lY As Long, ByVal bIsDark As Boolean) As Long
    Dim lQrSize         As Long
    
    lQrSize = baQrCode(0)
    If 0 <= lX And lX < lQrSize And 0 <= lY And lY < lQrSize Then
        pvSetModuleBounded baQrCode, lX, lY, bIsDark
    End If
End Function

Private Function pvGetBit(ByVal lX As Long, ByVal lIdx As Long) As Boolean
    pvGetBit = (lX And LNG_POW2(lIdx)) <> 0
End Function

Private Function pvCalcSegmentBitLength(ByVal eMode As QRCodegenMode, ByVal lNumChars As Long) As Long
    If lNumChars > INT16_MAX Then
        pvCalcSegmentBitLength = -1
    Else
        pvCalcSegmentBitLength = lNumChars
        Select Case eMode
        Case QRCodegenMode_NUMERIC
            pvCalcSegmentBitLength = (pvCalcSegmentBitLength * 10 + 2) \ 3
        Case QRCodegenMode_ALPHANUMERIC
            pvCalcSegmentBitLength = (pvCalcSegmentBitLength * 11 + 1) \ 2
        Case QRCodegenMode_BYTE
            pvCalcSegmentBitLength = pvCalcSegmentBitLength * 8
        Case QRCodegenMode_KANJI
            pvCalcSegmentBitLength = pvCalcSegmentBitLength * 13
        Case QRCodegenMode_ECI
            Debug.Assert lNumChars = 0
            pvCalcSegmentBitLength = 3 * 8
        End Select
        If pvCalcSegmentBitLength > INT16_MAX Then
            pvCalcSegmentBitLength = -1
        End If
    End If
End Function

Private Function pvGetBufferLenForVersion(ByVal lVersion As Long) As Long
    pvGetBufferLenForVersion = (((lVersion * 4 + 17) * (lVersion * 4 + 17) + 7) \ 8 + 1)
End Function

Private Sub pvAppendBitsToBuffer(ByVal lVal As Long, ByVal lNumBits As Long, baBuffer() As Byte, lBitLen As Long)
    Dim lIdx            As Long
    
    Debug.Assert 0 <= lNumBits And lNumBits <= 16
    For lIdx = lNumBits - 1 To 0 Step -1
        If (lVal And LNG_POW2(lIdx)) <> 0 Then
            baBuffer(lBitLen \ 8) = baBuffer(lBitLen \ 8) Or LNG_POW2(7 - (lBitLen And 7))
        End If
        lBitLen = lBitLen + 1
    Next
End Sub

Private Function pvToUtf8Array(sText As String) As Byte()
    Const CP_UTF8       As Long = 65001
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    pvToUtf8Array = baRetVal
End Function
