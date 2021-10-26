Attribute VB_Name = "Module2"
Option Explicit
Const CSIDL_TEMPLATES = &H15&
Const MAX_PATH = 260
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public doslast As Double
Dim FontList As FastCollection
Public Trush() As VarItem
Public TrushCount As Long, TrushWait As Boolean
Public Nonbsp As Boolean
Const b123 = vbCr + "'\"
Const b1234 = vbCr + "'\:"
Public k1 As Long, Kform As Boolean
Private Const doc = "Document"
Public Check2SaveModules As Boolean
Dim ObjectCatalog As FastCollection, loadcatalog As New FastCollection
Public tracecode As String, lasttracecode As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpszOp As Long, ByVal lpszFile As Long, ByVal lpszParams As Long, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableW" (ByVal lpFile As Long, ByVal lpDirectory As Long, ByVal lpResult As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal Hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Declare Function GetLocaleInfoW Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&) ' not NT?
Private Const DWL_ANYTHREAD& = 0
Const LOCALE_ILANGUAGE = 1
Private Const LOCALE_SENGLANGUAGE& = 4097&
Private Const LOCALE_SLANGUAGE& = &H2& '  localized name of language

'Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
'Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public stackshowonly As Boolean, NoBackFormFirstUse As Boolean
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Declare Function ExpandEnvironmentStrings _
   Lib "kernel32" Alias "ExpandEnvironmentStringsW" _
   (ByVal lpSrc As Long, ByVal lpDst As Long, _
   ByVal nSize As Long) As Long
Private Declare Function GetTempFileNameW Lib "kernel32" _
    (ByVal lpszPath As Long, ByVal lpPrefixString As Long, _
     ByVal wUnique As Long, ByVal lpTempFileName As Long) _
     As Long
Private Const UNIQUE_NAME = &H0&
Public Pow2(33) As Currency, Pow2minusOne(33) As Currency
Public Enum Ftypes
    FnoUse
    Finput
    Foutput
    Fappend
    Frandom
End Enum
Public FLEN(512) As Long, FKIND(512) As Ftypes
Public Type Counters
    k1 As Currency
    RRCOUNTER As Currency
End Type
Public Type basket
    used As Long
    x As Long  ' for hotspot
    y As Long  '
    XGRAPH As Long  ' graphic cursor
    YGRAPH As Long
    MAXXGRAPH As Long
    MAXYGRAPH As Long
    dv15 As Long  ' not used
    curpos As Long   ' text cursor
    currow As Long
    mypen As Long
    mypentrans As Long
    mysplit As Long
    Paper As Long
    italics As Boolean  ' removed from process, only in current
    bold As Boolean
    double As Boolean
    osplit As Long  '(for double size letters)
    Column As Long
    OCOLUMN As Long
    pageframe As Long
    basicpageframe As Long
    MineLineSpace As Long
    uMineLineSpace As Long
    LastReportLines As Double
    SZ As Single
    UseDouble As Single
    Xt As Long
    Yt As Long
    mx As Long
    My As Long
    FontName As String
    charset As Long
    FTEXT As Long
    FTXT As String
    lastprint As Boolean  ' if true then we have to place letters using currentX
    ' gdi drawing enabled Smooth On, disabled with Smooth Of
    NoGDI As Boolean
    IamEmf As Boolean
    pathgdi As Long  ' only for gdi+
    pathcolor As Long ' only for gdi+
    pathfillstyle As Integer
    LastIcon As Integer  ' 1..   / 99 loaded
    LastIconPic As StdPicture
    HideIcon As Boolean
    ReportTab As Long
    overrideTextHeight As Long
    HotSpotX As Long
    HotSpotY As Long
End Type
Private stopwatch As Long
Private Const myArray = "mArray"
Private Const LOCALE_SYSTEM_DEFAULT As Long = &H800
Private Const LOCALE_USER_DEFAULT As Long = &H800
Private Const C3_DIACRITIC As Long = &H2
Private Const CT_CTYPE3 As Byte = &H4
Private Declare Function GetStringTypeExW Lib "kernel32.dll" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByRef lpCharType As Integer) As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal Hdc As Long, ByVal nCharExtra As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GdiFlush Lib "gdi32" () As Long
Public iamactive As Boolean
Declare Function MultiByteToWideChar& Lib "kernel32" (ByVal CodePage&, ByVal dwFlags&, MultiBytes As Any, ByVal cBytes&, ByVal pWideChars&, ByVal cWideChars&)
Private Declare Function FillRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal Hdc As Long, _
         hRgn As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Type RECT
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type
Public Const LOCALE_SDECIMAL = &HE&
Public Const LOCALE_SGROUPING = &H10&
Public Const LOCALE_STHOUSAND = &HF&
Public Const LOCALE_SMONDECIMALSEP = &H16&
Public Const LOCALE_SMONTHOUSANDSEP = &H17&
Public Const LOCALE_SMONGROUPING = &H18&
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Const DT_BOTTOM As Long = &H8&
Private Const DT_CALCRECT As Long = &H400&
Private Const DT_CENTER As Long = &H1&
Private Const DT_EDITCONTROL As Long = &H2000&
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_EXPANDTABS As Long = &H40&
Private Const DT_EXTERNALLEADING As Long = &H200&
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const DT_INTERNAL As Long = &H1000&
Private Const DT_LEFT As Long = &H0&
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_NOCLIP As Long = &H100&
Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
Private Const DT_NOPREFIX As Long = &H800&
Private Const DT_PATH_ELLIPSIS As Long = &H4000&
Private Const DT_PREFIXONLY As Long = &H200000
Private Const DT_RIGHT As Long = &H2&
Private Const DT_SINGLELINE As Long = &H20&
Private Const DT_TABSTOP As Long = &H80&
Private Const DT_TOP As Long = &H0&
Private Const DT_VCENTER As Long = &H4&
Private Const DT_WORDBREAK As Long = &H10&
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Public Declare Function DestroyCaret Lib "user32" () As Long
Public Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetCaretPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Const dv = 0.877551020408163
Public QUERYLIST As String
Public LASTQUERYLIST As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public releasemouse As Boolean
Public LASTPROG$
Public NORUN1 As Boolean
Public UseEnter As Boolean
Public dv20 As Single  ' = 24.5
Public dv15 As Long
Public mHelp As Boolean
Public abt As Boolean
Public vH_title$
Public vH_doc$
Public vH_x As Long
Public vH_y As Long
Public ttl As Boolean
Public Const SRCCOPY = &HCC0020
Public Release As Boolean
Private Declare Function SetBkMode Lib "gdi32" (ByVal Hdc As Long, ByVal nBkMode As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function RoundRect Lib "gdi32" (ByVal Hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ScrollDC Lib "user32" (ByVal Hdc As Long, ByVal dX As Long, ByVal dY As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public LastErName As String
Public LastErNameGR As String
Public LastErNum As Long
Public LastErNum1 As Long, LastErNum2 As Long
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Byte)

Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal Hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As POINTAPI) As Long

Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function PaintDesktop Lib "user32" (ByVal Hdc As Long) As Long
Declare Function SelectClipPath Lib "gdi32" (ByVal Hdc As Long, ByVal iMode As Long) As Long
  Public Const RGN_AND = 1
    Public Const RGN_COPY = 5
    Public Const RGN_DIFF = 4
    Public Const RGN_MAX = RGN_COPY
    Public Const RGN_MIN = RGN_AND
    Public Const RGN_OR = 2
    Public Const RGN_XOR = 3
Declare Function StrokePath Lib "gdi32" (ByVal Hdc As Long) As Long
Declare Function Polygon Lib "gdi32" (ByVal Hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function PolyBezier Lib "gdi32.dll" (ByVal Hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal Hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Declare Function BeginPath Lib "gdi32" (ByVal Hdc As Long) As Long
Declare Function EndPath Lib "gdi32" (ByVal Hdc As Long) As Long
Declare Function FillPath Lib "gdi32" (ByVal Hdc As Long) As Long
Declare Function StrokeAndFillPath Lib "gdi32" (ByVal Hdc As Long) As Long

Public PLG() As POINTAPI
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public lckfrm As Long
Public NERR As Boolean
Public moux As Single, mouy As Single, MOUB As Long
Public mouxb As Single, mouyb As Single, MOUBb As Long
Public vol As Long
Public MyFont As String, myCharSet As Integer, myBold As Boolean
Public FFONT As String

Public escok As Boolean
Public NOEDIT As Boolean
Public CancelEDIT As Boolean

Global Const HWND_TOP = 0

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long)
Declare Function ExtFloodFill Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1
Public Const FLOODFILLBORDER = 0

Public avifile As String
Public BigPi As Variant
Public Const Pi = 3.14159265358979
Public Const PI2 = 6.28318530717958
Public EditTabWidth As Long, ReportTabWidth As Long
Public Result As Long, Use13 As Boolean
Public mcd As String
Public NOEXECUTION As Boolean, RoundDouble As Boolean
Public QRY As Boolean, GFQRY As Boolean
Public nomore As Boolean
Private Declare Function CallWindowProc _
 Lib "user32.dll" Alias "CallWindowProcW" ( _
 ByVal lpPrevWndFunc As Long, _
 ByVal hWnd As Long, _
 ByVal Msg As Long, _
 ByVal wParam As Long, _
 ByVal lParam As Long) As Long

'== MCI Wave API Declarations ================================================
Public ExTarget As Boolean
''Public pageframe As Long
''Public basicpageframe As Long

Public q() As target
Public Targets As Boolean
Public SzOne As Single
Public PenOne As Long
Public NoAction As Boolean
Public StartLine As Boolean
Public www&
Public WWX&, ins&
Public INK$, MINK$
Public MKEY$
Public Type target
    Comm As String
    Tag As String ' specified by id
    id As Long ' function id
    ' THIS IS POINTS AT CHARACTER RESOLUTION
    SZ As Single
    ' SO WE NEED SZ
    Lx As Long
    ly As Long
    tx As Long
    ty As Long
    back As Long 'background fill color' -1 no fill
    fore As Long 'border line ' -1 no line
    Enable As Boolean ' in use
    pen As Long
    layer As Long
    Xt As Long
    Yt As Long
    sUAddTwipsTop As Long
End Type

Public here$, PaperOne As Long
Const PROOF_QUALITY = 2
Const NONANTIALIASED_QUALITY = 3
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare Function PathToRegion Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long

' OCTOBER 2000
Public dstyle As Long
' Jule 2001
Const DC_ACTIVE = &H1
Const DC_ICON = &H4
Const DC_TEXT = &H8
Const BDR_SUNKENOUTER = &H2
Const BDR_RAISEDINNER = &H4
Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const DFC_BUTTON = 4
Const DFC_POPUPMENU = 5            'Only Win98/2000 !!
Const DFCS_BUTTON3STATE = &H10
Const DC_GRADIENT = &H20          'Only Win98/2000 !!

Private Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal Hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal Hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal Hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal Hdc As Long, ByVal lpsz As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long, ByVal lpDrawTextParams As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
''API declarations
' old api..
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceW" (ByVal lpFileName As Long) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceW" (ByVal lpFileName As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Long
Public TextEditLineHeight As Long
Public LablelEditLineHeight As Long
Private Const Utf8CodePage As Long = 65001
Public Type DRAWTEXTPARAMS
     cbSize As Long
     iTabLength As Long
     iLeftMargin As Long
     iRightMargin As Long
     uiLengthDrawn As Long
End Type
Public tParam As DRAWTEXTPARAMS

Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_RTLREADING = &H100&
Public Declare Function SetTextJustification Lib "gdi32" (ByVal Hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal Hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutW" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal nTabPositions As Long, ByRef lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutW" (ByVal Hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Public Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentW" (ByVal Hdc As Long, ByVal lpString As Long, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
Function CheckItemType(bstackstr As basetask, v As Variant, a$, R$, Optional ByVal wasarr As Boolean = False, Optional UseCase As Boolean) As Boolean
Dim usehandler As mHandler, fastcol As FastCollection, pppp As mArray, w1 As Long, p As Variant, s$
UseCase = False
CheckItemType = True
Dim vv
If MyIsObject(v) Then
    Set vv = v
Else
   R$ = Typename(v)
   CheckItemType = FastSymbol(a$, ")")
   Exit Function
End If
againtype:
        R$ = Typename(vv)
        If R$ = "mHandler" Then
            Set usehandler = vv
            Select Case usehandler.t1
            Case 1
            If TypeOf usehandler.objref Is mHandler Then
            Set fastcol = usehandler.objref.objref
            Else
            Set fastcol = usehandler.objref
            End If
                If FastSymbol(a$, ",") Then
contHandler:
                    If IsExp(bstackstr, a$, p) Then
                        If Not fastcol.Find(p) Then GoTo keynotexist
                        If fastcol.IsObj Then
                            Set vv = fastcol.ValueObj
                            GoTo againtype
                        Else
                            wasarr = True
                            GoTo checkit
                        End If
                    ElseIf IsStrExp(bstackstr, a$, s$, Len(bstackstr.tmpstr) = 0) Then
                        If fastcol.IsObj Then
                            Set vv = fastcol.ValueObj
                            GoTo againtype
                        Else
                            If fastcol.StructLen > 0 Then GoTo checkit
                                R$ = Typename(fastcol.Value)
                            End If
                        Else
                            MissParam a$
                            CheckItemType = False
                            Exit Function
keynotexist:
                            indexout a$
                            CheckItemType = False
                            Exit Function
                    End If
                ElseIf FastSymbol(a$, ")(", , 2) Then
                    GoTo contHandler
   
                Else
                    ' new
checkit:
                    If fastcol.StructLen > 0 Then
                                    Select Case fastcol.sValue
                                    Case Is < 0
                                        R$ = "String"
                                    Case 1
                                        R$ = "Byte"
                                    Case 2
                                        R$ = "Integer"
                                    Case 4
                                        R$ = "Long"
                                    Case 8
                                        R$ = "LongLong"  ' can be double or two longs or ...etc
                                    Case Else
                                        R$ = "Structure"
                                    End Select
                    ElseIf wasarr Then
                    R$ = Typename(fastcol.Value)
                    
                    ElseIf FastSymbol(a$, "!") Then
                        If fastcol.IsQueue Then
                            R$ = "Queue"
                        Else
                            R$ = "List"
                        End If
                    Else
                        R$ = "Inventory"
                    End If
                End If
            Case 2
                R$ = "Buffer"
                
                
            Case 3
                w1 = usehandler.indirect
                If w1 > -1 And w1 <= var2used Then
                                R$ = Typename(var(w1))
                                If R$ = "mHandler" Then Set vv = var(w1): GoTo againtype
                    Else
                            R$ = Typename(usehandler.objref)
                                       If FastSymbol(a$, ",") Then
contarr0:
                                        If R$ = "mArray" Then

                                            Set pppp = usehandler.objref
                                                If IsExp(bstackstr, a$, p) Then
                                                   pppp.Index = p
                                                    If MyIsObject(pppp.Value) Then
                                                         Set vv = pppp.Value
                                                         wasarr = False
                                                         GoTo againtype
                                                    Else
                                                        R$ = Typename(pppp.Value)
                                                    End If
                                                Else
                                                MissParam a$
                                                CheckItemType = False
                                                Exit Function
                                            End If
                                        ElseIf FastSymbol(a$, ")(", , 2) Then
                                        GoTo contarr0
            
                                        Else
                                                MyEr "Use STACKTYPE$() ", " Χρησιμοποίησε την ΣΩΡΟΥΤΥΠΟΣ$()"
                                                CheckItemType = False
                                                Exit Function
                                        End If
                                        
                                        End If
                                        
                                        End If
                                    
            Case 4
                    UseCase = True
                    R$ = usehandler.objref.EnumName
            Case Else
                Set usehandler = vv
                R$ = Typename(usehandler.objref)
                Set usehandler = Nothing
            End Select
        ElseIf Typename(vv) = "PropReference" Then
                p = vv.Value
        
            R$ = Typename$(vv.lastobjfinal)
            If R$ = "Nothing" Then
            If VarType(p) <> vbEmpty Then
            R$ = Typename(p)
            End If
            End If
        ElseIf Typename(vv) = "mArray" Then
        
         If FastSymbol(a$, ",") Then
contarr1:
            Set pppp = vv

            If IsExp(bstackstr, a$, p) Then

                pppp.Index = p
                If MyIsObject(pppp.Value) Then
                     Set vv = pppp.Value
                     wasarr = False
                     GoTo againtype
                Else
                    R$ = Typename(pppp.Value)
                End If
                Else
                MissParam a$
                CheckItemType = False
                Exit Function
                End If
            ElseIf FastSymbol(a$, ")(", , 2) Then
          GoTo contarr1
            Else
            R$ = "mArray"
            End If
        ElseIf Typename(vv) = "lambda" Then
        If FastSymbol(a$, ")(", , 2) Then
            Set bstackstr.lastobj = vv
            Set bstackstr.lastpointer = Nothing
            s$ = BlockParam(a$)
            If Len(s$) > 0 Then Mid$(a$, 1, Len(s$)) = space$(Len(s$))
            s$ = s$ + ")"
            If CallLambdaASAP(bstackstr, s$, p, False) Then
                If bstackstr.lastobj Is Nothing Then
                    R$ = Typename$(p)
                Else
                    Set vv = bstackstr.lastobj
                    Set bstackstr.lastobj = Nothing
                    GoTo againtype
                End If
            Else
            Exit Function
            End If
            Else
            R$ = "lambda"
            End If
        ElseIf Typename(vv) = "Group" Then
        If FastSymbol(a$, ")(", , 2) Then
            s$ = BlockParam(a$)
            If Len(s$) > 0 Then Mid$(a$, 1, Len(s$)) = space$(Len(s$))
            If FastSymbol(s$, "@") Then
            s$ = NLtrim(s$)
                    If Len(s$) > 0 Then
                        Set pppp = New mArray
                        pppp.Arr = False
                        Set pppp.GroupRef = vv
                        Set vv = bstackstr.soros
                        Set bstackstr.Sorosref = New mStiva
                        
                        SpeedGroup bstackstr, pppp, "FOR", "", "{Push type$(" + String$(bstackstr.ForLevel + 1, ".") + s$ + ")}", -2
                        If bstackstr.soros.IsEmpty Then
                           Set bstackstr.Sorosref = vv
                           Exit Function
                        Else
                           R$ = bstackstr.soros.PopStr
                        End If
                        Set bstackstr.Sorosref = vv
                     Else
                        SyntaxError
                        Exit Function
                End If
                ' check member
            ElseIf vv.IamApointer Then
                R$ = "Group"   ' not decide yet
            ElseIf vv.HasStrValue Then
                R$ = "String"
            ElseIf vv.HasValue Then
            Set pppp = New mArray
            pppp.Arr = False
            Set pppp.GroupRef = vv
             If SpeedGroup(bstackstr, pppp, "VAL", "", s$ + ")", -2) = 1 Then
                If bstackstr.lastobj Is Nothing Then
                    R$ = Typename$(bstackstr.LastValue)
                Else
                    Set vv = bstackstr.lastobj
                    Set bstackstr.lastobj = Nothing
                    GoTo againtype
                End If

             End If

            End If
        End If
        End If
        Set bstackstr.lastobj = Nothing
        Set bstackstr.lastpointer = Nothing
        While FastSymbol(a$, "!")
        Wend
        CheckItemType = FastSymbol(a$, ")", True)
End Function

Sub NoValidCodePage()
    MyEr "Invalid code page", "ανύπαρκτη κωδικοσελίδα"
End Sub
Sub NoValidLocale()
    MyEr "Invalid Locale ID", "ανύπαρκτος κωδικός τοπικού"
End Sub
Public Function Utf16toUtf8(s As String) As Byte()
    ' code from vbforum
    ' UTF-8 returned to VB6 as a byte array (zero based) because it's pretty useless to VB6 as anything else.
    Dim iLen As Long
    Dim bbBuf() As Byte
    '
    iLen = WideCharToMultiByte(Utf8CodePage, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    If iLen = 0 Then bbBuf() = vbNullString: Exit Function
    ReDim bbBuf(0 To iLen - 1) ' Will be initialized as all &h00.
    iLen = WideCharToMultiByte(Utf8CodePage, 0, StrPtr(s), Len(s), VarPtr(bbBuf(0)), iLen, 0, 0)
    Utf16toUtf8 = bbBuf
End Function
Public Function KeyPressedLong(ByVal VirtKeyCode As Long) As Long
On Error GoTo KEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
KeyPressedLong = GetAsyncKeyState(VirtKeyCode)
End If
End If
KEXIT:
End Function
Public Function KeyPressed2(ByVal VirtKeyCode As Long, ByVal VirtKeyCode2 As Long) As Boolean
On Error GoTo KEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
KeyPressed2 = CBool((GetAsyncKeyState(VirtKeyCode) And &H8000&) = &H8000&) And CBool((GetAsyncKeyState(VirtKeyCode2) And &H8000&) = &H8000&)
End If
End If
KEXIT:
End Function
Public Function KeyPressed(ByVal VirtKeyCode As Long) As Boolean
On Error GoTo KEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
KeyPressed = CBool((GetAsyncKeyState(VirtKeyCode) And &H8000&) = &H8000&)
End If
End If
KEXIT:
End Function
Public Function mouse2() As Long
On Error GoTo MEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then

mouse2 = (UINT(GetAsyncKeyState((1))) And &HFF) + (UINT(GetAsyncKeyState((2))) And &HFF) * 2 + (UINT(GetAsyncKeyState((4))) And &HFF) * 4
End If
End If
MEXIT:
End Function
Public Function mouse() As Long
On Error GoTo MEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
''If Screen.ActiveForm Is Form1 Then If Form1.lockme Then Exit Function

mouse = -1 * CBool((GetAsyncKeyState(1) And &H8000&) = &H8000&) - 2 * CBool((GetAsyncKeyState(2) And &H8000&) = &H8000&) - 4 * CBool((GetAsyncKeyState(4) And &H8000&) = &H8000&)
End If
End If
MEXIT:
End Function

Public Function MOUSEX(Optional Offset As Long = 0) As Long
Static x As Long
On Error GoTo MOUSEX
Dim tp As POINTAPI
MOUSEX = x
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
   GetCursorPos tp
   x = tp.x * dv15 - Offset
  MOUSEX = x
  End If
End If
MOUSEX:
End Function
Public Function MOUSEY(Optional Offset As Long = 0) As Long
Static y As Long
On Error GoTo MOUSEY
Dim tp As POINTAPI
MOUSEY = y
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
   GetCursorPos tp
   y = tp.y * dv15 - Offset
   MOUSEY = y
  End If
End If
MOUSEY:
End Function
Public Sub OnlyInAGroup()
    MyEr "Only in a group", "Μόνο σε μια ομάδα"
End Sub
Public Sub WrongOperator()
MyEr "Wrong operator", "λάθος τελεστής"
End Sub
Public Sub NoOperatorForThatObject(ss$)
If ss$ = "g" Then ss$ = "<="
    MyEr "Object not support operator " + ss$, "Το αντικείμενο δεν υποστηρίζει το τελεστή " + ss$
End Sub
Public Sub NoStackObjectFound(a$)
    MyErMacro a$, "Not stack object found", "Δεν βρήκα αντικείμενο σωρού"
End Sub
Public Sub NoStackObjectToMerge()
    MyEr "Not stack object to merge", "Δεν βρήκα αντικείμενο σωρού να ενώσω"
End Sub
Public Sub Unsignlongnegative(a$)
    MyErMacro a$, "Unsigned long can't be negative", "Ο ακέραιος χωρίς προσημο δεν μπορεί να είναι αρνητικός"
End Sub
Public Sub Unsignlongfailed(a$)
MyErMacro a$, "Unsigned long to sign failed", "Η μετατροπή ακέραιου χωρίς πρόσημο σε ακέραιο με πρόσημο, απέτυχε"
End Sub
Public Sub NoProperObject()
MyEr "This object not supported", "Αυτό το αντικείμενο δεν υποστηρίζεται"
End Sub

Public Sub MyEr(er$, ergr$)
If Left$(LastErName, 1) = Chr(0) Then
    LastErName = vbNullString
    LastErNameGR = vbNullString
End If
If er$ = vbNullString Then
LastErNum = 0
LastErNum1 = 0
LastErName = vbNullString
LastErNameGR = vbNullString
Else
er$ = Split(er$, ChrW(&H1FFF))(0)
ergr$ = Split(ergr$, ChrW(&H1FFF))(0)
If rinstr(er$, " ") = 0 Then
LastErNum = 1001
Else

LastErNum = val(" " & Mid$(er$, rinstr(er$, " ")) + ".0")
End If
If LastErNum = 0 Then LastErNum = -1
LastErNum1 = LastErNum

If InStr("*" + LastErName, NLtrim$(er$)) = 0 Then
LastErName = RTrim$(LastErName) & " " & NLtrim$(er$)
LastErNameGR = RTrim$(LastErNameGR) & " " & NLtrim$(ergr$)
End If
End If
End Sub
Sub UnknownVariable1(a$, v$)
Dim i As Long
i = rinstr(v$, "." + ChrW(8191))
If i > 0 Then
    i = rinstr(v$, ".")
    MyErMacro a$, "Unknown Variable " & Mid$(v$, i), "’γνωστη μεταβλητή " & Mid$(v$, i)
Else
    i = rinstr(v$, "].")
    If i > 0 Then
        MyErMacro a$, "Unknown Variable " & Mid$(v$, i + 2), "’γνωστη μεταβλητή " & Mid$(v$, i + 2)
    Else
        i = rinstr(v$, ChrW(8191))
    If i > 0 Then
        i = InStr(i + 1, v$, ".")
        If i > 0 Then
            MyErMacro a$, "Unknown Variable " & Mid$(v$, i + 1), "’γνωστη μεταβλητή " & Mid$(v$, i + 1)
        Else
            MyErMacro a$, "Unknown Variable", "’γνωστη μεταβλητή"
        End If
    Else
        MyErMacro a$, "Unknown Variable " & v$, "’γνωστη μεταβλητή " & v$
    End If
    End If
End If

End Sub
Sub UnknownProperty1(a$, v$)
MyErMacro a$, "Unknown Property " & v$, "’γνωστη ιδιότητα " & v$
End Sub
Sub UnknownMethod1(a$, v$)
 MyErMacro a$, "unknown method/array  " & v$, "’γνωστη μέθοδος/πίνακας " & v$
End Sub
Sub UnknownFunction1(a$, v$)
 MyErMacro a$, "unknown function/array " & v$, "’γνωστη συνάρτηση/πίνακας " & v$
End Sub

Sub InternalError()
 MyEr "Internal Error", "Εσωτερικό Πρόβλημα"
End Sub
Public Function LoadFont(ByVal FntFileName As String) As Boolean
    Dim FntRC As Long
    If FontList Is Nothing Then
    Set FontList = New FastCollection
    End If
    FntFileName = mylcasefILE(FntFileName)
    If FontList.ExistKey(FntFileName) Then
        LoadFont = True
    Else
        FntRC = AddFontResource(StrPtr(FntFileName))
        If FntRC = 0 Then 'no success
         LoadFont = False
        Else 'success
        FontList.AddKey FntFileName
     LoadFont = True
    End If
        End If
End Function
'FntFileName includes also path
Public Function RemoveFont(ByVal FntFileName As String) As Boolean
     Dim rc As Long, Inc As Integer
     If FontList Is Nothing Then Exit Function
        FntFileName = mylcasefILE(FntFileName)
     If FontList.ExistKey(FntFileName) Then
     Do
       rc = RemoveFontResource(StrPtr(FntFileName))
       Inc = Inc + 1
     Loop Until rc = 0 Or Inc > 10
     If rc = 0 Then
        FontList.Remove (FntFileName)
        RemoveFont = True
     End If
    End If
End Function
Public Sub RemoveAllFonts()
Dim i As Long, FntFileName As String, Inc As Integer, rc As Long
If FontList Is Nothing Then Exit Sub
For i = 0 To FontList.count - 1
    FontList.Index = i
    FntFileName = FontList.KeyToString
    Inc = 0
    FntFileName = mylcasefILE(FntFileName)
    Do
      rc = RemoveFontResource(StrPtr(FntFileName))
      Inc = Inc + 1
    Loop Until rc = 0 Or Inc > 10
Next i
Set FontList = Nothing
End Sub



Sub myform(m As Object, x As Long, y As Long, x1 As Long, y1 As Long, Optional t As Boolean = False, Optional factor As Single = 1)
Dim hRgn As Long
m.move x, y, x1, y1
If Int(25 * factor) > 2 Then
m.ScaleMode = vbPixels

hRgn = CreateRoundRectRgn(0, 0, m.ScaleX(x1, 1, 3), m.ScaleY(y1, 1, 3), 25 * factor, 25 * factor)
SetWindowRgn m.hWnd, hRgn, t
DeleteObject hRgn
m.ScaleMode = vbTwips

m.Line (0, 0)-(m.Scalewidth - dv15, m.Scaleheight - dv15), m.backcolor, BF
End If
End Sub

Sub MyRect(m As Object, mb As basket, x1 As Long, y1 As Long, way As Long, par As Variant, Optional zoom As Long = 0)
Dim R As RECT, b$
With mb
Dim x0&, y0&, x As Long, y As Long
GetXYb m, mb, x0&, y0&
x = m.ScaleX(x0& * .Xt - DXP, 1, 3)
y = m.ScaleY(y0& * .Yt - DYP, 1, 3)
If x1 >= .mx Then x1 = m.ScaleX(m.Scalewidth, 1, 3) Else x1 = m.ScaleX(x1 * .Xt, 1, 3)
If y1 >= .My Then y1 = m.ScaleY(m.Scaleheight, 1, 3) Else y1 = m.ScaleY(y1 * .Yt + .Yt, 1, 3)

SetRect R, x + zoom, y + zoom, x1 - zoom, y1 - zoom
Select Case way
Case 0
DrawEdge m.Hdc, R, CLng(par) Mod 256, CLng(par) \ 256
Case 1
DrawCaption m.hWnd, m.Hdc, R, CLng(par)
Case 2
DrawEdge m.Hdc, R, CLng(par), BF_RECT
Case 3
DrawFocusRect m.Hdc, R
Case 4
DrawFrameControl m.Hdc, R, DFC_BUTTON, DFCS_BUTTON3STATE
Case 5
b$ = Replace(CStr(par), ChrW(&HFFFFF8FB), ChrW(&H2007))
DrawText m.Hdc, StrPtr(b$), Len(CStr(par)), R, DT_CENTER + DT_NOCLIP
Case 6
DrawFrameControl m.Hdc, R, CLng(par) Mod 256, CLng(par) \ 256
Case Else
k1 = 0
MyDoEvents1 Form1
End Select
LCTbasket m, mb, y0&, x0&
End With
End Sub
Sub MyFill(m As Object, x1 As Long, y1 As Long, way As Long, par As Variant, Optional zoom As Long = 0)
Dim R As RECT, b$
Dim x As Long, y As Long
Const sp$ = " "
With players(GetCode(m))
x1 = .XGRAPH + x1
y1 = .YGRAPH + y1
x1 = m.ScaleX(x1, 1, 3)
y1 = m.ScaleY(y1, 1, 3)
x = m.ScaleX(.XGRAPH, 1, 3)
y = m.ScaleY(.YGRAPH, 1, 3)
SetRect R, x + zoom, y + zoom, x1 - zoom, y1 - zoom
Select Case way
Case 0
DrawEdge m.Hdc, R, CLng(par) Mod 256, CLng(par) \ 256
Case 1
DrawCaption m.hWnd, m.Hdc, R, CLng(par)
Case 2
DrawEdge m.Hdc, R, CLng(par), BF_RECT
Case 3
DrawFocusRect m.Hdc, R
Case 4
DrawFrameControl m.Hdc, R, DFC_BUTTON, DFCS_BUTTON3STATE
Case 5
b$ = Replace(CStr(par), ChrW(&HFFFFF8FB), ChrW(&H2007))
DrawText m.Hdc, StrPtr(b$), Len(CStr(par)), R, DT_CENTER + DT_NOCLIP
Case 6
DrawFrameControl m.Hdc, R, CLng(par) Mod 256, CLng(par) \ 256
Case Else
k1 = 0
MyDoEvents1 Form1
End Select
End With
End Sub
' ***************


Public Sub TextColor(D As Object, tc As Long)
D.forecolor = tc And &HFFFFFF
End Sub
Public Sub TextColorB(D As Object, mb As basket, tc As Long)
D.forecolor = tc And &HFFFFFF
mb.mypen = D.forecolor
End Sub

Public Sub LCTNo(DqQQ As Object, ByVal y As Long, ByVal x As Long)

''DqQQ.CurrentX = x * Xt
''DqQQ.CurrentY = y * Yt + UAddTwipsTop
''xPos = x
''yPos = y
End Sub

Public Sub LCTbasketCur(DqQQ As Object, mybasket As basket)
With mybasket
DqQQ.currentX = .curpos * .Xt
DqQQ.currentY = .currow * .Yt + .uMineLineSpace

End With
End Sub
Public Sub LCTbasket(DqQQ As Object, mybasket As basket, ByVal y As Long, ByVal x As Long)
DqQQ.currentX = x * mybasket.Xt
DqQQ.currentY = y * mybasket.Yt + mybasket.uMineLineSpace
mybasket.curpos = x
mybasket.currow = y
End Sub
Public Sub nomoveLCTC(dqq As Object, mb As basket, y As Long, x As Long, t&)
Dim oldx&, oldy&
With mb
oldx& = dqq.currentX
oldy& = dqq.currentY
dqq.DrawMode = vbXorPen
If t& = 1 Then
dqq.Line (x * .Xt, Int(y * .Yt + .uMineLineSpace))-(x * .Xt + .Xt - DXP, y * .Yt - .uMineLineSpace + .Yt - DYP), (mycolor(.mypen) Xor dqq.backcolor), BF
Else
dqq.Line (x * .Xt, Int((y + 1) * .Yt - .uMineLineSpace - .Yt \ 6 - DYP))-(x * .Xt + .Xt - DXP, (y + 1) * .Yt - .uMineLineSpace - DYP), (mycolor(.mypen) Xor dqq.backcolor), BF
End If
dqq.DrawMode = vbCopyPen
dqq.currentX = oldx&
dqq.currentY = oldy&
End With
End Sub

Public Sub oldLCTCB(dqq As Object, mb As basket, t&)

dqq.DrawMode = vbXorPen
With mb
'QRY = Not QRY
If IsWine Then
If t& = 1 Then
dqq.Line (.curpos * .Xt, .currow * .Yt + .uMineLineSpace)-(.curpos * .Xt + .Xt, .currow * .Yt - .uMineLineSpace + .Yt), (mycolor(.mypen) Xor dqq.backcolor), BF
Else
dqq.Line (.curpos * .Xt, (dqq.ScaleY((.currow + 1) * .Yt - .uMineLineSpace, 1, 3) - .Yt \ DYP \ 6 - 1) * DYP)-(.curpos * .Xt + .Xt - DXP, (.currow + 1) * .Yt - .uMineLineSpace - DYP), (mycolor(.mypen) Xor dqq.backcolor), BF

End If
Else
If t& = 1 Then
dqq.Line (.curpos * .Xt, .currow * .Yt + .uMineLineSpace)-(.curpos * .Xt + .Xt, .currow * .Yt - .uMineLineSpace + .Yt), &HFFFFFF, BF
Else
dqq.Line (.curpos * .Xt, (dqq.ScaleY((.currow + 1) * .Yt - .uMineLineSpace, 1, 3) - .Yt \ DYP \ 6 - 1) * DYP)-(.curpos * .Xt + .Xt - DXP, (.currow + 1) * .Yt - .uMineLineSpace - DYP), &HFFFFFF, BF
End If
End If
End With
dqq.DrawMode = vbCopyPen
End Sub
Public Sub LCTCnew(dqq As Object, mb As basket, y As Long, x As Long)
DestroyCaret
With mb
CreateCaret dqq.hWnd, 0, dqq.ScaleX(.Xt, 1, 3), dqq.ScaleY((.Yt - .uMineLineSpace * 2) * 0.2, 1, 3)
SetCaretPos dqq.ScaleX(x * .Xt, 1, 3), dqq.ScaleY((y + 0.8) * .Yt, 1, 3)
End With
End Sub
Public Sub LCTCB(dqq As Object, mb As basket, t&)
With mb
If t& = -1 Or Not Form1.ActiveControl Is dqq Then
        If Not t& = -1 Then
        
        Else
        If Form1.ActiveControl Is Nothing Then
        Else
            CreateCaret Form1.ActiveControl.hWnd, 0, -1, 0
            End If
            CreateCaret dqq.hWnd, 0, -1, 0
        End If
        Exit Sub
End If

If t& = 1 Then
       ' CreateCaret dqq.hWnd, 0, dqq.ScaleX(.Xt, 1, 3), dqq.ScaleY((.Yt - .uMineLineSpace * 2), 1, 3)
       CreateCaret dqq.hWnd, 0, dqq.ScaleX(.Xt, 1, 3), dqq.ScaleY(.Yt - .uMineLineSpace * 2, 1, 3)
        SetCaretPos dqq.ScaleX(.curpos * .Xt, 1, 3), dqq.ScaleY(.currow * .Yt + .uMineLineSpace, 1, 3)
        On Error Resume Next
        If Not extreme Then If INK$ = vbNullString Then dqq.Refresh
Else
    CreateCaret dqq.hWnd, 0, dqq.ScaleX(.Xt, 1, 3), .Yt \ DYP \ 6 + 1
        
            SetCaretPos dqq.ScaleX(.curpos * .Xt, 1, 3), dqq.ScaleY((.currow + 1) * .Yt - .uMineLineSpace, 1, 3) - .Yt \ DYP \ 6 - 1
        On Error Resume Next
        If Not extreme Then If INK$ = vbNullString Then dqq.Refresh
End If
dqq.DrawMode = vbCopyPen
dqq.currentX = .curpos * .Xt
dqq.currentY = .currow * .Yt + .uMineLineSpace
End With
End Sub
Public Sub SetDouble(dq As Object)

SetTextSZ dq, players(GetCode(dq)).SZ, 2


End Sub

Public Sub SetNormal(dq As Object)
SetTextSZ dq, players(GetCode(dq)).SZ, 1
End Sub

Sub BoxBigNew(dqq As Object, mb As basket, x1&, y1&, c As Long)
With mb
 If TypeOf dqq Is MetaDc Then
 dqq.Line2 .curpos * .Xt - DXP, .currow * .Yt - DYP, x1& * .Xt - DXP + .Xt, y1& * .Yt + .Yt - DYP, mycolor(c), , True
 
 Else
dqq.Line (.curpos * .Xt - DXP, .currow * .Yt - DYP)-(x1& * .Xt - DXP + .Xt, y1& * .Yt + .Yt - DYP), mycolor(c), B
End If
End With

End Sub
Sub CircleBig(dqq As Object, mb As basket, x1&, y1&, c As Long, el As Boolean)
Dim x&, y&
With mb
x& = .curpos
y& = .currow
dqq.fillcolor = mycolor(c)
dqq.fillstyle = vbFSSolid
If TypeOf dqq Is MetaDc Then
If el Then
DrawCircleApi dqq, Form1.ScaleX(((x& + x1& + 1) / 2 * .Xt) - DXP, 1, 3), Form1.ScaleY(((y& + y1& + 1) / 2 * .Yt) - DYP, 1, 3), Form1.ScaleX(RMAX((x1& - x& + 1) * .Xt, (y1& - y& + 1) * .Yt) / 2 - DYP, 1, 3), mycolor(c), ((y1& - y& + 1) * .Yt - DYP) / ((x1& - x& + 1) * .Xt - DXP)
Else
DrawCircleApi dqq, Form1.ScaleX(((x& + x1& + 1) / 2 * .Xt) - DXP, 1, 3), Form1.ScaleY(((y& + y1& + 1) / 2 * .Yt) - DYP, 1, 3), Form1.ScaleX(RMAX((x1& - x& + 1) * .Xt, (y1& - y& + 1) * .Yt) / 2 - DYP, 1, 3), mycolor(c)
End If
Else
If el Then
dqq.Circle (((x& + x1& + 1) / 2 * .Xt) - DXP, ((y& + y1& + 1) / 2 * .Yt) - DYP), RMAX((x1& - x& + 1) * .Xt, (y1& - y& + 1) * .Yt) / 2 - DYP, mycolor(c), , , ((y1& - y& + 1) * .Yt - DYP) / ((x1& - x& + 1) * .Xt - DXP)
Else
dqq.Circle (((x& + x1& + 1) / 2 * .Xt) - DXP, ((y& + y1& + 1) / 2 * .Yt) - DYP), (RMIN((x1& - x& + 1) * .Xt, (y1& - y& + 1) * .Yt) / 2 - DYP), mycolor(c)

End If
End If
dqq.fillstyle = vbFSTransparent
End With
End Sub
Sub Ffill(dqq As Object, x1 As Long, y1 As Long, c As Long, v As Boolean)
Dim osm
With players(GetCode(dqq))

If Not .IamEmf Then
osm = dqq.ScaleMode
dqq.ScaleMode = vbPixels
End If
dqq.fillcolor = mycolor(c)
dqq.fillstyle = vbFSSolid
If v Then
ExtFloodFill dqq.Hdc, dqq.ScaleX(x1, 1, 3), dqq.ScaleY(y1, 1, 3), dqq.point(dqq.ScaleX(x1, 1, 3), dqq.ScaleY(y1, 1, 3)), FLOODFILLSURFACE
Else
ExtFloodFill dqq.Hdc, dqq.ScaleX(x1, 1, 3), dqq.ScaleY(y1, 1, 3), mycolor(.mypen), FLOODFILLBORDER
End If
If Not .IamEmf Then
dqq.ScaleMode = osm
End If
dqq.fillstyle = vbFSTransparent
End With
'LCT Dqq, y&, x&
End Sub

Sub BoxColorNew(dqq As Object, mb As basket, x1&, y1&, c As Long)
Dim addpixels As Long
With mb
'If InternalLeadingSpace() = 0 And .MineLineSpace = 0 Then
addpixels = 2

'Else
'addpixels = 2
'End If
If TypeOf dqq Is MetaDc Then
dqq.Line2 .curpos * .Xt, .currow * .Yt, x1& * .Xt + .Xt - 2 * DXP, y1& * .Yt + .Yt - addpixels * DYP, mycolor(c), True, True
Else
dqq.Line (.curpos * .Xt, .currow * .Yt)-(x1& * .Xt + .Xt - 2 * DXP, y1& * .Yt + .Yt - addpixels * DYP), mycolor(c), BF
End If
End With
End Sub
Sub BoxImage(d1 As Object, mb As basket, x1&, y1&, F As String, df&, s As Boolean)
'
Dim p As Picture, scl As Double, x2&, dib As Object, aPic As StdPicture

If df& > 0 Then
df& = df& * DXP '* 20

Else

df& = 0
End If
With mb
x1& = .curpos + x1& - 1
x2& = x1&
y1& = .currow + y1& - 1
On Error Resume Next
 If (Left$(F$, 4) = "cDIB" And Len(F$) > 12) Then
   Set dib = New cDIBSection
  If Not cDib(F$, dib) Then
    dib.create x1&, y1&
    dib.Cls d1.backcolor
  End If
      Set p = dib.Picture
    Set dib = Nothing
 Else
        If ExtractType(F, 0) = vbNullString Then
        F = F + ".bmp"
        End If
        FixPath F
        
    If CFname(F) <> "" Then
    F = CFname(F)
    Set aPic = LoadMyPicture(GetDosPath(F$))
     If aPic Is Nothing Then Exit Sub
    Set p = aPic
                                            

    Else
    Set dib = New cDIBSection
    dib.create x1&, y1&
    dib.Cls d1.backcolor
    Set p = dib.Picture
    Set dib = Nothing
    End If
End If

If Err.Number > 0 Then Exit Sub

If s Then
scl = (y1& - .currow + 1) * .Yt - df&
If p.Type = vbPicTypeBitmap Then
d1.PaintPicture p, .curpos * .Xt, .currow * .Yt, (x1& - .curpos + 1) * .Xt - df&, scl, , , , , vbSrcCopy
Else
d1.PaintPicture p, .curpos * .Xt, .currow * .Yt, (x1& - .curpos + 1) * .Xt - df&, scl
End If
Else
scl = p.Height * ((x1& - .curpos + 1) * .Xt - df&) / p.Width
If p.Type = vbPicTypeBitmap Then
d1.PaintPicture p, .curpos * .Xt, .currow * .Yt, (x1& - .curpos + 1) * .Xt - df&, scl, , , , , vbSrcCopy
Else
d1.PaintPicture p, .curpos * .Xt, .currow * .Yt, (x1& - .curpos + 1) * .Xt - df&, scl
End If
End If
y1& = -Int(-((scl) / .Yt))
Set p = Nothing
''LCT d1, .currow, .curpos
End With
End Sub

Sub sprite(bstack As basetask, ByVal F As String, rst As String)

On Error GoTo SPerror
Dim d1 As Object, amask$, aPic As StdPicture
Set d1 = bstack.Owner
Dim raster As New cDIBSection
Dim p As Double, i As Long, rot As Double, sp As Double
Dim Pcw As Long, Pch As Long, blend As Double, NoUseBack As Boolean

If Not cDib(F, raster) Then
    If CFname(F) <> "" Then
        F = CFname(F)
        Set aPic = LoadMyPicture(GetDosPath(F$))
        If aPic Is Nothing Then Exit Sub
        If aPic.Type = 4 Then
        With players(GetCode(bstack.Owner))
        raster.emfSizeFactor = 100 '.MAXXGRAPH \ dv15
        End With
        
        
        End If
        
        raster.CreateFromPicture aPic
        If raster.bitsPerPixel <> 24 Then
            Conv24 raster
        Else
            CheckOrientation raster, F
        End If
    Else
        
        BACKSPRITE = vbNullString
        Exit Sub
    End If
End If
If raster.Width = 0 Then
    BACKSPRITE = vbNullString
    Set raster = Nothing
    Set d1 = Nothing
    Exit Sub
End If
i = -1
sp = 100!
blend = 100!
If FastSymbol(rst$, ",") Then
    If IsExp(bstack, rst$, p, , True) Then i = CLng(p) Else i = -players(GetCode(d1)).Paper
    If FastSymbol(rst$, ",") Then
        If IsExp(bstack, rst$, p, , True) Then rot = p
        If FastSymbol(rst$, ",") Then
            If Not IsExp(bstack, rst$, sp) Then sp = 100!
            If FastSymbol(rst$, ",") Then
                If IsExp(bstack, rst$, blend) Then
                    blend = Abs(Int(blend)) Mod 101
                    If FastSymbol(rst$, ",") Then GoTo cont0
                ElseIf IsStrExp(bstack, rst$, amask$, Len(bstack.tmpstr) = 0) Then
                    blend = 100!
                    If FastSymbol(rst$, ",") Then GoTo cont0
                ElseIf FastSymbol(rst$, ",") Then
                blend = 100!
cont0:
                    If Not IsExp(bstack, rst$, p, , True) Then
                            MyEr "missing parameter", "λείπει παράμετρος"
                            Exit Sub
                    End If
                    NoUseBack = CBool(p)
                Else
                    MyEr "missing parameter", "λείπει παράμετρος"
                End If
                
                
            End If
            End If
        End If
Else
        Pcw = raster.Width \ 2
        Pch = raster.Height \ 2
        With players(GetCode(d1))
        raster.PaintPicture d1.Hdc, Int(d1.ScaleX(.XGRAPH, 1, 3) - Pcw), Int(d1.ScaleX(.YGRAPH, 1, 3) - Pch)
        End With
    GoTo cont1
End If
If sp <= 0 Then sp = 0
If i > 0 Then i = QBColor(i) Else i = -i
RotateDib bstack, raster, rot, sp, i, NoUseBack, (blend), amask$
Pcw = raster.Width \ 2
Pch = raster.Height \ 2
With players(GetCode(d1))
raster.PaintPicture d1.Hdc, Int(d1.ScaleX(.XGRAPH, 1, 3) - Pcw), Int(d1.ScaleX(.YGRAPH, 1, 3) - Pch)
End With
cont1:
If Not bstack.toprinter Then
GdiFlush
End If
Set raster = Nothing
'MyDoEvents1 d1
Set d1 = Nothing
Exit Sub
SPerror:
 BACKSPRITE = vbNullString
Set raster = Nothing
End Sub
Sub spriteGDI(bstack As basetask, rst As String)
Dim NoUseBack As Boolean, usehandler As mHandler
If bstack.lastobj Is Nothing Then
Err1:
    MyEr "Expecting a memory Buffer", "Περίμενα διάρθρωση μνήμης"
    Exit Sub
End If
If Not TypeOf bstack.lastobj Is mHandler Then GoTo Err1
Set usehandler = bstack.lastobj
If Not usehandler.t1 = 2 Then GoTo Err1
Dim d1 As Object
Set d1 = bstack.Owner
Dim p, i As Long, mem As MemBlock, blend, sp, rot As Single
Set mem = usehandler.objref
i = -1
sp = 100!
blend = 0!
If FastSymbol(rst$, ",") Then
    If IsExp(bstack, rst$, p, , True) Then i = CLng(p) Else i = -players(GetCode(d1)).Paper
    If FastSymbol(rst$, ",") Then
        If IsExp(bstack, rst$, p, , True) Then rot = p
        If FastSymbol(rst$, ",") Then
            If Not IsExp(bstack, rst$, sp, , True) Then sp = 100!
            If FastSymbol(rst$, ",") Then
                If IsExp(bstack, rst$, blend) Then blend = 100 - Abs(Int(blend)) Mod 101
                If FastSymbol(rst$, ",") Then
                    If Not IsExp(bstack, rst$, p) Then
                        MyEr "missing parameter", "λείπει παράμετρος"
                        Exit Sub
                    End If
                    NoUseBack = CBool(p)
                End If
            End If
        End If
    End If
End If
If sp <= 0 Then sp = 0
If i > 0 Then i = QBColor(i Mod 16) Else i = -i
If Not bstack.toprinter Then
GdiFlush
End If
mem.DrawSpriteToHdc bstack, NoUseBack, rot, (sp), (blend), i

'MyDoEvents1 d1
Set d1 = Nothing
Set bstack.lastobj = Nothing
Exit Sub
SPerror:
Set bstack.lastobj = Nothing
 BACKSPRITE = vbNullString
End Sub

Sub ThumbImage(d1 As Object, x1 As Long, y1 As Long, F As String, border As Long, tpp As Long, ttl1$)
On Error Resume Next
With players(GetCode(d1))
If Left$(F, 4) = "cDIB" And Len(F) > 12 Then
Dim ph As New cDIBSection
If cDib(F, ph) Then
ph.ThumbnailPartPaint d1, x1 / tpp, y1 / tpp, 0, 0, border <> 0, , ttl1$, .XGRAPH / tpp, .YGRAPH / tpp
End If
End If
End With
End Sub
Sub ThumbImageDib(d1 As Object, x1 As Long, y1 As Long, ph As Object, border As Long, tpp As Long, ttl1$)
On Error Resume Next
Dim pointer2dib As cDIBSection
Set pointer2dib = ph
With players(GetCode(d1))
    pointer2dib.ThumbnailPartPaint d1, x1 / tpp, y1 / tpp, 0, 0, border <> 0, , ttl1$, .XGRAPH / tpp, .YGRAPH / tpp
End With
Set pointer2dib = Nothing
End Sub
Sub SImage(d1 As Object, x1 As Long, y1 As Long, F As String)
'
Dim p As Picture, aPic As StdPicture
On Error Resume Next
With players(GetCode(d1))
If Left$(F, 4) = "cDIB" And Len(F) > 12 Then
Dim ph As New cDIBSection
If cDib(F, ph) Then
If x1 = 0 Then
ph.PaintPicture d1.Hdc, CLng(d1.ScaleX(.XGRAPH, 1, 3)), CLng(d1.ScaleX(.YGRAPH, 1, 3))
Exit Sub
Else
If y1 = 0 Then y1 = Abs(ph.Height * x1 / ph.Width)
ph.StretchPictureH d1.Hdc, CLng(d1.ScaleX(.XGRAPH, 1, 3)), CLng(d1.ScaleX(.YGRAPH, 1, 3)), CLng(d1.ScaleX(x1, 1, 3)), CLng(d1.ScaleX(y1, 1, 3))
Exit Sub
End If
End If
ElseIf CFname(F) <> "" Then
    F = CFname(F)
     Set aPic = LoadMyPicture(GetDosPath(F$), , , True)
     If aPic Is Nothing Then Exit Sub
     Set p = aPic
Else
If y1 = 0 Then y1 = x1
d1.Line (.XGRAPH, .YGRAPH)-(x1, y1), .Paper, BF
d1.currentX = .XGRAPH
d1.currentY = .YGRAPH
Exit Sub
End If
If x1 = 0 Then
x1 = d1.ScaleX(p.Width, vbHimetric, vbTwips)

If y1 = 0 Then y1 = p.Height * d1.ScaleX(p.Width, vbHimetric, vbTwips) / p.Width
Else
If y1 = 0 Then y1 = p.Height * x1 / p.Width
End If
If Err.Number > 0 Then Exit Sub

If p.Type = vbPicTypeBitmap Then
d1.PaintPicture p, .XGRAPH, .YGRAPH, x1, y1, , , , , vbSrcCopy
Else
d1.PaintPicture p, .XGRAPH, .YGRAPH, x1, y1
End If
Set p = Nothing
End With
' UpdateWindow d1.hwnd
End Sub
Public Function LoadMyPicture(s1$, Optional useback As Boolean = False, Optional bColor As Variant = 0&, Optional includeico As Boolean = False) As StdPicture
Dim s As String
Err.Clear
   On Error Resume Next
                    If s1$ <> vbNullString Then
                        s$ = UCase(ExtractType(s1$))
                        If LenB(s$) = 0 Then s$ = "Bmp": s1$ = s1$ + ".bmp"
                        Select Case s
                        Case "JPG", "BMP", "WMF", "EMF", "ICO", "DIB"
                        
                           Set LoadMyPicture = LoadPicture(s1$)
                           If Err.Number > 0 Then
                           Err.Clear
                           If useback Then
                              Set LoadMyPicture = LoadPictureGDIPlus(s1$, , , bColor)
                           Else
                              Set LoadMyPicture = LoadPictureGDIPlus(s1$)
                            End If
                           ElseIf useback Then
                                Set LoadMyPicture = LoadPictureGDIPlus(s1$, , , bColor)
                                If Err.Number > 0 Then Err.Clear
                           End If
                           If Err.Number > 0 Then
                           Err.Clear
                           
                           Set LoadMyPicture = LoadPicture("")
                           End If
                           If LoadMyPicture Is Nothing Then
                           Set LoadMyPicture = LoadPicture("")
                           End If
                        Case Else
                            If includeico And Not useback Then
                            Set LoadMyPicture = LoadPicture(s1$)
                                If Err.Number > 0 Then
                                    Err.Clear
                                    GoTo conthere
                                End If
                            Else
conthere:
                          If useback Then
                              Set LoadMyPicture = LoadPictureGDIPlus(s1$, , , bColor)
                           Else
                              Set LoadMyPicture = LoadPictureGDIPlus(s1$)
                            End If
                            End If
                            If Err.Number > 0 Then
                           Err.Clear
                          
                           Set LoadMyPicture = LoadPicture("")
                           End If
                           If LoadMyPicture Is Nothing Then
                           Set LoadMyPicture = LoadPicture("")
                           End If
                        End Select
                    End If
                          
End Function

Public Function MyTextWidth(D As Object, a$) As Long
Dim nr As RECT
CalcRect D.Hdc, a$, nr
'MyTextWidth = nr.Right * d.ScaleX(1, 3, 1)
MyTextWidth = nr.Right * DXP
End Function
Public Sub CalcRect(mHdc As Long, c As String, R As RECT)
R.top = 0
R.Left = 0
R.Right = 20000
R.Bottom = 20000
DrawTextEx mHdc, StrPtr(c), -1, R, DT_CALCRECT Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
End Sub
Public Sub CalcRectNoSingle(mHdc As Long, c As String, R As RECT)
R.top = 0
R.Left = 0
R.Right = 20000
R.Bottom = 20000
DrawTextEx mHdc, StrPtr(c), -1, R, DT_CALCRECT Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
End Sub

Public Sub PrintLineControlSingle(mHdc As Long, c As String, R As RECT)
    DrawTextEx mHdc, StrPtr(c), -1, R, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
    End Sub
'
Public Sub MyPrintNew(ddd As Object, UAddTwipsTop, s$, Optional cr As Boolean = False, Optional fake As Boolean = False)

Dim nr As RECT, nl As Long, mytop As Long
mytop = ddd.currentY
If s$ = vbNullString Then
nr.Left = 0: nr.Right = 0: nr.top = 0: nr.Bottom = 0
CalcRect ddd.Hdc, " ", nr
nr.Left = ddd.currentX / dv15
nr.Right = nr.Right + nr.Left
nr.top = ddd.currentY / dv15
nr.Bottom = nr.top + nr.Bottom
nl = (nr.Bottom + 1) * dv15
If cr Then
ddd.currentY = (nr.Bottom + 1) * dv15 + UAddTwipsTop ''2
ddd.currentX = 0
Else
ddd.currentX = nr.Right * dv15
End If
Else
nr.Left = 0: nr.Right = 0: nr.top = 0: nr.Bottom = 0
CalcRect ddd.Hdc, s$, nr
nr.Left = ddd.currentX / dv15
nr.Right = nr.Right + nr.Left
nr.top = ddd.currentY / dv15
nr.Bottom = nr.top + nr.Bottom
nl = (nr.Bottom + 1) * dv15
If Not fake Then
If nr.Left * dv15 < ddd.Width Then PrintLineControlSingle ddd.Hdc, s$, nr
End If
If cr Then
ddd.currentY = nl + UAddTwipsTop ''* 2
ddd.currentX = 0
Else
ddd.currentY = mytop
ddd.currentX = nr.Right * dv15
End If
End If

End Sub
Public Sub MyPrint(ddd As Object, s$)
Dim nr As RECT, nl As Long
If s$ = vbNullString Then
    nr.Left = 0: nr.Right = 0: nr.top = 0: nr.Bottom = 0
    CalcRect ddd.Hdc, " ", nr
    nr.Left = ddd.currentX / dv15
    nr.Right = nr.Right + nr.Left
    nr.top = ddd.currentY / dv15
    nr.Bottom = nr.top + nr.Bottom
    nl = (nr.Bottom + 1) * dv15
    ddd.currentY = (nr.Bottom + 1) * dv15
    ddd.currentX = 0
Else
nr.Left = 0: nr.Right = 0: nr.top = 0: nr.Bottom = 0
CalcRect ddd.Hdc, s$, nr
nr.Left = ddd.currentX / dv15
nr.Right = nr.Right + nr.Left
nr.top = ddd.currentY / dv15
nr.Bottom = nr.top + nr.Bottom
nl = (nr.Bottom + 1) * dv15
If nr.Left * dv15 <= ddd.Width Then PrintLineControlSingle ddd.Hdc, s$, nr
ddd.currentY = nl
ddd.currentX = 0
End If
End Sub

Public Function TextWidth(ddd As Object, a$) As Long
Dim nr As RECT
CalcRect ddd.Hdc, a$, nr
TextWidth = nr.Right * dv15
End Function
Public Function TextWidth2(ddd As Object, a$) As Long
Dim nr As RECT
CalcRectNoSingle ddd.Hdc, a$, nr
TextWidth2 = nr.Right * dv15

End Function
Public Function TextWidthPixels(ddd As Object, a$) As Long
Dim nr As RECT
CalcRect ddd.Hdc, a$, nr
TextWidthPixels = nr.Right
End Function
Private Function TextHeight(ddd As Object, a$) As Long
Dim nr As RECT
CalcRect ddd.Hdc, a$, nr

TextHeight = nr.Bottom * dv15
End Function
Private Function TextHeight2(ddd As Object, a$) As Long
Dim nr As RECT
CalcRectNoSingle ddd.Hdc, a$, nr

TextHeight2 = nr.Bottom * dv15
End Function

Public Sub PlainBaSket(ddd As Object, mybasket As basket, ByVal what As String, Optional ONELINE As Boolean = False, Optional nocr As Boolean = False, Optional plusone As Long = 2, Optional clearline As Boolean = False, Optional processcr As Boolean = False, Optional semicolon As Boolean = False)

Dim PX As Long, PY As Long, R As Long, p$, c$, LEAVEME As Boolean, nr As RECT, nr2 As RECT, w As Integer
Dim p2 As Long, mUAddPixelsTop As Long
Dim pixX As Long, pixY As Long
Dim rTop As Long, rBottom As Long
Dim lenw&, realR&, realstop&, R1 As Long, WHAT1$, ff As Long, LL$(), must As Long
If processcr Then
    If Len(what$) = 0 Then Exit Sub
    LL$() = Split(what, vbLf)
    what = LL$(0)
    ff = 0
End If

Dim a1() As Integer, A2() As Integer
'' LEAVEME = False -  NOT NEEDED
again:
nr.Left = 0
realR& = 0

With mybasket
    mUAddPixelsTop = mybasket.uMineLineSpace \ dv15  ' for now
    PX = .curpos
    PY = .currow
    If PY = .My And .double Then
        If ddd Is Form1.PrinterDocument1 Then
            getnextpage
            With nr
                .top = PY * pixY + mUAddPixelsTop
                .Bottom = .top + pixY - p2
            End With
            PY = 0
            .currow = 0
        Else
            ScrollUpNew ddd, mybasket
        End If
        PY = .currow
    End If
    p2 = mUAddPixelsTop * 2
    pixX = .Xt / dv15
    pixY = .Yt / dv15
    With nr
        .Left = PX * pixX
        .Right = .Left + pixX
        .top = PY * pixY + mUAddPixelsTop
        .Bottom = .top + pixY - mUAddPixelsTop * 2
    End With
    rTop = PY * pixY
    rBottom = rTop + pixY - plusone
    lenw& = RealLen(what)
    WHAT1$ = what + " "
    ReDim a1(Len(WHAT1$) + 10)
    ReDim A2(Len(WHAT1$) + 10)
    Dim skip As Boolean
    skip = GetStringTypeExW(&HB, 4, StrPtr(WHAT1$), Len(WHAT1$), a1(0)) = 0
    skip = GetStringTypeExW(&HB, 2, StrPtr(WHAT1$), Len(WHAT1$), A2(0)) = 0 Or skip
    Dim ii As Long, mark1 As Long, mr As Long, ML As Long
    
    Do While (lenw& - R) >= .mx - PX And (.mx - PX) > 0
     If ddd.FontTransparent = False Then
        With nr2
            .Left = PX * pixX
            .Right = mybasket.mx * pixX + 1
            .top = rTop
            .Bottom = rBottom
        End With
         FillBack ddd.Hdc, nr2, .Paper
         End If
        ddd.currentX = PX * .Xt
        ddd.currentY = PY * .Yt + .uMineLineSpace
        R1 = .mx - PX - 1 + R

        If ddd.currentX = 0 And clearline Then
            If Not TypeOf ddd Is MetaDc Then ddd.Line (0&, PY * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (PY) * .Yt + .Yt - 1 * DYP), .Paper, BF
        End If
        
        Do
            If ONELINE And nocr And PX > .mx Then what = vbNullString: Exit Do
            c$ = Mid$(WHAT1$, R + 1, 1)
            w = AscW(c$)
            If w > -10241 And w < -9984 Then
                c$ = Mid$(WHAT1$, R + 1, 2)
                R = R + 1
                GoTo checkcombine
            ElseIf A2(R) = 0 And a1(R) = 0 Then
                R = R + 1
                GoTo cont0
            ElseIf (A2(R) And 254) = 2 And (a1(R) And &H8000) <> 0 Then
                mark1 = R + 1
                If processcr Then
                For ii = R + 2 To Len(what$)
                    If Not A2(ii) > 2 Then If (A2(ii) And 2) <> 2 And (a1(ii) And 7) = 0 Then Exit For
                   If (RealLen(Mid$(what$, mark1, ii - mark1 + 2)) + .curpos) > .mx Then
                   If TextWidth(ddd, Mid$(what$, mark1, ii - mark1 + 2)) \ .Xt > (.mx - .curpos - 1) Then
                   c$ = Mid$(what$, mark1, ii - mark1 + 1)
                
                LL(ff) = Mid$(what$, mark1 + Len(c$) + 1)
                If Len(LL(ff)) > 0 Then ff = ff - 1
                lenw& = R + Len(c$) - 1
             
                   
                   Exit For
                   End If
                   End If
                Next ii
                Else
                For ii = R + 2 To Len(what$)
                    If Not A2(ii) > 2 Then If (A2(ii) And 2) <> 2 And (a1(ii) And 7) = 0 Then Exit For
                Next ii
                End If
                c$ = Mid$(what$, mark1, ii - mark1 + 1)
                R = R + Len(c$): If ii > mark1 Then R = R - 1
                mark1 = nr.Right
                nr.Right = (PX + Len(what$)) * pixX + 1
                DrawText ddd.Hdc, StrPtr(c$), -1, nr, DT_SINGLELINE Or DT_NOPREFIX + DT_NOCLIP
                mark1 = TextWidth(ddd, c$)
                mark1 = mark1 \ .Xt - (mark1 Mod .Xt > 0)
                nr.Right = nr.Left + mark1 * pixX + 1
                realR& = realR + mark1
                ddd.currentX = nr.Right * DXP
                If processcr Then
                
                .curpos = 0
                End If
                ElseIf nounder32(c$) Then
checkcombine:
                If Not skip Then
                    If (a1(R + 1) And &H87F8) = 0 And (a1(R + 1) And 7) <> 0 Then
                        Do
                            p$ = Mid$(WHAT1$, R + 2, 1)
                            If Not nounder32(p$) Then Mid$(WHAT1$, R + 2, 1) = " ": Exit Do
                            c$ = c$ + p$
                            R = R + 1
                            If R >= R1 Then Exit Do
                         Loop Until (a1(R + 1) And 7) = 0
                     End If
                 End If
                 DrawText ddd.Hdc, StrPtr(c$), -1, nr, DT_SINGLELINE Or DT_CENTER Or DT_NOPREFIX + DT_NOCLIP
            Else
                If c$ = Chr$(7) Then
                    If Not ddd Is Form1.PrinterDocument1 Then Beep
                    R = R + 1: realR = realR - 1:
                    GoTo cont0
                End If
                If processcr Then
                    realR& = realR + 1
                    If c$ = ChrW(9) Then
                        what$ = space$(.Column - (PX + realR - 1) Mod (.Column + 1)) + Mid$(WHAT1$, R + 2)
                        R = 0
                        .curpos = PX + realR
                        If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
                        GoTo again
                    ElseIf c$ = ChrW(13) Then
                        If Mid$(WHAT1$, R + 2, 1) = ChrW(10) Then R = R + 1
                        .curpos = 0
                        If PY + 1 >= .My Then
                            If ddd Is Form1.PrinterDocument1 Then
                                getnextpage
                                With nr
                                    .top = PY * pixY + mUAddPixelsTop
                                    .Bottom = .top + pixY - p2
                                End With
                                PY = 0
                                .currow = 0
                            Else
                                ScrollUpNew ddd, mybasket
                            End If
                        Else
                            .currow = PY + 1
                        End If
                        ff = ff + 1
                        If ff < UBound(LL) Then
                            If Right$(LL$(ff), 1) <> vbCr Then
                                what = LL$(ff) + vbCr
                            Else
                                what = LL$(ff)
                            End If
                            R = 0
                            GoTo again
                        ElseIf ff = UBound(LL) Then
                            what = LL$(ff)
                            R = 0
                            GoTo again
                        Else
                            Exit Do
                        End If
                    ElseIf c$ = ChrW(10) Then
                        .curpos = 0
                        If Not TypeOf ddd Is MetaDc Then
                        If PY + 1 = .My Then
                            If ddd Is Form1.PrinterDocument1 Then
                                getnextpage
                                With nr
                                    .top = PY * pixY + mUAddPixelsTop
                                    .Bottom = .top + pixY - p2
                                End With
                                PY = 0
                                .currow = 0
                            Else
                                ScrollUpNew ddd, mybasket
                            End If
                        Else
                            .currow = PY + 1
                        End If
                        Else
                            .currow = PY + 1
                        End If
                        what$ = Mid$(WHAT1$, R + 2)
                        If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
                        R = 0
                        GoTo again
                    End If
                End If
            End If
            R = R + 1
            With nr
                .Left = .Right
                .Right = .Left + pixX
            End With
cont0:
            ddd.currentX = (PX + realR) * .Xt
            realR = realR + 1
            If R >= lenw& Then
                R = lenw& + 1
                lenw& = lenw& - 1
                Exit Do
            End If
            If realR > .mx - PX - 1 Then Exit Do
        Loop
        If realR < .mx - PX - 1 Then GoTo cont1
        .curpos = PX + realR
        If Not ONELINE Then PX = 0
        If nocr Then GoTo jumpexit Else PY = PY + 1
        If PY >= .My And Not ONELINE Then
        If processcr Then
        If ff < UBound(LL) Then
        GoTo skipthis
        End If
        End If
        If Not TypeOf ddd Is MetaDc Then
        If ddd Is Form1.PrinterDocument1 Then
                getnextpage
                With nr
                    .top = PY * pixY + mUAddPixelsTop
                    .Bottom = .top + pixY - p2
                End With
                PY = 0
                .currow = 0
            Else
                ScrollUpNew ddd, mybasket
            End If
            PY = PY - 1
        End If
        End If
skipthis:
        
        If ONELINE Then
            LEAVEME = True
            Exit Do
        Else
            With nr
               .Left = PX * pixX
               .Right = .Left + pixX
               .top = PY * pixY + mUAddPixelsTop
               .Bottom = .top + pixY - p2
            End With
            rTop = PY * pixY
            rBottom = rTop + pixY - plusone
        End If
        realR& = 0
    Loop
    If LEAVEME Then
        With mybasket
            .curpos = PX
            .currow = PY
        End With
        GoTo jumpexit
    End If
    If ddd.FontTransparent = False Then
        With nr2
            .Left = PX * pixX
            .Right = (PX + Len(what$)) * pixX + 1
            .top = rTop
            .Bottom = rBottom
        End With
        FillBack ddd.Hdc, nr2, mybasket.Paper
    End If
    realR& = 0
    If Len(what$) > R Then
        ddd.currentX = PX * .Xt
        ddd.currentY = PY * .Yt + .uMineLineSpace
        If ddd.currentX = 0 And clearline Then
            If Not TypeOf ddd Is MetaDc Then ddd.Line (0&, PY * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (PY) * .Yt + .Yt - 1 * DYP), .Paper, BF
        End If
        R1 = Len(what$) - 1
        For R = R To R1
            c$ = Mid$(WHAT1$, R + 1, 1)
            w = AscW(c$)
            If w > -10241 And w < -9984 Then
                c$ = Mid$(WHAT1$, R + 1, 2)
                R = R + 1
                GoTo checkcombine1
            ElseIf A2(R) = 0 And a1(R) = 0 Then
                R = R + 1
                GoTo cont1
            ElseIf (A2(R) And 254) = 2 And (a1(R) And &H8000) <> 0 Then
                mark1 = R + 1
                For ii = R + 2 To Len(what$)
                    If Not A2(ii) > 2 Then If (A2(ii) And 2) <> 2 And (a1(ii) And 7) = 0 Then Exit For
                Next ii
                c$ = Mid$(what$, mark1, ii - mark1 + 1)
                R = R + Len(c$)
                If ii > mark1 Then R = R - 1
                mark1 = TextWidth(ddd, c$) \ DXP
                nr.Right = nr.Left + mark1 + 1
                DrawTextEx ddd.Hdc, StrPtr(c$), -1, nr, DT_SINGLELINE Or DT_NOPREFIX + DT_NOCLIP, 0
                .curpos = nr.Right \ pixX
                realR& = realR + mark1 \ pixX - (mark1 Mod pixX > 0) * 1
                ddd.currentX = .curpos * .Xt
                If Not processcr Then GoTo contNew
                
                GoTo again
            ElseIf nounder32(c$) Then
checkcombine1:
                If Not skip Then
                    If (a1(R + 1) And &H87F8) = 0 And (a1(R + 1) And 7) <> 0 Then
                        Do
                            p$ = Mid$(WHAT1$, R + 2, 1)
                            If Not nounder32(p$) Then Mid$(WHAT1$, R + 2, 1) = " ": Exit Do
                            c$ = c$ + p$
                            R = R + 1
                            If R >= R1 Then Exit Do
                        Loop Until (a1(R + 1) And 7) = 0
                    End If
                End If
                ddd.currentX = ddd.currentX + .Xt
            Else
CHECK1:
                If c$ = Chr$(7) Then
                    If Not ddd Is Form1.PrinterDocument1 Then Beep
                    GoTo cont1
                End If
                If processcr Then
                    realR& = realR + 1
                    If c$ = ChrW(9) Then
                        what$ = space$(.Column - (PX + realR - 1) Mod (.Column + 1)) + Mid$(WHAT1$, R + 2)
                        R = 0
                        .curpos = PX + realR
                        If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
                        GoTo again
                    ElseIf c$ = ChrW(13) Then
                        If Mid$(WHAT1$, R + 2, 1) = ChrW(10) Then R = R + 1
                        .curpos = 0
                        PX = 0
                        ddd.currentX = 0
                        realR& = 0
                        c$ = ""
                        If Not TypeOf ddd Is MetaDc Then
                            If PY + 1 = .My Then
                                If ddd Is Form1.PrinterDocument1 Then
                                    getnextpage
                                    With nr
                                        .top = PY * pixY + mUAddPixelsTop
                                        .Bottom = .top + pixY - p2
                                    End With
                                    PY = 0
                                    .currow = 0
                                Else
                                    ScrollUpNew ddd, mybasket
                                End If
                            Else
                                .currow = PY + 1
                            End If
                        Else
                                .currow = PY + 1
                        End If
                        ff = ff + 1
                        If ff < UBound(LL) Then
                            If Right$(LL$(ff), 1) <> vbCr Then
                                what = LL$(ff) + vbCr
                            Else
                                what = LL$(ff)
                            End If
                            R = 0
                            GoTo again
                        ElseIf ff = UBound(LL) Then
                            what = LL$(ff)
                            R = 0
                            GoTo again
                        Else
                            Exit For
                        End If
                    ElseIf c$ = ChrW(10) Then
                        .curpos = 0
                        If Not TypeOf ddd Is MetaDc Then
                        If PY + 1 >= .My Then
                            If ddd Is Form1.PrinterDocument1 Then
                                getnextpage
                                With nr
                                    .top = PY * pixY + mUAddPixelsTop
                                    .Bottom = .top + pixY - p2
                                End With
                                PY = 0
                                .currow = 0
                            Else
                                ScrollUpNew ddd, mybasket
                            End If
                        Else
                            .currow = PY + 1
                        End If
                        Else
                            .currow = PY + 1
                        End If
                        what$ = Mid$(WHAT1$, R + 2)
                        If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
                        R = 0
                        If Len(what$) = 0 Then GoTo contNew
                        GoTo again
                    End If
                End If
            End If
            DrawText ddd.Hdc, StrPtr(c$), -1, nr, DT_SINGLELINE Or DT_CENTER Or DT_NOPREFIX + DT_NOCLIP
            realR& = realR + 1
contNew:
            With nr
               .Left = .Right
               .Right = .Left + pixX
            End With
cont1:
        Next R
        If Not processcr Then
            If semicolon Then
            ElseIf lenw& > realR& Then
                realR& = ((realR& + .Column + 1) \ (.Column + 1)) * (.Column + 1)
            End If
        ElseIf ff < UBound(LL) Then
        If ff < UBound(LL) Then
            R = 0
            what = Chr$(13)
            GoTo again
        End If
        End If
        .curpos = PX + realR
        .currow = PY
        GoTo jumpexit
    ElseIf processcr Then
        If ff < UBound(LL) Then
            R = 0
            what = Chr$(13)
            GoTo again
        End If
    End If
    
    .curpos = PX
    .currow = PY
    End With
jumpexit:

End Sub


Public Function nTextY(basestack As basetask, ByVal what As String, ByVal Font As String, ByVal Size As Single, Optional ByVal degree As Double = 0#, Optional ByVal ExtraWidth As Long = 0)
Dim ddd As Object
Set ddd = basestack.Owner
Dim PX As Long, PY As Long, OLDFONT As String, OLDSIZE As String, DE#
Dim F As LOGFONT, hPrevFont As Long, hFont As Long, fline$
Dim BFONT As String
Dim prive As Long
prive = GetCode(ddd)
ExtraWidth = ExtraWidth \ dv15
If ExtraWidth <> 0 Then
SetTextCharacterExtra ddd.Hdc, ExtraWidth
End If
On Error Resume Next
With players(prive)
BFONT = ddd.Font.Name
If Font <> "" Then
If Size = 0 Then Size = ddd.FontSize
StoreFont Font, Size, .charset
ddd.Font.charset = 0
ddd.FontSize = 9
ddd.FontName = .FontName
ddd.Font.charset = .charset
ddd.FontSize = Size
Else
Font = .FontName
End If

DE# = 0 '(degree) * 180# / Pi
   F.lfItalic = Abs(.italics)
F.lfWeight = Abs(.bold) * 800
  F.lfEscapement = CLng(10 * DE#)
  F.lfFaceName = Left$(Font, 30) + Chr$(0)
  F.lfCharSet = .charset
  F.lfQuality = 3 ' PROOF_QUALITY
  F.lfHeight = (Size * -20) / DYP

  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(ddd.Hdc, hFont)
 what$ = Replace$(what, vbCrLf, vbCr) + vbCr
Dim textmetrics As POINTAPI, Max, maxx As Long, sumy As Long
Do While what$ <> ""
If Left$(what$, 1) = vbCr Then
fline$ = vbNullString
what$ = Mid$(what$, 2)
Else
fline$ = GetStrUntil(vbCr, what$)
End If
If Len(what$) = 0 And Len(fline$) = 0 Then If sumy > 0 Then Exit Do
  
textmetrics.x = 0
textmetrics.y = 0
    If Len(fline$) = 0 Then
        fline$ = " "
        GetTextExtentPoint32 ddd.Hdc, StrPtr(fline$), Len(fline$), textmetrics
        textmetrics.x = 0
    Else
        GetTextExtentPoint32 ddd.Hdc, StrPtr(fline$), Len(fline$), textmetrics
    End If
sumy = sumy + textmetrics.y
If maxx < textmetrics.x Then maxx = textmetrics.x
Loop
nTextY = Int(Abs(maxx * dv15 * Sin(degree)) + Abs(sumy * dv15 * Cos(degree)))
  hFont = SelectObject(ddd.Hdc, hPrevFont)
  DeleteObject hFont
If ExtraWidth <> 0 Then SetTextCharacterExtra ddd.Hdc, 0
End With
PlaceBasket ddd, players(prive)


End Function
Public Function nText(basestack As basetask, ByVal what As String, ByVal Font As String, ByVal Size As Single, Optional ByVal degree As Double = 0#, Optional ByVal ExtraWidth As Long = 0)
Dim ddd As Object
Set ddd = basestack.Owner
Dim PX As Long, PY As Long, OLDFONT As String, OLDSIZE As String, DE#
Dim F As LOGFONT, hPrevFont As Long, hFont As Long, fline$
Dim BFONT As String
Dim prive As Long
prive = GetCode(ddd)
On Error Resume Next
With players(prive)
ExtraWidth = ExtraWidth \ dv15
If ExtraWidth <> 0 Then
SetTextCharacterExtra ddd.Hdc, ExtraWidth
End If
BFONT = ddd.Font.Name
If Font <> "" Then
If Size = 0 Then Size = ddd.FontSize
StoreFont Font, Size, .charset
ddd.Font.charset = 0
ddd.FontSize = 9
ddd.FontName = .FontName
ddd.Font.charset = .charset
ddd.FontSize = Size
Else
Font = .FontName
End If


DE# = 0 '(degree) * 180# / Pi
   F.lfItalic = Abs(.italics)
F.lfWeight = Abs(.bold) * 800
    F.lfEscapement = 0
  'F.lfEscapement = CLng(10 * DE#)
  F.lfFaceName = Left$(Font, 30) + Chr$(0)
  F.lfCharSet = .charset
  F.lfQuality = 3 ' NONANTIALIASED_QUALITY
  F.lfHeight = (Size * -20) / DYP

  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(ddd.Hdc, hFont)
  what$ = Replace$(what, vbCrLf, vbCr) + vbCr
Dim textmetrics As POINTAPI, Max, maxx As Long, sumy As Long
Do While what$ <> ""
If Left$(what$, 1) = vbCr Then
fline$ = vbNullString
what$ = Mid$(what$, 2)
Else
fline$ = GetStrUntil(vbCr, what$)
End If
If Len(what$) = 0 And Len(fline$) = 0 Then If sumy > 0 Then Exit Do
  
textmetrics.x = 0
textmetrics.y = 0
    If Len(fline$) = 0 Then
        fline$ = " "
        GetTextExtentPoint32 ddd.Hdc, StrPtr(fline$), Len(fline$), textmetrics
        textmetrics.x = 0
    Else
        GetTextExtentPoint32 ddd.Hdc, StrPtr(fline$), Len(fline$), textmetrics
    End If
sumy = sumy + textmetrics.y
If maxx < textmetrics.x Then maxx = textmetrics.x
Loop

nText = Int(Abs(maxx * dv15 * Cos(degree)) + Abs(sumy * dv15 * Sin(degree)))

  hFont = SelectObject(ddd.Hdc, hPrevFont)
  DeleteObject hFont
If ExtraWidth <> 0 Then SetTextCharacterExtra ddd.Hdc, 0

End With
PlaceBasket ddd, players(prive)


End Function
Public Sub fullPlain(dd As Object, mb As basket, ByVal wh$, ByVal wi, Optional fake As Boolean = False, Optional nocr As Boolean = False)
Dim whNoSpace$, Displ As Long, DisplLeft As Long, i As Long, whSpace$, INTD As Long, MinDispl As Long, some As Long
Dim st As Long, CROP As Long, curx As Long
st = DXP
MinDispl = (TextWidth(dd, "A") \ 2) \ st
If MinDispl <= 1 Then MinDispl = 3
MinDispl = st * MinDispl
INTD = TextWidth(dd, space$(MyTrimL3Len(wh$)))
dd.currentX = dd.currentX + INTD

wi = wi - INTD
wh$ = NLTrim2$(wh$)
If Len(wh$) > 1 And Right$(wh$, 1) = "_" Then
wh$ = Left$(wh$, Len(wh$) - 1) + "-"

End If
INTD = wi + dd.currentX
CROP = rinstr(wh$, Chr$(9))
If CROP > 0 Then
curx = dd.currentX
MyPrintNew dd, mb.uMineLineSpace, Left$(wh$, CROP), nocr, fake
wi = wi - (dd.currentX - curx)
wh$ = Mid$(wh$, CROP + 1)
End If
whNoSpace$ = Replace$(wh$, " ", "")
Dim magicratio As Double, whsp As Long, whl As Double

If whNoSpace$ = wh$ Then
If CROP > 0 Then
dd.currentX = dd.currentX + (wi - TextWidth(dd, whNoSpace))
End If
MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake

    'dd.Print wh$
Else



 If Len(whNoSpace$) > 0 Then
   whSpace$ = space$(Len(Trim$(wh$)) - Len(whNoSpace$))
   
        Displ = st * ((wi - (TextWidth(dd, whNoSpace))) \ (Len(whSpace)) \ st)
        some = (wi - TextWidth(dd, whNoSpace) - Len(whSpace) * Displ) \ st
        magicratio = some / Len(whNoSpace)
        whsp = Len(whSpace)
                whNoSpace$ = vbNullString
        For i = 1 To Len(wh$)
            If Mid$(wh$, i, 1) = " " Then
            whsp = whsp - 1
            
               If whNoSpace$ <> "" Then
               whl = Len(whNoSpace$) * magicratio + whl

                    MyPrintNew dd, mb.uMineLineSpace, whNoSpace$, , fake
                whNoSpace$ = vbNullString
                End If
                If some > 0 Then
                '
                some = some - whl
                dd.currentX = ((dd.currentX + Displ) \ st) * st + CLng(whl) * st
                whl = whl - CLng(whl)
                Else
             dd.currentX = ((dd.currentX + Displ) \ st) * st
              End If
           
            Else
                whNoSpace$ = whNoSpace$ & Mid$(wh$, i, 1)
            End If
        Next i

          whl = Len(whNoSpace$) * magicratio + whl
      dd.currentX = dd.currentX + CLng(whl) * st
          MyPrintNew dd, mb.uMineLineSpace, whNoSpace$, , fake
    Else

            MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake
    End If
End If
End Sub
Public Sub fullPlainWhere(dd As Object, mb As basket, ByVal wh$, ByVal wi As Long, whr As Long, Optional fake As Boolean = False, Optional nocr As Boolean = False)
Dim whNoSpace$, Displ As Long, DisplLeft As Long, i As Long, whSpace$, INTD As Long, MinDispl As Long
Dim stdisp As Long, ratio As Double

'If Left$(LTrim(wh$), 1) = Chr$(9) Then
'MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake
'Exit Sub
'End If
If TextWidth(dd, "W") = TextWidth(dd, " ") Then
MinDispl = TextWidth(dd, " ")
Else
MinDispl = (TextWidth(dd, "A") \ 2) \ DXP

If MinDispl <= 1 Then MinDispl = 3

MinDispl = DXP * MinDispl
End If


If whr = 3 Or whr = 0 Then INTD = TextWidth(dd, space$(MyTrimL3Len(wh$)))
dd.currentX = dd.currentX + INTD
wi = wi - INTD
If Len(wh$) > 1 And Right$(wh$, 1) = "_" Then
wh$ = Left$(wh$, Len(wh$) - 1) + "-"
ElseIf Len(wh$) > 1 And Right$(wh$, 1) = Chr$(9) Then
wh$ = Left$(wh$, Len(wh$) - 1)
End If
wh$ = NLTrim2$(wh$)
INTD = wi + dd.currentX
whNoSpace$ = Replace$(Replace$(wh$, Chr$(9), ""), " ", "")
If whr = 2 Then
    wh$ = Trim$(wh$)
    whNoSpace$ = Replace$(Replace$(wh$, Chr$(9), ""), " ", "")
    dd.currentX = dd.currentX + wi \ 2 - (TextWidth(dd, wh$) + (Len(wh$) - Len(Replace$(wh$, " ", ""))) * (MinDispl - TextWidth(dd, " "))) / 2
ElseIf whr = 1 Then
    dd.currentX = dd.currentX + wi - TextWidth(dd, wh$) - (Len(wh$) - Len(Replace$(wh$, " ", ""))) * (MinDispl - TextWidth(dd, " "))
Else

INTD = (wi - TextWidth(dd, whNoSpace)) * 0.2 + dd.currentX

End If

If whNoSpace$ = wh$ Then
 MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake
Else
 If Len(whNoSpace$) > 0 Then
   whSpace$ = space$(Len(Trim$(wh$)) - Len(whNoSpace$))
   INTD = TextWidth(dd, Trim$(wh$)) - TextWidth(dd, whNoSpace$) + dd.currentX
  ' INTD = TextWidth(dd, whSpace$) + dd.CurrentX
   
   wh$ = Trim$(wh$)
   Displ = MinDispl
   If Displ * Len(whSpace$) + TextWidth(dd, whNoSpace$) > wi Then
   Displ = (wi - TextWidth(dd, whNoSpace$)) / (Len(wh$))
   
   End If
     
    stdisp = dd.currentX
                whNoSpace$ = vbNullString
        For i = 1 To Len(wh$)
            If Mid$(wh$, i, 1) = " " Then
            whSpace$ = Mid$(whSpace$, 2)
               If whNoSpace$ <> "" Then
               
                 MyPrintNew dd, mb.uMineLineSpace, whNoSpace$, , fake
                whNoSpace$ = vbNullString
                
                End If
              dd.currentX = dd.currentX + Displ
 
            ElseIf Mid$(wh$, i, 1) = Chr$(9) Then
           
             whSpace$ = Mid$(whSpace$, 2)
                 MyPrintNew dd, mb.uMineLineSpace, whNoSpace$ + Chr$(9), , fake
                dd.currentX = TextWidth(dd, Left$(wh$, i)) + stdisp
                whNoSpace$ = vbNullString
               
            Else
                whNoSpace$ = whNoSpace$ & Mid$(wh$, i, 1)
            End If
        Next i

          MyPrintNew dd, mb.uMineLineSpace, whNoSpace$, Not nocr, fake
    Else
    
    MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake
    
    End If
End If
End Sub

Public Sub wPlain(ddd As Object, mb As basket, ByVal what As String, ByVal wi&, ByVal Hi&, Optional nocr As Boolean = False)
Dim PX As Long, PY As Long, ttt As Long, ruller&
Dim buf$, b$, npy As Long ', npx As long

With mb
PlaceBasket ddd, mb
tParam.iTabLength = .ReportTab
If what = vbNullString Then Exit Sub
PX = .curpos
PY = .currow
If .mx - PX < wi& Then wi& = .mx - PX
If .My - PY < Hi& Then Hi& = .My - PY
If wi& = 0 Or Hi& < 0 Then Exit Sub
npy = PY
ruller& = wi&
For ttt = 1 To Len(what)
    b$ = Mid$(what, ttt, 1)
   ' If nounder32(b$) Then
   
   If Not (b$ = vbCr Or b$ = vbLf) Then
    If TextWidth(ddd, buf$ & b$) <= (wi& * .Xt) Then
    buf$ = buf$ & b$
    End If
    ElseIf b$ = vbCr Then
    
    If nocr Then Exit For
    MyPrintNew ddd, mb.uMineLineSpace, buf$, Not nocr
    
    
    buf$ = vbNullString
    Hi& = Hi& - 1
    npy = npy + 1
    LCTbasket ddd, mb, npy, PX
    End If
    If Hi& < 0 Then Exit For
Next ttt
If Hi& >= 0 And buf$ <> "" Then MyPrintNew ddd, mb.uMineLineSpace, buf$, Not nocr
If Not nocr Then LCTbasket ddd, mb, PY, PX
End With
End Sub
Public Sub FeedFont2Stack(basestack As basetask, ok As Boolean)
Dim mS As New mStiva
If ok Then
mS.PushVal CDbl(ReturnBold)
mS.PushVal CDbl(ReturnItalic)
mS.PushVal CDbl(ReturnCharset)
mS.PushVal CDbl(ReturnSize)
mS.PushStr ReturnFontName
mS.PushVal CDbl(1)
Else
mS.PushVal CDbl(0)
End If
basestack.soros.MergeTop mS
End Sub
Public Sub nPlain(basestack As basetask, ByVal what As String, ByVal Font As String, ByVal Size As Single, Optional ByVal degree As Double = 0#, Optional ByVal JUSTIFY As Long = 0, Optional ByVal qual As Boolean = True, Optional ByVal ExtraWidth As Long = 0)
Dim ddd As Object
Set ddd = basestack.Owner
Dim PX As Long, PY As Long, OLDFONT As String, OLDSIZE As Long, DEGR As Double
Dim F As LOGFONT, hPrevFont As Long, hFont As Long, fline$, TT As Long
Dim BFONT As String

On Error Resume Next
BFONT = ddd.Font.Name
If ExtraWidth <> 0 Then
SetTextCharacterExtra ddd.Hdc, ExtraWidth
End If
Dim icx As Long, icy As Long, x As Long, y As Long, icH As Long

If JUSTIFY < 0 Then degree = 0
DEGR = (degree) * 180# / Pi
  F.lfItalic = Abs(basestack.myitalic)
  F.lfWeight = Abs(basestack.myBold) * 800
  F.lfEscapement = 0
  F.lfFaceName = Left$(Font, 30) + Chr$(0)
  F.lfCharSet = basestack.myCharSet
  If qual Then
    F.lfQuality = PROOF_QUALITY 'NONANTIALIASED_QUALITY '
  Else
    F.lfQuality = NONANTIALIASED_QUALITY
  End If
  F.lfHeight = (Size * -20) / DYP
  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(ddd.Hdc, hFont)
  icH = TextHeight(ddd, "fq")
  hFont = SelectObject(ddd.Hdc, hPrevFont)
  DeleteObject hFont
  F.lfItalic = Abs(basestack.myitalic)
  F.lfWeight = Abs(basestack.myBold) * 800
  F.lfEscapement = CLng(10 * DEGR)
  F.lfFaceName = Left$(Font, 30) + Chr$(0)
  F.lfCharSet = basestack.myCharSet
  If qual Then
  F.lfQuality = PROOF_QUALITY 'NONANTIALIASED_QUALITY '
  Else
  F.lfQuality = NONANTIALIASED_QUALITY
  End If
  F.lfHeight = (Size * -20) / DYP
  

  
    hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(ddd.Hdc, hFont)

TT = ExtraWidth \ 2
icy = CLng(Cos(degree) * icH)
icx = CLng(Sin(degree) * icH)

With players(GetCode(ddd))
If JUSTIFY < 0 Then
JUSTIFY = Abs(JUSTIFY) - 1
If JUSTIFY = 0 Then
y = .YGRAPH - icy
x = .XGRAPH - icx * 2
ElseIf JUSTIFY = 1 Then
y = .YGRAPH
x = .XGRAPH
Else
y = .YGRAPH - icy / 2
x = .XGRAPH - icx
End If
Else
y = .YGRAPH - icy
x = .XGRAPH - icx

End If
End With
If TT > 0 Then
x = x + (Cos(degree) * TT * dv15)
y = y - (Sin(degree) * TT * dv15)
End If
what$ = Replace$(what, vbCrLf, vbCr) + vbCr
Dim textmetrics As POINTAPI

Do While what$ <> ""
If Left$(what$, 1) = vbCr Then
fline$ = vbNullString
what$ = Mid$(what$, 2)
Else
fline$ = GetStrUntil(vbCr, what$)
End If
textmetrics.x = 0
textmetrics.y = 0
GetTextExtentPoint32 ddd.Hdc, StrPtr(fline$), Len(fline$), textmetrics
x = x + icx
y = y + icy
If JUSTIFY = 1 Then
    ddd.currentX = x - Int((textmetrics.x * Cos(degree) + textmetrics.y * Sin(degree)) * dv15)
    ddd.currentY = y + Int((textmetrics.x * Sin(degree) - textmetrics.y * Cos(degree)) * dv15)
ElseIf JUSTIFY = 2 Then
'If tt <> 0 Then textmetrics.X = textmetrics.X - tt * 1.5

     ddd.currentX = x - Int((textmetrics.x * Cos(degree) + textmetrics.y * Sin(degree)) * dv15) \ 2
    ddd.currentY = y + Int((textmetrics.x * Sin(degree) - textmetrics.y * Cos(degree)) * dv15) \ 2
Else

    ddd.currentX = x
    ddd.currentY = y
End If
MyPrint ddd, fline$
Loop
  hFont = SelectObject(ddd.Hdc, hPrevFont)
  DeleteObject hFont
If ExtraWidth <> 0 Then SetTextCharacterExtra ddd.Hdc, 0
End Sub

Public Sub nForm(bstack As basetask, TheSize As Single, nW As Long, nH As Long, myLineSpace As Long)
    On Error Resume Next
    StoreFont bstack.Owner.Font.Name, TheSize, bstack.myCharSet
    nH = fonttest.TextHeight("Wq") + myLineSpace * 2
    nW = fonttest.TextWidth("W") + dv15
End Sub
Sub crNew(bstack As basetask, mb As basket)

Dim D As Object
Set D = bstack.Owner
With mb
Dim PX As Long, PY As Long, R As Long
PX = .curpos
PY = .currow
PX = 0
PY = PY + 1
If TypeOf D Is MetaDc Then

ElseIf PY >= .My Then
    If Not bstack.toprinter Then
        ScrollUpNew D, mb
        PY = .My - 1
    Else
        PY = 0
        PX = 0
        getnextpage
    End If
End If
.curpos = PX
.currow = PY

End With
End Sub

Public Sub CdESK()
Dim x, y, ff As Form, useform1 As Boolean
If Form1.Visible Then
    If Form5.Visible Then
    Set ff = Form5
    Form5.RestoreSizePos
    Form5.backcolor = 0
    useform1 = True
    Else
    Set ff = Form1
    End If
    x = ff.Left / DXP
    y = ff.top / DYP
    If useform1 Then Form1.Visible = False
    ff.Hide
    Sleep 50
    MyDoEvents1 ff, True
    
    Dim aa As New cDIBSection
    aa.CreateFromPicture hDCToPicture(GetDC(0), x, y, ff.Width / DXP, ff.Height / DYP)
    aa.ThumbnailPaint ff
    GdiFlush
      ff.Visible = True
      
    If useform1 Then Form1.Visible = True
    
End If
Set ff = Nothing
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub

Public Sub ScrollUpNew(D As Object, mb As basket)
If TypeOf D Is MetaDc Then Exit Sub
Dim ar As RECT, R As Long
Dim p As Long
With mb
ar.Left = 0
ar.Bottom = D.Height / dv15
ar.Right = D.Width / dv15
ar.top = .mysplit * .Yt / dv15
p = .Yt / dv15
R = BitBlt(D.Hdc, CLng(ar.Left), CLng(ar.top), CLng(ar.Right), CLng(ar.Bottom - p), D.Hdc, CLng(ar.Left), CLng(ar.top + p), SRCCOPY)

 
  ar.top = ar.Bottom - p
FillBack D.Hdc, ar, .Paper
.curpos = 0
.currow = .My - 1
End With
GdiFlush
End Sub
Public Sub ScrollDownNew(D As Object, mb As basket)
If TypeOf D Is MetaDc Then Exit Sub
Dim ar As RECT, R As Long
Dim p As Long
With mb
ar.Left = 0
ar.Bottom = D.ScaleY(D.Height, 1, 3)
ar.Right = D.ScaleX(D.Width, 1, 3)
ar.top = D.ScaleY(.mysplit * .Yt, 1, 3)
p = D.ScaleY(.Yt, 1, 3)
R = BitBlt(D.Hdc, CLng(ar.Left), CLng(ar.top + p), CLng(ar.Right), CLng(ar.Bottom - p), D.Hdc, CLng(ar.Left), CLng(ar.top), SRCCOPY)
D.Line (0, .mysplit * .Yt)-(D.Scalewidth, .mysplit * .Yt + .Yt), .Paper, BF
.currow = .mysplit
.curpos = 0
End With
End Sub

Public Sub SetText(dq As Object, Optional alinespace As Long = -1, Optional ResetColumns As Boolean = False)
' can be used for first time also
Dim mymul As Long
On Error Resume Next
With players(GetCode(dq))
If .FontName = vbNullString Or alinespace = -2 Then
' we have to make it
If alinespace = -2 Then alinespace = 0
ResetColumns = True
.FontName = dq.FontName
.charset = dq.Font.charset
.SZ = dq.FontSize
Else
If Not (fonttest.FontName = .FontName And fonttest.Font.charset = dq.Font.charset And fonttest.Font.Size = .SZ) Then
fonttest.Font.charset = .charset
If fonttest.Font.charset = .charset Then
StoreFont .FontName, .SZ, .charset
dq.Font.charset = 0
dq.FontSize = 9
dq.FontName = .FontName
dq.Font.charset = .charset
dq.FontSize = .SZ
End If
End If
End If
If alinespace <> -1 Then
If .uMineLineSpace = .MineLineSpace * 2 And .MineLineSpace <> 0 Then
.MineLineSpace = alinespace
.uMineLineSpace = alinespace * 2
Else
.MineLineSpace = alinespace
.uMineLineSpace = alinespace ' so now we have normal
End If
End If
.SZ = dq.FontSize
.Xt = fonttest.TextWidth("W") + dv15
.Yt = fonttest.TextHeight("fj")
.overrideTextHeight = .Yt
.mx = Int(dq.Width / .Xt)
.My = Int(dq.Height / (.Yt + .uMineLineSpace * 2))
''.Paper = dq.BackColor
If .My <= 0 Then .My = 1
If .mx <= 0 Then .mx = 1
.Yt = .Yt + .uMineLineSpace * 2
If ResetColumns Then
mymul = Int(.mx / 8)
If mymul = 1 Then mymul = 2
If mymul = 0 Then
.Column = .mx \ 2 - 1
Else
.Column = Int(.mx / mymul)
While (.mx Mod .Column) > 0 And (.mx / .Column >= 3)
.Column = .Column + 1
Wend
End If
If .Column = 0 Then .Column = .mx
.Column = .Column - 1
If .Column < 4 Then .Column = 4
End If
.MAXXGRAPH = dq.Width
.MAXYGRAPH = dq.Height
End With

End Sub

Public Sub SetTextSZ(dq As Object, mSz As Single, Optional factor As Single = 1, Optional AddTwipsTop As Long = -1)
' Used for making specific basket
On Error Resume Next
With players(GetCode(dq))
If AddTwipsTop < 0 Then
    If .double And factor = 1 Then
    .mysplit = .osplit
    .Column = .OCOLUMN
    .currow = (.currow + 1) * 2 - 2
    .curpos = .curpos * 2
    mSz = .SZ / 2
    .uMineLineSpace = .MineLineSpace
    .double = False
    ElseIf factor = 2 And Not .double Then
     .osplit = .mysplit
     .OCOLUMN = .Column
     .Column = .Column / 2
     .mysplit = .mysplit / 2
     .currow = (.currow + 1) / 2
     .curpos = .curpos / 2
     mSz = .SZ * 2
    .uMineLineSpace = .MineLineSpace * 2
    .double = True
    End If
Else

mSz = mSz * factor
.MineLineSpace = AddTwipsTop
.uMineLineSpace = AddTwipsTop * factor
.double = factor <> 1
End If
dq.FontSize = mSz

StoreFont dq.Font.Name, mSz, dq.Font.charset
If .double Then
    Dim nowtextheight As Long
    nowtextheight = fonttest.TextHeight("fj")
    If .MineLineSpace = 0 Then
    Else
    If (.Yt - .MineLineSpace * 2) * 2 <> nowtextheight Then
    .uMineLineSpace = Int((.MAXYGRAPH - nowtextheight * .My / 2) / .My)
    End If
    
    End If
End If
SetText dq



If .My <= 0 Then .My = 1
If .mx <= 0 Then .mx = 1
.SZ = dq.FontSize
.MAXXGRAPH = dq.Width
.MAXYGRAPH = dq.Height
End With

End Sub

Public Sub SetTextBasketBack(dq As Object, mb As basket)
' set minimum display parameters for current object
' need an already filled basket
On Error Resume Next
With mb

If Not (dq.FontName = .FontName And dq.Font.charset = .charset And dq.Font.Size = .SZ) Then

StoreFont .FontName, .SZ, .charset
dq.Font.charset = 0
dq.FontSize = 9
dq.FontName = .FontName
dq.Font.charset = .charset
dq.FontSize = .SZ
End If
dq.forecolor = .mypen And &HFFFFFF

If Not dq.backcolor = .Paper Then
   dq.backcolor = .Paper
End If
End With
End Sub

Function gf$(bstack As basetask, ByVal y&, ByVal x&, ByVal a$, c&, F&, Optional STAR As Boolean = False)
On Error Resume Next
Dim cLast&, b$, cc$, dq As Object, ownLinespace, oldrefresh As Double, noinp As Double
oldrefresh = REFRESHRATE
Dim mybasket As basket, addpixels As Long
GFQRY = True
Set dq = bstack.Owner
SetText dq
mybasket = players(GetCode(dq))

With mybasket
If InternalLeadingSpace() = 0 And .MineLineSpace = 0 Then
addpixels = 0
Else
addpixels = 2
End If
If dq.Visible = False Then dq.Visible = True
If exWnd = 0 Then dq.SetFocus
dq.FontTransparent = False
LCTbasket dq, mybasket, y&, x&
Dim o$
o$ = a$
If a$ = vbNullString Then a$ = " "
INK$ = vbNullString

Dim xx&
xx& = x&

x& = x& - 1

cLast& = Len(a$)
'*****************
If cLast& + x& >= .mx Then
MyDoEvents
If dq.Font.charset = 161 Then
b$ = InputBoxN("Εισαγωγή Τιμής Πεδίου", MesTitle$, a$, noinp)
Else
b$ = InputBoxN("Input Field Value", MesTitle$, a$, noinp)
End If
If noinp <> 1 Then b$ = a$
If MyTrim(b$) < "A" Then b$ = Right$(String$(cLast&, " ") + b$, cLast&) Else b$ = Left$(b$ + String$(cLast&, " "), cLast&)
gf$ = b$
If xx& < .mx Then
dq.FontTransparent = False
If STAR Then
PlainBaSket dq, mybasket, StarSTR(Left$(b$, .mx - x&)), True, , addpixels
Else
PlainBaSket dq, mybasket, Left$(b$, .mx - x&), True, , addpixels
End If
End If
GoTo GFEND
Else
dq.FontTransparent = False
If STAR Then
PlainBaSket dq, mybasket, StarSTR(a$), True, , addpixels
Else
PlainBaSket dq, mybasket, a$, True, , addpixels
End If
End If

'************
b$ = a$
.currow = y&
.curpos = c& + x&
LCTCB dq, mybasket, ins&

Do
MyDoEvents1 Form1, , True
If bstack.IamThread Then If myexit(bstack) Then GoTo contgfhere
If Not TaskMaster Is Nothing Then
If TaskMaster.QueueCount > 0 Then
dq.FontTransparent = True
TaskMaster.RestEnd1
TaskMasterTick
End If
End If
 cc$ = INKEY$
 If cc$ <> "" Then
If Not TaskMaster Is Nothing Then TaskMaster.rest
SetTextBasketBack dq, mybasket
 Else
If Not TaskMaster Is Nothing Then TaskMaster.RestEnd
SetTextBasketBack dq, mybasket
        If iamactive Then
           If Screen.ActiveForm Is Nothing Then
                            DestroyCaret
                      nomoveLCTC dq, mybasket, y&, c& + x&, ins&
                      iamactive = False
           Else
                If Not (GetForegroundWindow = Screen.ActiveForm.hWnd And Screen.ActiveForm.Name = "Form1") Then
                 
                      DestroyCaret
                      nomoveLCTC dq, mybasket, y&, c& + x&, ins&
                      iamactive = False
             Else
                         If ShowCaret(dq.hWnd) = 0 Then
                                   HideCaret dq.hWnd
                                   .currow = y&
                                   .curpos = c& + x&
                                   LCTCB dq, mybasket, ins&
                                   ShowCaret dq.hWnd
                         End If
                End If
                End If
     Else
  If Not Screen.ActiveForm Is Nothing Then
            If GetForegroundWindow = Screen.ActiveForm.hWnd And Screen.ActiveForm.Name = "Form1" Then
           
                          nomoveLCTC dq, mybasket, y&, c& + x&, ins&
                             iamactive = True
                              If ShowCaret(dq.hWnd) = 0 And Screen.ActiveForm.Name = "Form1" Then
                                   HideCaret dq.hWnd
                                   .currow = y&
                                   .curpos = c& + x&
                                   LCTCB dq, mybasket, ins&
                                   ShowCaret dq.hWnd
                         End If
                         End If
            End If
     End If

 End If

 
        If NOEXECUTION Then
        If KeyPressed(&H1B) Then
                       F& = 99 'ESC  ****************
                        c& = 1
                        gf$ = o$
                        b$ = o$
                                          NOEXECUTION = False
                                         BLOCKkey = True
                                    While KeyPressed(&H1B)
                                    If Not TaskMaster Is Nothing Then
                             If TaskMaster.Processing Then
                                                TaskMaster.RestEnd1
                                                TaskMaster.TimerTick
                                                TaskMaster.rest
                                                MyDoEvents1 dq
                                                Else
                                                MyDoEvents
                                                
                                                End If
                                                Else
                                                DoEvents
                                                End If
'''sleepwait 1
                                    Wend
                                                                        BLOCKkey = False
                                                                        End If
                 Exit Do
        End If
        Select Case Len(cc$)
        Case 0
        If Fkey > 0 Then
        If FK$(Fkey) <> "" And Fkey <> 13 Then
            cc$ = FK$(Fkey)
            interpret Basestack1, cc$
        
        End If
        Fkey = 0
        Else
        
        End If
        
        Case 1
        If STAR And cc$ = " " Then cc$ = Chr$(127)
                Select Case AscW(cc$)
                Case 8
                        If c& > 1 Then
                        Mid$(b$, c& - 1) = Mid$(b$, c&) & " "
                         c& = c& - 1
                         dq.FontTransparent = False
                                   .currow = y&
                                   .curpos = c& + x&
                                   LCTCB dq, mybasket, ins&
                        If STAR Then
                        PlainBaSket dq, mybasket, StarSTR(Mid$(b$, c&)), True, , addpixels
                        Else
                        PlainBaSket dq, mybasket, Mid$(b$, c&), True, , addpixels
                        End If
                         dq.Refresh
                                   .currow = y&
                                   .curpos = c& + x&
                                   LCTCB dq, mybasket, ins&
                        End If
                Case 6
                F& = -1
                 gf$ = b$
                Exit Do
                Case 13, 9
                F& = 1 'NEXT  *************
                gf$ = b$
                Exit Do

                Case 27
                        F& = 99 'ESC  ****************
                        c& = 1
                        gf$ = o$
                        b$ = o$
                                    NOEXECUTION = False
                                    BLOCKkey = True
                                    While KeyPressed(&H1B)
                                    If Not TaskMaster Is Nothing Then
                                    If TaskMaster.Processing Then
                                            TaskMaster.RestEnd1
                                            TaskMaster.TimerTick
                                            TaskMaster.rest
                                            MyDoEvents1 dq
                                            Else
                                            MyDoEvents
                                            
                                            End If
                                            Else
                                            DoEvents
                                            End If
                                    '''
                                    Wend
                                                                        BLOCKkey = False
                        NOEXECUTION = False
                        Exit Do
                       Case 32 To 126, Is > 128
           
                        .currow = y&
                        .curpos = c& + x&
                        LCTCB dq, mybasket, ins&
                        If ins& = 1 Then
                          If AscW(cc$) = 32 And STAR Then
                If AscW(Mid$(b$, c& + 1)) > 32 Then
                 Mid$(b$, c&) = Mid$(b$, c& + 1) & " "
                End If
                
                
                Else
                        
                                                
                        Mid$(b$, c&, 1) = cc$
                        dq.FontTransparent = False
                        If STAR Then
                        PlainBaSket dq, mybasket, StarSTR(Mid$(b$, c&)), True, , addpixels
                        Else
                        PlainBaSket dq, mybasket, Mid$(b$, c&), True, , addpixels
                        End If
                         dq.Refresh
                        End If
                        If c& < Len(b$) Then c& = c& + 1
                                   .currow = y&
                                   .curpos = c& + x&
                                   LCTCB dq, mybasket, ins&
                        Else
                                 If AscW(cc$) = 32 And STAR Then
            
                
                
                Else
                     
                        LSet b$ = Left$(b$, c& - 1) + cc$ & Mid$(b$, c&)
                        dq.FontTransparent = False
                        If STAR Then
                        PlainBaSket dq, mybasket, StarSTR(Mid$(b$, c&)), True, , addpixels
                        Else
                        PlainBaSket dq, mybasket, Mid$(b$, c&), True, , addpixels
                        End If
                         dq.Refresh
                        'LCTC Dq, Y&, X& + C& + 1, INS&
                        End If
                        If c& < cLast& Then c& = c& + 1
                                .currow = y&
                                .curpos = c& + x&
                                LCTCB dq, mybasket, ins&
                        End If
                End Select
        Case 2
                Select Case AscW(Right$(cc$, 1))
                Case 81
                F& = 10 ' exit - pagedown ***************
                gf$ = b$
                Exit Do
                Case 73
                F& = -10 ' exit - pageup
                gf$ = b$
                Exit Do
                Case 79
                F& = 20 ' End
                gf$ = b$
                Exit Do
                Case 71
                F& = -20 ' exit - home
                gf$ = b$
                Exit Do
                Case 75 'LEFT
                        If c& > 1 Then
                                   .currow = y&
                                .curpos = c& + x&
                                LCTCB dq, mybasket, ins&
                        c& = c& - 1:
                        .currow = y&
                        .curpos = c& + x&
                        LCTCB dq, mybasket, ins&
                        End If
                Case 77 'RIGHT
                        If c& < cLast& Then
                      
                If Not (AscW(Mid$(b$, c&)) = 32 And STAR) Then
                
             
                                    .currow = y&
                                .curpos = c& + x&
                                LCTCB dq, mybasket, ins&
                        c& = c& + 1:
                        .currow = y&
                                .curpos = c& + x&
                                LCTCB dq, mybasket, ins&
                        End If
                        End If
                Case 72 ' EXIT UP
                F& = -1 ' PREVIUS ***************
                gf$ = b$
                Exit Do
                Case 80 'EXIT DOWN OR ENTER OR TAB
                F& = 1 'NEXT  *************
                gf$ = b$
                Exit Do
                Case 82
                            .currow = y&
                                .curpos = c& + x&
                                LCTCB dq, mybasket, ins&
                ins& = 1 - ins&
                           .currow = y&
                                .curpos = c& + x&
                                LCTCB dq, mybasket, ins&
                Case 83
                        Mid$(b$, c&) = Mid$(b$, c& + 1) & " "
                        dq.FontTransparent = False
                        LCTbasket dq, mybasket, y&, c& + x&
                        If STAR Then
                        PlainBaSket dq, mybasket, StarSTR(Mid$(b$, c&)), True, , addpixels
                        Else
                        PlainBaSket dq, mybasket, Mid$(b$, c&), True, , addpixels
                        End If
                               .currow = y&
                                .curpos = c& + x&
                                LCTCB dq, mybasket, ins&
                     dq.Refresh
                End Select
        End Select
      If GFQRY = False Then Exit Do
Loop

GFEND:
REFRESHRATE = oldrefresh
LCTbasket dq, mybasket, y&, x& + 1
If x& < .mx And Not xx& > .mx Then
If STAR Then
 PlainBaSket dq, mybasket, StarSTR(b$), True, , addpixels
Else
PlainBaSket dq, mybasket, b$, True, , addpixels
End If
contgfhere:
 dq.Refresh
If Not TaskMaster Is Nothing Then If TaskMaster.QueueCount > 0 Then TaskMaster.RestEnd
End If
dq.FontTransparent = True
 DestroyCaret
Set dq = Nothing
TaskMaster.RestEnd1
GFQRY = False
End With
End Function

Function StarSTR(ByVal sStr As String) As String
Dim l As Long, s As Long
l = Len(sStr)
sStr = RTrim$(sStr)
s = l - Len(sStr)
StarSTR = String$(l - s, "*") + String$(s, " ")

End Function
Function ProcBeep(bstack As basetask, rest$) As Boolean
Dim p As Variant
If IsExp(bstack, rest$, p) Then
MessageBeep CLng(p)
Else
Beep
End If
ProcBeep = True
End Function
Public Sub ResetPrefresh()
Dim i As Long
For i = -2 To 131
    Prefresh(i).k1 = 0
    Prefresh(i).RRCOUNTER = 0
Next i

End Sub

Sub original(bstack As basetask, COM$)
Dim D As Object, b$

If Len(COM$) > 0 Then QUERYLIST = vbNullString
If Form1.Visible Then REFRESHRATE = 25: ResetPrefresh
If bstack.toprinter Then
bstack.toprinter = False
Form1.PrinterDocument1.Cls
Set D = bstack.Owner
Else
Set D = bstack.Owner
End If
On Error Resume Next
Dim basketcode As Long
basketcode = GetCode(D)


Form1.IEUP ""
Form1.KeyPreview = True
Dim dummy As Boolean, rs As String, mPen As Long, ICO As Long, BAR As Long, bar2 As Long
BAR = 1
Form1.DIS.Visible = True
GDILines = False  ' reset to normal ' use Smooth on to change this to true
If COM$ <> "" Then D.Visible = False
ClrSprites
mPen = PenOne
D.Font.bold = bstack.myBold
D.Font.Italic = bstack.myitalic
GetMonitorsNow
Console = FindFormSScreen(Form1)
'' Console = FindPrimary
With ScrInfo(Console)
If SzOne < 4 Then SzOne = 4
    'Form1.Visible = False
   ' If IsWine Then
    Sleep 30
    .Width = .Width - 1
    .Height = .Height - 1
   ' End If
    If Not Form1.WindowState = 0 Then Form1.WindowState = 0
    Sleep 10
    If Form1.WindowState = 0 Then
        Form1.move .Left, .top, .Width - 1, .Height - 1
        If Form1.top <> .Left Or Form1.Left <> .top Then
            Form1.Cls
            Form1.move .Left, .top, .Width - 1, .Height - 1
        End If
    Else
        Sleep 100
        On Error Resume Next
        Form1.WindowState = 0
        Form1.move .Left, .top, .Width - 1, .Height - 1
        If Form1.top <> .top Or Form1.Left <> .Left Then
        Form1.Cls
        Form1.move .Left, .top, .Width - 1, .Height - 1
        End If
    End If
NoBackFormFirstUse = False
If players(-1).MAXXGRAPH <> 0 Then ClearScrNew Form1, players(-1), 0&
Form1.DIS.Visible = True
FrameText D, SzOne, (.Width + .Left - 1 - Form1.Left), (.Height + .top - 1 - Form1.top), PaperOne
End With
Form1.DIS.backcolor = mycolor(PaperOne)
If lckfrm = 0 Then
SetText D
bstack.Owner.Font.charset = bstack.myCharSet
StoreFont bstack.Owner.Font.Name, SzOne, bstack.myCharSet
 
 With players(basketcode)
.mypen = PenOne
.XGRAPH = 0
.YGRAPH = 0
.bold = bstack.myBold '' I have to change that
.italics = bstack.myitalic
.FontName = bstack.Owner.FontName
.SZ = SzOne
.charset = bstack.myCharSet
.MAXXGRAPH = Form1.Width
.MAXYGRAPH = Form1.Height
.Paper = bstack.Owner.backcolor
.mypen = mycolor(PenOne)
End With


 
' check to see if
Dim ss$, skipthat As Boolean
If Not IsSupervisor Then
    ss$ = ReadUnicodeOrANSI(userfiles & "desktop.inf")
    LastErNum = 0
    If ss$ <> "" Then
     skipthat = interpret(bstack, ss$)
     If mycolor(PenOne) <> D.forecolor Then
     PenOne = -D.forecolor
     End If
    End If
End If
If SzOne < 36 And D.Height / SzOne > 250 Then SetDouble D: BAR = BAR + 1
If SzOne < 83 Then

If bstack.myCharSet = 161 Then
b$ = "ΠΕΡΙΒΑΛΛΟΝ "
Else
b$ = "ENVIRONMENT "
End If
D.forecolor = mycolor(PenOne)
LCTbasket D, players(DisForm), 0, 0
wwPlain2 bstack, players(DisForm), b$ & "M2000", D.Width, 0, 0 '',True
ICO = TextWidth(D, b$ & "M2000") + 100
' draw graphic'
Dim iX As Long, iY As Long
With players(DisForm)
iX = (.Xt \ 25) * 25
iY = Form1.icon.Height * iX / Form1.icon.Width
If IsWine Then
Form1.DIS.PaintPicture Form1.icon, ICO, (.Yt - iY) / 2, iX, iY
Form1.DIS.PaintPicture Form1.icon, ICO, (.Yt - iY) / 2, iX, iY
Else
Dim myico As New cDIBSection
myico.backcolor = Form1.DIS.backcolor
myico.CreateFromPicture Form1.icon
Form1.DIS.PaintPicture myico.Picture(1), ICO, (.Yt - iY) / 2, iX, iY
End If
End With

' ********
SetNormal D
   Dim osbit As String
   If Is64bit Then osbit = " (64-bit)" Else osbit = " (32-bit)"
        LCTbasket D, players(basketcode), BAR, 0
        rs = vbNullString
            If bstack.myCharSet = 161 Then
            If Revision = 0 Then
            wwPlain2 bstack, players(DisForm), "Έκδοση Διερμηνευτή: " & CStr(VerMajor) & "." & CStr(VerMinor), D.Width, 0, True
            Else
                    wwPlain2 bstack, players(DisForm), "Έκδοση Διερμηνευτή: " & CStr(VerMajor) & "." & Left$(CStr(VerMinor), 1) & " (" & CStr(Revision) & ")", D.Width, 0, True
                End If
                   wwPlain2 bstack, players(DisForm), "Λειτουργικό Σύστημα: " & os & osbit, D.Width, 0, True
            
                      wwPlain2 bstack, players(DisForm), "Όνομα Χρήστη: " & Tcase(Originalusername), D.Width, 0, True
                
            Else
             If Revision = 0 Then
              wwPlain2 bstack, players(DisForm), "Interpreter Version: " & CStr(VerMajor) & "." & CStr(VerMinor), D.Width, 0, True
             Else
                    wwPlain2 bstack, players(DisForm), "Interpreter Version: " & CStr(VerMajor) & "." & Left$(CStr(VerMinor), 1) & " rev. (" & CStr(Revision) & ")", D.Width, 0, True
                 End If
              
                      wwPlain2 bstack, players(DisForm), "Operating System: " & os & osbit, D.Width, 0, True
                
                   wwPlain2 bstack, players(DisForm), "User Name: " & Tcase(Originalusername), D.Width, 0, True
        
                 End If
                        '    cr bstack
            GetXYb D, players(basketcode), bar2, BAR
             players(basketcode).curpos = bar2
            players(basketcode).currow = BAR
           BAR = BAR + 1
            If BAR >= players(basketcode).My Then ScrollUpNew D, players(basketcode)
                    LCTbasket D, players(basketcode), BAR, 0
                    players(basketcode).curpos = 0
            players(basketcode).currow = BAR
    End If
If Not ASKINUSE Then Load NeoMsgBox
If Not skipthat Then

ProcPen bstack, CStr(mPen) + ", 255"
ProcCls bstack, "," + CStr(BAR)
' dummy = interpret(bstack, "PEN " & CStr(mPen) & ":CLS ," & CStr(BAR))
End If
If Not ASKINUSE Then Unload NeoMsgBox
End If
If Not skipthat Then
    If Len(COM$) > 0 Then dummy = interpret(bstack, COM$)
End If
'cr bstack
End Sub
Sub ClearScr(D As Object, c1 As Long)
Dim aa As Long
With players(GetCode(D))
.Paper = c1
.curpos = 0
.currow = 0
.lastprint = False
End With
If Not TypeOf D Is MetaDc Then
D.Line (0, 0)-(D.Scalewidth - dv15, D.Scaleheight - dv15), c1, BF
End If
D.currentX = 0
D.currentY = 0

End Sub
Sub ClearScrNew(D As Object, mb As basket, c1 As Long)
Dim im As New StdPicture, spl As Long
If TypeOf D Is MetaDc Then
D.backcolor = c1
Exit Sub
End If
With mb
spl = .mysplit * .Yt
Set im = D.Image
.Paper = c1

If TypeOf D Is GuiM2000 Then
If .mysplit = 0 Then
    If Not D.backcolor = c1 Then D.backcolor = c1
    D.Cls
Else
    D.Line (0, spl)-(D.Scalewidth - dv15, D.Scaleheight - dv15), .Paper, BF
    End If
    .currow = .mysplit
ElseIf D.Name = "Form1" Or mb.used Then
D.Line (0, spl)-(D.Scalewidth - dv15, D.Scaleheight - dv15), .Paper, BF
.curpos = 0
.currow = .mysplit
Else
D.backcolor = c1
If spl > 0 Then D.PaintPicture im, 0, 0, D.Width, spl, 0, 0, D.Width, spl, vbSrcCopy
.curpos = 0
.currow = .mysplit

End If
.lastprint = False
D.currentX = 0
D.currentY = 0
End With
End Sub
Function iText(bb As basetask, ByVal v$, wi&, Hi&, aTitle$, n As Long, Optional NumberOnly As Boolean = False, Optional UseIntOnly As Boolean = False, Optional curset = -1&) As String
Dim x&, y&, dd As Object, wh&, shiftlittle As Long, OLDV$
Set dd = bb.Owner
With players(GetCode(dd))
If .lastprint Then
x& = (dd.currentX + .Xt - dv15) \ .Xt
y& = dd.currentY \ .Yt
shiftlittle = x& * .Xt - dd.currentX
If y& > .mx Then
y& = .mx - 1
crNew bb, players(GetCode(dd))

End If
Else
x& = .curpos
y& = .currow
End If
If .mx - x& - 1 < wi& Then wi& = .mx - x&
If .My - y& - 1 < Hi& Then Hi& = .My - y& - 1
If wi& = 0 Or Hi& < 0 Then
iText = v$
Exit Function
End If
wi& = wi& + x&
Hi& = Hi& + y&
Form1.EditTextWord = True
wh& = CLng(curset)
Dim oldshow As Boolean
With Form1.TEXT1
     oldshow = .showparagraph
    .showparagraph = False
    
    If n <= 0 Then .Title = aTitle$ + " ": If wh& = -1& Then wh& = Abs(n - 1)
    If NumberOnly Then
     .glistN.UseTab = False
        .NumberOnly = True
        .NumberIntOnly = UseIntOnly
        OLDV$ = v$
        ScreenEdit bb, v$, x&, y&, wi& - 1, Hi&, wh&, , n, shiftlittle
        If Result = 99 Then v$ = OLDV$
        .NumberIntOnly = False
        .NumberOnly = False
    Else
    .glistN.UseTab = True
        OLDV$ = v$
        ScreenEdit bb, v$, x&, y&, wi& - 1, Hi&, wh&, , n, shiftlittle
        If Result = 99 And Hi& = wi& Then v$ = OLDV$
    End If
    .showparagraph = oldshow
    .glistN.UseTab = UseTabInForm1Text1
End With
iText = v$
End With
End Function
Sub ScreenEditDOC(bstack As basetask, aaa As Variant, x&, y&, x1&, y1&, Optional l As Long = 0, Optional usecol As Boolean = False, Optional Col As Long)
On Error Resume Next
Dim ot As Boolean, back As New Document, i As Long, D As Object
Dim prive As basket
Set D = bstack.Owner
prive = players(GetCode(D))
With prive
Dim oldesc As Boolean
oldesc = escok
escok = False
' we have a limit here
If Not aaa.IsEmpty Then
For i = 1 To aaa.DocParagraphs
back.AppendParagraph aaa.TextParagraph(i)
Next i
End If
i = back.LastSelStart
Dim Aaaa As Document, tcol As Long, trans As Boolean
If usecol Then tcol = mycolor(Col) Else tcol = D.backcolor
If Not Form1.Visible Then newshow Basestack1

'd.Enabled = False
If Not bstack.toback Then D.TabStop = False
If D Is Form1 Then
D.lockme = True
Else
D.Parent.lockme = True
End If
If y1& - y& = 0 Then y& = y& - 1: If y1& < 0 Then y& = y& + 1: y1& = y1& + 1
TextEditLineHeight = y1& - y& + 1

With Form1.TEXT1

ProcTask2 bstack
.glistN.UseTab = True
Hook Form1.hWnd, Nothing '.glistN
.AutoNumber = Not Form1.EditTextWord

.UsedAsTextBox = False
.glistN.LeftMarginPixels = 10
.glistN.maxchar = 0
If D.forecolor = tcol Then
Set Form1.Point2Me = D
If D.Name = "Form1" Then
.glistN.SkipForm = False
Else
.glistN.SkipForm = True
End If
Form1.TEXT1.glistN.BackStyle = 1
End If
Dim scope As Long
scope = ChooseByHue(D.forecolor, rgb(16, 12, 8), rgb(253, 245, 232))
If D.backcolor = ChooseByHue(scope, D.backcolor, rgb(128, 128, 128)) Then
If lightconv(scope) > 192 Then
scope = lightconv(scope) - 128
.glistN.CapColor = rgb(128 + scope / 2, 128 + scope / 2, 128 + scope / 2)
Else
.glistN.CapColor = scope
End If
Else
scope = lightconv(scope) - 128

If scope > 0 Then
.glistN.CapColor = rgb(128 + scope / 2, 128 + scope / 2, 128 + scope / 2)
Else
.glistN.CapColor = rgb(128, 128, 128)
End If
End If
.SelectionColor = .glistN.CapColor
.glistN.addpixels = 2 * prive.uMineLineSpace / dv15
.EditDoc = True
.enabled = True
.glistN.ZOrder 0

.backcolor = tcol

.forecolor = D.forecolor
Form1.SetText1
.glistN.overrideTextHeight = prive.overrideTextHeight '         fonttest.TextHeight("fj")
.Font.Name = D.Font.Name
.Font.Size = D.Font.Size ' SZ 'Int(d.font.Size) Why
.Font.charset = D.Font.charset
.Font.Italic = D.Font.Italic
.Font.bold = D.Font.bold
.Font.Name = D.Font.Name
.Font.charset = D.Font.charset
.Font.Size = prive.SZ
With prive
If bstack.toback Then

Form1.TEXT1.move x& * .Xt, y& * .Yt, (x1& - x&) * .Xt + .Xt, (y1& - y&) * .Yt + .Yt
Else
Form1.TEXT1.move x& * .Xt + D.Left, y& * .Yt + D.top, (x1& - x&) * .Xt + .Xt, (y1& - y&) * .Yt + .Yt
End If
End With
If D.forecolor = tcol Then
If D.Name = "Form1" Then
Form1.TEXT1.glistN.RepaintFromOut D.Image, D.Left, D.top
Else
Form1.TEXT1.glistN.RepaintFromOut D.Image, 0, 0
End If

End If

Set .mDoc = aaa
.mDoc.ColorEvent = True
.nowrap = False


With Form1.TEXT1
.Form1mn1Enabled = False
.Form1mn2Enabled = False
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
End With

Form1.KeyPreview = False
NOEDIT = False

.WrapAll
.Render

.Visible = True
.SetFocus
If l <> 0 Then
    If l > 0 Then
        If aaa.SizeCRLF < l Then l = aaa.SizeCRLF
        
        .SelStart = l
        Else
        .SelStart = 0
    End If
Else
If aaa.SizeCRLF < .LastSelStart Then
.SelStart = 1
Else
 .SelStart = .LastSelStart
End If
End If
    .ResetUndoRedo

End With
'
ProcTask2 bstack
CancelEDIT = False
Do
BLOCKkey = False

 If bstack.IamThread Then If myexit(bstack) Then GoTo contScreenEditThere1

ProcTask2 bstack


'End If

Loop Until NOEDIT
 NOEXECUTION = False
 BLOCKkey = True
While KeyPressed(&H1B)
ProcTask2 bstack

Wend
BLOCKkey = False
contScreenEditThere1:
TaskMaster.RestEnd1
If Form1.TEXT1.Visible Then Form1.TEXT1.Visible = False
 l = Form1.TEXT1.LastSelStart


If D Is Form1 Then
D.lockme = False
Else
D.Parent.lockme = False
End If
If Not CancelEDIT Then

Else
Set aaa = back
back.LastSelStart = i
End If
Set Form1.TEXT1.mDoc = New Document
Form1.TEXT1.glistN.UseTab = UseTabInForm1Text1
Form1.TEXT1.glistN.BackStyle = 0
Set Form1.Point2Me = Nothing
UnHook Form1.hWnd
Form1.KeyPreview = True

INK$ = vbNullString
escok = oldesc
Set D = Nothing
End With
End Sub
Sub ScreenEdit(bstack As basetask, a$, x&, y&, x1&, y1&, Optional l As Long = 0, Optional changelinefeeds As Long = 0, Optional maxchar As Long = 0, Optional ExcludeThisLeft As Long = 0, Optional internal As Boolean = False)
On Error Resume Next
' allways a$ enter with crlf,but exit with crlf or cr or lf depents from changelinefeeds
Dim oldesc As Boolean, D As Object
Set D = bstack.Owner

''SetTextSZ d, Sz

Dim prive As basket
prive = players(GetCode(D))
oldesc = escok
escok = False
Dim ot As Boolean

If Not bstack.toback Then
D.TabStop = False
D.Parent.lockme = True
Else
D.lockme = True
End If
If Not Form1.Visible Then newshow Basestack1
D.Visible = True
If D.Visible Then D.SetFocus
With Form1.TEXT1

ProcTask2 bstack
Hook Form1.hWnd, Nothing
'.Filename = VbNullString
.AutoNumber = Not Form1.EditTextWord

If maxchar > 0 Then
ot = .glistN.DragEnabled
 .glistN.DragEnabled = True
y1& = y&
TextEditLineHeight = 1
.glistN.BorderStyle = 0
.glistN.BackStyle = 1
Set Form1.Point2Me = D
If D.Name = "Form1" Then
.glistN.SkipForm = False
Else
.glistN.SkipForm = True
End If

.glistN.HeadLine = vbNullString
.glistN.HeadLine = vbNullString
.glistN.LeftMarginPixels = 1
.glistN.maxchar = maxchar
.nowrap = True
If Len(a$) > maxchar Then
a$ = Left$(a$, maxchar)
End If
If l = -1 Then
l = Len(a$) + 1
Else
l = 1
End If

.UsedAsTextBox = True

Else
.glistN.BorderStyle = 0
.glistN.BackStyle = 0

If y1& - y& = 0 Then y& = y& - 1: If y1& < 0 Then y& = y& + 1: y1& = y1& + 1
TextEditLineHeight = y1& - y& + 1
.UsedAsTextBox = False
.glistN.LeftMarginPixels = 10
.glistN.maxchar = 0

End If

If Form1.EditTextWord Then
.glistN.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", """", " ", "+", "-", "/", "*", "^", "$", "%", "_", "@")
.glistN.WordCharRight = ConCat(".", ":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", """", " ", "+", "-", "/", "*", "^", "$", "%", "_")
.glistN.WordCharRightButIncluded = vbNullString
.glistN.WordCharLeftButIncluded = vbNullString

Else
.glistN.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", "@", Chr$(9), "#", "%", "&")
.glistN.WordCharRight = ConCat(":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", Chr$(9), "#")
.glistN.WordCharRightButIncluded = "(" ' so aaa(sdd) give aaa( as word
.glistN.WordCharLeftButIncluded = "#"
End If

Dim scope As Long
scope = ChooseByHue(D.forecolor, rgb(16, 12, 8), rgb(253, 245, 232))
If D.backcolor = ChooseByHue(scope, D.backcolor, rgb(128, 128, 128)) Then
If lightconv(scope) > 192 Then
scope = lightconv(scope) - 128
.glistN.CapColor = rgb(128 + scope / 2, 128 + scope / 2, 128 + scope / 2)
Else
.glistN.CapColor = scope
End If
Else
scope = lightconv(scope) - 128

If scope > 0 Then
.glistN.CapColor = rgb(128 + scope / 2, 128 + scope / 2, 128 + scope / 2)
Else
.glistN.CapColor = rgb(128, 128, 128)
End If
End If
.SelectionColor = .glistN.CapColor
.glistN.addpixels = 2 * prive.uMineLineSpace / dv15
.enabled = False
.EditDoc = True
.enabled = True
'.glistN.AddPixels = 0
.glistN.ZOrder 0
.backcolor = D.backcolor
.forecolor = D.forecolor
.Font.Name = D.Font.Name
Form1.SetText1
.glistN.overrideTextHeight = prive.overrideTextHeight '   fonttest.TextHeight("fj")
.Font.Size = D.Font.Size ' SZ 'Int(d.font.Size) Why
.Font.charset = D.Font.charset
.Font.Italic = D.Font.Italic
.Font.bold = D.Font.bold

.Font.Name = D.Font.Name

.Font.charset = D.Font.charset
.Font.Size = prive.SZ 'Int(d.font.Size)
If bstack.toback Then
If maxchar > 0 Then

.move x& * prive.Xt - ExcludeThisLeft, y& * prive.Yt, (x1& - x&) * prive.Xt + prive.Xt, (y1& - y&) * prive.Yt + prive.Yt
If D.Name = "Form1" Then
.glistN.RepaintFromOut D.Image, D.Left, D.top
Else
.glistN.RepaintFromOut D.Image, 0, 0
End If
Else
.move x& * prive.Xt, y& * prive.Yt, (x1& - x&) * prive.Xt + prive.Xt, (y1& - y&) * prive.Yt + prive.Yt
End If
Else
If maxchar > 0 Then
.move x& * prive.Xt + D.Left - ExcludeThisLeft, y& * prive.Yt + D.top, (x1& - x&) * prive.Xt + prive.Xt, (y1& - y&) * prive.Yt + prive.Yt
If D.Name = "Form1" Then
.glistN.RepaintFromOut D.Image, D.Left, D.top
Else
.glistN.RepaintFromOut D.Image, 0, 0
End If
Else
.move x& * prive.Xt + D.Left, y& * prive.Yt + D.top, (x1& - x&) * prive.Xt + prive.Xt, (y1& - y&) * prive.Yt + prive.Yt
End If
End If
If a$ <> "" Then
If .Text <> a$ Then .LastSelStart = 0
If internal Then
.Text2 = a$
Else
.Text = a$
End If
Else
.Text = vbNullString
.LastSelStart = 0
End If
'.glistN.NoFreeMoveUpDown = True

'With Form1.TEXT1
.Form1mn1Enabled = False
.Form1mn2Enabled = False
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
'End With

Form1.KeyPreview = False

NOEDIT = False

If maxchar = 0 Then
If .nowrap Then
.nowrap = False
End If

.Charpos = 1
If Len(a$) < 100000 Then .Render
Else
.Render
End If
If l <> 0 Then
    If l > 0 Then
        If Len(a$) < l Then l = Len(a$) Else l = l - 1
       
        .SelStart = l
                Else
        .SelStart = 0
    End If
Else
If Len(a$) < .LastSelStart Then
.SelStart = 1
l = Len(a$)
Else
    .SelStart = .LastSelStart
End If
End If
.Visible = True
'
ProcTask2 bstack
.SetFocus


    .ResetUndoRedo



End With

ProcTask2 bstack
CancelEDIT = False
Dim Timeout As Long


Do
BLOCKkey = False

 If bstack.IamThread Then If myexit(bstack) Then GoTo contScreenEditThere

ProcTask2 bstack

 Loop Until NOEDIT
 NOEXECUTION = False
 BLOCKkey = True
While KeyPressed(&H1B)
'
ProcTask2 bstack


Wend
BLOCKkey = False
contScreenEditThere:
TaskMaster.RestEnd1
If Form1.TEXT1.Visible Then Form1.TEXT1.Visible = False

 l = Form1.TEXT1.LastSelStart + 1

If bstack.toback Then
D.lockme = False
Else
D.Parent.lockme = False
End If
If Not CancelEDIT Then

If changelinefeeds > 10 Then
a$ = Form1.TEXT1.TextFormatBreak(vbCr)
ElseIf changelinefeeds > 9 Then
a$ = Form1.TEXT1.TextFormatBreak(vbLf)
Else
If changelinefeeds = -1 Then changelinefeeds = 0
a$ = Form1.TEXT1.Text
End If
Else
changelinefeeds = -1
End If

Form1.KeyPreview = True
If maxchar > 0 Then Form1.TEXT1.glistN.DragEnabled = ot

UnHook Form1.hWnd
INK$ = vbNullString
Form1.TEXT1.glistN.UseTab = False
escok = oldesc
Set D = Nothing
End Sub

Function blockCheck(ByVal s$, ByVal Lang As Long, countlines As Long, Optional ByVal sbname$ = vbNullString, Optional Column As Long) As Boolean
If s$ = vbNullString Then blockCheck = True: Exit Function
Dim i As Long, j As Long, c As Long, b$, resp&
Dim openpar As Long, oldi As Long, lastlabel$, oldjump As Boolean, st As Long, stc As Long
Dim paren As New mStiva2
countlines = 1
Column = 0
Lang = Not Lang
Dim a1 As Boolean
Dim jump As Boolean
If Trim$(s$) = vbNullString Then Exit Function
c = Len(s$)
a1 = True
i = 1
Do
Column = Column + 1
Select Case AscW(Mid$(s$, i, 1))
Case 10
Column = 0
Case 13
lastlabel$ = ""
If openpar <> 0 Then
GoTo pareprob
End If
oldjump = False
jump = False
If Len(s$) > i + 1 Then countlines = countlines + 1
Column = 0
Case 58
lastlabel$ = ""
oldjump = False
jump = False
Case 32, 160, 9
If Len(lastlabel$) > 0 Then
lastlabel$ = myUcase(lastlabel$)
If Not ismine1(lastlabel$) Then
If Not ismine2(lastlabel$) Then
If Not ismine22(lastlabel$) Then
    jump = Not oldjump
Else
    oldjump = True
    jump = False
End If
Else
oldjump = True
jump = False
End If
Else
oldjump = False
jump = False
End If
lastlabel$ = ""
End If
Case 34
lastlabel$ = ""
oldi = i
Do While i < c
i = i + 1
Select Case AscW(Mid$(s$, i, 1))
Case 34
Exit Do
Case 13

checkit:
    If Not Lang Then
        b$ = sbname$ + "Problem in string in paragraph " + CStr(countlines)
    Else
        b$ = sbname$ + "Πρόβλημα με το αλφαριθμητικό στη παράγραφο " + CStr(countlines)
    End If
    resp& = ask(b$, True)
If resp& <> 4 Then
blockCheck = True
End If
Exit Function
End Select

Loop
If oldi <> i Then
Else
i = oldi + 1
GoTo checkit
End If

Case 40
lastlabel$ = ""
jump = True
openpar = openpar + 1
paren.PushVal countlines
Case 41
lastlabel$ = ""
openpar = openpar - 1
If openpar = 0 Then jump = False
If openpar < 0 Then Exit Do
paren.drop 1
Case 47
If Mid$(s$, i + 1, 1) = "/" Then i = i + 1: GoTo a1111
Case 39, 92
a1111:
lastlabel$ = ""
Do While i < c
i = i + 1
If Mid$(s$, i, 2) = vbCrLf Then Exit Do
Loop
countlines = countlines + 1
If openpar > 0 Then Exit Do
Case 61, 43, 44
lastlabel$ = ""
jump = True
Case 123
If Len(lastlabel$) > 0 Then
lastlabel$ = myUcase(lastlabel$)
If Not ismine1(lastlabel$) Then
If Not ismine2(lastlabel$) Then
If Not ismine22(lastlabel$) Then
    jump = Not oldjump
Else
    oldjump = True
    jump = False
End If
Else
oldjump = True
jump = False
End If
Else
oldjump = False
jump = False
End If
lastlabel$ = ""
End If

If jump Then
jump = False
' we have a multiline text
Dim target As Long
target = j
st = countlines
stc = Column
    Do
    Select Case AscW(Mid$(s$, i, 1))
            Case 34
            Do While i < c
            i = i + 1
            If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
            If AscW(Mid$(s$, i, 1)) = 13 Then GoTo checkit
            Loop
        Case 13
        countlines = countlines + 1
        Case 123
        j = j - 1
        Case 125
        j = j + 1: If j = target Then Exit Do
    End Select
    i = i + 1
    Loop Until i > c
    If j <> target Then
    countlines = st
    Column = st
    Exit Do
    End If
    Else
j = j - 1
oldjump = False
End If
Case 13

Case 125
If openpar <> 0 And j > 0 Then
pareprob:
If paren.count > 0 Then countlines = paren.PopVal
If Not Lang Then
        b$ = sbname$ + "Problem in parenthesis in paragraph" + Str$(countlines)
    Else
        b$ = sbname$ + "Πρόβλημα με τις παρενθέσεις στη παράγραφο" + Str$(countlines)
    End If
    resp& = ask(b$, True)
If resp& <> 4 Then
blockCheck = True
End If
    Exit Function

End If
j = j + 1: If j = 1 Then Exit Do
Case 65 To 93, 97 To 122, Is > 127
jump = False
lastlabel$ = lastlabel$ + Mid$(s$, i, 1)
Case 46
jump = False
lastlabel$ = lastlabel$ + Mid$(s$, i, 1)

Case 48 To 57, 95
jump = False
If Len(lastlabel$) > 0 Then lastlabel$ = lastlabel$ + Mid$(s$, i, 1)
Case Else
jump = False
lastlabel$ = ""
End Select
i = i + 1
Loop Until i > c
If openpar <> 0 Then
GoTo pareprob
End If
If j = 0 Then

ElseIf j < 0 Then
    If Not Lang Then
        b$ = sbname$ + "Problem in blocks - look } are less " + CStr(Abs(j))
    Else
        b$ = sbname$ + "Πρόβλημα με τα τμήματα - δες τα } είναι λιγότερα " + CStr(Abs(j))
    End If
resp& = ask(b$, True)
Else
If Not Lang Then
b$ = sbname$ + "Problem in blocks - look { are less " + CStr(j)
Else
b$ = sbname$ + "Πρόβλημα με τα τμήματα - δες τα { είναι λιγότερα " + CStr(j)
End If
resp& = ask(b$, True)
End If
If resp& <> 4 Then
blockCheck = True
End If

End Function

Sub ListChoise(bstack As basetask, a$, x&, y&, x1&, y1&)
On Error Resume Next
Dim D As Object, oldh As Long
Dim s$, prive As basket
If NOEXECUTION Then Exit Sub
Set D = bstack.Owner
prive = players(GetCode(D))
Dim ot As Boolean, drop
With Form1.List1
.Font.Name = D.Font.Name
Form1.Font.charset = D.Font.charset
Form1.Font.Strikethrough = False
.Font.Size = D.Font.Size
.Font.Name = D.Font.Name
Form1.Font.charset = D.Font.charset
.Font.Size = D.Font.Size
If LEVCOLMENU < 2 Then .backcolor = D.forecolor
If LEVCOLMENU < 3 Then .forecolor = D.backcolor
.Font.bold = D.Font.bold
.Font.Italic = D.Font.Italic
.addpixels = 2 * prive.uMineLineSpace / dv15
.VerticalCenterText = True
If D.Visible = False Then D.Visible = True
.StickBar = True
s$ = .HeadLine
.HeadLine = vbNullString
.HeadLine = s$
.enabled = False
If .Visible Then
If .BorderStyle = 0 Then

Else
End If

Else

If .BorderStyle = 0 Then
.move x& * prive.Xt + D.Left, y& * prive.Yt + D.top, (x1& - x&) * prive.Xt + prive.Xt, (y1& - y&) * prive.Yt + prive.Yt + .HeadlineHeight * dv15
Else
.move x& * prive.Xt - dv15 + D.Left, y& * prive.Yt - dv15 + D.top, (x1& - x&) * prive.Xt + prive.Xt + 2 * dv15, (y1& - y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15
End If
End If
.enabled = True
.ShowBar = False

If .LeaveonChoose Then
.CalcAndShowBar
Exit Sub
End If



ot = Targets
Targets = False

.PanPos = 0

If .ListIndex < 0 Then
.ShowThis 1
Else
.ShowThis .ListIndex + 1
End If
.Visible = True
.ZOrder 0
NOEDIT = False
.Tag = a$

If a$ = vbNullString Then
    drop = mouse
    MyDoEvents
    ' Form1.KeyPreview = False
    .enabled = True
    .SetFocus
    .LeaveonChoose = True
    If .HeadLine <> "" Then
    oldh = 0
    Else
    oldh = .HeadlineHeight
    End If
    Else
        .enabled = True
    .SetFocus
    .LeaveonChoose = False
    
    End If
    .ShowMe
            If bstack.TaskMain Or TaskMaster.Processing Then
            If TaskMaster.QueueCount > 0 Then
            mywait bstack, 100
              Else
            MyDoEvents
            End If
        Else
         DoEvents
         Sleep 1
         End If

    If .HeadlineHeight <> oldh Then
    If .BorderStyle = 0 Then
    If ((y1& - y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15) + .top > ScrY() Then
    .move .Left, .top - (((y1& - y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15) + .top - ScrY()), (x1& - x&) * prive.Xt + prive.Xt, (y1& - y&) * prive.Yt + prive.Yt + .HeadlineHeight * dv15
    Else
.move .Left, .top, (x1& - x&) * prive.Xt + prive.Xt, (y1& - y&) * prive.Yt + prive.Yt + .HeadlineHeight * dv15
End If
Else
If ((y1& - y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15) + .top > ScrY() Then
.move .Left, .top - (((y1& - y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15) + .top - ScrY()), (x1& - x&) * prive.Xt + prive.Xt + 2 * dv15, (y1& - y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15
Else
.move .Left, .top, (x1& - x&) * prive.Xt + prive.Xt + 2 * dv15, (y1& - y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15
End If
End If
  
oldh = .HeadlineHeight
    End If
    .FloatLimitTop = Form1.Height - prive.Yt * 2
     .FloatLimitLeft = Form1.Width - prive.Xt * 2
    MyDoEvents
    End With
If a$ = vbNullString Then
    Do
        If bstack.TaskMain Or TaskMaster.Processing Then
            If TaskMaster.QueueCount > 0 Then
          mywait bstack, 2
             TaskMaster.RestEnd1
   TaskMaster.TimerTick
TaskMaster.rest
''SleepWait 1
  Sleep 1
              Else
            MyDoEvents
            End If
        Else
         DoEvents
                  Sleep 1
         End If
    
    Loop Until Form1.List1.Visible = False
    If Not NOEXECUTION Then MOUT = False
    Do
    drop = mouse
    MyDoEvents
    Loop Until drop = 0 Or MOUT
    MOUT = False
    While KeyPressed(&H1B)
ProcTask2 bstack
Wend
MOUT = False: NOEXECUTION = False
    If Form1.List1.ListIndex >= 0 Then
    a$ = Form1.List1.list(Form1.List1.ListIndex)
    Else
    a$ = vbNullString
    End If
   Form1.List1.enabled = False
    Else
        Form1.List1.enabled = True
    
  If a$ = vbNullString Then
  Form1.List1.SetFocus
  Form1.List1.LeaveonChoose = True
  Else
  D.TabStop = True
  End If
  End If
NOEDIT = True
Set D = Nothing
Form1.KeyPreview = True
Targets = ot
End Sub
Private Sub mywait11(bstack As basetask, pp As Double)
Dim p As Boolean, e As Boolean
On Error Resume Next
If bstack.Process Is Nothing Then
''If extreme Then MyDoEvents
If pp = 0 Then Exit Sub
Else

Err.Clear
p = bstack.Process.Done
If Err.Number = 0 Then
e = True
If p <> 0 Then
Exit Sub
End If
End If
End If
pp = pp + CDbl(timeGetTime)

Do


If TaskMaster.Processing And Not bstack.TaskMain Then
        If Not bstack.toprinter Then bstack.Owner.Refresh
        'If TaskMaster.tickdrop > 0 Then TaskMaster.tickdrop
        TaskMaster.TimerTick  'Now
       ' SleepWait 1
       MyDoEvents
       
Else
        ' SleepWait 1
        MyDoEvents
        End If
If e Then
p = bstack.Process.Done
If Err.Number = 0 Then
If p <> 0 Then
Exit Do
End If
End If
End If
Loop Until pp <= CDbl(timeGetTime) Or NOEXECUTION

                       If exWnd <> 0 Then
                MyTitle$ bstack
                End If
End Sub
Public Sub WaitDialog(bstack As basetask)
Dim oldesc As Boolean
oldesc = escok
escok = False
Dim D As Object
Set D = bstack.Owner
Dim ot As Boolean, drop
ot = Targets
Targets = False  ' do not use targets for now
'NOEDIT = False
    drop = mouse
    ''SleepWait3 100
    Sleep 1
    If bstack.ThreadsNumber = 0 Then
    If Not (bstack.toback Or bstack.toprinter) Then If bstack.Owner.Visible Then bstack.Owner.Refresh

    End If
    Dim mycode As Double, oldcodeid As Double
mycode = Rnd * 1233312231
oldcodeid = Modalid
Dim x As Form, zz As Form
Set zz = Screen.ActiveForm
For Each x In Forms
        If x.Visible And x.Name = "GuiM2000" Then
                                   If x.Enablecontrol Then
                                        x.Modal = mycode
                                        x.Enablecontrol = False
                                    End If
        End If
Next x
On Error Resume Next
If zz.enabled Then zz.SetFocus
Set zz = Nothing
      Do
   

            mywait11 bstack, 5
      Sleep 1
    
    Loop Until loadfileiamloaded = False Or LastErNum <> 0
    Modalid = mycode
    MOUT = False
    Do
    drop = mouse Or KeyPressed(&H1B)
    MyDoEvents

    Loop Until drop = 0 Or MOUT Or LastErNum <> 0
 ' NOEDIT = True
 BLOCKkey = True

While KeyPressed(&H1B)

ProcTask2 bstack
NOEXECUTION = False
Wend
Dim z As Form
Set z = Nothing

           For Each x In Forms
            If x.Visible And x.Name = "GuiM2000" Then
                If Not x.Enablecontrol Then
                        x.TestModal mycode
                End If
            End If
            Next x
          Modalid = oldcodeid

BLOCKkey = False
escok = oldesc
INK$ = vbNullString
If Form1.Visible Then Form1.KeyPreview = Not Form1.gList1.Visible
Targets = ot
 mywait11 bstack, 5
End Sub

Public Sub FrameText(dd As Object, ByVal Size As Single, x As Long, y As Long, cc As Long, Optional myCut As Boolean = False)
Dim i As Long, mymul As Long

If dd Is Form1.PrinterDocument1 Then
' check this please
dd.Width = x
dd.Height = y
Pr_Back dd, Size
Exit Sub
End If


Dim basketcode As Long
basketcode = GetCode(dd)
With players(basketcode)
.curpos = 0
.currow = 0
.XGRAPH = 0
.YGRAPH = 0
If x = 0 Then
x = dd.Width
y = dd.Height
End If

.mysplit = 0

''dd.BackColor = 0 '' mycolor(cc)    ' check if paper...

.Paper = mycolor(cc)
dd.currentX = 0
dd.currentY = 0

''ClearScreenNew dd, mybasket, cc
dd.currentY = 0
dd.Font.Size = Size
Size = dd.Font.Size

''Sleep 1  '' USED TO GIVE TIME TO LOAD FONT
If fonttest.FontName = dd.Font.Name And dd.Font.Size = fonttest.Font.Size Then
Else
StoreFont dd.Font.Name, Size, dd.Font.charset
End If
.Yt = fonttest.TextHeight("fj")
.Xt = fonttest.TextWidth("W")

While TextHeight(fonttest, "fj") / (.Yt / 2 + dv15) < dv
Size = Size + 0.2
fonttest.Font.Size = Size
Wend
dd.Font.Size = Size
.overrideTextHeight = fonttest.TextHeight("fj")
.Yt = TextHeight(fonttest, "fj")
.Xt = fonttest.TextWidth("W") + dv15

.mx = Int(x / .Xt)
.My = Int(y / (.Yt + .MineLineSpace * 2))
.Yt = .Yt + .MineLineSpace * 2
If .mx < 2 Then .mx = 2: x = 2 * .Xt
If .My < 2 Then .My = 2: y = 2 * .Yt
If (.mx Mod 2) = 1 And .mx > 1 Then
.mx = .mx - 1
End If
mymul = Int(.mx / 8)
If mymul = 1 Then mymul = 2
If mymul = 0 Then
.Column = .mx \ 2 - 1
Else
.Column = Int(.mx / mymul)

While (.mx Mod .Column) > 0 And (.mx / .Column >= 3)
.Column = .Column + 1
Wend
End If
If .Column = 0 Then .Column = .mx
' second stage
If .mx Mod .Column > 0 Then


If .mx Mod 4 <> 0 Then .mx = 4 * (.mx \ 4)
If .mx < 4 Then .mx = 4
'.My = Int(y / (.Yt + .MineLineSpace * 2))
'.Yt = .Yt + .MineLineSpace * 2
If .mx < 2 Then .mx = 2: x = 2 * .Xt
If .My < 2 Then .My = 2: y = 2 * .Yt
If (.mx Mod 2) = 1 And .mx > 1 Then
.mx = .mx - 1
End If
mymul = Int(.mx / 8)
If mymul = 1 Then mymul = 2
If mymul = 0 Then
.Column = .mx \ 2 - 1
Else
.Column = Int(.mx / mymul)

While (.mx Mod .Column) > 0 And (.mx / .Column >= 3)
.Column = .Column + 1
Wend
End If
If .Column = 0 Then .Column = .mx

End If

.Column = .Column - 1 ' FOR PRINT 0 TO COLUMN-1

If .Column < 4 Then .Column = 4


.SZ = Size

If dd.Name = "Form1" Then
' no change
Else
If dd.Name <> "dSprite" And Typename(dd) <> "GuiM2000" And Not TypeOf dd Is MetaDc Then
Dim mmxx As Long, mmyy As Long, xx As Long, YY As Long
mmxx = .mx * CLng(.Xt)
mmyy = .My * CLng(.Yt)

xx = (dd.Parent.Scalewidth - mmxx) \ 2
YY = (dd.Parent.Scaleheight - mmyy) \ 2
dd.move xx, YY, mmxx, mmyy
ElseIf myCut And Not TypeOf dd Is MetaDc Then
Dim mmxx1, mmyy1
mmxx1 = .mx * .Xt
mmyy1 = .My * .Yt
dd.move dd.Left, dd.top, mmxx1, mmyy1
End If

End If

.MAXXGRAPH = dd.Width
.MAXYGRAPH = dd.Height
.FTEXT = 0
.FTXT = vbNullString

Form1.MY_BACK.ClearUp
If dd.Visible Then
ClearScr dd, .Paper
Else
dd.backcolor = .Paper
End If
End With



End Sub

Sub Pr_Back(dd As Object, Optional msize As Single = 0)

SetText dd
If msize <> 0 Then players(GetCode(dd)).SZ = msize
If msize > 0 Then
SetTextSZ dd, msize
End If

End Sub
Function INKEY$()
Dim w As Long
If MKEY$ <> "" Then
    INK$ = MKEY$ & INK$
    MKEY$ = vbNullString
End If
If INK$ <> "" Then
' ειδική περίπτωση αν έχουμε 0 στο πρώτο Byte, έχουμε ειδικό κ
        w = AscW(INK$)
    If w = 0 Then
        INKEY$ = Left$(INK$, 2)
        INK$ = Mid$(INK$, 3)
    ElseIf w > -10241 And w < -9984 Then
    INKEY$ = Left$(INK$, 2)
    INK$ = Mid$(INK$, 3)
    Else
    ' αλλιώς σηκώνουμε ένα χαρακτήρα με ότι έχει ακόμα
    INKEY$ = PopOne(INK$)
    
   
        
    End If
Else
    'Αν δεν έχουμε τίποτα...δεν κάνουμε τίποτα...γυρίζουμε το τίποτα!
    INKEY$ = vbNullString
End If

End Function
Function UINKEY$()
Dim w As Long
' mink$ used for reinput keystrokes
' MINK$ = MINK$ & UINK$
If UKEY$ <> "" Then MINK$ = MINK$ + UKEY$: UKEY$ = vbNullString
If MINK$ <> "" Then
w = AscW(MINK$)
If w = 0 Then
    UINKEY$ = Left$(MINK$, 2)
    MINK$ = Mid$(MINK$, 3)
ElseIf w > -10241 And w < -9984 Then
    UINKEY$ = Left$(MINK$, 2)
    MINK$ = Mid$(MINK$, 3)
Else
    UINKEY$ = Left$(MINK$, 1)
    MINK$ = Mid$(MINK$, 2)
End If
Else
    UINKEY$ = vbNullString
End If

End Function

Function QUERY(bstack As basetask, Prompt$, s$, m&, Optional USELIST As Boolean = True, Optional endchars As String = vbCr, Optional excludechars As String = vbNullString, Optional checknumber As Boolean = False) As String
'NoAction = True
On Error Resume Next
Dim dX As Long, dY As Long, safe$, oldREFRESHRATE As Double, AUX As Long
oldREFRESHRATE = REFRESHRATE

If excludechars = vbNullString Then excludechars = Chr$(0)
If QUERYLIST = vbNullString Then QUERYLIST = Chr$(13): LASTQUERYLIST = 1
Dim q1 As Long, sp$, once As Boolean, dq As Object
 
Set dq = bstack.Owner
SetText dq
Dim basketcode As Long, prive As basket
prive = players(GetCode(dq))
With prive
If .currow >= .My Or .lastprint Then crNew bstack, prive: .lastprint = False
LCTbasketCur dq, prive
ins& = 0
Dim fr1 As Long, fr2 As Long, p As Double
UseEnter = False
If dq.Name = "DIS" Then
If Form1.Visible = False Then
    If Not Form3.Visible Then
        Form1.Hide: Sleep 100
    Else
        'Form3.PREPARE
    End If

    If Form1.WindowState = vbMinimized Then Form1.WindowState = vbNormal
    Form1.Show , Form5
    If ttl Then
    If Form3.Visible Then
    If Not Form3.WindowState = 0 Then
    Form3.skiptimer = True: Form3.WindowState = 0
    End If
    End If
    End If
    MyDoEvents
    Sleep 100
    End If
Else
    Console = FindFormSScreen(Form1)
If Form1.top >= VirtualScreenHeight() Then Form1.move ScrInfo(Console).Left, ScrInfo(Console).top
End If
If dq.Visible = False Then dq.Visible = True
If exWnd = 0 Then Form1.KeyPreview = True
QRY = True
If GetForegroundWindow = Form1.hWnd Then
If exWnd = 0 Then dq.SetFocus
End If


Dim DE$

PlainBaSket dq, prive, Prompt$, , , 0
dq.Refresh

 

INK$ = vbNullString
dq.FontTransparent = False

Dim a$
s$ = vbNullString
oldLCTCB dq, prive, 0
Do
If Not once Then
If USELIST Then
 DoEvents
  If Not iamactive Then
  Sleep 1
  Else
  If Not (bstack.IamChild Or bstack.IamAnEvent) Then Sleep 1
  End If
 ''If MKEY$ = VbNullString Then Dq.refresh
Else
If Not bstack.IamThread Then

 If Not iamactive Then
 If Not Form1.Visible Then
 If Form1.WindowState = 1 Then Form1.WindowState = 0
 If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top
 Form1.Visible = True
 If Form3.Visible Then Form3.skiptimer = True: Form3.WindowState = 0
 End If
 k1 = 0: MyDoEvents1 Form1, , True
 End If
If LastErNum <> 0 Then
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0
Exit Do
End If
 Else
 
LCTbasketCur dq, prive                       ' here
 End If
 End If
 End If
If Not QRY Then HideCaret dq.hWnd:   Exit Do

 BLOCKkey = False
 If USELIST Then

 If Not once Then
 once = True

 If QUERYLIST <> "" Then  ' up down
 
    If INK = vbNullString Then MyDoEvents
If clickMe = 38 Then

 If Len(QUERYLIST) < LASTQUERYLIST Then LASTQUERYLIST = 2
  q1 = InStr(LASTQUERYLIST, QUERYLIST, vbCr)
         If q1 < 2 Or q1 <= LASTQUERYLIST Then
         q1 = 1: LASTQUERYLIST = 1
         End If
        MKEY$ = vbNullString
        INK = String$(Len(s$), 8) + Mid$(QUERYLIST, LASTQUERYLIST, q1 - LASTQUERYLIST)
        LASTQUERYLIST = q1 + 1

    ElseIf clickMe = 40 Then
    
    If LASTQUERYLIST < 3 Then LASTQUERYLIST = Len(QUERYLIST)
    q1 = InStrRev(QUERYLIST, vbCr, LASTQUERYLIST - 2)
         If q1 < 2 Then
                   q1 = Len(QUERYLIST)
         End If
         If q1 > 1 Then
         LASTQUERYLIST = InStrRev(QUERYLIST, vbCr, q1 - 1) + 1
         If LASTQUERYLIST < 2 Then LASTQUERYLIST = 2
         
        MKEY$ = vbNullString
        INK = String$(RealLen(s$), 8) + Mid$(QUERYLIST, LASTQUERYLIST, q1 - LASTQUERYLIST)
   LASTQUERYLIST = q1 + 1

      End If
 End If
 clickMe = -2
 End If
 
 ElseIf INK <> "" Then
 MKEY$ = vbNullString
 Else
 clickMe = 0
 once = False
 End If
 End If

  
againquery:
 a$ = INKEY$
 
If a$ = vbNullString Then
If TaskMaster Is Nothing Then Set TaskMaster = New TaskMaster
    If TaskMaster.QueueCount > 0 Then
  ProcTask2 bstack
  If Not NOEDIT Or Not QRY Then
  LCTCB dq, prive, -1: DestroyCaret
   oldLCTCB dq, prive, 0
  Exit Do
  End If
  SetText dq

LCTbasket dq, prive, .currow, .curpos
    Else
  
   End If
      If iamactive Then
 If ShowCaret(dq.hWnd) = 0 Then
 
   LCTCB dq, prive, 0
  End If
If Not bstack.IamThread Then

MyDoEvents1 Form1, , True
End If

 If Screen.ActiveForm Is Nothing Then
 iamactive = False:  If ShowCaret(dq.hWnd) <> 0 Then HideCaret dq.hWnd
Else
 
    If Not GetForegroundWindow = Screen.ActiveForm.hWnd Then
    iamactive = False:  If ShowCaret(dq.hWnd) <> 0 Then HideCaret dq.hWnd
  
    End If
    End If
    End If

  End If
    If bstack Is Nothing Then
    Set bstack = Basestack1
    NOEXECUTION = True
    MOUT = True
     Modalid = 0
                         ShutEnabledGuiM2000
                         MyDoEvents
                         GoTo contqueryhere
    End If
   If bstack.IamThread Then If myexit(bstack) Then GoTo contqueryhere

If Screen.ActiveForm Is Nothing Then
iamactive = False
Else
If Screen.ActiveForm.Name <> "Form1" Then
iamactive = False
Else
iamactive = GetForegroundWindow = Screen.ActiveForm.hWnd
End If
End If
If Fkey > 0 Then
If FK$(Fkey) <> "" Then
s$ = FK$(Fkey)
Fkey = 0
             ''  here
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0
 Exit Do
End If
End If


dq.FontTransparent = False
If RealLen(a$) = 1 Or Len(a$) = 1 Or (RealLen(a$) = 0 And Len(a$) = 1 And Len(s$) > 1) Then
   '
   
   If Len(a$) = 1 Then
    If InStr(endchars, a$) > 0 Then
     If a$ = vbCr Then
        If a$ <> Left$(endchars, 1) Then
            a$ = Left$(endchars, 1)
            If checknumber And Len(s$) = 0 Then
                s$ = "0": PlainBaSket dq, prive, "0", , , 0
                oldLCTCB dq, prive, 0
                LCTCB dq, prive, 0
            End If
        Else
            If checknumber And Len(s$) = 0 Then
                s$ = "0": PlainBaSket dq, prive, "0", , , 0
                oldLCTCB dq, prive, 0
                LCTCB dq, prive, 0
            End If
            LCTCB dq, prive, -1: DestroyCaret
            oldLCTCB dq, prive, 0
            Exit Do
        End If
        Else
                If checknumber And Len(s$) = 0 Then
                s$ = "0": PlainBaSket dq, prive, "0", , , 0
                oldLCTCB dq, prive, 0
                LCTCB dq, prive, 0
            End If
     End If
     End If
     End If
    If Asc(a$) = 27 And (escok Or Not checknumber) Then
        
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0
    s$ = vbNullString
    'If ExTarget Then End
    Result = 99
    Exit Do
ElseIf Asc(a$) = 27 Then
a$ = Chr$(0)
End If
If a$ = Chr(8) Then
DE$ = " "
    If Len(s$) > 0 Then
    AUX = RealLen(s$)
    
        ExcludeOne s$
             LCTCB dq, prive, -1: DestroyCaret
            oldLCTCB dq, prive, 0

        
        .curpos = .curpos - 1
        If .curpos < 0 Then
            .curpos = .mx - 1: .currow = .currow - 1

            If .currow < .mysplit Then
                ScrollDownNew dq, prive
                
                PlainBaSket dq, prive, RealRight(Prompt$ & s$, .mx - 1), , , 0
                DE$ = vbNullString
            End If
        End If

       LCTbasketCur dq, prive
        dX = .curpos
        dY = .currow
       PlainBaSket dq, prive, DE$, , , 0
       .curpos = dX
       .currow = dY
         
         
            oldLCTCB dq, prive, 0
            
    End If
End If
If safe$ <> "" Then
        a$ = 65
End If
w = AscW(a$)
If w < 0 Then
GoTo cont12345

ElseIf w > 31 And (RealLen(s$) < m& Or RealLen(a$, True) = 0) Then
If RealLen(a$, True) = 0 Then
    If Asc(a$) = 63 And s$ <> "" Then
        s$ = s$ & a$: a$ = s$: ExcludeOne s$: a$ = Mid$(a$, Len(s$) + 1)
        s$ = s$ + a$
        MKEY$ = vbNullString
        'UINK = VbNullString
        safe$ = a$
       INK = Chr$(8) + INK
    ElseIf a$ = vbNullString Then
        a$ = a$ + s$
        safe$ = a$
       
    Else
       ' If s$ = vbNullString Then a$ = " "
        GoTo cont12345
    End If
Else
cont12345:
    If InStr(excludechars, a$) > 0 Then

    Else
            If checknumber Then
                    fr1 = 1
                    If (s$ = vbNullString And a$ = "-") Or IsNumberQuery(s$ + a$, fr1, p, fr2) Then
                            If fr2 - 1 = RealLen(s$) + 1 Or (s$ = vbNullString And a$ = "-") Then
   If ShowCaret(dq.hWnd) <> 0 Then DestroyCaret
                If a$ = "." Then
                If Not NoUseDec Then
                    If OverideDec Then
                    PlainBaSket dq, prive, NowDec$, , , 0
                    Else
                    PlainBaSket dq, prive, ".", , , 0
                    End If
                Else
                    PlainBaSket dq, prive, QueryDecString, , , 0
                End If
                Else
                   PlainBaSket dq, prive, a$, , , 0
                   End If
                   s$ = s$ & a$
                 
              oldLCTCB dq, prive, 0
                  LCTCB dq, prive, 0
GdiFlush
                            End If
                    
                    End If
            Else
            If ShowCaret(dq.hWnd) <> 0 Then DestroyCaret
                   If safe$ <> "" Then
        a$ = safe$: safe$ = vbNullString
End If
 If InStr(endchars, a$) = 0 Then PlainBaSket dq, prive, a$, , , 0: s$ = s$ & a$
              If .curpos >= .mx Then
                                .curpos = 0
                                .currow = .currow + 1
                            End If
              oldLCTCB dq, prive, 0
                  LCTCB dq, prive, 0
                  GdiFlush
                
            End If
    End If
End If
If InStr(endchars, a$) > 0 Then
    If a$ >= " " Then
                     PlainBaSket dq, prive, a$, , , 0
              
      LCTCB dq, prive, -1: DestroyCaret
                                GdiFlush
                                End If
QUERY = a$
Exit Do
End If
 .pageframe = 0
 End If
End If
If Not QRY Then
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0
Exit Do
''HideCaret dq.hWnd:


End If
Loop


 
If Not QRY Then s$ = vbNullString
dq.FontTransparent = True
SetBkMode dq.Hdc, 1
QRY = False

Call mouse

If s$ <> "" And USELIST Then
q1 = InStr(QUERYLIST, Chr$(13) + s$ & Chr$(13))
If q1 = 1 Then ' same place
ElseIf q1 > 1 Then ' reorder
sp$ = Mid$(QUERYLIST, q1 + RealLen(s$) + 1)
QUERYLIST = Chr$(13) + s$ & Mid$(QUERYLIST, 1, q1 - 1) + sp$
Else ' insert
QUERYLIST = Chr$(13) + s$ & QUERYLIST
End If
LASTQUERYLIST = 2
End If
End With
contqueryhere:
If Not bstack.IamThread Then
MyDoEvents1 Form1, , True
End If
REFRESHRATE = oldREFRESHRATE
If TaskMaster Is Nothing Then Exit Function
If TaskMaster.QueueCount > 0 Then TaskMaster.RestEnd
players(GetCode(dq)) = prive
Set dq = Nothing
TaskMaster.RestEnd1

End Function


Public Sub GetXYb(dd As Object, mb As basket, x As Long, y As Long)
With mb
If dd.currentY Mod .Yt <= dv15 Then
y = (dd.currentY) \ .Yt
Else
y = (dd.currentY - .uMineLineSpace) \ .Yt
End If
x = dd.currentX \ .Xt

''
End With
End Sub
Public Sub GetXYb2(dd As Object, mb As basket, x As Long, y As Long)
With mb
x = dd.currentX \ .Xt
y = Int((dd.currentY / .Yt) + 0.5)
End With
End Sub
Sub Gradient(TheObject As Object, ByVal F&, ByVal t&, ByVal xx1&, ByVal xx2&, ByVal yy1&, ByVal yy2&, ByVal hor As Boolean, ByVal all As Boolean)
    Dim Redval&, Greenval&, Blueval&
    Dim R1&, G1&, b1&, sr&, SG&, sb&
    Dim obj As MetaDc
    F& = F& Mod &H1000000
    t& = t& Mod &H1000000
    Redval& = F& And &H10000FF
    Greenval& = (F& And &H100FF00) / &H100
    Blueval& = (F& And &HFF0000) / &H10000
    R1& = t& And &H10000FF
    G1& = (t& And &H100FF00) / &H100
    b1& = (t& And &HFF0000) / &H10000
    sr& = (R1& - Redval&) * 1000 / 127
    SG& = (G1& - Greenval&) * 1000 / 127
    sb& = (b1& - Blueval&) * 1000 / 127
    Redval& = Redval& * 1000
    
    Greenval& = Greenval& * 1000
    Blueval& = Blueval& * 1000
    Dim Step&, Reps&, FillTop As Single, FillLeft As Single, FillRight As Single, FillBottom As Single
     If TypeOf TheObject Is MetaDc Then
    If hor Then
    yy2& = TheObject.Height + 2 * dv15 - yy2&
    If all Then
    Step = ((yy2& - yy1&) / 127)
    Else
    Step = ((TheObject.Height + 2 * dv15) / 127)
    End If
    If all Then
    FillTop = yy1&
    Else
    FillTop = 0
    End If
    FillLeft = xx1&
    FillRight = TheObject.Width + 4 * dv15 - xx2&
    FillBottom = FillTop + Step * 2
    
    Else ' vertical
    
        xx2& = TheObject.Width + 4 * dv15 - xx2&
    If all Then
    Step = ((xx2& - xx1&) / 127)
    Else
    Step = ((TheObject.Width + 4 * dv15) / 127)
    End If
    If all Then
    FillLeft = xx1&
    Else
    FillLeft = 0
    End If
    FillTop = yy1&
    FillBottom = TheObject.Height + 2 * dv15 - yy2&
    FillRight = FillLeft + Step * 2
    
    End If
    Else
    If hor Then
    yy2& = TheObject.Height - yy2&
    If all Then
    Step = ((yy2& - yy1&) / 127)
    Else
    Step = ((TheObject.Height) / 127)
    End If
    If all Then
    FillTop = yy1&
    Else
    FillTop = 0
    End If
    FillLeft = xx1&
    FillRight = TheObject.Width - xx2&
    FillBottom = FillTop + Step * 2
    
    Else ' vertical
    
        xx2& = TheObject.Width - xx2&
    If all Then
    Step = ((xx2& - xx1&) / 127)
    Else
    Step = (TheObject.Width / 127)
    End If
    If all Then
    FillLeft = xx1&
    Else
    FillLeft = 0
    End If
    FillTop = yy1&
    FillBottom = TheObject.Height - yy2&
    FillRight = FillLeft + Step * 2
    
    End If
    
    
    
    
    End If
    If TypeOf TheObject Is MetaDc Then
    Set obj = TheObject
    
    For Reps = 1 To 127
    If hor Then
    
        If FillTop <= yy2& And FillBottom >= yy1& Then
        obj.Line2 FillLeft, RMAX(FillTop, yy1&), FillRight, RMIN(FillBottom, yy2&), rgb(Redval& / 1000, Greenval& / 1000, Blueval& / 1000), True
        End If
        Redval& = Redval& + sr&
        Greenval& = Greenval& + SG&
        Blueval& = Blueval& + sb&
        FillTop = FillBottom
        FillBottom = FillTop + Step
    Else
        If FillLeft <= xx2& And FillRight >= xx1& Then
        obj.Line2 RMAX(FillLeft, xx1&), FillTop, RMIN(FillRight, xx2&), FillBottom, rgb(Redval& / 1000, Greenval& / 1000, Blueval& / 1000), True
        End If
        Redval& = Redval& + sr&
        Greenval& = Greenval& + SG&
        Blueval& = Blueval& + sb&
        FillLeft = FillRight
        FillRight = FillRight + Step
    End If
    Next
        
    Else
    
    For Reps = 1 To 127
    If hor Then
        If FillTop <= yy2& And FillBottom >= yy1& Then
        TheObject.Line (FillLeft, RMAX(FillTop, yy1&))-(FillRight, RMIN(FillBottom, yy2&)), rgb(Redval& / 1000, Greenval& / 1000, Blueval& / 1000), BF
        End If
        Redval& = Redval& + sr&
        Greenval& = Greenval& + SG&
        Blueval& = Blueval& + sb&
        FillTop = FillBottom
        FillBottom = FillTop + Step
    Else
        If FillLeft <= xx2& And FillRight >= xx1& Then
        TheObject.Line (RMAX(FillLeft, xx1&), FillTop)-(RMIN(FillRight, xx2&), FillBottom), rgb(Redval& / 1000, Greenval& / 1000, Blueval& / 1000), BF
        End If
        Redval& = Redval& + sr&
        Greenval& = Greenval& + SG&
        Blueval& = Blueval& + sb&
        FillLeft = FillRight
        FillRight = FillRight + Step
    End If
    Next
    End If
End Sub
Function mycolor(q)
If Abs(q) > 2147483392# Then
If q < 0 Then
mycolor = GetSysColor(q And &HFF) And &HFFFFFF
Else
mycolor = GetSysColor((q - 4294967296#) And &HFF) And &HFFFFFF
End If
Exit Function
End If
If q = 0 Then
mycolor = 0
ElseIf q < 0 Or q > 15 Then

 mycolor = Abs(q) And &HFFFFFF
Else
mycolor = QBColor(q Mod 16)
End If
End Function




Sub ICOPY(d1 As Object, x1 As Long, y1 As Long, w As Long, H As Long)
Dim sV As Long
With players(GetCode(d1))
sV = BitBlt(d1.Hdc, CLng(d1.ScaleX(x1, 1, 3)), CLng(d1.ScaleY(y1, 1, 3)), CLng(d1.ScaleX(w, 1, 3)), CLng(d1.ScaleY(H, 1, 3)), d1.Hdc, CLng(d1.ScaleX(.XGRAPH, 1, 3)), CLng(d1.ScaleY(.YGRAPH, 1, 3)), SRCCOPY)
'sv = UpdateWindow(d1.hwnd)
End With
End Sub

Sub sHelp(Title$, doc$, x As Long, y As Long)
vH_title$ = Title$
vH_doc$ = doc$
vH_x = x
vH_y = y
End Sub

Sub vHelp(Optional ByVal bypassshow As Boolean = False)
Dim huedif As Long
Dim UAddPixelsTop As Long, monitor As Long

If abt Then
If vH_title$ = lastAboutHTitle Then Exit Sub
vH_title$ = lastAboutHTitle
vH_doc$ = LastAboutText
Else
If vH_title$ = vbNullString Then Exit Sub
End If
If bypassshow Then
monitor = FindMonitorFromMouse
Else
monitor = FindFormSScreen(Form4)
End If
If Not Form4.Visible Then Form4.Show , Form1: bypassshow = True

If bypassshow Then


If (ScrInfo(monitor).Height - vH_y * Helplastfactor + ScrInfo(monitor).top) < 0 Then
Helplastfactor = (ScrInfo(monitor).Height + ScrInfo(monitor).top) / vH_y
End If

If (ScrInfo(monitor).Width - vH_x * Helplastfactor + ScrInfo(monitor).Left) < 0 Then
 Helplastfactor = (ScrInfo(monitor).Left + ScrInfo(monitor).Width) / vH_x
End If
If VirtualScreenWidth * 0.8 < vH_x * Helplastfactor Then
Helplastfactor = VirtualScreenWidth * 0.8 / vH_x
End If

myform Form4, ScrInfo(monitor).Width - vH_x * Helplastfactor + ScrInfo(monitor).Left, ScrInfo(monitor).Height - vH_y * Helplastfactor + ScrInfo(monitor).top, vH_x * Helplastfactor, vH_y * Helplastfactor, True, Helplastfactor
Else
If Screen.Width <= Form4.Left - ScrInfo(monitor).Left Then
myform Form4, Screen.Width - vH_x * Helplastfactor + ScrInfo(monitor).Left, Form4.top, vH_x * Helplastfactor, vH_y * Helplastfactor, True, Helplastfactor
Else
myform Form4, Form4.Left, Form4.top, vH_x * Helplastfactor, vH_y * Helplastfactor, True, Helplastfactor
End If
End If
Form4.moveMe
If Form1.Visible Then
If Form1.DIS.Visible Then
  ''  If Abs(Val(hueconvSpecial(mycolor(uintnew(&H80000018)))) - Val(hueconvSpecial(-Paper))) > Abs(Val(hueconvSpecial(mycolor(uintnew(&H80000003)))) - Val(hueconvSpecial(-Paper))) Then
  If Abs(hueconv(mycolor(uintnew(&H80000018))) - val(hueconv(players(0).Paper))) > 10 And Not Abs(lightconv(mycolor(uintnew(&H80000018))) - val(lightconv(players(0).Paper))) < 50 Then
    Form4.backcolor = &H80000018
    Form4.label1.backcolor = &H80000018
    
    Else
    
    Form4.backcolor = &H80000003
    Form4.label1.backcolor = &H80000003
    End If

Else
''If Abs(Val(hueconvSpecial(mycolor(&H80000018))) - Val(hueconvSpecial(Form1.BackColor))) > Abs(Val(hueconvSpecial(mycolor(&H80000003))) - Val(hueconvSpecial(Form1.BackColor))) Then
     If Abs(hueconv(mycolor(uintnew(&H80000018))) - val(hueconv(Form1.backcolor))) > 10 And Not Abs(lightconv(mycolor(uintnew(&H80000018))) - val(lightconv(Form1.backcolor))) < 50 Then

    Form4.backcolor = &H80000018
    Form4.label1.backcolor = &H80000018
    Else
    
    Form4.backcolor = &H80000003
    Form4.label1.backcolor = &H80000003
    End If
End If
End If
With Form4.label1
.Visible = True
.enabled = False
.Text = vH_doc$
If Not bypassshow Then .SetRowColumn 1, 0
.EditDoc = False
.NoMark = True
If abt Then
.glistN.WordCharLeft = "["
.glistN.WordCharRight = "]"
.glistN.WordCharRightButIncluded = vbNullString
.glistN.WordCharLeftButIncluded = vbNullString
Else
.glistN.WordCharRightButIncluded = ChrW(160) + "#("
.glistN.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", "@", Chr$(9), "#", "%", "&", "$")
.glistN.WordCharRight = ConCat(":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", Chr$(9), "#")
.glistN.WordCharLeftButIncluded = "#$@~"

End If
.enabled = True
.NewTitle vH_title$, (4 + UAddPixelsTop) * Helplastfactor
If Not bypassshow Then .glistN.ShowMe
End With


'Form4.ZOrder
Form4.label1.glistN.DragEnabled = Not abt
If exWnd = 0 Then If Form1.Visible Then Form1.SetFocus
End Sub

Function FileNameType(extension As String) As String
Dim i As Long, fs, b
 strTemp = String(200, Chr$(0))
    'Get
    GetTempPath 200, StrPtr(strTemp)
    strTemp = LONGNAME(mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1)))
    If strTemp = vbNullString Then
     strTemp = mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1))
     If Right$(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
    End If
    
    i = FreeFile
    Open strTemp & "dummy." & extension For Output As i
    Print #i, "test"
    Close #i
    Sleep 10
    Set fs = CreateObject("Scripting.FileSystemObject")
  Set b = fs.GetFile(strTemp & "dummy." & extension)
    FileNameType = b.Type
    KillFile strTemp & "dummy." & extension
End Function
Function mylcasefILE(ByVal a$) As String
If a$ = vbNullString Then Exit Function
If casesensitive Then
' no case change
mylcasefILE = a$
Else
 mylcasefILE = LCase(a$)
 End If

End Function
Function myUcase(ByVal a$, Optional convert As Boolean = False) As String
Dim i As Long, w As Integer
If a$ = vbNullString Then Exit Function
 If AscW(a$) > 255 Or convert Then
 For i = 1 To Len(a$)
 w = AscW(Mid$(a$, i, 1))
 If w > 901 And w < 975 Then
 Select Case w
Case 902
Mid$(a$, i, 1) = ChrW(913)
Case 904
Mid$(a$, i, 1) = ChrW(917)
Case 906
Mid$(a$, i, 1) = ChrW(921)
Case 912
Mid$(a$, i, 1) = ChrW(921)
Case 905
Mid$(a$, i, 1) = ChrW(919)
Case 908
Mid$(a$, i, 1) = ChrW(927)
Case 911
Mid$(a$, i, 1) = ChrW(937)
Case 910
Mid$(a$, i, 1) = ChrW(933)
Case 940
Mid$(a$, i, 1) = ChrW(913)
Case 941
Mid$(a$, i, 1) = ChrW(917)
Case 943
Mid$(a$, i, 1) = ChrW(921)
Case 942
Mid$(a$, i, 1) = ChrW(919)
Case 972
Mid$(a$, i, 1) = ChrW(927)
Case 974
Mid$(a$, i, 1) = ChrW(937)
Case 973
Mid$(a$, i, 1) = ChrW(933)
Case 962
Mid$(a$, i, 1) = ChrW(931)
End Select
End If
Next i
End If
myUcase = UCase(a$)
End Function

Function myUcase1(ByVal a$, Optional convert As Boolean = False) As String
Dim i As Long, w As Integer
If a$ = vbNullString Then Exit Function
 
 For i = 1 To Len(a$)
 w = AscW(Mid$(a$, i, 1))
 If w >= 0 And (w < 65 Or w > 974) Then
 ElseIf w < 0 Then
 If w > -10241 And w < -9984 Then
 i = i + 1
 End If
 ElseIf w > 901 Or convert Then
 Select Case w
Case 902
Mid$(a$, i, 1) = ChrW(913)
Case 904
Mid$(a$, i, 1) = ChrW(917)
Case 906
Mid$(a$, i, 1) = ChrW(921)
Case 912
Mid$(a$, i, 1) = ChrW(921)
Case 905
Mid$(a$, i, 1) = ChrW(919)
Case 908
Mid$(a$, i, 1) = ChrW(927)
Case 911
Mid$(a$, i, 1) = ChrW(937)
Case 910
Mid$(a$, i, 1) = ChrW(933)
Case 940
Mid$(a$, i, 1) = ChrW(913)
Case 941
Mid$(a$, i, 1) = ChrW(917)
Case 943
Mid$(a$, i, 1) = ChrW(921)
Case 942
Mid$(a$, i, 1) = ChrW(919)
Case 972
Mid$(a$, i, 1) = ChrW(927)
Case 974
Mid$(a$, i, 1) = ChrW(937)
Case 973
Mid$(a$, i, 1) = ChrW(933)
Case 962
Mid$(a$, i, 1) = ChrW(931)
End Select
ElseIf Not convert Then
Exit For
End If
Next i

myUcase1 = UCase(a$)
End Function
Sub myUcase2(a$)
Dim i As Long, w As Integer
If a$ = vbNullString Then Exit Sub
 
 For i = 1 To Len(a$)
 w = AscW(Mid$(a$, i, 1))
  If w > 901 Or w < 975 Then
 Select Case w
Case 902
Mid$(a$, i, 1) = ChrW(913)
Case 904
Mid$(a$, i, 1) = ChrW(917)
Case 906
Mid$(a$, i, 1) = ChrW(921)
Case 912
Mid$(a$, i, 1) = ChrW(921)
Case 905
Mid$(a$, i, 1) = ChrW(919)
Case 908
Mid$(a$, i, 1) = ChrW(927)
Case 911
Mid$(a$, i, 1) = ChrW(937)
Case 910
Mid$(a$, i, 1) = ChrW(933)
Case 940
Mid$(a$, i, 1) = ChrW(913)
Case 941
Mid$(a$, i, 1) = ChrW(917)
Case 943
Mid$(a$, i, 1) = ChrW(921)
Case 942
Mid$(a$, i, 1) = ChrW(919)
Case 972
Mid$(a$, i, 1) = ChrW(927)
Case 974
Mid$(a$, i, 1) = ChrW(937)
Case 973
Mid$(a$, i, 1) = ChrW(933)
Case 962
Mid$(a$, i, 1) = ChrW(931)
End Select
 ElseIf w > -10241 And w < -9984 Then
 i = i + 1
End If
Next i
LSet a$ = UCase(a$)
End Sub
Function myLcase(ByVal a$) As String
If a$ = vbNullString Then Exit Function

a$ = LCase(a$)
Dim i As Long, ok As Boolean, w As Integer
For i = 1 To Len(a$)
w = AscW(Mid$(a$, i, 1))
If w > -10241 And w < -9984 Then
i = i + 1
ElseIf w = 963 Then
ok = True
Exit For
End If
Next i
 If ok Then
a$ = a$ & Chr(0)
' Here are greek letters for proper case conversion
a$ = Replace(a$, "σ" & Chr(0), "ς")
a$ = Replace(a$, Chr(0), "")
a$ = Replace(a$, "σ ", "ς ")
a$ = Replace(a$, "σ$", "ς$")
a$ = Replace(a$, "σ&", "ς&")
a$ = Replace(a$, "σ.", "ς.")
a$ = Replace(a$, "σ(", "ς(")
a$ = Replace(a$, "σ_", "ς_")
a$ = Replace(a$, "σ/", "ς/")
a$ = Replace(a$, "σ\", "ς\")
a$ = Replace(a$, "σ-", "ς-")
a$ = Replace(a$, "σ+", "ς+")
a$ = Replace(a$, "σ*", "ς*")
a$ = Replace(a$, "σ" & vbCr, "ς" & vbCr)
a$ = Replace(a$, "σ" & vbLf, "ς" & vbLf)
End If
myLcase = a$
End Function
Sub myLcase2(a$)
If a$ = vbNullString Then Exit Sub

a$ = LCase(a$)
Dim i As Long, ok As Boolean, w As Integer
For i = 1 To Len(a$)
w = AscW(Mid$(a$, i, 1))
If w > -10241 And w < -9984 Then
i = i + 1
ElseIf w = 963 Then
ok = True
Exit For
End If
Next i
 If ok Then
a$ = a$ & Chr(0)
' Here are greek letters for proper case conversion
a$ = Replace(a$, "σ" & Chr(0), "ς")
a$ = Replace(a$, Chr(0), "")
a$ = Replace(a$, "σ ", "ς ")
a$ = Replace(a$, "σ$", "ς$")
a$ = Replace(a$, "σ&", "ς&")
a$ = Replace(a$, "σ.", "ς.")
a$ = Replace(a$, "σ(", "ς(")
a$ = Replace(a$, "σ_", "ς_")
a$ = Replace(a$, "σ/", "ς/")
a$ = Replace(a$, "σ\", "ς\")
a$ = Replace(a$, "σ-", "ς-")
a$ = Replace(a$, "σ+", "ς+")
a$ = Replace(a$, "σ*", "ς*")
a$ = Replace(a$, "σ" & vbCr, "ς" & vbCr)
a$ = Replace(a$, "σ" & vbLf, "ς" & vbLf)
End If

End Sub
Function MesTitle$()
On Error Resume Next
If ttl Then
If Form1.Caption = vbNullString Then
If here$ = vbNullString Then
MesTitle$ = "M2000"
' IDE
Else
If LASTPROG$ <> "" Then
MesTitle$ = ExtractNameOnly(LASTPROG$, True)
Else
MesTitle$ = "M2000"
End If
End If
Else
MesTitle$ = Form1.Caption
End If
Else

If Typename$(Screen.ActiveForm) = "GuiM2000" Then
MesTitle$ = Screen.ActiveForm.Title
Else
If here$ = vbNullString Or LASTPROG$ = vbNullString Then
MesTitle$ = "M2000"
Else
If Not UseMe Is Nothing Then
If UseMe.AppTitle <> vbNullString Then
MesTitle$ = UseMe.AppTitle & " " & here$
Else
MesTitle$ = ExtractNameOnly(LASTPROG$, True) & " " & here$
End If
Else
MesTitle$ = ExtractNameOnly(LASTPROG$, True) & " " & here$
End If
End If
End If
End If
End Function
Public Function holdcontrol(wh As Object, mb As basket) As Long
Dim x1 As Long, y1 As Long
If TypeOf wh Is MetaDc Then holdcontrol = 100000: Exit Function
With mb
If .pageframe = 0 Then

If .mysplit > 0 Then .pageframe = (.My - .mysplit) * 4 / 5 Else .pageframe = Fix(.My * 4 / 5)
If .pageframe < 1 Then .pageframe = 1
.basicpageframe = .pageframe
holdcontrol = .pageframe
Else

holdcontrol = .basicpageframe
End If
End With
End Function
Public Sub HoldReset(Col As Long, mb As basket)
With mb
.basicpageframe = Col
If .basicpageframe <= 0 Then .basicpageframe = .pageframe
End With
End Sub
Public Sub gsb_file(Optional assoc As Boolean = True)
   Dim CD As String
     CD = App.Path
        AddDirSep CD

        If assoc Then
          associate ".gsb", "M2000 Ver" & Str$(VerMajor) & "." & CStr(VerMinor \ 100) & " User Module", CD & "M2000.EXE"
        Else
      deassociate ".gsb", "M2000 Ver" & Str$(VerMajor) & "." & CStr(VerMinor \ 100) & " User Module", CD & "M2000.EXE"
   End If
End Sub
Public Sub Switches(s$, Optional fornow As Boolean = False)
Dim cc As cRegistry
Set cc = New cRegistry
cc.Temp = fornow
cc.ClassKey = HKEY_CURRENT_USER
cc.SectionKey = basickey
Dim D$, w$, p As Long, b As Long
If s$ <> "" Then
    Do While FastSymbol(s$, "-")
        If IsLabel(Basestack1, s$, D$) > 0 Then
            D$ = UCase(D$)
            If D$ = "TEST" Then
                STq = False
                STEXIT = False
                STbyST = True
                Form2.Show , Form1
                Form2.label1(0) = vbNullString
                Form2.label1(1) = vbNullString
                Form2.label1(2) = vbNullString
                TestShowSub = vbNullString
                TestShowStart = 0
                stackshow Basestack1
                Form1.Show , Form5
                If Form3.Visible Then Form3.skiptimer = True: Form3.WindowState = 0
                trace = True
            ElseIf D$ = "NORUN" Then
                If ttl Then Form3.WindowState = vbNormal Else Form1.Show , Form5
                NORUN1 = True
            ElseIf D$ = "INP" Then
                Use13 = False
            ElseIf D$ = "FONT" Then
            ' + LOAD NEW
                cc.ValueKey = "FONT"
                cc.ValueType = REG_SZ
                cc.Value = "Monospac821Greek BT"
            ElseIf D$ = "SEC" Then
                cc.ValueKey = "NEWSECURENAMES"
                cc.ValueType = REG_DWORD
                cc.Value = 0
                SecureNames = False
            ElseIf D$ = "DIV" Then
                cc.ValueKey = "DIV"
                cc.ValueType = REG_DWORD
                cc.Value = 0
                UseIntDiv = False
            ElseIf D$ = "LINESPACE" Then
                cc.ValueKey = "LINESPACE"
                cc.ValueType = REG_DWORD
                cc.Value = 0
            ElseIf D$ = "SIZE" Then
                cc.ValueKey = "SIZE"
                cc.ValueType = REG_DWORD
                cc.Value = 15
            ElseIf D$ = "PEN" Then
                cc.ValueKey = "PEN"
                cc.ValueType = REG_DWORD
                cc.Value = 0
                cc.ValueKey = "PAPER"
                cc.ValueType = REG_DWORD
                cc.Value = 7
            ElseIf D$ = "BOLD" Then
                cc.ValueKey = "BOLD"
                cc.ValueType = REG_DWORD
                cc.Value = 0
            ElseIf D$ = "PAPER" Then
                cc.ValueKey = "PAPER"
                cc.ValueType = REG_DWORD
                cc.Value = 7
                cc.ValueKey = "PEN"
                cc.ValueType = REG_DWORD
                cc.Value = 0
            ElseIf D$ = "GREEK" Then
                If Not fornow Then
                    cc.ValueKey = "COMMAND"
                    cc.ValueType = REG_SZ
                    cc.Value = "LATIN"
                End If
                pagio$ = "LATIN"
            ElseIf D$ = "DARK" Then
                If Not fornow Then
                     cc.ValueKey = "HTML"
                     cc.ValueType = REG_SZ
                     cc.Value = "BRIGHT"
                End If
                pagiohtml$ = "BRIGHT"
            ElseIf D$ = "CASESENSITIVE" Then
                If Not fornow Then
                    cc.ValueKey = "CASESENSITIVE"
                    cc.ValueType = REG_SZ
                    cc.Value = "NO"
                End If
                casesensitive = False
            ElseIf D$ = "NBS" Then
                Nonbsp = True
            ElseIf D$ = "RDB" Then
                RoundDouble = False
            ElseIf D$ = "EXT" Then
                wide = False
            ElseIf D$ = "TAB" Then
                UseTabInForm1Text1 = False
            ElseIf D$ = "SBL" Then
                ShowBooleanAsString = False
            ElseIf D$ = "DIM" Then
                DimLikeBasic = False
            ElseIf D$ = "FOR" Then
                ForLikeBasic = False
            ElseIf D$ = "PRI" Then
                cc.ValueKey = "PRIORITY-OR"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(0)  ' FALSE IS WRONG VALUE HERE
                priorityOr = False
            ElseIf D$ = "REG" Then
                gsb_file False
            ElseIf D$ = "DEC" Then
                cc.ValueKey = "DEC"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(0)
                mNoUseDec = False
                CheckDec
            ElseIf D$ = "TXT" Then
                cc.ValueKey = "TEXTCOMPARE"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(0)
                mTextCompare = False
            ElseIf D$ = "REC" Then
                cc.ValueKey = "FUNCDEEP"  ' RESET
                cc.ValueType = REG_DWORD
                cc.Value = 300
                If m_bInIDE Then funcdeep = 128
                 ' funcdeep not used - but functionality stay there for old dll's
                ClaimStack
                If findstack - 100000 > 0 Then
                    stacksize = findstack - 100000
                End If
            ElseIf D$ = "MDB" Then
                cc.ValueKey = "MDBHELP"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(False)
                UseMDBHELP = False
            Else
                s$ = "-" & D$ & s$
                Exit Do
            End If
        Else
            Exit Do
        End If
        Sleep 2
    Loop
    Do While FastSymbol(s$, "+")
        If IsLabel(Basestack1, s$, D$) > 0 Then
            D$ = UCase(D$)
            If D$ = "TEST" Then
                STq = False
                STEXIT = False
                STbyST = True
                Form2.Show , Form1
                Form2.label1(0) = vbNullString
                Form2.label1(1) = vbNullString
                Form2.label1(2) = vbNullString
                TestShowSub = vbNullString
                TestShowStart = 0
                stackshow Basestack1
                Form1.Show , Form5
                If Form3.Visible Then Form3.skiptimer = True: Form3.WindowState = 0
                trace = True
                ElseIf D$ = "REG" Then
                gsb_file
            ElseIf D$ = "INP" Then
                Use13 = True
            ElseIf D$ = "FONT" Then
                ' + LOAD NEW
                cc.ValueKey = "FONT"
                cc.ValueType = REG_SZ
                If ISSTRINGA(s$, w$) Then cc.Value = w$
            ElseIf D$ = "SEC" Then
                cc.ValueKey = "NEWSECURENAMES"
                cc.ValueType = REG_DWORD
                cc.Value = -1
                SecureNames = True
            ElseIf D$ = "DIV" Then
                cc.ValueKey = "DIV"
                cc.ValueType = REG_DWORD
                cc.Value = -1
                UseIntDiv = True
            ElseIf D$ = "LINESPACE" Then
                cc.ValueKey = "LINESPACE"
                cc.ValueType = REG_DWORD
                If IsNumberLabel(s$, w$) Then If val(w$) >= 0 And val(w$) <= 60 * dv15 Then cc.Value = CLng(val(w$) * 2)
            ElseIf D$ = "SIZE" Then
                cc.ValueKey = "SIZE"
                cc.ValueType = REG_DWORD
                If IsNumberLabel(s$, w$) Then If val(w$) >= 8 And val(w$) <= 48 Then cc.Value = CLng(val(w$))
            ElseIf D$ = "PEN" Then
                cc.ValueKey = "PAPER"
                cc.ValueType = REG_DWORD
                p = cc.Value
                cc.ValueKey = "PEN"
                cc.ValueType = REG_DWORD
                If IsNumberLabel(s$, w$) Then
                    If p = val(w$) Then p = 15 - p Else p = val(w$) Mod 16
                    cc.Value = CLng(val(p))
                End If
            ElseIf D$ = "BOLD" Then
                cc.ValueKey = "BOLD"
                cc.ValueType = REG_DWORD
                cc.Value = 1
                If IsNumberLabel(s$, w$) Then cc.Value = CLng(val(w$) Mod 16)
            ElseIf D$ = "PAPER" Then
                cc.ValueKey = "PEN"
                cc.ValueType = REG_DWORD
                p = cc.Value
                cc.ValueKey = "PAPER"
                cc.ValueType = REG_DWORD
                If IsNumberLabel(s$, w$) Then
                    If p = val(w$) Then p = 15 - p Else p = val(w$) Mod 16
                    cc.Value = CLng(val(p))
                End If
            ElseIf D$ = "GREEK" Then
                If Not fornow Then
                    cc.ValueKey = "COMMAND"
                    cc.ValueType = REG_SZ
                    cc.Value = "GREEK"
                End If
                pagio$ = "GREEK"
            ElseIf D$ = "DARK" Then
                If Not fornow Then
                    cc.ValueKey = "HTML"
                    cc.ValueType = REG_SZ
                    cc.Value = "DARK"
                End If
                pagiohtml$ = "DARK"
            ElseIf D$ = "CASESENSITIVE" Then
                If Not fornow Then
                    cc.ValueKey = "CASESENSITIVE"
                    cc.ValueType = REG_SZ
                    cc.Value = "YES"
                End If
                casesensitive = True
            ElseIf D$ = "NBS" Then
                Nonbsp = False
            ElseIf D$ = "RDB" Then
                RoundDouble = True
            ElseIf D$ = "EXT" Then
                wide = True
            ElseIf D$ = "TAB" Then
                UseTabInForm1Text1 = True
            ElseIf D$ = "SBL" Then
                ShowBooleanAsString = True
            ElseIf D$ = "DIM" Then
                DimLikeBasic = True
            ElseIf D$ = "FOR" Then
                ForLikeBasic = True
            ElseIf D$ = "PRI" Then
                cc.ValueKey = "PRIORITY-OR"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(True)
                priorityOr = True
            ElseIf D$ = "TXT" Then
                cc.ValueKey = "TEXTCOMPARE"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(True)
                mTextCompare = True
            ElseIf D$ = "DEC" Then
                cc.ValueKey = "DEC"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(True)
                mNoUseDec = True
                CheckDec
            ElseIf D$ = "REC" Then
                cc.ValueKey = "FUNCDEEP"  ' RESET
                cc.ValueType = REG_DWORD
                funcdeep = 3260
                cc.Value = 3260 ' SET REVISION DEFAULT
                ClaimStack
                If findstack - 100000 > 0 Then
                    stacksize = findstack - 100000
                End If
            ElseIf D$ = "MDB" Then
                cc.ValueKey = "MDBHELP"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(True)
                UseMDBHELP = True
            Else
                s$ = "+" & D$ & s$
                Exit Do
            End If
        Else
            Exit Do
        End If
        Sleep 2
    Loop
End If
End Sub
    
    
Sub myesc(b$)
MyErMacro b$, "Escape", "Διακοπή εκτέλεσης"
End Sub
Sub wrongsizeOrposition(a$)
    MyErMacro a$, "Wrong Size-Position for reading buffer", "Λάθος Μέγεθος-θέση, για διάβασμα Διάρθρωσης"
End Sub
Sub wrongweakref(a$)
MyErMacro a$, "Wrong weak reference", "λάθος ισχνής αναφοράς"
End Sub
Sub negsqrt(a$)
MyErMacro a$, "negative number for root", "αρνητικός σε ρίζα"
End Sub
Sub expecteddecimal(a$)
MyErMacro a$, "Expected decimal separator char", "Περίμενα χαρακτήρα διαχωρισμού δεκαδικών"
End Sub
Sub wrongexprinstring(a$)
MyErMacro a$, "Wrong expression in string", "λάθος μαθηματική έκφραση στο αλφαριθμητικό"
End Sub
Sub unknownoffset(a$, s$)
MyErMacro a$, "Unknown Offset " & s$, "’γνωστη Μετάθεση " & s$
End Sub
Sub wronguseofenum(a$)
MyErMacro a$, "Wrong use of enumerator", "λάθος χρήση απαριθμητή"
End Sub
Sub nosuchfile()
MyEr "No such file", "Δεν υπάρχει τέτοιο αρχείο"
End Sub

Public Sub MyDoEvents()
On Error GoTo there
If TaskMaster Is Nothing Then
    DoEvents
    Exit Sub
ElseIf Not TaskMaster.Processing And TaskMaster.QueueCount = 0 Then
    DoEvents
    Exit Sub
Else
    If TaskMaster.PlayMusic Then
        TaskMaster.OnlyMusic = True
        TaskMaster.TimerTick
        TaskMaster.OnlyMusic = False
    End If
    TaskMaster.StopProcess
    TaskMaster.TimerTick
    DoEvents
    TaskMaster.StartProcess
End If
Exit Sub
there:
If Not TaskMaster Is Nothing Then TaskMaster.RestEnd1
End Sub

Public Function ContainsUTF16(ByRef Source() As Byte, Optional maxsearch As Long = -1) As Long
  Dim i As Long, lUBound As Long, lUBound2 As Long, lUBound3 As Long
  Dim CurByte As Byte, CurByte1 As Byte
  Dim CurBytes As Long, CurBytes1 As Long
    lUBound = UBound(Source)
    If lUBound > 4 Then
    CurByte = Source(0)
    CurByte1 = Source(1)
    If maxsearch = -1 Then
    maxsearch = lUBound - 1
    ElseIf maxsearch < 8 Or maxsearch > lUBound - 1 Then
    maxsearch = lUBound - 1
    End If
    
    
    
    For i = 2 To maxsearch Step 2
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
Function ReadUnicodeOrANSI(FileName As String, Optional ByVal EnsureWinLFs As Boolean, Optional feedback As Long) As String
Dim i&, FNr&, BLen&, WChars&, bom As Integer, BTmp As Byte, b() As Byte
Dim mLof As Long, nobom As Long
nobom = 1
' code from Schmidt, member of vbforums
If FileName = vbNullString Then Exit Function
On Error Resume Next
If GetDosPath(FileName) = vbNullString Then MissFile: Exit Function
 On Error GoTo ErrHandler
  BLen = FileLen(GetDosPath(FileName))
'  If Err.Number = 53 Then missfile: Exit Function
 
  If BLen = 0 Then Exit Function
  
  FNr = FreeFile
  Open GetDosPath(FileName) For Binary Access Read As FNr
      Get FNr, , bom
    Select Case bom
      Case &HFEFF, &HFFFE 'one of the two possible 16 Bit BOMs
        If BLen >= 3 Then
          ReDim b(0 To BLen - 3): Get FNr, 3, b 'read the Bytes
utf16conthere:
          feedback = 0
          If bom = &HFFFE Then 'big endian, so lets swap the byte-pairs
          feedback = 1
            For i = 0 To UBound(b) Step 2
              BTmp = b(i): b(i) = b(i + 1): b(i + 1) = BTmp
            Next
          End If
          ReadUnicodeOrANSI = b
        End If
      Case &HBBEF 'the start of a potential UTF8-BOM
        Get FNr, , BTmp
        If BTmp = &HBF Then 'it's indeed the UTF8-BOM
        feedback = 2
          If BLen >= 4 Then
            ReDim b(0 To BLen - 4): Get FNr, 4, b 'read the Bytes
            WChars = MultiByteToWideChar(65001, 0, b(0), BLen - 3, 0, 0)
            ReadUnicodeOrANSI = space$(WChars)
            MultiByteToWideChar 65001, 0, b(0), BLen - 3, StrPtr(ReadUnicodeOrANSI), WChars
          End If
        Else 'not an UTF8-BOM, so read the whole Text as ANSI
        feedback = 3
        
          ReadUnicodeOrANSI = StrConv(space$(BLen), vbFromUnicode)
          Get FNr, 1, ReadUnicodeOrANSI
        End If
        
      Case Else 'no BOM was detected, so read the whole Text as ANSI
        feedback = 3
       mLof = LOF(FNr)
       Dim buf() As Byte
       If mLof > 1000 Then
       ReDim buf(1000)
       Else
       ReDim buf(mLof)
       End If
       Get FNr, 1, buf()
       Seek FNr, 1
       Dim notok As Boolean
      If ContainsUTF8(buf()) Then 'maybe is utf-8
      feedback = 2
      nobom = -1
        ReDim b(0 To BLen - 1): Get FNr, 1, b
            WChars = MultiByteToWideChar(65001, 0, b(0), BLen, 0, 0)
            ReadUnicodeOrANSI = space$(WChars)
            MultiByteToWideChar 65001, 0, b(0), BLen, StrPtr(ReadUnicodeOrANSI), WChars
        Else
        notok = True
        
        
            Select Case ContainsUTF16(buf())
        Case 1
            nobom = -1
            bom = &HFEFF
            ReDim b(0 To BLen - 1): Get FNr, 1, b 'read the Bytes
            GoTo utf16conthere
        Case 2
            nobom = -1
            bom = &HFEFF
            ReDim b(0 To BLen - 1): Get FNr, 1, b 'read the Bytes
            GoTo utf16conthere
        End Select
        End If
        If notok Then
        ReDim b(0 To BLen - 1): Get FNr, 1, b
        If BLen Mod 2 = 1 Then
        ReadUnicodeOrANSI = StrConv(space$(BLen), vbFromUnicode)
        Else
        ReadUnicodeOrANSI = space$(BLen \ 2)
        End If
         CopyMemory ByVal StrPtr(ReadUnicodeOrANSI), b(0), BLen
         
         Clid = FoundLocaleId(Left$(ReadUnicodeOrANSI, 500))
         
         
         
        ReadUnicodeOrANSI = StrConv(ReadUnicodeOrANSI, vbUnicode, Clid)
        'End If
        End If
    End Select
    
    If InStr(ReadUnicodeOrANSI, vbCrLf) = 0 Then
      If InStr(ReadUnicodeOrANSI, vbLf) Then
      feedback = feedback + 10
   If EnsureWinLFs Then ReadUnicodeOrANSI = Replace(ReadUnicodeOrANSI, vbLf, vbCrLf)
      ElseIf InStr(ReadUnicodeOrANSI, vbCr) Then
      feedback = feedback + 20
      
    If EnsureWinLFs Then ReadUnicodeOrANSI = Replace(ReadUnicodeOrANSI, vbCr, vbCrLf)
      End If
    End If
    feedback = nobom * feedback
ErrHandler:
If FNr Then Close FNr
If Err Then
'MyEr Err.Description, Err.Description
Err.Raise Err.Number, Err.Source & ".ReadUnicodeOrANSI", Err.Description
End If
End Function

Public Function SaveUnicode(ByVal FileName As String, ByVal buf As String, mode2save As Long, Optional Append As Boolean = False) As Boolean
' using doc as extension you can read it from word...with automatic conversion to unicode
' OVERWRITE ALWAYS
Dim w As Long, a() As Byte, F$, i As Long, bb As Byte, yesswap As Boolean
On Error GoTo t12345
If Not Append Then
If Not NeoUnicodeFile(FileName) Then Exit Function
Else
If Not CanKillFile(FileName$) Then Exit Function
End If
F$ = GetDosPath(FileName)
If Err.Number > 0 Or F$ = vbNullString Then Exit Function
w = FreeFile
MyDoEvents
Open F$ For Binary As w
' mode2save
' 0 is utf-le
If Append Then Seek #w, LOF(w) + 1
mode2save = mode2save Mod 10
If mode2save = 0 Then
If Not Append Then
    a() = ChrW(&HFEFF)
    Put #w, , a()
End If
ElseIf mode2save = 1 Then
a() = ChrW(&HFFFE) ' big endian...need swap
If Not Append Then Put #w, , a()
yesswap = True
ElseIf Abs(mode2save) = 2 Then  'utf8
If mode2save > 0 And Not Append Then

        Put #w, , CByte(&HEF)
        Put #w, , CByte(&HBB)
        Put #w, , CByte(&HBF)
        End If
        Put #w, , Utf16toUtf8(buf)
        Close w
    SaveUnicode = True
        Exit Function
ElseIf mode2save = 3 Then ' ascii
Dim buf1() As Byte
buf1 = StrConv(buf, vbFromUnicode, Clid)
Put #w, , buf1()
      Close w
    SaveUnicode = True
        Exit Function
End If

Dim maxmw As Long, iPos As Long
iPos = 1
maxmw = 32000 ' check it with maxmw 20 OR 1
If yesswap Then
For iPos = 1 To Len(buf) Step maxmw
a() = Mid$(buf, iPos, maxmw)
For i = 0 To UBound(a()) - 1 Step 2
bb = a(i): a(i) = a(i + 1): a(i + 1) = bb
Next i
Put #w, 3, a()
Next iPos
Else
For iPos = 1 To Len(buf) Step maxmw
a() = Mid$(buf, iPos, maxmw)
Put #w, , a()
Next iPos
End If
Close w
SaveUnicode = True
t12345:
End Function
Public Sub getUniString(F As Long, s As String)
Dim a() As Byte
a() = s
Get #F, , a()
s = a()
End Sub
Public Function getUniStringNoUTF8(F As Long, s As String) As Boolean
Dim a() As Byte
a() = s
Get #F, , a()
If UBound(a) > 4 Then If Not ContainsUTF16(a(), 256) = 1 Then MyEr "No UTF16LE", "Δεν βρήκα UTF16LE": Exit Function
s = a()
getUniStringNoUTF8 = True
End Function
Public Sub putUniString(F As Long, s As String)
Dim a() As Byte
a() = s

Put #F, , a()
End Sub
Public Sub putANSIString(F As Long, s As String)
Dim a() As Byte
a() = StrConv(s, vbFromUnicode, Clid)

Put #F, , a()
End Sub
Public Function getUniStringlINE(F As Long, s As String) As Boolean
' 2 bytes a time... stop to line end and advance to next line

Dim a() As Byte, s1 As String, ss As Long, lbreak As String
a = " "
On Error GoTo a11
Do While Not (LOF(F) < Seek(F))
Get #F, , a()

s1 = a()
If s1 <> vbCr And s1 <> vbLf Then
s = s + s1
'If Asc(s1) = 63 And (AscW(a()) <> 63 And AscW(a()) <> -257) Then
'If AscW(a()) < &H4000 Then Exit Function
''End If
Else
If Not (LOF(F) < Seek(F)) Then
ss = Seek(F)
lbreak = s1
Get #F, , a()
s1 = a()
If s1 <> vbCr And s1 <> vbLf Or lbreak = s1 Then
Seek #F, ss  ' restore it
End If
End If
Exit Do
End If
Loop
getUniStringlINE = True
a11:
End Function

Public Sub getAnsiStringlINE(F As Long, s As String)
' 2 bytes a time... stop to line end and advance to next line
Dim a As Byte, s1 As String, ss As Long, lbreak As String
'a = " "
On Error GoTo a11
Do While Not (LOF(F) < Seek(F))
Get #F, , a

s1 = ChrW(AscW(ChrW(AscW(StrConv(ChrW(a), vbUnicode, Clid)))))
If s1 <> vbCr And s1 <> vbLf Then
s = s + s1
Else
If Not (LOF(F) < Seek(F)) Then
ss = Seek(F)
Get #F, , a
lbreak = s1
s1 = ChrW(AscW(ChrW(AscW(StrConv(ChrW(a), vbUnicode, Clid)))))

If s1 <> vbCr And s1 <> vbLf Or lbreak = s1 Then
Seek #F, ss  ' restore it
End If
End If
Exit Do
End If
Loop
'S = StrConv(S, vbUnicode)
a11:
End Sub
Public Sub getUniStringComma(F As Long, s As String, Optional nochar34 As Boolean)
' sring must be in quotes
' 2 bytes a time... stop to line end and advance to next line
' use numbers with . as decimal not ,
Dim a() As Byte, s1 As String, ss As Long, inside As Boolean
s = vbNullString

a = " "
On Error GoTo a1115

Do While Not (LOF(F) < Seek(F))
    Get #F, , a()
    s1 = a()
    If s1 <> " " Then
    If nochar34 Then s = s1: Exit Do
    If s1 = """" Then inside = True: Exit Do
    End If
Loop
' we throw the first
If Not nochar34 Then If s1 <> """" Then Exit Sub

Do While Not (LOF(F) < Seek(F))
    Get #F, , a()
    
    s1 = a()
    If s1 <> vbCr And s1 <> vbLf And nochar34 And Not s1 = inpcsvsep$ Then
        s = s + s1
    ElseIf s1 <> vbCr And s1 <> vbLf And s1 <> """" And Not nochar34 Then
        s = s + s1
    Else
        If nochar34 Then
        GoTo there
        ElseIf s1 = """" Then
            If s = vbNullString Then ' is the first we have empty string
                inside = False
            Else
            ' look if we have one  more
                If Not (LOF(F) < Seek(F)) Then
                    ss = Seek(F)
                    Get #F, , a()
                    If a(0) = 34 Then
                        s = s + Chr(34)
                        GoTo nn1
                    Else
                        Seek #F, ss
                    End If
                End If
            End If
            inside = False
            Do While Not (LOF(F) < Seek(F))
            Get #F, , a()
            s1 = a()
            
            If s1 = vbCr Or s1 = vbLf Or s1 = inpcsvsep$ Then Exit Do
            Loop
there:
            If s1 = inpcsvsep$ Then Exit Do
        End If
        If s1 <> inpcsvsep$ And (Not (LOF(F) < Seek(F))) And (Not inside) Then
            ss = Seek(F)
            Get #F, , a()
            s1 = a()
            If s1 <> vbCr And s1 <> vbLf Then Seek #F, ss             ' restore it
        End If
        If Not inside Then Exit Do Else s = s + s1
    End If
nn1:
Loop
a1115:
End Sub
Public Sub getAnsiStringComma(F As Long, s As String, Optional nochar34 As Boolean)
' sring must be in quotes
' 2 bytes a time... stop to line end and advance to next line
' use numbers with . as decimal not ,
Dim a As Byte, s1 As String, ss As Long, inside As Boolean
s = vbNullString

On Error GoTo a1111

Do While Not (LOF(F) < Seek(F))
Get #F, , a
s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, Clid)))
If s1 <> " " Then
If nochar34 Then s = s1: Exit Do
If s1 = """" Then inside = True: Exit Do

End If
Loop
' we throw the first
If Not nochar34 Then If s1 <> """" Then Exit Sub

Do While Not (LOF(F) < Seek(F))
Get #F, , a

s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, Clid)))
If s1 <> vbCr And s1 <> vbLf And nochar34 And Not s1 = inpcsvsep$ Then
    s = s + s1
ElseIf s1 <> vbCr And s1 <> vbLf And s1 <> """" And Not nochar34 Then
    s = s + s1
Else
If nochar34 Then
        GoTo there
        ElseIf s1 = """" Then
If s = vbNullString Then ' is the first we have empty string
inside = False
Else
' look if we have one  more
If Not (LOF(F) < Seek(F)) Then
ss = Seek(F)

Get #F, , a
If a = 34 Then
s = s + Chr(34)
GoTo nn1
Else
Seek #F, ss
End If
End If

End If
inside = False
Do While Not (LOF(F) < Seek(F))
Get #F, , a
s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, Clid)))

If s1 = vbCr Or s1 = vbLf Or s1 = inpcsvsep$ Then Exit Do

Loop
there:
If s1 = inpcsvsep$ Then Exit Do
End If
If s1 <> inpcsvsep$ And (Not (LOF(F) < Seek(F))) And (Not inside) Then
    ss = Seek(F)
    Get #F, , a
    s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, Clid)))
    If s1 <> vbCr And s1 <> vbLf Then
    Seek #F, ss  ' restore it
    End If
    End If
If Not inside Then Exit Do Else s = s + s1

End If
nn1:
Loop

a1111:
End Sub
Public Sub getUniRealComma(F As Long, s$)
' 2 bytes a time... stop to line end and advance to next line
' use numbers with . as decimal not ,
Dim a() As Byte, s1 As String, ss As Long
s$ = vbNullString
a = " "
On Error GoTo a111
Do While Not LOF(F) < Seek(F)
Get #F, , a()

s1 = a()
If s1 <> vbCr And s1 <> vbLf And s1 <> inpcsvsep$ Then
s = s + s1
Else
If s1 <> inpcsvsep$ And Not (LOF(F) < Seek(F)) Then
    ss = Seek(F)
    Get #F, , a()
    s1 = a()
    If s1 <> vbCr And s1 <> vbLf Then
    Seek #F, ss  ' restore it
    End If
End If
Exit Do
End If
Loop
s$ = MyTrim$(s$)
If LenB(s$) = 0 Then s$ = "0"
a111:


End Sub
Public Sub getAnsiRealComma(F As Long, s$)
' 2 bytes a time... stop to line end and advance to next line
' use numbers with . as decimal not ,
Dim a As Byte, s1 As String, ss As Long
s$ = vbNullString


On Error GoTo a112
Do While Not LOF(F) < Seek(F)
Get #F, , a

s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, Clid)))
If s1 <> vbCr And s1 <> vbLf And s1 <> inpcsvsep$ Then
s = s + s1
Else
If s1 <> inpcsvsep$ And Not (LOF(F) < Seek(F)) Then
    ss = Seek(F)
    Get #F, , a
    s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, Clid)))
    If s1 <> vbCr And s1 <> vbLf Then
    Seek #F, ss  ' restore it
    End If
End If
Exit Do
End If
Loop
s$ = MyTrim$(s$)
If LenB(s$) = 0 Then s$ = "0"
a112:


End Sub

Public Function PopOne(s$) As String
Dim i&, LL As Long, n As Long
LL = Len(s): If LL = 0 Then Exit Function
Dim a1() As Integer, A2() As Integer
ReDim a1(LL + 6)
ReDim A2(LL + 6)
Dim skip As Boolean
skip = GetStringTypeExW(&HB, 4, StrPtr(s$), LL, a1(0)) = 0
skip = GetStringTypeExW(&HB, 2, StrPtr(s$), LL, A2(0)) = 0 Or skip
If skip Then
 PopOne = Left$(s$, 1)
 s$ = Mid$(s$, 2)
Else
  i& = LL - 1
  LL = 0
  For n = 0 To i&
  If a1(n) = 2048 And A2(n) = 1 Then
  If LL = 2 Then Exit For
    LL = LL + 1
  ElseIf a1(n) = 4096 And A2(n) = 0 Then
    If LL = 2 Then Exit For
     LL = LL + 1
     ElseIf a1(n) = 3 And A2(n) = 11 Then
        If LL < 2 Then
            PopOne = Left$(s, 1)
            s = Mid$(s, 2)
            Exit Function
        End If
     ElseIf (a1(n) = 0 And A2(n) = 0) Or a1(n) = 1 Then
        If LL < 2 Then
            PopOne = Left$(s, 1)
            s = Mid$(s, 2)
            Exit Function
        End If
    Else
    If LL = 2 Then Exit For
    LL = LL + 2
   End If
  Next n
  If LL < 2 Then LL = 1 Else LL = LL \ 2
 PopOne = Left$(s$, LL)
s$ = Mid$(s$, LL + 1)
End If


End Function
Public Sub ExcludeOne(s$)
Dim i&, LL As Long, n As Long
LL = Len(s): If LL = 0 Then Exit Sub
Dim a1() As Integer, A2() As Integer
ReDim a1(LL + 6)
ReDim A2(LL + 6)
Dim skip As Boolean
skip = GetStringTypeExW(&HB, 4, StrPtr(s$), LL, a1(0)) = 0
skip = GetStringTypeExW(&HB, 2, StrPtr(s$), LL, A2(0)) = 0 Or skip
If skip Then
 s$ = Left$(s$, Len(s$) - 1)
Else
  i& = LL - 1
  LL = 0
  For n = i& To 0 Step -1
  If a1(n) = 2048 And A2(n) = 1 Then
  If LL = 2 Then Exit For
    LL = LL + 1
  ElseIf a1(n) = 4096 And A2(n) = 0 Then
    If LL = 2 Then Exit For
     LL = LL + 1
     ElseIf a1(n) = 3 And A2(n) = 11 Then
    ElseIf a1(n) = 0 And A2(n) = 0 Then
    ElseIf a1(n) = 1 Then
    Else
    If LL = 2 Then Exit For
    LL = LL + 2
   End If
  Next n
     s$ = Left$(s$, n + 1)
End If
End Sub
Function RealRight(s$, ByVal many As Long) As String
Dim i&, LL As Long, n As Long
LL = Len(s): If LL = 0 Then Exit Function
If many >= LL Then RealRight = s$: Exit Function

Dim a1() As Integer, A2() As Integer
ReDim a1(LL + 6)
ReDim A2(LL + 6)
Dim skip As Boolean
skip = GetStringTypeExW(&HB, 4, StrPtr(s$), LL, a1(0)) = 0
skip = GetStringTypeExW(&HB, 2, StrPtr(s$), LL, A2(0)) = 0 Or skip
If skip Then
 RealRight = Right$(s$, many)
Else
  i& = LL - 1
  LL = -(many - 1) * 2 + 2
  For n = i& To 0 Step -1
  If a1(n) = 2048 And A2(n) = 1 Then
  If LL = 2 Then Exit For
    LL = LL + 1
  ElseIf a1(n) = 4096 And A2(n) = 0 Then
    If LL = 2 Then Exit For
     LL = LL + 1
     ElseIf a1(n) = 3 And A2(n) = 11 Then
     ElseIf a1(n) = 0 And A2(n) = 0 Then
     ElseIf a1(n) = 1 Then
    Else
    If LL = 2 Then Exit For
    LL = LL + 2
   End If
  Next n
     RealRight = Mid$(s$, n + 1)
End If
End Function
Function RealLeft(s$, ByVal many As Long) As String
Dim i&, LL As Long, n As Long
LL = Len(s): If LL = 0 Then Exit Function
If many >= LL Then RealLeft = s$: Exit Function

Dim a1() As Integer, A2() As Integer
ReDim a1(LL + 6)
ReDim A2(LL + 6)
Dim skip As Boolean
skip = GetStringTypeExW(&HB, 4, StrPtr(s$), LL, a1(0)) = 0
skip = GetStringTypeExW(&HB, 2, StrPtr(s$), LL, A2(0)) = 0 Or skip
If skip Then
 RealLeft = Left$(s$, many)
Else
  i& = LL - 1
  LL = -(many - 1) * 2 + 2
  For n = 0 To i&
  If a1(n) = 2048 And A2(n) = 1 Then
  If LL = 2 Then Exit For
    LL = LL + 1
  ElseIf a1(n) = 4096 And A2(n) = 0 Then
    If LL = 2 Then Exit For
     LL = LL + 1
     ElseIf a1(n) = 3 And A2(n) = 11 Then
     ElseIf a1(n) = 0 And A2(n) = 0 Then
     ElseIf a1(n) = 1 Then
    Else
    If LL = 2 Then Exit For
    LL = LL + 2
   End If
  Next n
     RealLeft = Mid$(s$, 1, n + 1)
End If
End Function

Function Tcase(s$) As String
Dim a() As String, i As Long
If s$ = vbNullString Then Exit Function
a() = Split(s$, " ")
For i = 0 To UBound(a())
a(i) = myUcase(Left$(a(i), 1), True) + Mid$(myLcase(a(i)), 2)
Next i
If UBound(a()) > 0 Then
Tcase = Join(a(), " ")
Else
Tcase = a(0)
End If
End Function
Public Sub choosenext()
Dim catchit As Boolean
On Error Resume Next
If Not Screen.ActiveForm Is Nothing Then

    Dim x As Form
     For Each x In Forms
     If x.Name = "Form1" Or x.Name = "GuiM2000" Or x.Name = "Form2" Or x.Name = "Form4" Then
         If x.Visible And x.enabled Then
             If catchit Then x.SetFocus: Exit Sub
             If x.hWnd = GetForegroundWindow Then
             catchit = True
             End If
         End If
    End If
         
     Next x
     Set x = Nothing
     For Each x In Forms
     If x.Name = "Form1" Or x.Name = "GuiM2000" Or x.Name = "Form2" Or x.Name = "Form4" Then
         If x.Visible And x.enabled Then x.SetFocus: Exit Sub
             
             
         End If
     Next x
     Set x = Nothing
    End If

End Sub

Public Function CheckLastHandler(obj As Object) As Boolean
Dim oldobj As Object, first As Object, usehandler As mHandler
If obj Is Nothing Then Exit Function
Set first = obj

Dim kk As Long
again:
If kk > 20 Then Set obj = first: Exit Function
If TypeOf obj Is mHandler Then
Set usehandler = obj
        If usehandler.indirect >= 0 And usehandler.indirect <= var2used Then
                Set oldobj = obj
                Set obj = var(usehandler.indirect)
                kk = kk + 1
                GoTo again
        Else
                kk = kk + 1
                Set oldobj = obj
                Set obj = usehandler.objref
                GoTo again
        End If

    'End If
    
End If
If Not oldobj Is Nothing Then Set obj = oldobj: Set oldobj = Nothing: CheckLastHandler = True: Exit Function
Set obj = first
End Function
Public Function CheckLastHandlerVariant(obj) As Boolean
Dim oldobj As Object, first As Object, usehandler As mHandler
If obj Is Nothing Then Exit Function
Set first = obj

Dim kk As Long
again:
If kk > 20 Then Set obj = first: Exit Function
If obj Is Nothing Then Exit Function
If TypeOf obj Is mHandler Then
    Set usehandler = obj
        If usehandler.indirect >= 0 And usehandler.indirect <= var2used Then
                Set oldobj = obj
                Set obj = var(usehandler.indirect)
                kk = kk + 1
                GoTo again
        Else
                kk = kk + 1
                Set oldobj = obj
                Set obj = usehandler.objref
                GoTo again
        End If

    'End If
    
End If
If Not oldobj Is Nothing Then Set obj = oldobj: Set oldobj = Nothing: CheckLastHandlerVariant = True: Exit Function
Set obj = first
End Function
Public Function CheckLastHandlerOrIterator(obj As Object, lastindex As Long) As Boolean
Dim oldobj As Object, first As Object, usehandler As mHandler
If obj Is Nothing Then Exit Function
Set first = obj
lastindex = -1
Dim kk As Long
again:
If kk > 20 Then Set obj = first: Exit Function
If TypeOf obj Is mHandler Then
Set usehandler = obj
        If usehandler.UseIterator Then lastindex = usehandler.index_cursor
        If usehandler.indirect >= 0 And usehandler.indirect <= var2used Then
                Set oldobj = obj
                Set obj = var(usehandler.indirect)
                kk = kk + 1
                GoTo again
        Else
                kk = kk + 1
                Set oldobj = obj
                Set obj = usehandler.objref
                GoTo again
        End If

End If
    

If Not oldobj Is Nothing Then Set obj = oldobj: Set oldobj = Nothing: CheckLastHandlerOrIterator = True: Exit Function
Set obj = first
End Function
Public Function IfierVal()
If LastErNum <> 0 Then LastErNum = 0: IfierVal = True
End Function
Public Sub OutOfLimit()
  MyEr "Out of limit", "Εκτός ορίου"
End Sub
Public Sub stackproblem()
MyEr "Problem in return stack", "Πρόβλημα στον σωρό επιστροφής"
End Sub
Public Sub PlaceAcommaBefore()
MyEr "Place a comma before", "Βάλε ένα κόμμα πριν"
End Sub
Public Sub unknownid(b$, w$)
MyErMacro b$, "unknown identifier " & w$, "’γνωστο αναγνωριστικό " & w$
End Sub
Public Sub VarNull()
MyEr "Vatiable is Null", "Η μεταβλητή έχει την καμία τιμή - Null"
End Sub

Public Sub NoRename()
    MyEr "Nothing renamed", "Δεν έγινε μετονομασία"
End Sub
Public Sub MissCdib()
  MyEr "Missing IMAGE", "Λείπει εικόνα"
End Sub
Public Sub MissFile()
 MyEr "File not found", "Δεν βρέθηκε ο αρχείο"
End Sub
Public Sub BadObjectDecl()
  MyEr "Bad object declaration - use Clear Command for Gui Elements", "Λάθος όρισμα αντικειμένου - χρησιμοποίησε Καθαρό για να καθαρίσεις τυχόν στοιχεία του γραφικού περιβάλλοντος"
End Sub
Public Sub NoEnumerator()
  MyEr " - No enumerator found for this object", " - Δεν βρήκα δρομέα συλλογής για αυτό το αντικείμενο"
End Sub
Public Sub AssigntoNothing()
  MyEr "Bad object declaration - use Declare command", "Λάθος όρισμα αντικειμένου - χρησιμοποίησε την Όρισε"
End Sub
Public Sub Overflow()
 MyEr "Overflow", "υπερχείλιση"
End Sub
Public Sub MissCdibStr()
  MyEr "Missing IMAGE in string", "Λείπει εικόνα στο αλφαριθμητικό"
End Sub
Public Sub MissStackStr()
  MyEr "Missing string value from stack", "Λείπει αλφαριθμητικό από το σωρό"
End Sub
Public Sub WrongFileHandler()
MyEr "Wrong File Handler", "Λάθος Χειριστής Αρχείου"
End Sub

Public Sub MissStackItem()
 MyEr "Missing item from stack", "Λείπει κάτι από το σωρό"
End Sub
Public Sub MissStackNumber()
 MyEr "Missing number value from stack", "Λείπει αριθμός από το σωρό"
End Sub
Public Sub MissMic()
MyEr "No mic found", "Δεν βρήκα είσοδο ηχογράφησης"
End Sub
Public Sub missNumber()
MyEr "Only number allowed", "Μόνο αριθμός επιτρέπεται"
End Sub
Public Sub MissNumExpr()
MyEr "Missing number expression", "Λείπει αριθμητική παράσταση"
End Sub
Public Sub MissLicense()
MyEr "Missing License", "Λείπει ’δεια"
End Sub
Public Sub MissStringExpr()
MyEr "Missing string expression", "Λείπει αλφαριθμητική παράσταση"
End Sub
Public Sub MissString()
MyEr "Missing string", "Λείπει αλφαριθμητικό"
End Sub
Public Sub MissStringNumber()
MyEr "Missing string or number", "Λείπει αλφαριθμητικό ή αριθμός"
End Sub

Public Sub NoCreateFile()
    MyEr "Can't create file", "Δεν μπορώ να φτιάξω αρχείο"
End Sub
Public Sub BadFilename()
MyEr "Bad filename", "Λάθος στο όνομα αρχείου"
End Sub
Public Sub ReadOnly()
MyEr "Read Only", "Μόνο για ανάγνωση"
End Sub
Public Sub MissDir()
MyEr "Missing directory name", "Λείπει όνομα φακέλου"
End Sub
Public Sub MissType()
MyEr "Wrong data type", "’λλος τύπος μεταβλητής"
End Sub

Public Sub BadPath()
MyEr "Bad Path name", "Λάθος στο όνομα φακέλου (τόπο)"
End Sub
Public Sub BadReBound()
MyEr "Can't commit a reference here", "Δεν μπορώ να αναθέσω εδώ μια αναφορά"
End Sub
Public Sub oxiforPrinter()
MyEr "Not allowed this command for printer", "Δεν επιτρέπεται αυτή η εντολή για τον εκτυπωτή"
End Sub
Public Sub ResourceLimit()
MyEr "No more Graphic Resource for forms - 100 Max", "Δεν έχω άλλο χώρο για γραφικά σε φόρμες - 100 Μεγιστο"
End Sub
Public Sub oxiforforms()
MyEr "Not allowed this command for forms", "Δεν επιτρέπεται αυτή η εντολή για φόρμες"
End Sub
Public Sub oxiforMetaFiles()
MyEr "Not allowed this command for drawings", "Δεν επιτρέπεται αυτή η εντολή για σχέδια"
End Sub
Public Sub SyntaxError()
If LastErName = vbNullString Then
MyEr "Syntax Error", "Συντακτικό Λάθος"
Else
If LastErNum = 0 Then LastErNum = -1 ' general
LastErNum1 = LastErNum
End If
End Sub
Public Sub MissingnumVar()
MyEr "missing numeric variable", "λείπει αριθμητική μεταβλητή"
End Sub
Public Sub BadGraphic()
MyEr "Can't operate graphic", "δεν μπορώ να χειριστώ το γραφικό"
End Sub
Public Sub SelectorInUse()
MyEr "File/Folder Selector in Use", "Η φόρμα επιλογής αρχείων/φακέλων είναι σε χρήση"
End Sub
Public Sub MissingDoc()  ' this is for identifier or execute part
MyEr "missing document type variable", "λείπει μεταβλητή τύπου εγγράφου"
End Sub
Public Sub MissingDocOrArrayOrInventory()  ' this is for identifier or execute part
MyEr "missing document or Array or Inventory type ", "λείπει τύπος έγγραφο ή πίνακα ή κατάσταση"
End Sub
Public Sub MissingArrayOrInventory()  ' this is for identifier or execute part
MyEr "missing Array or Inventory type ", "λείπει τύπος πίνακα ή κατάσταση"
End Sub
Public Sub MissingLabel()
MyEr "Missing label/Number line", "Λείπει Ετικέτα/Αριθμός γραμμής"
End Sub
Public Sub MissFuncParammeterdOCVar(ar$)
MyEr "Not a Document variable " + ar$, "Δεν είναι μεταβλητή τύπου εγγράφου " + ar$
End Sub
Public Sub MissingBlock()  ' this is for identifier or execute part
MyEr "missing block {} or string expression", "λείπει κώδικας σε {} η αλφαριθμητική έκφραση"
End Sub
Public Sub MissingBlockCode()
MyEr "missing block {}", "λείπει κώδικας σε μπλοκ {}"
End Sub
Public Sub OnlyOneLineAllowed()
MyEr "Use block {} in starting line only", "Χρησιμοποίησε μπλοκ {} στην αρχική γραμμή"
End Sub
Public Function CheckBlock(once As Boolean) As Long
                                    If once Then
                                        OnlyOneLineAllowed
                                    Else
                                        MissingBlockCode
                                    End If
End Function

Public Sub MissingEnumBlock()
MyEr "missing block {} for enumeration constants", "λείπει μπλοκ {} για σταθερές απαρίθμησης "
End Sub
Public Sub MissingCodeBlock()
MyEr "missing block {}", "λείπει μπλοκ κώδικα σε {}"
End Sub
Public Sub MissingArray(w$)
MyEr "Can't find array " & w$ & ")", "Δεν βρίσκω πίνακα " & w$ & ")"
End Sub
Public Sub ErrNum()
MyEr "Error in number", "Λάθος στον αριθμό"
End Sub
Public Sub CantAssignValue()
MyEr "Can't assign value to constant", "Δεν μπορώ να βάλω τιμή σε σταθερά"
End Sub
Public Sub ExpectedEnumType()
 MyEr "Expected Enumaration Type", "Περίμενα τύπο απαρίθμησης"
End Sub

Public Sub ExpectedVariable()
 MyEr "Expected variable", "Περίμενα μεταβλητή"
End Sub
Public Sub Expected(w1$, w2$)
 MyEr "Expected object type " + w1$, "Περίμενα αντικείμενο τύπου " + w2$
End Sub
Public Sub ExpectedCaseorElseorEnd2()
MyEr "Expected Case or Else or End Select", "Περίμενα Με ή Αλλιώς ή Τέλος Επιλογής"
End Sub
Public Sub ExpectedCaseorElseorEnd()
 MyEr "Expected Case or Else or End Select, for two or more commands use {}", "Περίμενα Με ή Αλλιώς ή Τέλος Επιλογής, για δυο ή περισσότερες εντολές χρησιμοποίησε { }"
End Sub
Public Sub ExpectedCommentsOnly()
 MyEr "Expected comments (using ' or \) or new line", "Περίμενα σημειώσεις (με ' ή \) ή αλλαγή γραμής"
End Sub

Public Sub ExpectedEndSelect()
 MyEr "Expected Εnd Select", "Περίμενα Τέλος Επιλογής"
End Sub
Public Sub ExpectedEndSelect2()
 MyEr "Expected Εnd Select, for two or more commands use {}", "Περίμενα Τέλος Επιλογής, για δυο ή περισσότερες εντολές χρησιμοποίησε { }"
End Sub
Public Sub LocalAndGlobal()
MyEr "Global and local together;", "Γενική και τοπική μαζί!"
End Sub
Public Sub UnknownProperty(w$)
MyEr "Unknown Property " & w$, "’γνωστη ιδιότητα " & w$
End Sub
Public Sub UnknownVariable(v$)
Dim i As Long
i = rinstr(v$, "." + ChrW(8191))
If i > 0 Then
    i = rinstr(v$, ".")
    MyEr "Unknown Variable " & Mid$(v$, i), "’γνωστη μεταβλητή " & Mid$(v$, i)
Else
    i = rinstr(v$, "].")
    If i > 0 Then
        MyEr "Unknown Variable " & Mid$(v$, i + 2), "’γνωστη μεταβλητή " & Mid$(v$, i + 2)
    Else
        i = rinstr(v$, ChrW(8191))
    If i > 0 Then
        i = InStr(i + 1, v$, ".")
        If i > 0 Then
            MyEr "Unknown Variable " & Mid$(v$, i + 1), "’γνωστη μεταβλητή " & Mid$(v$, i + 1)
        Else
            MyEr "Unknown Variable", "’γνωστη μεταβλητή"
        End If
    Else
        MyEr "Unknown Variable " & v$, "’γνωστη μεταβλητή " & v$
    End If
    End If
End If
End Sub
Sub indexout(a$)
MyErMacro a$, "Index out of limits", "Δείκτης εκτός ορίων"
End Sub

Sub wrongfilenumber(a$)
 MyErMacro a$, "not valid file number", "λάθος αριθμός αρχείου"
End Sub
Public Sub WrongArgument(a$)
MyErMacro a$, Err.Description, "Λάθος όρισμα"
End Sub
Public Sub UnKnownWeak(w$)
 MyEr "Unknown Weak " & w$, "’γνωστη ισχνή " & w$
End Sub
Public Sub InternalEror()
MyEr "Internal error", "Εσωτερικό λάθος"
End Sub
Sub NegativeIindex(a$)
MyErMacro a$, "negative index", "αρνητικός δείκτη"
End Sub
Sub joypader(a$, R)
MyErMacro a$, "Joypad number " & CStr(R) & " isn't ready", "Το νούμερο Λαβής " & CStr(R) & " δεν είναι έτοιμο"
End Sub
Sub noImage(a$)
MyErMacro a$, "Νο image in string", "Δεν υπάρχει εικόνα στο αλφαριθμητικό"
End Sub
Sub noImageInBuffer(a$)
MyErMacro a$, "No Image in Buffer", "Δεν έχει εικόνα η Διάρθρωση"
End Sub

Sub WrongJoypadNumber(a$)
MyErMacro a$, "Joypad number 0 to 15", "Αριθμός Λαβής από 0 έως 15"
End Sub
Sub CantFindArray(a$, s$)
MyErMacro a$, "Can't find array " & s$, "Δεν βρίσκω πίνακα " & s$
End Sub
Sub CantReadDimension(a$, s$)
 MyErMacro a$, "Can't read dimension index from array " & s$, "Δεν μπορώ να διαβάσω τον δείκτη διάστασης του πίνακα " & s$

End Sub
Sub cantreadlib(a$)
MyErMacro a$, "Can't Read TypeLib", "Δεν μπορώ να διαβάσω τους τύπους των παραμέτρων"
End Sub
Public Sub NotForArray()
MyEr "not for array items", "όχι για στοιχεία πίνακα"
End Sub
Public Sub NotArray()  ' this is for identifier or execute part
MyEr "Expected Array", "Περίμενα πίνακα"
End Sub
Public Sub NotExistArray()  ' this is for identifier or execute part
MyEr "Array not exist", "Δεν υπάρχει τέτοιος πίνακας"
End Sub
Public Sub MissingGroup()  ' this is for identifier or execute part
MyEr "missing group type variable", "λείπει μεταβλητή τύπου ομάδας"
End Sub
Public Sub MissingGroupExp()  ' this is for identifier or execute part
MyEr "missing group type expression", "λείπει έκφραση τύπου ομάδας"
End Sub
Public Sub BadGroupHandle()  ' this is for identifier or execute part
MyEr "group isn't variable", "η ομάδα δεν είναι μεταβλητή"
End Sub
Public Sub MissingDocRef()  ' this is for identifier or execute part
MyEr "invalid document pointer", "μη έγκυρος δείκτης εγγράφου"
End Sub
Public Sub MissingObjReturn()
MyEr "Missing Object", "Δεν βρήκα αντικείμενο"
End Sub
Public Sub NoNewLambda()
    MyEr "No New statement for lambda", "Όχι δήλωση νέου για λαμδα"
End Sub
Public Sub ExpectedObj(nn$)
MyEr "Expected object type " + nn$, "Περίμενα αντικείμενο τύπου " + nn$
End Sub
Public Sub MisOperatror(ss$)
MyEr "Group not support operator " + ss$, "Η ομάδα δεν υποστηρίζει το τελεστή " + ss$
End Sub
Public Sub CantReadFileTimeStap(a$)
MyErMacro a$, "Can't Read File TimeStamp", "Δεν μπορώ να διαβάσω την Χρονοσήμανση του αρχείου"
End Sub

Public Sub ExpectedObjInline(nn$)
MyErMacro nn$, "Expected Object", "Περίμενα αντικείμενο"
End Sub
Public Sub MissingObj()
MyEr "missing object type variable", "λείπει μεταβλητή τύπου αντικειμένου"
End Sub
Public Sub BadGetProp()
MyEr "Can't Get Property", "Δεν μπορώ να διαβάσω αυτή την ιδιότητα"
End Sub
Public Sub BadLetProp()
MyEr "Can't Let Property", "Δεν μπορώ να γράψω αυτή την ιδιότητα"
End Sub
Public Sub NoNumberAssign()
MyEr "Can't assign number to object", "Δεν μπορώ να δώσω αριθμό στο αντικείμενο"
End Sub
Public Sub NoAssignThere()
MyEr "Use Return Object to change items", "Χρησιμοποίησε την Επιστροφή αντικείμενο για να επιστρέψεις τιμές"
End Sub
Public Sub NoObjectpAssignTolong()
MyEr "Can't assign object to long", "Δεν μπορώ να δώσω αντικείμενο στον μακρυ"
End Sub
Public Sub NoObjectpAssignToInteger()
MyEr "Can't assign object to Integer", "Δεν μπορώ να δώσω αντικείμενο στον ακέραιο"
End Sub
Public Sub NoObjectAssign()
MyEr "Can't assign object", "Δεν μπορώ να δώσω αντικείμενο"
End Sub
Public Sub NoNewStatFor(w1$, w2$)
MyEr "No New statement for " + w1$, "Όχι δήλωση νέου για " + w2$
End Sub
Public Sub NoThatOperator(ss$)
    MyEr ss$ + " operator not allowed in group definition", " Ο τελεστής " + ss$ + " δεν επιτρεπεται σε ορισμό ομάδας"
End Sub
Public Sub MissingObjRef()
MyEr "invalid object pointer", "μη έγκυρος δείκτης αντικειμένου"
End Sub
Public Sub MissingStrVar()  ' this is for identifier or execute part
MyEr "missing string variable", "λείπει αλφαριθμητική μεταβλητή"
End Sub
Public Sub NoSwap(nameOfvar$)
MyEr "Can't swap ", "Δεν μπορώ να αλλάξω τιμές "
End Sub
Public Sub Nosuchvariable(nameOfvar$)
MyEr "No such variable " + nameOfvar$, "δεν υπάρχει τέτοια μεταβλητή " + nameOfvar$
End Sub
Public Sub NoValueForVar(w$)
If LastErNum = 0 Then
MyEr "No value for variable " & w$, "Χωρίς τιμή η μεταβλητή " & w$
End If
End Sub
Public Sub NoReference()
   MyEr "No reference exist", "Δεν υπάρχει αναφορά"
End Sub
Public Sub NoCommandOrBlock()
MyEr "Expected in Select Case a Block or a Command", "Περίμενα στην Επίλεξε Με μια εντολή ή ένα μπλοκ εντολών)"
End Sub

Public Sub NoSecReF()
MyEr "No reference allowed - use new variable", "Δεν δέχεται αναφορά - χρησιμοποίησε νέα μεταβλητή"
End Sub
Public Sub MissSymbolMyEr(wht$)   ' not the macro one
MyEr "missing " & wht$, "λείπει " & wht$
End Sub
Public Sub MissTHENELSE()
    MyEr "missing THEN or ELSE", "δεν βρήκα ΤΟΤΕ ή ΑΛΛΙΩΣ"
End Sub

Public Sub MissENDIF()
    MyEr "missing END IF", "δεν βρήκα ΤΕΛΟΣ ΑΝ"
End Sub
Public Sub MissIF()
    MyEr "No IF for END IF", "δεν βρήκα ΑΝ για την ΤΕΛΟΣ ΑΝ"
End Sub

Public Sub BadCommand()
 MyEr "Command for supervisor rights", "Εντολή μόνο για επόπτη"
End Sub
Public Sub NoClauseInThread()
MyEr "can't find ERASE or HOLD or RESTART or INTERVAL clause", "Δεν μπορώ να βρω όρο όπως το ΣΒΗΣΕ ή το ΚΡΑΤΑ ή το ΞΕΚΙΝΑ ή το ΚΑΘΕ"
End Sub
Public Sub NoThisInThread()
MyEr "Clause This can't used outside a thread", "Ο όρος ΑΥΤΟ δεν μπορεί να χρησιμοποιηθεί έξω από ένα νήμα"
End Sub
Public Sub MisInterval()
MyEr "Expected number for interval, miliseconds", "Περίμενα αριθμό για ορισμό τακτικού διαστήματος εκκίνησης νήματος (χρόνο σε χιλιοστά δευτερολέπτου)"
End Sub
Public Sub NoRef2()
MyEr "No with reference in left side of assignment", "Όχι με αναφορά στην εκχώρηση τιμής"
End Sub
Public Sub WrongObject()
MyEr "Wrong object type", "λάθος τύπος αντικειμένου"
End Sub
Public Sub NullObject()
MyEr "object type is Nothing", "O τύπος αντικειμένου είναι Τίποτα"
End Sub
Public Sub WrongType()
MyEr "Wrong type", "λάθος τύπος"
End Sub
Public Sub GroupWrongUse()
MyEr "Something wrong with group", "Κάτι πάει στραβά με την ομάδα"
End Sub
Public Sub GroupCantSetValue()
    MyEr "Group can't set value", "Η ομάδα δεν μπορεί να θέσει τιμή"
End Sub
Public Sub PropCantChange()
MyEr "Property can't change", "Η ιδιότητα δεν μπορεί να αλλάξει"
End Sub
Public Sub NeedAGroupFromExpression()
MyEr "Need a group from expression", "Χρειάζομαι μια ομάδα από την έκφραση"
End Sub
Public Sub NeedAGroupInRightExpression()
MyEr "Need a group from right expression", "Χρειάζομαι μια ομάδα από την δεξιά έκφραση"
End Sub
Public Sub NotAfter(a$)
MyErMacro a$, "not an expression after not operator", "δεν υπάρχει παράσταση δεξιά τού τελεστή όχι"
End Sub
Public Sub EmptyArray()
MyEr "Empty Array", "’δειος Πίνακας"
End Sub
Public Sub EmptyStack(a$)
 MyErMacro a$, "Stack is empty", "O σωρός είναι άδειος"
End Sub
Public Sub StackTopNotArray(a$)
 MyErMacro a$, "Stack top isn't array", "Η κορυφή του σωρού δεν είναι πίνακας"
End Sub

Public Sub StackTopNotGroup(a$)
MyErMacro a$, "Stack top isn't group", "Η κορυφή του σωρού δεν είναι ομάδα"
End Sub
Public Sub StackTopNotNumber(a$)
MyErMacro a$, "Stack top isn't number", "Η κορυφή του σωρού δεν είναι αριθμός"
End Sub
Public Sub NeedAnArray(a$)
MyErMacro a$, "Need an Array", "Χρειάζομαι ένα πίνακα"
End Sub
Public Sub noref()
MyEr "No with reference (&)", "Όχι με αναφορά (&)"
End Sub
Public Sub NoMoreDeep(deep As Variant)
MyEr "No more" + Str(deep) + " levels gosub allowed", "Δεν επιτρέπονται πάνω από" + Str(deep) + " επίπεδα για εντολή ΔΙΑΜΕΣΟΥ"
End Sub
Public Sub CantFind(w$)
MyEr "Can't find " + w$ + " or type name", "Δεν μπορώ να βρω το " + w$ + " ή όνομα τύπου"
End Sub
Public Sub OverflowLong(Optional b As Boolean = False)
If b Then
MyEr "OverFlow Integer", "Yπερχείλιση ακεραίου"
Else
MyEr "OverFlow Long", "Yπερχείλιση μακρύ"
End If
End Sub
Public Sub BadUseofReturn()
MyEr "Wrong Use of Return", "Κακή χρήση της επιστροφής"
End Sub
Public Sub DevZero()
    MyEr "division by zero", "διαίρεση με το μηδέν"
End Sub
Public Sub DevZeroMacro(aa$)
    MyErMacro aa$, "division by zero", "διαίρεση με το μηδέν"
End Sub
Public Sub ErrInExponet(a$)
MyErMacro a$, "Error in exponet", "Λάθος στον εκθέτη"
End Sub

Public Sub LambdaOnly(a$)
MyErMacro a$, "Only in lambda function", "Μόνο σε λάμδα συνάρτηση"
End Sub
Public Sub FilePathNotForUser()
MyEr "Filepath is not valid for user", "Ο τόπος του αρχείου δεν είναι έγκυρος για τον χρήστη"
End Sub

' used to isnumber
Public Sub MyErMacro(wher$, en$, gr$)
If stackshowonly Then
LastErNum = -2
wher$ = " : ERROR -2" & Sput(en$) + Sput(gr$) + wher$
Else
MyEr en$, gr$
End If
End Sub
Public Sub MyErMacroStr(wher$, en$, gr$)
If stackshowonly Then
LastErNum = -2
wher$ = " : ERROR -2" & Sput(en$) + Sput(gr$) + wher$
Else
MyEr en$, gr$
End If
End Sub
Public Sub ZeroParam(ar$)   ' we use MyErMacro in isNumber and isString
MyErMacro ar$, "Empty parameter", "Μηδενική παράμετρος"
End Sub
Public Sub MissPar()
MyEr "missing parameter", "λείπει παράμετρος"
End Sub
Public Sub MissModuleName()
MyEr "Missing module name", "Λείπει όνομα τμήματος"
End Sub
Public Sub nonext()
MyEr "NEXT without FOR", "ΕΠΟΜΕΝΟ χωρίς ΓΙΑ"
End Sub
Public Sub MissWhile()
MyEr "Missing the End While", "Έχασα το Τέλος Ενώ"
End Sub

Public Sub MissUntil()
MyEr "Missing the Until or Always", "Έχασα το Μέχρι ή το Πάντα"
End Sub

Public Sub MissNext()
MyEr "Missing the right NEXT", "Έχασα το σωστό ΕΠΟΜΕΝΟ"
End Sub
Public Sub MissVarName()
MyEr "Missing variable name", "Λείπει όνομα μεταβλητής"
End Sub
Public Sub MissParamref(ar$)
MyErMacro ar$, "missing by reference parameter", "λείπει με αναφορά παράμετρος"
End Sub
Public Sub MissParam(ar$)
MyErMacro ar$, "missing parameter", "λείπει παράμετρος"
End Sub
Public Sub MissFuncParameterStringVar()
MyEr "Not a string variable", "Δεν είναι αλφαριθμητική μεταβλητή"
End Sub
Public Sub MissFuncParameterStringVarMacro(ar$)
MyErMacro ar$, "Not a string variable", "Δεν είναι αλφαριθμητική μεταβλητή"
End Sub
Public Sub NoSuchFolder()
MyEr "No such folder", "Δεν υπάρχει τέτοιος φάκελος"
End Sub
Public Sub MissSymbol(wht$)
MyEr "missing " & wht$, "λείπει " & wht$
End Sub
Public Sub ClearSpace(nm$)
Dim i As Long
Do
    i = 1
    If FastOperator(nm$, vbCrLf, i, 2, False) Then
        SetNextLine nm$
    ElseIf FastOperator(nm$, "/", i) Then
        SetNextLine nm$
    ElseIf FastOperator(nm$, "\", i) Then
        SetNextLine nm$
    ElseIf FastOperator(nm$, "'", i) Then
        SetNextLine nm$
    Else
    Exit Do
    End If
Loop
End Sub
Public Function StringToEscapeStr(RHS As String, Optional Json As Boolean = False) As String
Dim i As Long, cursor As Long, ch As String
cursor = 0
Dim del As String
Dim H9F As String
For i = 1 To Len(RHS)
                ch = Mid$(RHS, i, 1)
                cursor = cursor + 1
                Select Case AscW(ch)
                    Case 92:        ch = "\\"
                   ' Case """":       ch = "\"""
                    Case 34
                    If Json Then
                        ch = "\"""
                    Else
                        ch = "\u0022"
                    End If
                    Case 10:       ch = "\n"
                    Case 13:       ch = "\r"
                    Case 9:      ch = "\t"
                    Case 8:     ch = "\b"
                    Case 12: ch = "\f"
                    Case 0 To 31, 127 To &H9F
                        ch = "\u" & Right$("000" & Hex$(AscW(ch)), 4)
                    Case Is > 255
                       If Json Then ch = "\u" & Right$("000" & Hex$(AscW(ch)), 4)
                End Select
                If cursor + Len(ch) > Len(StringToEscapeStr) Then StringToEscapeStr = StringToEscapeStr + space$(500)
                Mid$(StringToEscapeStr, cursor, Len(ch)) = ch
                cursor = cursor + Len(ch) - 1
Next
If cursor > 0 Then StringToEscapeStr = Left$(StringToEscapeStr, cursor)

End Function
Public Function EscapeStrToString(RHS As String) As String
Dim i As Long, cursor As Long, ch As String
     For cursor = 1 To Len(RHS)
        ch = Mid$(RHS, cursor, 1)
        i = i + 1
        Select Case ch
            Case """": GoTo ok1
            Case "\":
                cursor = cursor + 1
                ch = Mid$(RHS, cursor, 1)
                Select Case LCase$(ch) 'We'll make this forgiving though lowercase is proper.
                    Case "\", "/": ch = ch
                    Case """":      ch = """"
                    Case "a":       ch = Chr$(7)
                    Case "n":      ch = vbLf
                    Case "r":      ch = vbCr
                    Case "t":      ch = vbTab
                    Case "b":      ch = vbBack
                    Case "f":      ch = vbFormFeed
                    Case "u":      ch = ParseHexChar(RHS, cursor, Len(RHS))
                End Select
        End Select
                If i + Len(ch) > Len(EscapeStrToString) Then EscapeStrToString = EscapeStrToString + space$(500)
                Mid$(EscapeStrToString, i, Len(ch)) = ch
                i = i + Len(ch) - 1
    Next
ok1:
    If i > 0 Then EscapeStrToString = Left$(EscapeStrToString, i)
End Function

Private Function ParseHexChar( _
    ByRef Text As String, _
    ByRef cursor As Long, _
    ByVal LenOfText As Long) As String
    
    Const ASCW_OF_ZERO As Long = &H30&
    Dim Length As Long
    Dim ch As String
    Dim DigitValue As Long
    Dim Value As Long

    For cursor = cursor + 1 To LenOfText
        ch = Mid$(Text, cursor, 1)
        Select Case ch
            Case "0" To "9", "A" To "F", "a" To "f"
                Length = Length + 1
                If Length > 4 Then Exit For
                If ch > "9" Then
                    DigitValue = (AscW(ch) And &HF&) + 9
                Else
                    DigitValue = AscW(ch) - ASCW_OF_ZERO
                End If
                Value = Value * &H10& + DigitValue
            Case Else
                Exit For
        End Select
    Next
    If Length = 0 Then Err.Raise 5 'No hex digits at all.
    cursor = cursor - 1
    ParseHexChar = ChrW$(Value)
End Function

Public Function ReplaceSpace(a$) As String
Dim i As Long, j As Long
i = 1
Do
i = InStr(i, a$, "[")
If i > 0 Then
    i = i + 1
    j = InStr(i, a$, "]")
    If j > 0 Then
    j = j - i
    Mid$(a$, i, j) = Replace(Mid$(a$, i, j), " ", ChrW(160))
    i = i + j
    End If
Else
    Exit Do
End If
Loop
ReplaceSpace = a$
End Function
Function GetReturnArray(bstack As basetask, x1 As Long, b$, p As Variant, ss$, pppp As mArray) As Boolean ' true is error

Do
        Set bstack.lastobj = Nothing
        If IsExp(bstack, b$, p) Then
        If x1 = 0 Then If lookOne(b$, ",") Then x1 = 1: Set pppp = New mArray: pppp.PushDim (1): pppp.PushEnd
        If x1 = 0 Then
                If Len(bstack.originalname$) > 3 Then
                        If Mid$(bstack.originalname$, Len(bstack.originalname$) - 2, 1) = "$" Then
                            MissStringExpr
                            Exit Do
                        End If
                    End If
                 If Right$(bstack.originalname$, 3) = "%()" Then p = MyRound(p)
                 Set bstack.FuncObj = bstack.lastobj
                 Set bstack.lastobj = Nothing
                 bstack.FuncValue = p
        Else
                pppp.SerialItem 0, x1 + 1, 9
                If bstack.lastobj Is Nothing Then
                    pppp.item(x1 - 1) = p
                Else
                    Set pppp.item(x1 - 1) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                End If
                bstack.FuncValue = p
                x1 = x1 + 1
                             
        End If
        ElseIf Not bstack.lastobj Is Nothing Then
cont1:
        If x1 = 0 Then If lookOne(b$, ",") Then x1 = 1: Set pppp = New mArray: pppp.PushDim (1): pppp.PushEnd
        If x1 = 0 Then
            Set bstack.FuncObj = bstack.lastobj
            Set bstack.lastobj = Nothing
            bstack.FuncValue = vbNullString
        Else
                pppp.SerialItem 0, x1 + 1, 9
                If bstack.lastobj Is Nothing Then
                    pppp.item(x1 - 1) = p
                Else
                    Set pppp.item(x1 - 1) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                End If
                bstack.FuncValue = p
                x1 = x1 + 1
        End If
        ElseIf IsStrExp(bstack, b$, ss$, Len(bstack.tmpstr) = 0) Then
            If x1 = 0 Then If lookOne(b$, ",") Then x1 = 1: Set pppp = New mArray: pppp.PushDim (1): pppp.PushEnd
            If x1 = 0 Then
                If Len(bstack.originalname$) > 3 Then
                    If Mid$(bstack.originalname$, Len(bstack.originalname$) - 2, 1) <> "$" Then
                    
misum:                      MissNumExpr
                         GetReturnArray = True
                         Exit Function
                    End If
                Else
                    GoTo misum
                End If
                Set bstack.FuncObj = bstack.lastobj
                Set bstack.lastobj = Nothing
                bstack.FuncValue = ss$
            Else
                pppp.SerialItem 0, x1 + 1, 9
                If bstack.lastobj Is Nothing Then
                    pppp.item(x1 - 1) = ss$
                Else
                    Set pppp.item(x1 - 1) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                End If
                x1 = x1 + 1
                bstack.FuncValue = ss$
            End If
            ElseIf Not bstack.lastobj Is Nothing Then
           GoTo cont1
        End If
        Loop Until Not FastSymbol(b$, ",")
        If x1 > 0 Then
         pppp.SerialItem 0, x1, 9
         Set bstack.FuncObj = pppp
         Set pppp = New mArray
         Set bstack.lastobj = Nothing
         If VarType(bstack.FuncValue) = 5 Then
         bstack.FuncValue = 0
         Else
         bstack.FuncValue = vbNullString
         End If
        End If
        x1 = 0
End Function


Function MergeOperators(ByVal a$, ByVal b$) As String
If a$ = vbNullString Then MergeOperators = b$: Exit Function
If b$ = vbNullString Then MergeOperators = a$: Exit Function
If a$ = b$ Then MergeOperators = a$: Exit Function
Dim BR() As String, i As Long
If Len(a$) > Len(b$) Then
BR() = Split("[]" + b$ + "[]", "][")
For i = 1 To UBound(BR) - 1
If InStr(a$, "[" + BR(i) + "]") = 0 Then a$ = a$ + "[" + BR(i) + "]"
Next i
MergeOperators = a$
Else
BR() = Split("[]" + a$ + "[]", "][")
For i = 1 To UBound(BR) - 1
If InStr(b$, "[" + BR(i) + "]") = 0 Then b$ = b$ + "[" + BR(i) + "]"
Next i
MergeOperators = b$
End If
End Function
Public Sub GarbageFlush()
ReDim Trush(500) As VarItem
Dim i As Long
For i = 1 To 500
   Set Trush(i) = New VarItem
Next i
TrushCount = 500
End Sub
Public Sub GarbageFlush2()
ReDim Trush(500) As VarItem
Dim i As Long
For i = 1 To 500
   Set Trush(i) = New VarItem
Next i
TrushCount = 500
End Sub
Function PointPos(F$) As Long
Dim er As Long, er2 As Long
While FastSymbol(F$, Chr(34))
F$ = GetStrUntil(Chr(34), F$)
Wend
Dim i As Long, j As Long, oj As Long
If F$ = vbNullString Then
PointPos = 1
Else
er = 3
er2 = 3
For i = 1 To Len(F$)
er = er + 1
er2 = er2 + 1
Select Case Mid$(F$, i, 1)
Case "."
oj = j: j = i
Case "\", "/", ":", Is = Chr(34)
If er = 2 Then oj = 0: j = i - 2: Exit For
er2 = 1
oj = j: j = 0
If oj = 0 Then oj = i - 1: If oj < 0 Then oj = 0
Case " ", ChrW(160), vbTab
If j > 0 Then Exit For
If er2 = 2 Then oj = 0: j = i - 1: Exit For
er = 1
Case "|", "'"
j = i - 1
Exit For
Case Is > " "

If j > 0 Then oj = j Else oj = 0
Case Else
If oj <> 0 Then j = oj Else j = i
Exit For
End Select
Next i
If j = 0 Then
If oj = 0 Then
j = Len(F$) + 1
Else
j = oj
End If
End If
While Mid$(F$, j, i) = " "
j = j - 1
Wend
PointPos = j
End If
End Function
Public Function ExtractType(F$, Optional JJ As Long = 0, Optional simple As Boolean = True) As String
Dim i As Long, j As Long, D$
If FastSymbol(F$, Chr(34)) Then F$ = GetStrUntil(Chr(34), F$)
If F$ = vbNullString Then ExtractType = vbNullString: Exit Function
If simple Then
    j = SimplePointPos(F$)
ElseIf JJ > 0 Then
    j = JJ
Else
    j = PointPos(F$)
End If
D$ = F$ & " "
If j < Len(D$) Then
For i = j To Len(D$)
Select Case Mid$(D$, i, 1)
Case "/", "|", "'", " ", Is = Chr(34)
i = i + 1
Exit For
End Select
Next i
If (i - j - 2) < 1 Then
ExtractType = vbNullString
Else
ExtractType = mylcasefILE(Mid$(D$, j + 1, i - j - 2))
End If
Else
ExtractType = vbNullString
End If
End Function


Public Function CFname(a$, Optional TS As Variant, Optional createtime As Variant) As String
If Len(a$) > 2000 Then Exit Function
Dim b$
Dim mDir As New recDir
If Not IsMissing(createtime) Then
mDir.UseUTC = createtime <= 0
End If
Sleep 1
If a$ <> "" Then
On Error GoTo 1
b$ = mDir.Dir1(a$, GetCurDir)
If b$ = vbNullString Then b$ = mDir.Dir1(a$, mDir.GetLongName(App.Path))
If b$ <> "" Then
CFname = mylcasefILE(b$)
If Not IsMissing(TS) Then
If Not IsMissing(createtime) Then
If Abs(createtime) = 1 Then
TS = CDbl(mDir.lastTimeStamp2)
Else
TS = CDbl(mDir.lastTimeStamp)
End If
Else
TS = CDbl(mDir.lastTimeStamp)
End If
End If
End If
Exit Function
End If
1:
CFname = vbNullString
End Function

Public Function LONGNAME(sPath As String) As String
LONGNAME = ExtractPath(sPath, , True)
End Function
Public Function ExpEnvirStr(strInput) As String
Dim Result As Long
Dim strOutput As String
'' Two calls required, one to get expansion buffer length first then do expansion
strOutput = space$(1000)
Result = ExpandEnvironmentStrings(StrPtr(strInput), StrPtr(strOutput), Result)
strOutput = space$(Result)
Result = ExpandEnvironmentStrings(StrPtr(strInput), StrPtr(strOutput), Result)
ExpEnvirStr = StripTerminator(strOutput)
End Function

Public Function ExtractPath(ByVal F$, Optional Slash As Boolean = True, Optional existonly As Boolean = False) As String
If F$ = vbNullString Then Exit Function
Dim i As Long, j As Long, test$
test$ = F$ & " \/:": i = InStr(test$, " "): j = InStr(test$, "\")
If i < j Then j = InStr(test$, "/"): If i < j Then j = InStr(test$, ":"): If i < j Then Exit Function
If Right(F$, 1) = "\" Or Right(F$, 1) = "/" Then F$ = F$ & " a"
j = PointPos(F$)
If Mid$(F$, j, 1) = "." Then j = j - 1
If Len(F$) < j Then
If ExtractType(Mid$(F$, j) & "\.10", , False) = "10" Then j = j - 1 Else Exit Function
Else

End If

j = j - Len(ExtractNameOnly(F$))
If j <= 3 Then
If Mid$(F$, 2, 1) = ":" Then
If Slash Then
ExtractPath = mylcasefILE(Left$(F$, 2)) & "\"
Else
ExtractPath = mylcasefILE(Left$(F$, 2))
End If
Else
ExtractPath = vbNullString
End If
Else
If Slash Then
ExtractPath = mylcasefILE(Left$(F$, j))
Else
ExtractPath = mylcasefILE(Left$(F$, j - 1))
End If
End If

If existonly Then
ExtractPath = mylcasefILE(StripTerminator(GetLongName(ExpEnvirStr(ExtractPath))))
Else
ExtractPath = ExpEnvirStr(ExtractPath)
End If
Dim ccc() As String, c$
ccc() = Split(ExtractPath, "\..")
If UBound(ccc()) > LBound(ccc()) Then
c$ = vbNullString
For i = LBound(ccc()) To UBound(ccc()) - 1
If ccc(i) = vbNullString Then
c$ = ExtractPath(ExtractPath(c$, False))
Else
c$ = c$ & ExtractPath(ccc(i), True)
End If

Next i
If Left$(ccc(i), 1) = "\" Then
ExtractPath = c$ & Mid$(ccc(i), 2)
Else
ExtractPath = c$ & ccc(i)
End If
End If
End Function
Function SimplePointPos(s$)
Dim a As Long, b As Long, c As Long, D As Long
a = rinstr(s$, ".")
b = rinstr(s$, "\")
c = rinstr(s$, "/")
D = rinstr(s$, ":")
If b < a And c < a And D < a Then
    SimplePointPos = a
Else
    SimplePointPos = Len(s$)
End If

End Function
Public Function ExtractName(F$, Optional simple As Boolean = False) As String
Dim i As Long, j As Long, k$
If F$ = vbNullString Then Exit Function
If simple Then
j = SimplePointPos(F$)
Else
j = PointPos(F$)
If j > Len(F$) Then j = Len(F$)
End If
If Mid$(F$, j, 1) = "." Then
k$ = ExtractType(F$, j, simple)
Else
j = Len(F$)
End If
For i = j To 1 Step -1
Select Case Mid$(F$, i, 1)
Case Is < " ", "\", "/", ":"
Exit For
End Select
Next i
If k$ = vbNullString Then
If Mid$(F$, i + j - i, 1) = "." Then
ExtractName = mylcasefILE(Mid$(F$, i + 1, j - i - 1))
Else
ExtractName = mylcasefILE(Mid$(F$, i + 1, j - i))

End If
Else
ExtractName = mylcasefILE(Mid$(F$, i + 1, j - i)) + k$
End If

'ExtractName = mylcasefILE(Trim$(Mid$(f$, I + 1, j - I)))

End Function
Public Function ExtractNameOnly(ByVal F$, Optional simple As Boolean = False) As String
Dim i As Long, j As Long
If F$ = vbNullString Then Exit Function
If simple Then
j = SimplePointPos(F$)
Else
j = PointPos(F$)
If j > Len(F$) Then j = Len(F$)
End If
For i = j To 1 Step -1
Select Case Mid$(F$, i, 1)
Case Is < " ", "\", "/", ":"
Exit For
End Select
Next i
If Mid$(F$, i + j - i, 1) = "." Then
ExtractNameOnly = mylcasefILE(Mid$(F$, i + 1, j - i - 1))
Else
ExtractNameOnly = mylcasefILE(Mid$(F$, i + 1, j - i))
End If
End Function
Public Function GetCurDir(Optional AppPath As Boolean = False) As String
Dim a$, CD As String

If AppPath Then
CD = App.Path
AddDirSep CD
a$ = mylcasefILE(CD)
Else
AddDirSep mcd
a$ = mylcasefILE(mcd)

End If
'If Right$(a$, 1) <> "\" Then a$ = a$ & "\"
GetCurDir = a$
End Function
Sub MakeGroupPointer(bstack As basetask, v, Optional usethisname As String = "", Optional glob As Boolean)
Dim varv As New Group
    With varv
        .IamGlobal = v.IamGlobal
        .IamApointer = True
        .BeginFloat 2
        Set .Sorosref = v.soros
        If v.IamFloatGroup Then
        v.ToDelete = False
        Else
            If Len(usethisname) > 0 Then
                If glob Then
                    .IamGlobal = True
                Else
                    .lasthere = here$
                End If
                .GroupName = usethisname
            Else
                If Not .IamGlobal Then
                    .lasthere = here$
                End If
                If Len(v.GroupName) > 1 Then
                    .GroupName = Mid$(v.GroupName, 1, Len(v.GroupName) - 1)
                End If
            End If
        End If
    End With
     Set varv.LinkRef = v
Set bstack.lastpointer = varv
Set bstack.lastobj = varv
End Sub
Function PreparePointer(bstack As basetask) As Boolean
Dim a As Group, pppp As mArray
    If bstack.lastpointer Is Nothing Then
    
    Else
        Set a = bstack.lastpointer
        
            Set pppp = New mArray
            pppp.PushDim 1
            pppp.PushEnd
            pppp.Arr = True
            Set pppp.item(0) = a
            Set bstack.lastpointer = pppp
            PreparePointer = True
  
    End If
    
End Function
Function BoxGroupVar(aGroup As Variant) As mArray
            Dim bGroup As Group
            Set bGroup = aGroup
            Set BoxGroupVar = New mArray
            BoxGroupVar.PushDim 1
            BoxGroupVar.PushEnd
            BoxGroupVar.Arr = True
            bGroup.ToDelete = True
            Set BoxGroupVar.item(0) = aGroup
End Function

Function BoxGroupObj(aGroup As Object) As mArray
            Dim bGroup As Group
            Set bGroup = aGroup
            Set BoxGroupObj = New mArray
            BoxGroupObj.PushDim 1
            BoxGroupObj.PushEnd
            BoxGroupObj.Arr = True
            bGroup.ToDelete = True
            Set BoxGroupObj.item(0) = aGroup
End Function

Sub monitor(bstack As basetask, prive As basket, Lang As Long)
    Dim ss$, di As Object
    Set di = bstack.Owner
    Dim primarymonitor As Long
    primarymonitor = FindPrimary
    If Lang = 0 Then
        wwPlain2 bstack, prive, "Εξ ορισμού κωδικοσελίδα: " & GetACP, bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Φάκελος εφαρμογής", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, PathFromApp("m2000"), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Καταχώρηση gsb", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, myRegister("gsb"), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Φάκελος προσωρινών αρχείων", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, LONGNAME(strTemp), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Τρέχον φάκελος", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, mcd, bstack.Owner.Width, 1000, True
        If m_bInIDE Then
        wwPlain2 bstack, prive, "Όριο Αναδρομής για Συναρτήσεις " + CStr(stacksize \ 2948 - 1), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Όριο Αναδρομής Συναρτήσεων/Τμημάτων με την Κάλεσε " + CStr(stacksize \ 1772 - 1), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Όριο κλήσεων για Τμήματα " + CStr(stacksize \ 1254 - 1), bstack.Owner.Width, 1000, True
        Else
        wwPlain2 bstack, prive, "Όριο Αναδρομής για Συναρτήσεις " + CStr(stacksize \ 9832 - 1), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Όριο Αναδρομής Συναρτήσεων/Τμημάτων με την Κάλεσε " + CStr(stacksize \ 5864), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Όριο κλήσεων για Τμήματα  " + CStr(stacksize \ 5004), bstack.Owner.Width, 1000, True
        End If
        If OverideDec Then wwPlain2 bstack, prive, "Αλλαγή Τοπικού " + CStr(Clid), bstack.Owner.Width, 1000, True
        If UseIntDiv Then ss$ = "+DIV" Else ss$ = "-DIV"
        If priorityOr Then ss$ = ss$ + " +PRI" Else ss$ = ss$ + " -PRI"
        If Not mNoUseDec Then ss$ = ss$ + " -DEC" Else ss$ = ss$ + " +DEC"
        If mNoUseDec <> NoUseDec Then ss$ = ss$ + "(παράκαμψη)"
GoSub part2
        wwPlain2 bstack, prive, "Διακόπτες " + ss$, bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Περί διακοπτών: χρησιμοποίησε την εντολή Βοήθεια Διακόπτες", bstack.Owner.Width, 1000, True
        
        wwPlain2 bstack, prive, "Οθόνες:" + Str$(DisplayMonitorCount()) + "  η βασική :" + Str$(primarymonitor + 1) + " Εντολή: Παράθυρο Τύπος," + Str$(primarymonitor), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Αυτή η φόρμα είναι στην οθόνη:" + Str$(FindFormSScreen(di) + 1), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Η κονσόλα είναι στην οθόνη:" + Str$(Console + 1), bstack.Owner.Width, 1000, True

    Else
        wwPlain2 bstack, prive, "Default Code Page:" & GetACP, bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "App Path", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, PathFromApp("m2000"), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Register gsb", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, myRegister("gsb"), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Temporary", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, LONGNAME(strTemp), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Current directory", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, mcd, bstack.Owner.Width, 1000, True
        If m_bInIDE Then
        wwPlain2 bstack, prive, "Max Limit for Function Recursion " + CStr(stacksize \ 2948 - 1), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Max Limit for Function/Module Recursion using Call " + CStr(stacksize \ 1772 - 1), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Max Limit for calling modules in depth " + CStr(stacksize \ 1254 - 1), bstack.Owner.Width, 1000, True
        Else
        wwPlain2 bstack, prive, "Max Limit for Function Recursion " + CStr(stacksize \ 9832 - 1), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Max Limit for Function/Module Recursion using Call " + CStr(stacksize \ 5864), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Max Limit for calling modules in depth " + CStr(stacksize \ 5004), bstack.Owner.Width, 1000, True
        End If
        If OverideDec Then wwPlain2 bstack, prive, "Locale Overide " + CStr(Clid), bstack.Owner.Width, 1000, True
        If UseIntDiv Then ss$ = "+DIV" Else ss$ = "-DIV"
        If priorityOr Then ss$ = ss$ + " +PRI" Else ss$ = ss$ + " -PRI"
        If Not mNoUseDec Then ss$ = ss$ + " -DEC" Else ss$ = ss$ + " +DEC"
        If mNoUseDec <> NoUseDec Then ss$ = ss$ + "(bypass)"
        GoSub part2
        wwPlain2 bstack, prive, "Switches " + ss$, bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "About Switches: use command Help Switches", bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Screens:" + Str$(DisplayMonitorCount()) + "  Primary is:" + Str$(primarymonitor + 1) + " Command: Window Mode," + Str$(primarymonitor), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "This form is in screen:" + Str$(FindFormSScreen(di) + 1), bstack.Owner.Width, 1000, True
        wwPlain2 bstack, prive, "Console is in screen:" + Str$(Console + 1), bstack.Owner.Width, 1000, True
    End If
Exit Sub
part2:
        If mTextCompare Then ss$ = ss$ + " +TXT" Else ss$ = ss$ + " -TXT"
        If ForLikeBasic Then ss$ = ss$ + " +FOR" Else ss$ = ss$ + " -FOR"
        If DimLikeBasic Then ss$ = ss$ + " +DIM" Else ss$ = ss$ + " -DIM"
        If ShowBooleanAsString Then ss$ = ss$ + " +SBL" Else ss$ = ss$ + " -SBL"
        If wide Then ss$ = ss$ + " +EXT" Else ss$ = ss$ + " -EXT"
        If RoundDouble Then ss$ = ss$ + " +RDB" Else ss$ = ss$ + " -RDB"
        If SecureNames Then ss$ = ss$ + " +SEC" Else ss$ = ss$ + " -SEC"
        If UseTabInForm1Text1 Then ss$ = ss$ + " +TAB" Else ss$ = ss$ + " -TAB"
        If UseMDBHELP Then ss$ = ss$ + " +MDB" Else ss$ = ss$ + " -MDB"
        If Use13 Then ss$ = ss$ + " +INP" Else ss$ = ss$ + " -INP"
        If Nonbsp Then ss$ = ss$ + " -NBS" Else ss$ = ss$ + " +NBS"
        Return
End Sub
Sub NeoSwap(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MySwap(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoComm(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyRead(3, ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoRef(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyRead(2, ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoRead(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyRead(1, ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoReport(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyReport(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoDeclare(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyDeclare(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoMethod(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyMethod(ObjFromPtr(basestackLP), rest$, Lang)
If LastErNum = -1 Then resp = False
End Sub
Sub NeoWith(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyWith(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoSprite(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim s$, p, bstack As basetask
Set bstack = ObjFromPtr(basestackLP)
If IsExp(bstack, rest$, p) Then
spriteGDI bstack, rest$
ElseIf IsStrExp(bstack, rest$, s$, Len(bstack.tmpstr) = 0) Then
sprite bstack, s$, rest$
End If
resp = LastErNum1 = 0
End Sub

Sub NeoPlayer(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPlayer(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoPrinter(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPrinter(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoPage(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
ProcPage ObjFromPtr(basestackLP), rest$, Lang
resp = True
End Sub
Sub NeoCompact(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
BaseCompact ObjFromPtr(basestackLP), rest$
resp = True
End Sub
Sub NeoLayer(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcLayer(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoOrder(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyOrder(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoDelete(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = DELfields(ObjFromPtr(basestackLP), rest$)
'resp = True  '' maybe this can be change
End Sub
Sub NeoAppend(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim s$, p As Variant, bstack As basetask
Set bstack = ObjFromPtr(basestackLP)
If IsExp(bstack, rest$, p) Then
resp = AddInventory(bstack, rest$)
ElseIf IsStrExp(bstack, rest$, s$, Len(bstack.tmpstr) = 0) Then
resp = append_table(bstack, s$, rest$, False)
Else
SyntaxError
resp = False
End If
End Sub
Sub NeoSearch(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
getrow ObjFromPtr(basestackLP), rest$, , "", Lang
resp = True
End Sub
Sub NeoRetr(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
getrow ObjFromPtr(basestackLP), rest$, , , Lang
resp = True
End Sub
Sub NeoExecute(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
If IsLabelSymbolNew(rest$, "ΚΩΔΙΚΑ", "CODE", Lang) Then
 resp = ExecCode(ObjFromPtr(basestackLP), rest$)
 Else
CommExecAndTimeOut ObjFromPtr(basestackLP), rest$
resp = True
End If

End Sub

Sub NeoTable(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = NewTable(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoBase(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = NewBase(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoHold(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcHold(ObjFromPtr(basestackLP))
End Sub
Sub NeoRelease(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcRelease(ObjFromPtr(basestackLP))
End Sub
Sub NeoSuperClass(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcClass(ObjFromPtr(basestackLP), rest$, Lang, True)
End Sub
Sub NeoClass(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcClass(ObjFromPtr(basestackLP), rest$, Lang, False)
End Sub
Sub NeoDIM(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyDim(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoPathDraw(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPath(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoCreateEmf(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCreateEmf(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoDrawings(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyDrawings(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoFill(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcFill(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoFloodFill(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcFLOODFILL(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoTextCursor(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyCursor(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoMouseIcon(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
i3MouseIcon ObjFromPtr(basestackLP), rest$, Lang
resp = True
End Sub
Sub NeoDouble(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim bstack As basetask
Set bstack = ObjFromPtr(basestackLP)
SetDouble bstack.Owner
Set bstack = Nothing
resp = True
End Sub
Sub NeoNormal(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim bstack As basetask
Set bstack = ObjFromPtr(basestackLP)
SetNormal bstack.Owner
Set bstack = Nothing
resp = True
End Sub
Sub NeoSort(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcSort(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoImage(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcImage(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoBitmaps(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyBitmaps(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoDef(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDef(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoMovies(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyMovies(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoSounds(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MySounds(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoPen(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPen(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoCls(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCls(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoDesktop(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDesktop(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoStructure(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = myStructure(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoInput(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyInput(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoEvent(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = myEvent(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoProto(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcProto(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoEnum(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcEnum(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoPset(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyPset(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoModule(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyModule(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoModules(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyModules(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoGroup(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcGroup(0, ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoClipBoard(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyClipboard(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoBack(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
ProcBackGround ObjFromPtr(basestackLP), rest$, Lang, resp
End Sub
Sub NeoOver(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcOver(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoDrop(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDrop(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoShift(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcShift(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoShiftBack(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcShiftBack(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoLoad(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcLoad(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoText(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcText(ObjFromPtr(basestackLP), False, rest$)
End Sub
Sub NeoHtml(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcText(ObjFromPtr(basestackLP), True, rest$)
End Sub

Sub NeoCurve(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCurve(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoPoly(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPoly(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoCircle(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCircle(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoNew(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyNew(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoTitle(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcTitle(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoDraw(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDraw(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoWidth(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDrawWidth(ObjFromPtr(basestackLP), rest$)
End Sub

Sub NeoMove(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcMove(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoStep(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcStep(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoPrint(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = RevisionPrint(ObjFromPtr(basestackLP), rest$, 0, Lang)
End Sub
Sub NeoCopy(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyCopy(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoPrinthEX(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = RevisionPrint(ObjFromPtr(basestackLP), rest$, 1, Lang)
End Sub
Sub NeoRem(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
    Dim i As Long
    If FastSymbol(rest$, "{") Then
    i = blockLen(rest$)
    If i > 0 Then rest$ = Mid$(rest$, i + 1) Else rest$ = vbNullString
    Else
    SetNextLineNL rest$
    End If
    resp = True
End Sub
Sub NeoPush(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyPush(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoData(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyData(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoClear(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyClear(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoLinespace(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = procLineSpace(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoSet(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim i As Long, s$
aheadstatusANY rest$, i
s$ = Left$(rest$, i - 1)
resp = interpret(ObjFromPtr(basestackLP), s$)
If resp Then
rest$ = Mid$(rest$, i)
Else
rest$ = s$ + Mid$(rest$, i)
End If
End Sub


Sub NeoBold(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
ProcBold ObjFromPtr(basestackLP), rest$
resp = True
End Sub
Sub NeoChooseObj(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
    resp = ProcChooseObj(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoChooseFont(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
    ProcChooseFont ObjFromPtr(basestackLP), Lang
    resp = True
End Sub
Sub NeoFont(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
    ProcChooseFont ObjFromPtr(basestackLP), Lang
    resp = True
End Sub
Sub NeoScore(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyScore(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoPlayScore(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyPlayScore(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoMode(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcMode(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoGradient(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcGradient(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoFunction(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyFunction(0, ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoFiles(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcFiles(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoCat(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCat(ObjFromPtr(basestackLP), rest$, Lang)
End Sub


Function CallAsk(bstack As basetask, a$, v$) As Boolean
Dim s$
    If UCase(v$) = "ASK(" Then
        DialogSetupLang 1
    Else
        DialogSetupLang 0
    End If
    If AskText$ = vbNullString Then: ZeroParam a$: Exit Function
    If FastSymbol(a$, ",") Then IsStrExp bstack, a$, AskTitle$
    If FastSymbol(a$, ",") Then
        IsStrExp bstack, a$, s$ 'AskOk$
        If s$ = "" Then
        AskOk$ = ""
        ElseIf s$ = "*" Then
            AskOk$ = "*" + AskOk$
        Else
            AskOk$ = s$
        End If
    End If
    If FastSymbol(a$, ",") Then
        IsStrExp bstack, a$, s$ ' AskCancel$
        If s$ = "" Then
        AskCancel$ = ""
        ElseIf s$ = "*" Then
            AskCancel$ = "*" + AskCancel$
        Else
            AskCancel$ = s$
        End If
    End If
    If FastSymbol(a$, ",") Then IsStrExp bstack, a$, AskDIB$
    
    If FastSymbol(a$, ",") Then IsStrExp bstack, a$, AskStrInput$: AskInput = True

olamazi
CallAsk = True
End Function
Public Sub olamazi()
If Form4Loaded Then Exit Sub
If Form4.Visible Then
Form4.Visible = False
If Form1.Visible Then
   
   ' If Form2.Visible Then Form2.ZOrder
    If Form1.TEXT1.Visible Then
        Form1.TEXT1.SetFocus
    Else
        Form1.SetFocus
    End If
    End If
    End If
End Sub
Sub GetGuiM2000(R$)
Dim aaa As GuiM2000
If TypeOf Screen.ActiveForm Is GuiM2000 Then
Set aaa = Screen.ActiveForm
                  If aaa.Index > -1 Then
                  R$ = myUcase(aaa.MyName$ + "(" + CStr(aaa.Index) + ")", True)
                  Else
                  R$ = myUcase(aaa.MyName$, True)
                  End If
Else
                R$ = vbNullString
End If

End Sub
Public Function IsSupervisor() As Boolean

Dim ss$
                 ss$ = UCase(userfiles)
                    DropLeft "\M2000_USER\", ss$
IsSupervisor = ss$ = vbNullString
End Function


Public Function UserPath() As String

Dim ss$
                 ss$ = UCase(userfiles)
                    DropLeft "\M2000_USER\", ss$
        If ss$ <> "" Then
        If CanKillFile(mcd) Then
        DropLeft "\", ss$
UserPath = Mid$(mcd, Len(userfiles) - Len(ss$) + 1)
If UserPath = vbNullString Then
UserPath = "."
End If
Else
UserPath = mcd
End If
Else
UserPath = mcd
End If
End Function
Public Function UserPath2() As String

Dim ss$
                 ss$ = UCase(userfiles)
                    DropLeft "\M2000_USER\", ss$
        If ss$ <> "" Then
        If CanKillFile(mcd) Then
        DropLeft "\", ss$
UserPath2 = Mid$(mcd, Len(userfiles) - Len(ss$) + 1)
If UserPath2 = vbNullString Then
UserPath2 = "."
End If
Else
UserPath2 = mcd
End If
Else
UserPath2 = mcd
End If
If Right$(UserPath2, 1) = "\" Then UserPath2 = Left$(UserPath2$, Len(UserPath2$) - 1)


End Function
Function Fast2Label(a$, c$, cl As Long, D$, dl As Long, ahead&) As Boolean
Dim i As Long, Pad$, j As Long
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function
Pad$ = myUcase(Mid$(a$, i, ahead& + 1)) + " "
If j - i >= cl - 1 Then
If InStr(c$, Left$(Pad$, cl)) > 0 Then
If Mid$(Pad$, cl + 1, 1) Like "[0-9+.\( @-]" Then
a$ = Mid$(a$, MyTrimLi(a$, i + cl))
Fast2Label = True
End If
Exit Function
End If
End If
If j - i >= dl - 1 Then
If InStr(D$, Left$(Pad$, dl)) > 0 Then
If Mid$(Pad$, dl + 1, 1) Like "[0-9+.\( @-]" Then
a$ = Mid$(a$, MyTrimLi(a$, i + dl))
Fast2Label = True
End If
End If
End If
End Function
Function Fast2Symbol(a$, c$, k As Long, D$, l As Long) As Boolean
Dim i As Long, j As Long
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function
If j - i >= k - 1 Then
    If InStr(c$, Mid$(a$, i, k)) > 0 Then
    a$ = Mid$(a$, MyTrimLi(a$, i + k))
    Fast2Symbol = True
    Exit Function
    End If
End If
'If j - i >= Len(d$) - 1 Then
If j - i >= l - 1 Then
    If InStr(D$, Mid$(a$, i, l)) > 0 Then
    a$ = Mid$(a$, MyTrimLi(a$, i + l))
    Fast2Symbol = True
    Exit Function
    End If

End If
End Function
Function FastOperator2(a$, c$, i As Long) As Boolean
If Mid$(a$, i, 1) = c$ Then
Mid$(a$, i, 1) = " "
FastOperator2 = True
End If
End Function



Function FastSymbol(a$, c$, Optional mis As Boolean = False, Optional cl As Long = 1) As Boolean
Dim i As Long, j As Long
'If Len(c$) <> cl Then Stop  ; only for check
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function  ' this is not good
If j - i < cl - 1 Then
If mis Then MyEr "missing " & c$, "λείπει " & c$
Exit Function
End If
If c$ = Mid$(a$, i, cl) Then
'If InStr(c$, Mid$(a$, i, cl)) > 0 Then
a$ = Mid$(a$, MyTrimLi(a$, i + cl))
'Mid$(a$, i, cl) = Space$(cl)
FastSymbol = True
ElseIf mis Then
MyEr "missing " & c$, "λείπει " & c$
End If
End Function
Function FastSymbol1(s$, c$) As Boolean
Dim i&, l As Long, where As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s$): If l = 0 Then Exit Function
  p2 = StrPtr(s$): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 7, 9
    Case Else
    
     where = (i - p2) \ 2 + 1
     
     FastSymbol1 = p1 = AscW(c$)
     
   Exit For
  End Select
  Next i
  If FastSymbol1 Then
  s$ = Mid$(s$, where + 1)
  ElseIf where > 0 Then
  s$ = Mid$(s$, where)
  End If
  
End Function

Function FastSymbol2(s$, c$) As Boolean
Dim i&, l As Long, where As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s$): If l = 0 Then Exit Function
  p2 = StrPtr(s$): l = l - 1
  p4 = p2 + l * 2
  For i = p2 To p4 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 7, 9
    Case Else
    
     where = (i - p2) \ 2 + 1
     FastSymbol2 = Mid$(s$, where, 2) = c$
     
   Exit For
  End Select
  Next i
  If FastSymbol2 Then
  s$ = Mid$(s$, where + 2)
  ElseIf where > 1 Then
  s$ = Mid$(s$, where)
  End If
  
End Function





Function lookA123(s) As Boolean
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then Exit Function
  p2 = StrPtr(s): l = l - 1
  GetMem2 p2, p1
  If p1 <> 58 Then Exit Function
  For i = p2 + 2 To p2 + l * 2 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 7, 9
    Case 13, 39, 47, 92
    lookA123 = True
    Case Else
   Exit Function
  End Select
  Next i
End Function

Function lookB123(s) As Boolean
Dim i&, l As Long
Dim p2 As Long, p1 As Integer, p4 As Long
  l = Len(s): If l = 0 Then Exit Function
  p2 = StrPtr(s): l = l - 1

  For i = p2 To p2 + l * 2 Step 2
  GetMem2 i, p1
  Select Case p1
    Case 32, 160, 7, 9
    Case 13, 47, 39, 92
    lookB123 = True
    Case Else
   Exit Function
  End Select
  Next i
End Function


Function IsLabelSymbolNewExp(a$, gre$, Eng$, code As Long, usethis$) As Boolean
' code 2  gre or eng, set new value to code 1 or 0
' 0 for gre
' 1 for eng
' return true if we have label
If Len(usethis$) = 0 Then
Dim what As Boolean
Select Case code
Case 0
IsLabelSymbolNewExp = IsLabelSymbol3(1032, a$, gre$, usethis$, False, False, False, True)
Case 1
IsLabelSymbolNewExp = IsLabelSymbol3(1033, a$, Eng$, usethis$, False, False, False, True)
Case 2
what = IsLabelSymbol3(1032, a$, gre$, usethis$, False, False, False, True)
If what Then
code = 0
IsLabelSymbolNewExp = what
Exit Function
End If
what = IsLabelSymbol3(1033, a$, Eng$, usethis$, False, False, False, True)
If what Then code = 1
IsLabelSymbolNewExp = what
End Select
Else
Select Case code
Case 0, 2
IsLabelSymbolNewExp = gre$ = usethis$
Case 1
IsLabelSymbolNewExp = Eng$ = usethis$
End Select
If IsLabelSymbolNewExp Then a$ = Mid$(a$, MyTrimL(a$) + Len(usethis$))
End If
If IsLabelSymbolNewExp Then
usethis$ = vbNullString
End If
End Function


Function IsLabelSymbol(a$, c$, Optional mis As Boolean = False, Optional ByVal ByPass As Boolean = False, Optional checkonly As Boolean = False) As Boolean
Dim test$, what$, Pass As Long
If ByPass Then Exit Function

  If a$ <> "" And c$ <> "" Then
test$ = a$
Pass = Len(c$)

IsLabelSymbol = IsLabelSYMB33(test$, what$, Pass)
If Len(what$) <> Len(c$) Then IsLabelSymbol = False
If Not IsLabelSymbol Then
     If mis Then
                 MyEr "missing " & c$, "λείπει " & c$
              End If
Exit Function
End If

        If myUcase(what$) = c$ Then
        If checkonly Then
     '   A$ = what$ & " " & TEST$
        Else
                    a$ = Mid$(test$, Pass)
          End If
  
             Else
             If mis Then
                 MyEr "missing " & c$, "λείπει " & c$
              End If
            IsLabelSymbol = False
            End If

End If
End Function
Function MakeEmf(bstack As basetask, b$, Lang As Long, Data$, Optional ww = 0, Optional hh = 0) As Boolean
Dim w$, x1 As Long, label1$, usehandler As mHandler, par As Boolean, pppp As mArray, p As Variant, it&, x2&
x2 = Len(b$)
If IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
            w$ = Funcweak(bstack, b$, x1, label1$)
            If LastErNum1 = -1 And x1 < 5 Then Exit Function
            If LenB(w$) = 0 Then
            If Len(bstack.UseGroupname) > 0 Then
                If Len(label1$) > Len(bstack.UseGroupname) Then
                    If bstack.UseGroupname = Left$(label1$, Len(bstack.UseGroupname)) Then
                        MyEr "No such member in this group", "Δεν υπάρχει τέτοιο μέλος σε αυτή την ομάδα"
                        Exit Function
                    End If
                End If
            ElseIf x1 = 1 Then
contvar1:
            x1 = globalvar(label1$, 0#)
            Set usehandler = New mHandler
                usehandler.t1 = 2

                Set usehandler.objref = ExecuteEmfBlock(bstack, Data$, it, ww, hh)


                    Set var(x1) = usehandler
                    MakeEmf = it <> 0
                    If it = 0 Then b$ = Data$ + space$(x2&)
                Exit Function
            ElseIf x1 = 5 Then
                If GetVar(bstack, label1$, x1) Then
                    If GetArrayReference(bstack, b$, label1$, var(x1), pppp, x1) Then
                        Set usehandler = New mHandler
                        usehandler.t1 = 2

                        Set usehandler.objref = ExecuteEmfBlock(bstack, Data$, it, ww, hh)
                        Set pppp.item(x1) = usehandler
                        MakeEmf = it <> 0
                        If it = 0 Then b$ = Data$ + space$(x2&) + b$
                    End If
                    Exit Function
            
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
            End If
            End If

            If x1 = 1 Then
            If GetVar(bstack, label1$, x1) Then
                Set usehandler = New mHandler
                usehandler.t1 = 2
        
                Set usehandler.objref = ExecuteEmfBlock(bstack, Data$, it, ww, hh)
          
                    Set var(x1) = usehandler
                    MakeEmf = it <> 0
                    If it = 0 Then b$ = Data$ + space$(x2&) + b$
            
                Exit Function
            Else
                GoTo contvar1
            End If
                ElseIf x1 = 5 Then
                If GetVar(bstack, label1$, x1) Then
                      DropLeft "(", w$
                    If GetArrayReference(bstack, w$, label1$, var(x1), pppp, x1) Then
                        Set usehandler = New mHandler
                        usehandler.t1 = 2
                        Set usehandler.objref = ExecuteEmfBlock(bstack, Data$, it, ww, hh)
                       
                        Set pppp.item(x1) = usehandler
                        MakeEmf = it <> 0
                        If it = 0 Then b$ = Data$ + space$(x2&) + b$
                    End If
                    Exit Function
            
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
            End If
        
        End If
End Function


Function GetRes(bstack As basetask, b$, Lang As Long, Data$) As Boolean
Dim w$, x1 As Long, label1$, usehandler As mHandler, par As Boolean, pppp As mArray, p As Variant
If IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
            w$ = Funcweak(bstack, b$, x1, label1$)
            If LastErNum1 = -1 And x1 < 5 Then Exit Function
            If LenB(w$) = 0 Then
            If Len(bstack.UseGroupname) > 0 Then
                If Len(label1$) > Len(bstack.UseGroupname) Then
                    If bstack.UseGroupname = Left$(label1$, Len(bstack.UseGroupname)) Then
                        MyEr "No such member in this group", "Δεν υπάρχει τέτοιο μέλος σε αυτή την ομάδα"
                        Exit Function
                    End If
                End If
            ElseIf x1 = 1 Then
contvar1:
            x1 = globalvar(label1$, 0#)
            Set usehandler = New mHandler
                usehandler.t1 = 2
        If FastSymbol(b$, ",") Then
        If IsExp(bstack, b$, p, , True) Then
         Set usehandler.objref = Decode64toMemBloc(Data$, par, CBool(p))
        Else
        GetRes = True
        MissParam Data$: Exit Function
        End If
        Else
                Set usehandler.objref = Decode64toMemBloc(Data$, par)
                End If
                If par Then
                    Set var(x1) = usehandler
                    GetRes = True
            
                Else
                    GoTo Err1
                End If
                Exit Function
            ElseIf x1 = 3 Then
                x1 = globalvar(label1$, vbNullString)
                var(x1) = Decode64(Data$, par)
                If Not par Then GoTo Err1
                GetRes = True
                Exit Function
            ElseIf x1 = 5 Then
                If GetVar(bstack, label1$, x1) Then
                    If GetArrayReference(bstack, b$, label1$, var(x1), pppp, x1) Then
                        Set usehandler = New mHandler
                        usehandler.t1 = 2
                        If Not par Then GoTo Err1
                        If FastSymbol(b$, ",") Then
        If IsExp(bstack, b$, p, , True) Then
         Set usehandler.objref = Decode64toMemBloc(Data$, par, CBool(p))
        Else
        GetRes = True
        MissParam Data$: Exit Function
        End If
        Else
                        Set usehandler.objref = Decode64toMemBloc(Data$, par)
                        End If
                    
                        Set pppp.item(x1) = usehandler
                        GetRes = True
                    End If
                    Exit Function
            
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
            ElseIf x1 = 6 Then
contstr1:
                If GetVar(bstack, label1$, x1) Then
                    If GetArrayReference(bstack, b$, label1$, var(x1), pppp, x1) Then
                        pppp.item(x1) = Decode64(Data$, par)
                        If Not par Then GoTo Err1
                        GetRes = True
                    End If
                    Exit Function
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
            End If
            End If

            If x1 = 1 Then
            If GetVar(bstack, label1$, x1) Then
            Set usehandler = New mHandler
                usehandler.t1 = 2
        
                Set usehandler.objref = Decode64toMemBloc(Data$, par)
                If par Then
                    Set var(x1) = usehandler
                    GetRes = True
            
                Else
Err1:
                    MyEr "Can't decode this resource", "Δεν μπορών να αποκωδικοποιήσω αυτό το πόρο"
                End If
                Exit Function
            Else
                GoTo contvar1
            End If
                ElseIf x1 = 3 Then
                
                If GetVar(bstack, label1$, x1) Then
                var(x1) = Decode64(Data$, par)
                If Not par Then GoTo Err1
                GetRes = True
                Exit Function
                End If
                ElseIf x1 = 5 Then
                If GetVar(bstack, label1$, x1) Then
                      DropLeft "(", w$
                    If GetArrayReference(bstack, w$, label1$, var(x1), pppp, x1) Then
                        Set usehandler = New mHandler
                        usehandler.t1 = 2
                        Set usehandler.objref = Decode64toMemBloc(Data$, par)
                        If Not par Then GoTo Err1
                        Set pppp.item(x1) = usehandler
                        GetRes = True
                    End If
                    Exit Function
            
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
                            ElseIf x1 = 6 Then
                               If GetVar(bstack, label1$, x1) Then
                            DropLeft "(", w$
                    If GetArrayReference(bstack, w$, label1$, var(x1), pppp, x1) Then
                        pppp.item(x1) = Decode64(Data$, par)
                        If Not par Then GoTo Err1
                        GetRes = True
                    End If
                    Exit Function
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
            End If
        
        End If
End Function

Function IsHILOWWORD(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
    Dim p As Variant
    If IsExp(bstack, a$, R, , True) Then
        If FastSymbol(a$, ",") Then
              If IsExp(bstack, a$, p, , True) Then
                    R = SG * (R * &H10000 + p)
                    
                     IsHILOWWORD = FastSymbol(a$, ")", True)
                  Else
                     
                    MissParam a$
                End If
        Else
             
             MissParam a$
        End If
     Else
             
             MissParam a$
      End If
     
End Function
Function IsBinaryNot(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
  If IsExp(bstack, a$, R, , True) Then
            On Error Resume Next
    If R < 0 Then R = R And &H7FFFFFFF
             R = SG * uintnew1(Not signlong2(R))
        If Err.Number > 0 Then
            
            WrongArgument a$
          
            Exit Function
            End If
    On Error GoTo 0
    
        IsBinaryNot = FastSymbol(a$, ")", True)
    Else
           MissParam a$
    
    End If
End Function
Function IsBinaryNeg(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
  If IsExp(bstack, a$, R, , True) Then
            On Error Resume Next
       
             R = SG * CDbl(Pow2minusOne(32) - uintnew(R))
        If Err.Number > 0 Then
        
            WrongArgument a$
        
            Exit Function
            End If
    On Error GoTo 0
    
        IsBinaryNeg = FastSymbol(a$, ")", True)
    Else
           MissParam a$
    
    End If
End Function
Function IsBinaryOr(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
        Dim p As Variant
     If IsExp(bstack, a$, R, , True) Then
        If FastSymbol(a$, ",") Then
        If IsExp(bstack, a$, p, , True) Then
            R = SG * uintnew1(signlong2(R) Or signlong2(p))
         IsBinaryOr = FastSymbol(a$, ")", True)
           Else
                
                MissParam a$
        End If
          Else
                MissParam a$
       End If
         Else
                MissParam a$
       End If
End Function
Function IsBinaryAdd(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
    Dim p As Variant
    On Error Resume Next
    If IsExp(bstack, a$, R, , True) Then
            If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p, , True) Then
                    R = add32(R, p)
                    If Err.Number Then R = add32(Int(R / 4294967296@), p): Err.Clear
                    While FastSymbol(a$, ",")
                    If Not IsExp(bstack, a$, p, , True) Then MissNumExpr: Exit Function
                    R = add32(R, p)
                    If Err.Number Then R = add32(R, Int(p / 4294967296@)): Err.Clear
                    Wend
                    If SG < 0 Then R = -R
                    IsBinaryAdd = FastSymbol(a$, ")", True)
                Else
                    
                    MissParam a$
                End If
            Else
                
                MissParam a$
            End If
        Else
            
            MissParam a$
       
       End If
End Function
Function IsBinaryAnd(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
    Dim p As Variant
    If IsExp(bstack, a$, R, , True) Then
            If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p, , True) Then
                    R = SG * uintnew1(signlong2(R) And signlong2(p))
                    
                    IsBinaryAnd = FastSymbol(a$, ")", True)
                Else
                    
                    MissParam a$
                End If
            Else
                
                MissParam a$
            End If
        Else
            
            MissParam a$
       
       End If
End Function
Function IsBinaryXor(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
    Dim p As Variant
        If IsExp(bstack, a$, R, , True) Then
            If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p, , True) Then
                    R = SG * uintnew1(signlong2(R) Xor signlong2(p))
                    
                    IsBinaryXor = FastSymbol(a$, ")", True)
                Else
                    
                    MissParam a$
                End If
            Else
                
                MissParam a$
            End If
        Else
            
            MissParam a$
       
       End If
End Function
Function IsBinaryShift(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim p As Variant
   If IsExp(bstack, a$, R, , True) Then
  
            If FastSymbol(a$, ",") Then
                    If IsExp(bstack, a$, p, , True) Then
                         If p > 31 Or p < -31 Then
                         
                         MyErMacro a$, "Shift from -31 to 31", "Ολίσθηση από -31 ως 31"
                         IsBinaryShift = False: Exit Function
                         Else
                               If p > 0 Then
                              
                                 R = SG * CCur((signlong(R) And signlong(Pow2minusOne(32 - p))) * Pow2(p))
                              
                              ElseIf p = 0 Then
                              If SG < 0 Then R = -CCur(R) Else R = CCur(R)
                              Else
                                    
                                 R = SG * CCur(Int(CCur(R) / Pow2(-p)))
                              End If
                              
                            IsBinaryShift = FastSymbol(a$, ")", True)
                    Exit Function
                         End If
                    Else
                          
                        MissParam a$
                    End If
            Else
                
                MissParam a$
            End If
    Else
            
            MissParam a$
   End If

End Function
Function IsBinaryRotate(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim p As Variant
        If IsExp(bstack, a$, R, , True) Then
             If FastSymbol(a$, ",") Then
                 If IsExp(bstack, a$, p, , True) Then
                        If p > 31 Or p < -31 Then
                            
                              MyErMacro a$, "Rotation from -31 to 31", "Περιστοφή από -31 ως 31"
                             IsBinaryRotate = False: Exit Function
                        Else
                             If p > 0 Then
                          
                                 R = SG * CCur((signlong(R) And signlong(Pow2minusOne(32 - p))) * Pow2(p) + Int(CCur(R) / Pow2(32 - p)))
                             ElseIf p = 0 Then
                                 If SG < 0 Then R = -CCur(R) Else R = CCur(R)
                             Else
                          
                                 R = SG * CCur((signlong(R) And signlong(Pow2minusOne(-p))) * Pow2(32 + p) + Int(CCur(R) / Pow2(-p)))
                             End If
                        End If
                     
                  Else
                    
                    MissParam a$
                 End If
             Else
                
                MissParam a$
            End If
        IsBinaryRotate = FastSymbol(a$, ")", True)
        Else
            
            MissParam a$
        End If
End Function
Function IsSin(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
   If IsExp(bstack, a$, R, , True) Then
    R = Sin(R * 1.74532925199433E-02)
    ''r = Sgn(r) * Int(Abs(r) * 10000000000000#) / 10000000000000#
    If Abs(R) < 1E-16 Then R = 0
    If SG < 0 Then R = -R
    
    
 IsSin = FastSymbol(a$, ")", True)
    Else
                
                MissParam a$
    
    End If
End Function
Function IsAbs(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
If IsExp(bstack, a$, R, , True) Then
    R = Abs(R)
    If SG < 0 Then R = -R
    
 IsAbs = FastSymbol(a$, ")", True)
    Else
                MissParam a$
    End If
End Function

Function IsCos(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
  If IsExp(bstack, a$, R, , True) Then

    R = Cos(R * 1.74532925199433E-02)
 
    If Abs(R) < 1E-16 Then R = 0
    If SG < 0 Then R = -R
    
    
  IsCos = FastSymbol(a$, ")", True)
    Else
                
                MissParam a$
    
    End If
End Function
Function IsTan(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
If IsExp(bstack, a$, R, , True) Then
     
     If R = Int(R) Then
        If R Mod 90 = 0 And R Mod 180 <> 0 Then
        MyErMacro a$, "Wrong Tan Parameter", "Λάθος παράμετρος εφαπτομένης"
        IsTan = False: Exit Function
        End If
        End If
    R = Sgn(R) * Tan(R * 1.74532925199433E-02)

     If Abs(R) < 1E-16 Then R = 0
     If Abs(R) < 1 And Abs(R) + 0.0000000000001 >= 1 Then R = Sgn(R)
   If SG < 0 Then R = -R
    
IsTan = FastSymbol(a$, ")", True)
     Else
                
                MissParam a$
    
    End If
End Function
Function IsAtan(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
 If IsExp(bstack, a$, R, , True) Then
     
     R = SG * Atn(R) * 180# / Pi
        
IsAtan = FastSymbol(a$, ")", True)
     Else
                
                MissParam a$
    
    End If
End Function
Function IsLn(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
  If IsExp(bstack, a$, R, , True) Then
    If R <= 0 Then
       MyErMacro a$, "Only > zero parameter", "Μόνο >0 παράμετρος"
        IsLn = False: Exit Function
    Else
    R = SG * Log(R)
    
    End If
    
 IsLn = FastSymbol(a$, ")", True)
     Else
                
                MissParam a$
    
    End If
End Function
Function IsLog(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
If IsExp(bstack, a$, R, , True) Then
        If R <= 0 Then
       MyErMacro a$, "Only > zero parameter", "Μόνο >0 παράμετρος"
        IsLog = False: Exit Function
    Else
    R = SG * Log(R) / 2.30258509299405
    
    End If
   IsLog = FastSymbol(a$, ")", True)
    Else
                
                MissParam a$
    
    End If
End Function
Function IsFreq(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim p As Variant
    If IsExp(bstack, a$, R, , True) Then
           If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p, , True) Then
                    R = SG * GetFrequency(CInt(R), CInt(p))
                    
                    IsFreq = FastSymbol(a$, ")", True)
                    Else
                
                MissParam a$
                End If
            Else
                
                MissParam a$
            End If
     Else
                
                MissParam a$
     End If
End Function
Function IsSqrt(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
    If IsExp(bstack, a$, R, , True) Then
    
    If R < 0 Then
    negsqrt a$
    Exit Function
   
    End If
  
    R = Sqr(R)
    If SG < 0 Then R = -R
    
   IsSqrt = FastSymbol(a$, ")", True)
    Else
                
                MissParam a$
    
    End If
End Function
Function GiveForm() As Form
If Form1.Visible Then
Set GiveForm = Form1
Else
Set GiveForm = Form3
End If
End Function
Function IsNumberD(a$, D As Double) As Boolean
Dim a1 As Long
If a$ <> "" Then
For a1 = 1 To Len(a$) + 1
Select Case Mid$(a$, a1, 1)
Case " ", ",", ChrW(160)
If a1 > 1 Then Exit For
Case Is = Chr(2)
If a1 = 1 Then Exit Function
Exit For
End Select
Next a1
If a1 > Len(a$) Then a1 = Len(a$) + 1
D = CDbl(val("0" & Left$(a$, a1 - 1)))
a$ = Mid$(a$, a1)
IsNumberD = True
Else
IsNumberD = False
End If
End Function
Function IsNumberLabel2(a$, Label$, a1 As Long, ByVal LI As Long) As Boolean
Dim A2 As Long
If LI > 0 Then
A2 = a1
If a1 > LI Then Exit Function
If LI > 5 + A2 Then LI = 4 + A2
If Mid$(a$, a1, 1) Like "[0-9]" Then
Do While a1 <= LI
a1 = a1 + 1
If Not Mid$(a$, a1, 1) Like "[0-9]" Then Exit Do

Loop
Label$ = Mid$(a$, A2, a1 - A2)
IsNumberLabel2 = True
End If
End If
End Function
Function IsNumberLabel(a$, Label$) As Boolean
Dim a1 As Long, LI As Long, A2 As Long
LI = Len(a$)

If LI > 0 Then

a1 = MyTrimL(a$)

A2 = a1
If a1 > LI Then a$ = vbNullString: Exit Function
If LI > 5 + A2 Then LI = 4 + A2
If Mid$(a$, a1, 1) Like "[0-9]" Then
Do While a1 <= LI
a1 = a1 + 1
If Not Mid$(a$, a1, 1) Like "[0-9]" Then Exit Do

Loop
Label$ = Mid$(a$, A2, a1 - A2): a$ = Mid$(a$, a1)
IsNumberLabel = True
End If

End If
End Function
Function IsNumberQuery(a$, fr As Long, R As Double, lr As Long) As Boolean
Dim SG As Long, sng As Long, n$, ig$, DE$, sg1 As Long, ex$, rr As Double
' ti kanei to e$
If a$ = vbNullString Then IsNumberQuery = False: Exit Function
SG = 1
sng = fr - 1
    Do While sng < Len(a$)
    sng = sng + 1
    Select Case Mid$(a$, sng, 1)
    Case " ", "+", ChrW(160)
    Case "-"
    SG = -SG
    Case Else
    Exit Do
    End Select
    Loop
n$ = Mid$(a$, sng)

If val("0" & Mid$(a$, sng, 1)) = 0 And Left(Mid$(a$, sng, 1), sng) <> "0" And Left(Mid$(a$, sng, 1), sng) <> "." Then
IsNumberQuery = False

Else
'compute ig$
    If Mid$(a$, sng, 1) = "." Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(a$)
        
        Select Case Mid$(a$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, sng, 1)
        Case "."
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    ' compute decimal part
    If DE$ <> "" Then
      sng = sng + 1
        Do While sng <= Len(a$)
       
        Select Case Mid$(a$, sng, 1)
        Case " ", ChrW(160), vbTab
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        DE$ = DE$ & Mid$(a$, sng, 1)
        End If
        Case "E", "e", "Ε", "ε" ' ************check it
             If ex$ = vbNullString Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
         ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng = sng + 1
        Loop
        If sg1 Then
            If Len(ex$) < 3 Then
                If ex$ = "E" Then
                    ex$ = " "
                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                    ex$ = "  "
                End If
            End If
        End If
    End If
    If ig$ = vbNullString Then
    IsNumberQuery = False
    lr = 1
    Else
    If SG < 0 Then ig$ = "-" & ig$
    Err.Clear
    On Error Resume Next
    n$ = ig$ & DE$ & ex$
    sng = Len(ig$ & DE$ & ex$)
    rr = val(ig$ & DE$ & ex$)
    If Err.Number > 0 Then
         lr = 0
    Else
        R = rr
       lr = sng - fr + 2
       IsNumberQuery = True
    End If
    
       
    
    End If
End If
End Function


Function IsNumberOnly(a$, fr As Long, R As Variant, lr As Long, Optional useRtypeOnly As Boolean = False, Optional usespecial As Boolean = False) As Boolean
Dim SG As Long, sng As Long, ig$, DE$, sg1 As Long, ex$, foundsign As Boolean
' ti kanei to e$
If a$ = vbNullString Then IsNumberOnly = False: Exit Function
SG = 1
sng = fr - 1
    Do While sng < Len(a$)
    sng = sng + 1
    Select Case Mid$(a$, sng, 1)
    Case " ", ChrW(160), vbTab
    Case "+"
    foundsign = True
    Case "-"
    SG = -SG
    foundsign = True
    Case Else
    Exit Do
    End Select
    Loop
If LCase(Mid$(a$, sng, 2)) Like "0[xχ]" Then
    If foundsign Then
    MyEr "no sign for hex values", "όχι πρόσημο για δεκαεξαδικούς"
    IsNumberOnly = False
    GoTo er111
    End If
    ig$ = vbNullString
    DE$ = vbNullString
    sng = sng + 1
    Do While MaybeIsSymbolNoSpace(Mid$(a$, sng + 1, 1), "[0-9A-Fa-f]")
    DE$ = DE$ + Mid$(a$, sng + 1, 1)
    sng = sng + 1
    If Len(DE$) = 8 Then Exit Do
    Loop
    sng = sng + 1
    SG = 1 ' no sign
    If LenB(DE$) = 0 Then
    MyEr "ivalid hex values", "λάθος όρισμα για δεκαεξαδικό"
    IsNumberOnly = False
    GoTo er111
    End If
    If MaybeIsSymbolNoSpace(Mid$(a$, sng, 1), "[&%]") Then
    
        sng = sng + 1
        ig$ = "&H" + DE$
        DE$ = vbNullString
        If Mid$(a$, sng - 1, 1) = "%" Then
        If Len(ig$) > 6 Then
        OverflowLong True
        IsNumberOnly = False
        GoTo er111
        Else
        R = CInt(0)
        End If
        Else
        R = CLng(0)
        End If
        GoTo conthere1
    ElseIf useRtypeOnly Then
        If VarType(R) = vbLong Or VarType(R) = vbInteger Then
        ig$ = "&H" + DE$
        DE$ = vbNullString
        GoTo conthere1
        End If
    End If
        DE$ = Right$("00000000" & DE$, 8)
        R = CDbl(UNPACKLNG(Left$(DE$, 4)) * 65536#) + CDbl(UNPACKLNG(Right$(DE$, 4)))
        GoTo contfinal
  
ElseIf val("0" & Mid$(a$, sng, 1)) = 0 And Left(Mid$(a$, sng, 1), sng) <> "0" And Left(Mid$(a$, sng, 1), sng) <> "." Then
IsNumberOnly = False

Else
'compute ig$
    If Mid$(a$, sng, 1) = "." Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(a$)
        
        Select Case Mid$(a$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, sng, 1)
        Case "."
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    ' compute decimal part
    If DE$ <> "" Then
      sng = sng + 1
        Do While sng <= Len(a$)
       
        Select Case Mid$(a$, sng, 1)
        Case " ", ChrW(160), vbTab
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        DE$ = DE$ & Mid$(a$, sng, 1)
        End If
        Case "E", "e", "Ε", "ε"  ' ************check it
            If ex$ = vbNullString Then
               sg1 = True
                ex$ = "E"
            Else
                Exit Do
            End If
        Case "+", "-"
            If sg1 And Len(ex$) = 1 Then
             ex$ = ex$ & Mid$(a$, sng, 1)
            Else
                Exit Do
            End If
        Case Else
            Exit Do
        End Select
        sng = sng + 1
        Loop
        If Len(ex$) < 3 Then
                If ex$ = "E" Then
                ex$ = "0"
                sng = sng + 1
                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                ex$ = "00"
                sng = sng + 2
                End If
                End If
    End If
    If ig$ = vbNullString Then
    IsNumberOnly = False
    lr = 1
    Else
    If SG < 0 Then ig$ = "-" & ig$
    On Error GoTo er111
     If useRtypeOnly Then GoTo conthere1
    If sng <= Len(a$) Then
    If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
    Select Case Mid$(a$, sng, 1)
    Case "@"
    R = CDec(ig$ & DE$)
    sng = sng + 1
    Case "&"
    R = CLng(ig$)
    sng = sng + 1
    Case "%"
    R = CInt(ig$)
    sng = sng + 1
    Case "~"
    R = CSng(ig$ & DE$ & ex$)
    sng = sng + 1
    Case "#"
    R = CCur(ig$ & DE$)
    sng = sng + 1
    Case Else
GoTo conthere
    End Select
    Else
conthere:
        If useRtypeOnly Then
conthere1:
        If usespecial Then
       If sng <= Len(a$) Then
            Select Case Mid$(a$, sng, 1)
            Case "@"
                R = CDec(0)
                sng = sng + 1
            Case "&"
                R = CLng(0)
                sng = sng + 1
            Case "~"
                R = CSng(0)
                sng = sng + 1
            Case "#"
                R = CCur(0)
                sng = sng + 1
            Case "%"
                R = CInt(0)
                sng = sng + 1
        End Select
        End If
        End If
         If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
        Select Case VarType(R)
        Case vbDecimal
        R = CDec(ig$ & DE$)
        Case vbLong
        R = CLng(ig$)
        Case vbInteger
        R = CInt(ig$)
        Case vbSingle
        R = CSng(ig$ & DE$ & ex$)
        Case vbCurrency
        R = CCur(ig$ & DE$)
        Case vbBoolean
        R = CBool(ig$ & DE$)
        Case Else
        R = CDbl(ig$ & DE$ & ex$)
        End Select
        Else
        If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
        R = val(ig$ & DE$ & ex$)
        End If
    End If
contfinal:
    lr = sng - fr + 1
    
    IsNumberOnly = True
    Exit Function
    End If
End If
er111:
    lr = sng - fr + 1
    Err.Clear
Exit Function

End Function


Function IsNumberD2(a$, D As Variant, Optional noendtypes As Boolean = False, Optional exceptspecial As Boolean) As Boolean
' for inline stacitems
If VarType(D) = vbEmpty Then D = 0#
Dim a1 As Long
If a$ <> "" Then
For a1 = 1 To Len(a$) + 1
Select Case Mid$(a$, a1, 1)
Case " ", ChrW(160), vbTab
If a1 > 1 Then Exit For
Case Is = Chr(2)
If a1 = 1 Then Exit Function
Exit For
End Select
Next a1
If a1 > Len(a$) Then a1 = Len(a$) + 1
    If IsNumberOnly(a$, 1, D, a1, noendtypes, exceptspecial) Then
        a$ = Mid$(a$, a1)
        IsNumberD2 = True
    ElseIf MaybeIsSymbol(a$, "ΑαΨψTtFf") Then
        If Fast3Varl(a$, "ΑΛΗΘΕΣ", 6, "ΑΛΗΘΗΣ", 6, "TRUE", 4, 6) Then
            D = True
            IsNumberD2 = True
        ElseIf Fast3Varl(a$, "ΨΕΥΔΕΣ", 6, "ΨΕΥΔΗΣ", 6, "FALSE", 5, 6) Then
            D = False
            IsNumberD2 = True
        Else
            IsNumberD2 = False
        End If
    Else
    IsNumberD2 = False
    End If
Else
    IsNumberD2 = False
End If

End Function

Function IsNumberD3(a$, fr As Long, a1 As Long) As Boolean
' for inline stacitems
Dim D As Double
If a$ <> "" Then
For a1 = fr To Len(a$) + 1
Select Case Mid$(a$, a1, 1)
Case " ", ChrW(160), vbTab
If a1 > fr Then Exit For
Case Is = Chr(2)
If a1 = fr Then Exit Function
Exit For
End Select
Next a1
If a1 > Len(a$) Then a1 = Len(a$) + 1
If IsNumberOnly(a$, fr, D, a1) Then
IsNumberD3 = True
ElseIf Fast3NoSpaceCheck(fr, a$, "ΑΛΗΘΕΣ", 6, "ΑΛΗΘΗΣ", 6, "TRUE", 4, 6) Then
D = True
IsNumberD3 = True
ElseIf Fast3NoSpaceCheck(fr, a$, "ΨΕΥΔΕΣ", 6, "ΨΕΥΔΗΣ", 6, "FALSE", 5, 6) Then
D = False
IsNumberD3 = True
Else
a1 = fr
IsNumberD3 = False
End If
Else
a1 = fr
IsNumberD3 = False
End If

End Function

Function IsNumberCheck(a$, R As Variant, Optional mydec$ = " ") As Boolean
Dim sng&, SG As Variant, ig$, DE$, sg1 As Boolean, ex$, s$
If mydec$ = " " Then mydec$ = "."
SG = 1
Do While sng& < Len(a$)
sng& = sng& + 1
Select Case Mid$(a$, sng&, 1)
Case "#"
    If Len(a$) > sng& Then
    If MaybeIsSymbolNoSpace(Mid$(a$, sng& + 1, 1), "[0-9A-Fa-f]") Then
    s$ = "0x00" + Mid$(a$, sng& + 1, 6)
    If Len(s$) < 10 Then Exit Function
        If IsNumberCheck(s$, R) Then
        If s$ <> "" Then
          
             
        Else
            s$ = Right$("00000000" & Mid$(a$, sng& + 1, 6), 8)
            a$ = Mid$(a$, sng& + 7)
   R = SG * -(CDbl(UNPACKLNG(Right$(s$, 2)) * 65536#) + CDbl(UNPACKLNG(Mid$(s$, 5, 2)) * 256#) + CDbl(UNPACKLNG(Mid$(s$, 3, 2))))
   IsNumberCheck = True
   Exit Function
        End If
        End If
        Else
        
    End If
    Else

    '' out
    End If
    Exit Function
Case " ", "+", ChrW(160)
Case "-"
SG = -SG
Case Else
Exit Do
End Select
Loop
a$ = Mid$(a$, sng&)
sng& = 1
If val("0" & Mid$(Replace(a$, mydec$, "."), sng&, 1)) = 0 And Left(Mid$(a$, sng&, 1), sng&) <> "0" And Left(Mid$(a$, sng&, 1), sng&) <> mydec$ Then
IsNumberCheck = False
Else

    If Mid$(a$, sng&, 1) = mydec$ Then

    ig$ = "0"
    DE$ = mydec$
    ElseIf LCase(Mid$(a$, sng&, 2)) Like "0[xχ]" Then
    ig$ = "0"
    DE$ = "0x"
  sng& = sng& + 1
Else
    Do While sng& <= Len(a$)
        
        Select Case Mid$(a$, sng&, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, sng&, 1)
        Case mydec$
        DE$ = mydec$
        Exit Do
        Case Else
        Exit Do
        End Select
       sng& = sng& + 1
    Loop
    End If
    ' compute decimal part
    If DE$ <> "" Then
      sng& = sng& + 1
        Do While sng& <= Len(a$)
       
        Select Case Mid$(a$, sng&, 1)
        Case " ", ChrW(160), vbTab
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "A" To "D", "a" To "d", "F", "f"
        If Left$(DE$, 2) = "0x" Then
        DE$ = DE$ & Mid$(a$, sng&, 1)
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(a$, sng&, 1)
        Else
        DE$ = DE$ & Mid$(a$, sng&, 1)
        End If
        Case "E", "e"
         If Left$(DE$, 2) = "0x" Then
         DE$ = DE$ & Mid$(a$, sng&, 1)
         Else
              If ex$ = vbNullString Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
        End If
        Case "Ε", "ε"
 If ex$ = vbNullString Then
          sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
         ex$ = ex$ & Mid$(a$, sng&, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng& = sng& + 1
        Loop
        If Len(ex$) < 3 Then
                If ex$ = "E" Then
                ex$ = "0"
                sng = sng + 1
                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                ex$ = "00"
                sng = sng + 2
                End If
                End If
    End If
    If ig$ = vbNullString Then
    IsNumberCheck = False
    Else

    If Left$(DE$, 2) = "0x" Then

            If Mid$(DE$, 3) = vbNullString Then
            R = 0
            Else
            DE$ = Right$("00000000" & Mid$(DE$, 3), 8)
            R = CDbl(UNPACKLNG(Left$(DE$, 4)) * 65536#) + CDbl(UNPACKLNG(Right$(DE$, 4)))
            End If
    Else
        If SG < 0 Then ig$ = "-" & ig$
                   On Error Resume Next
                        If ex$ <> "" Then
                        If Len(ex$) < 3 Then
                                If ex$ = "E" Then
                                ex$ = "0"
                                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                                ex$ = "00"
                                End If
                                End If
                               If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                               If val(Mid$(ex$, 2)) > 308 Or val(Mid$(ex$, 2)) < -324 Then
                               
                                   R = val(ig$ & DE$)
                                   sng = sng - Len(ex$)
                                   ex$ = vbNullString
                                   
                               Else
                                   R = val(ig$ & DE$ & ex$)
                               End If
                           Else
                       If sng <= Len(a$) Then
            Select Case Asc(Mid$(a$, sng, 1))
            Case 64
                Mid$(a$, sng, 1) = " "
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
                R = CDec(ig$ & DE$)
                If Err.Number = 6 Then
                Err.Clear
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                R = val(ig$ & DE$)
                End If
            Case 35
            Mid$(a$, sng, 1) = " "
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
                R = CCur(ig$ & DE$)
                If Err.Number = 6 Then
                Err.Clear
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                R = val(ig$ & DE$)
                End If
           Case 37
                Mid$(a$, sng, 1) = " "
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
                R = CInt(ig$)
                If Err.Number = 6 Then
                Err.Clear
                R = val(ig$)
                End If
           Case 38
                Mid$(a$, sng, 1) = " "
                R = CLng(ig$)
                If Err.Number = 6 Then
                    Err.Clear
                    R = val(ig$)
                End If
            Case 126
                Mid$(a$, sng, 1) = " "
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
                R = CSng(ig$ & DE$)
                If Err.Number = 6 Then
                Err.Clear
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                R = val(ig$ & DE$)
                End If
            Case Else
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                R = val(ig$ & DE$)
            End Select
            Else
            If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
            R = val(ig$ & DE$)
            End If
                           End If
                     If Err.Number = 6 Then
                         If Len(ex$) > 2 Then
                             ex$ = Left$(ex$, Len(ex$) - 1)
                             sng = sng - 1
                             Err.Clear
                             If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                             R = val(ig$ & DE$ & ex$)
                             If Err.Number = 6 Then
                                 sng = sng - Len(ex$)
                                 If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                                  R = val(ig$ & DE$)
                             End If
                         End If
                       MyEr "Error in exponet", "Λάθος στον εκθέτη"
                       IsNumberCheck = False
                       Exit Function
                     End If
           
         End If
           a$ = Mid$(a$, sng&)
           IsNumberCheck = True
End If
End If
End Function
Function utf8encode(a$) As String
Dim bOut() As Byte, lPos As Long
If LenB(a$) = 0 Then Exit Function
bOut() = Utf16toUtf8(a$)
lPos = UBound(bOut()) + 1
If lPos Mod 2 = 1 Then
    utf8encode = StrConv(String$(lPos, Chr(0)), vbFromUnicode)
Else
    utf8encode = String$((lPos + 1) \ 2, Chr(0))
    End If
    CopyMemory ByVal StrPtr(utf8encode), bOut(0), LenB(utf8encode)
End Function
Function utf8decode(a$) As String
Dim b() As Byte, BLen As Long, WChars As Long
BLen = LenB(a$)
If BLen = 0 Then Exit Function
            ReDim b(0 To BLen - 1)
            CopyMemory b(0), ByVal StrPtr(a$), BLen
            WChars = MultiByteToWideChar(65001, 0, b(0), (BLen), 0, 0)
            utf8decode = space$(WChars)
            MultiByteToWideChar 65001, 0, b(0), (BLen), StrPtr(utf8decode), WChars
End Function

Public Function ideographs(c$) As Boolean
Dim code As Long
If c$ = vbNullString Then Exit Function
code = AscW(c$)  '
ideographs = (code And &H7FFF) >= &H4E00 Or (-code > 24578) Or (code >= &H3400& And code <= &HEDBF&) Or (code >= -1792 And code <= -1281)
End Function
Public Function nounder32(c$) As Boolean
nounder32 = AscW(c$) > 31 Or AscW(c$) < 0
End Function

Function GetImageX(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray, usehandler As mHandler
GetImageX = False
If IsExp(bstack, a$, R) Then
      GetImageX = FastSymbol(a$, ")", True)
        If Not bstack.lastobj Is Nothing Then
           If TypeOf bstack.lastobj Is mHandler Then
              Set usehandler = bstack.lastobj
              Set bstack.lastobj = Nothing
              If usehandler.t1 = 2 Then
                  If usehandler.objref.ReadImageSizeX(R) Then
                  R = SG * bstack.Owner.ScaleX(R, 3, 1)
                          Set usehandler = Nothing
                      Exit Function
                  End If
              End If
           End If
        End If
            noImageInBuffer a$
            GetImageX = False
            R = 0#
            Set bstack.lastobj = Nothing
Else
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    R = cDIBwidth1(var(w1)) * DXP
                    If SG < 0 Then R = -R
                    GetImageX = FastSymbol(a$, ")", True)
                Else
                    noImage a$
                    Exit Function
                End If
            Else
                MissFuncParameterStringVarMacro a$
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    R = SG * cDIBwidth1(sV) * DXP
                    If SG < 0 Then R = -R
                    pppp.SwapItem w2, sV
                    GetImageX = FastSymbol(a$, ")", True)
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End If
    
 
End Function
Function GetImageY(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray, usehandler As mHandler
GetImageY = False
If IsExp(bstack, a$, R) Then
      GetImageY = FastSymbol(a$, ")", True)
        If Not bstack.lastobj Is Nothing Then
           If TypeOf bstack.lastobj Is mHandler Then
              Set usehandler = bstack.lastobj
              Set bstack.lastobj = Nothing
              If usehandler.t1 = 2 Then
                  If usehandler.objref.ReadImageSizeY(R) Then
                  R = SG * bstack.Owner.ScaleY(R, 3, 1)
                          Set usehandler = Nothing
                      Exit Function
                  End If
              End If
           End If
        End If
            noImageInBuffer a$
            GetImageY = False
            R = 0#
            Set bstack.lastobj = Nothing
Else
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    R = cDIBheight1(var(w1)) * DXP
                    If SG < 0 Then R = -R
                    GetImageY = FastSymbol(a$, ")", True)
                Else
                    noImage a$
                    Exit Function
                End If
            Else
                MissFuncParameterStringVarMacro a$
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    R = SG * cDIBheight1(sV) * DXP
                    If SG < 0 Then R = -R
                    pppp.SwapItem w2, sV
                    GetImageY = FastSymbol(a$, ")", True)
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End If
    
 
End Function
Function GetImageXpixels(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray, usehandler As mHandler
GetImageXpixels = False
If IsExp(bstack, a$, R) Then
      GetImageXpixels = FastSymbol(a$, ")", True)
        If Not bstack.lastobj Is Nothing Then
           If TypeOf bstack.lastobj Is mHandler Then
              Set usehandler = bstack.lastobj
              Set bstack.lastobj = Nothing
              If usehandler.t1 = 2 Then
                  If usehandler.objref.ReadImageSizeX(R) Then
                  R = SG * R
                          Set usehandler = Nothing
                      Exit Function
                  End If
              End If
           End If
        End If
            noImageInBuffer a$
            GetImageXpixels = False
            R = 0#
            Set bstack.lastobj = Nothing
Else
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    R = cDIBwidth1(var(w1))
                    If SG < 0 Then R = -R
                    GetImageXpixels = FastSymbol(a$, ")", True)
                Else
                    noImage a$
                    Exit Function
                End If
            Else
                MissFuncParameterStringVarMacro a$
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    R = SG * cDIBwidth1(sV)
                    If SG < 0 Then R = -R
                    pppp.SwapItem w2, sV
                    GetImageXpixels = FastSymbol(a$, ")", True)
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End If
    
 
End Function
Function GetImageYpixels(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray, usehandler As mHandler
GetImageYpixels = False
If IsExp(bstack, a$, R) Then
      GetImageYpixels = FastSymbol(a$, ")", True)
        If Not bstack.lastobj Is Nothing Then
           If TypeOf bstack.lastobj Is mHandler Then
              Set usehandler = bstack.lastobj
              Set bstack.lastobj = Nothing
              If usehandler.t1 = 2 Then
                  If usehandler.objref.ReadImageSizeY(R) Then
                  R = SG * R
                          Set usehandler = Nothing
                      Exit Function
                  End If
              End If
           End If
        End If
            noImageInBuffer a$
            GetImageYpixels = False
            R = 0#
            Set bstack.lastobj = Nothing
Else
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    R = cDIBheight1(var(w1))
                    If SG < 0 Then R = -R
                    GetImageYpixels = FastSymbol(a$, ")", True)
                Else
                    noImage a$
                    Exit Function
                End If
            Else
                MissFuncParameterStringVarMacro a$
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    R = SG * cDIBheight1(sV)
                    If SG < 0 Then R = -R
                    pppp.SwapItem w2, sV
                    GetImageYpixels = FastSymbol(a$, ")", True)
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End If
    
 
End Function

Function enthesi(bstack As basetask, rest$) As String
'first is the string "label {0} other {1}
Dim counter As Long, pat$, Final$, pat1$, pl1 As Long, pl2 As Long, pl3 As Long
Dim q$, p As Variant, p1 As Integer, pd$
If IsStrExp(bstack, rest$, Final$) Then
  If FastSymbol(rest$, ",") Then
    Do
                pl2 = 1
                    pat$ = "{" + CStr(counter)
                   pat1$ = pat$ + ":"
                    pat$ = pat$ + "}"
                    If IsExp(bstack, rest$, p, , True) Then
                    If VarType(p) = vbBoolean Then q$ = Format$(p, DefBooleanString): GoTo fromboolean
again1:
                    pl2 = InStr(pl2, Final$, pat1$)
                    If pl2 > 0 Then
                    pl1 = InStr(pl2, Final$, "}")
                    If Mid$(Final$, pl2 + Len(pat1$), 1) = ":" Then
                    p1 = 0
                    pl3 = val(Mid$(Final$, pl2 + Len(pat1$) + 1) + "}")
                    Else
                    p1 = val("0" + Mid$(Final$, pl2 + Len(pat1$)))
                    
                    pl3 = val(Mid$(Final$, pl2 + Len(pat1$) + Len(Str$(p1))) + "}")
                    If p1 < 0 Then p1 = 13 '22
                    If p1 > 13 Then p1 = 13
                  p = MyRound(p, p1)
                  End If
                  pd$ = LTrim$(Str(p))
                  
                  If InStr(pd$, "E") > 0 Or InStr(pd$, "e") > 0 Then '' we can change e to greek ε
                  pd$ = Format$(p, "0." + String$(p1, "0") + "E+####")
                       If Not NoUseDec Then
                               If OverideDec Then
                                pd$ = Replace$(pd$, GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                                pd$ = Replace$(pd$, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                                pd$ = Replace$(pd$, Chr(2), NowDec$)
                                pd$ = Replace$(pd$, Chr(3), NowThou$)
                                
                            ElseIf InStr(pd$, NowDec$) > 0 Then
                            pd$ = Replace$(pd$, NowDec$, Chr(2))
                            pd$ = Replace$(pd$, NowThou$, Chr(3))
                            pd$ = Replace$(pd$, Chr(2), ".")
                            pd$ = Replace$(pd$, Chr(3), ",")
                            
                            End If
                        End If
                  ElseIf p1 <> 0 Then
                   pd$ = Format$(p, "0." + String$(p1, "0"))
                           If Not NoUseDec Then
                            If OverideDec Then
                                pd$ = Replace$(pd$, GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                                pd$ = Replace$(pd$, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                                pd$ = Replace$(pd$, Chr(2), NowDec$)
                                pd$ = Replace$(pd$, Chr(3), NowThou$)
                            ElseIf InStr(pd$, NowDec$) > 0 Then
                            pd$ = Replace$(pd$, NowDec$, Chr(2))
                            pd$ = Replace$(pd$, NowThou$, Chr(3))
                            pd$ = Replace$(pd$, Chr(2), ".")
                            pd$ = Replace$(pd$, Chr(3), ",")
                            
                            End If
                        End If
                  End If
               
                  If pl3 <> 0 Then
                    If pl3 > 0 Then
                        pd$ = Left$(pd$ + space$(pl3), pl3)
                        Else
                        pd$ = Right$(space$(Abs(pl3)) + pd$, Abs(pl3))
                        End If
                  End If
                        Final$ = Replace$(Final$, Mid$(Final$, pl2, pl1 - pl2 + 1), pd$)
                        GoTo again1
                    Else
                    
                    If NoUseDec Then
                        Final$ = Replace$(Final$, pat$, CStr(p))
                    Else
                    pd$ = LTrim$(Str$(p))
                     If Left$(pd$, 1) = "." Then
                    pd$ = "0" + pd$
                    ElseIf Left$(pd$, 2) = "-." Then pd$ = "-0" + Mid$(pd$, 2)
                    End If
                    If OverideDec Then
                    Final$ = Replace$(Final$, pat$, Replace(pd$, ".", NowDec$))
                    Else
                    Final$ = Replace$(Final$, pat$, pd$)
                    End If
                    End If
                    
                    
                        End If
                        If Not FastSymbol(rest$, ",") Then Exit Do
                    
                    ElseIf IsStrExp(bstack, rest$, q$) Then
fromboolean:
                        Final$ = Replace$(Final$, pat$, q$)
AGAIN0:
                    pl2 = InStr(pl2, Final$, pat1$)
                      If pl2 > 0 Then
                       pl1 = InStr(pl2, Final$, "}")
                       pl3 = val(Mid$(Final$, pl2 + Len(pat1$)) + "}")
                       If pl3 <> 0 Then
                    If pl3 > 0 Then
                        pd$ = Left$(q$ + space$(pl3), pl3)
                        Else
                        pd$ = Right$(space$(Abs(pl3)) + q$, Abs(pl3))
                        End If
                  End If
                        Final$ = Replace$(Final$, Mid$(Final$, pl2, pl1 - pl2 + 1), pd$)
                        GoTo AGAIN0
                      End If
                        If Not FastSymbol(rest$, ",") Then Exit Do
                    Else
                        Exit Do
                    End If
                    counter = counter + 1
    Loop
    Else
    enthesi = EscapeStrToString(Final$)
    Exit Function
    End If
End If
enthesi = Final$
End Function

Public Function GetDeflocaleString(ByVal this As Long) As String
On Error GoTo 1234
    Dim Buffer As String, ret&, R&
    Buffer = String$(514, 0)
      
        ret = GetLocaleInfoW(0, this, StrPtr(Buffer), Len(Buffer))
    GetDeflocaleString = Left$(Buffer, ret - 1)
    
1234:
    
End Function
Function RetM2000array(var As Variant) As Variant
Dim ar As New mArray, v(), manydim As Long, probe As Long, probelow As Long
Dim j As Long
v() = var
On Error GoTo ma100
For j = 1 To 60
    probe = UBound(v, j)
    If Err Then Exit For
Next j
manydim = j - 1
On Error Resume Next
For j = manydim To 1 Step -1
    
    probe = UBound(v, j)
    If Err Then Exit For
    probelow = LBound(v, j)
    ar.PushDim probe - probelow + 1
Next j
ar.PushEnd
ar.RevOrder = True
ar.CopySerialize v()
ma100:
Set RetM2000array = ar

End Function

Public Sub DisableTargets(j() As target, ByVal myl As Long)
Dim iu&, id&, i&
iu& = LBound(j())
id& = UBound(j())
For i& = iu& To id&
 If j(i&).layer = myl Then j(i&).Enable = False
Next i&
End Sub
Function BoxTarget(DSTACK As basetask, ByVal xl&, ByVal yl&, ByVal b As Long, ByVal F As Long, ByVal Tag$, ByVal id&, ByVal COM$, XXT&, YYT&, Linespace&) As target
Dim x&, y&, D As Object
Set D = DSTACK.Owner
Dim prive As basket
prive = players(GetCode(D))
With prive

x& = .curpos
y& = .currow
xl& = xl& + x&
yl& = yl& + y& - 1
With BoxTarget
.SZ = prive.SZ
.Comm = COM$
.id = id&
.Tag = Tag$
.Lx = x&
.ly = y&
.tx = xl& - 1
.ty = yl&
.back = b
.fore = F
.Enable = True
.pen = prive.mypen
.Xt = XXT&
.Yt = YYT&
.sUAddTwipsTop = prive.uMineLineSpace
If D.Name = "DIS" Then
.layer = 0
ElseIf D.Name = "Form1" Then
.layer = -1
ElseIf D.Name = "dSprite" Then
.layer = D.Index
Else
.layer = GetCode(D)
End If
End With
If F <> &H81000000 Then BoxBigNew D, prive, xl& - 1, yl&, F
If b <> &H81000000 Then BoxColorNew D, prive, xl& - 1, yl&, b
If id& < 100 Then
    Tag$ = Left$(Tag$, xl& - x&)
    If Tag$ <> "" Then
    
    Select Case id& Mod 10
    Case 4, 5, 6
    y& = (yl& + y&) \ 2
    Case 7, 8, 9
    y& = yl&
    Case Else
    End Select
    Select Case id& Mod 10
    Case 2, 5, 8
    x& = (xl& + x& - Len(Tag$)) \ 2
    Case 3, 6, 9
    x& = xl& - Len(Tag$)
    Case Else
    End Select
    If (id& Mod 10) > 0 Then
    LCTbasket D, prive, y&, x&
    D.FontTransparent = True
    PlainBaSket D, prive, Tag$, True, True
    LCTbasket D, prive, BoxTarget.ly, BoxTarget.Lx
    End If
    End If
Else
    If Tag$ <> "" Then
    id& = id& Mod 100
    Select Case id& Mod 10
    Case 4, 5, 6
    y& = (yl& + y&) \ 2
    Case 7, 8, 9
    y& = yl&
    Case Else
    End Select
    F = 3
    Select Case id& Mod 10
    Case 2, 5, 8
    F = 2
    Case 3, 6, 9
    F = 1
    Case Else
    End Select
    
    If (id& Mod 10) > 0 Then
    LCTbasket D, prive, y&, x&
    D.FontTransparent = True
    D.currentX = D.currentX - dv15 * 2
    wwPlain2 DSTACK, prive, Tag$, xl& - x&, 10000, , True, F, , , True
    LCTbasket D, prive, BoxTarget.ly, BoxTarget.Lx
    End If
End If
    
End If
End With
End Function
Private Function MyMod(R1, po) As Variant
MyMod = R1 - Fix(R1 / po) * po
End Function
Sub dset()

'USING the temporary path
    strTemp = String(MAX_FILENAME_LEN, Chr$(0))
    'Get
    GetTempPath MAX_FILENAME_LEN, StrPtr(strTemp)
    strTemp = LONGNAME(mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1)))
    If strTemp = vbNullString Then
     strTemp = mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1))
    End If
' NOW COPY
' for mcd
Dim CD As String, dummy As Long, q$

''cd = App.Path
''AddDirSep cd
''mcd = mylcasefILE(cd)

' Return to standrad path...for all users
userfiles = GetSpecialfolder(CLng(26)) & "\M2000"
AddDirSep userfiles
If Not isdir(userfiles) Then
MkDir userfiles
End If

mcd = userfiles
DefaultDec$ = GetDeflocaleString(LOCALE_SDECIMAL)
If NowDec$ <> "" Then
ElseIf OverideDec Then
NowDec$ = GetlocaleString(LOCALE_SDECIMAL)
NowThou$ = GetlocaleString(LOCALE_STHOUSAND)
Else
NowDec$ = DefaultDec$
NowThou$ = GetDeflocaleString(LOCALE_STHOUSAND)
End If
CheckDec
cdecimaldot$ = GetDeflocaleString(LOCALE_SDECIMAL)
End Sub
Public Sub CheckDec()
OverideDec = False
NowDec$ = GetDeflocaleString(LOCALE_SDECIMAL)
NowThou$ = GetDeflocaleString(LOCALE_STHOUSAND)
If NowDec$ = "." Then
NoUseDec = False
Else
NoUseDec = mNoUseDec
End If
End Sub
Function ProcEnumGroup(bstack As basetask, rest$, Optional glob As Boolean = False) As Boolean

    Dim s$, w1$, v As Long, enumvalue As Long, myenum As Enumeration, mh As mHandler, v1 As Long
    Dim gr As Boolean
    enumvalue = 0
    If FastPureLabel(rest$, w1$, , , , , , gr) = 1 Then

        v = globalvar(bstack.GroupName + myUcase(w1$, gr), v, , glob)
        Set myenum = New Enumeration
        
        myenum.EnumName = w1$
        Else
        MyEr "No proper name for enumeration", "μη κανονικό όνομα για απαρίθμηση"
        Exit Function
    End If
    If FastSymbol(rest$, "{") Then
        s$ = block(rest$)
        
        Do
        If FastSymbol(s$, vbCrLf, , 2) Then
        While FastSymbol(s$, vbCrLf, , 2)
        Wend
        ElseIf FastPureLabel(s$, w1$, , , , , , gr) = 1 Then
            'w1 = myUcase(w1$)
            If FastSymbol(s$, "=") Then
            If IsExp(bstack, s$, enumvalue) Then
                If Not bstack.lastobj Is Nothing Then
                    MyEr "No Object allowed as enumeration value", "Δεν επιτρέπεται αντικείμενο για τιμή απαριθμητή"
                    Exit Function
                    End If
                End If
            Else
                    enumvalue = enumvalue + 1
            End If
            myenum.addone w1$, enumvalue
            Set mh = New mHandler
            Set mh.objref = myenum
            mh.t1 = 4
            mh.ReadOnly = True
            mh.index_cursor = enumvalue
            mh.index_start = myenum.count - 1
             v1 = globalvar(bstack.GroupName + myUcase(w1$, gr), v1, , glob)
             Set var(v1) = mh
            ProcEnumGroup = True
        Else
            Exit Do
        End If
        If FastSymbol(s$, ",") Then ProcEnumGroup = False
        Loop
        If v1 > v Then Set var(v) = var(v1) Else MyEr "Empty Enumeration", "’δεια Απαρίθμηση": Exit Function
        ProcEnumGroup = FastSymbol(rest$, "}", True)
    Else
        MissingEnumBlock
        Exit Function
    End If
    
    
End Function
Function ProcEnum(bstack As basetask, rest$, Optional glob As Boolean = False) As Boolean

    Dim s$, w1$, v As Long, enumvalue As Variant, myenum As Enumeration, mh As mHandler, v1 As Long, i As Long
    Dim gr As Boolean
    enumvalue = 0#
    If FastPureLabel(rest$, w1$, , , , , , gr) = 1 Then
       ' w1$ = myucase(w1$)
        v = globalvar(myUcase(w1$, gr), v, , glob)
        Set myenum = New Enumeration
        
        myenum.EnumName = w1$
        Else
        MyEr "No proper name for enumeration", "μη κανονικό όνομα για απαρίθμηση"
        Exit Function
    End If
    If FastSymbol(rest$, "{") Then
        s$ = block(rest$)
        
        Do
        If FastSymbol(s$, vbCrLf, , 2) Then
        While FastSymbol(s$, vbCrLf, , 2)
        Wend
        ElseIf MaybeIsSymbol(s$, "\'") Then
        
        SetNextLine s$
        ElseIf FastPureLabel(s$, w1$, , , , , , gr) = 1 Then
   
            If FastSymbol(s$, "=") Then
            If IsExp(bstack, s$, enumvalue) Then
                If Not bstack.lastobj Is Nothing Then
                    MyEr "No Object allowed as enumeration value", "Δεν επιτρέπεται αντικείμενο για τιμή απαριθμητή"
                    Exit Function
                   End If
            Else
                    MyEr "No String allowed as enumeration value", "Δεν επιτρέπεται αλφαριθμητικό για τιμή απαριθμητή"
                    Exit Function
            Exit Function
                End If
            Else
                    enumvalue = enumvalue + 1
            End If
            myenum.addone w1$, enumvalue
            w1$ = myUcase(w1$, gr)
            If numid.Find(w1$, i) Then If i > 0 Then numid.ItemCreator2 w1$, -1
            
            Set mh = New mHandler
            Set mh.objref = myenum
            mh.t1 = 4
            mh.ReadOnly = True
            mh.index_cursor = enumvalue
            mh.index_start = myenum.count - 1
            
             v1 = globalvar(w1$, v1, , glob)
             Set var(v1) = mh
            ProcEnum = True
        Else
            Exit Do
        End If
        If FastSymbol(s$, ",") Then ProcEnum = False
        Loop
        If v1 > v Then Set var(v) = var(v1) Else MyEr "Empty Enumeration", "’δεια Απαρίθμηση": Exit Function
        ProcEnum = FastSymbol(rest$, "}", True)
    Else
        MissingEnumBlock
        Exit Function
    End If
    
    
End Function
Function CallLambdaASAP(bstack As basetask, a$, R, Optional forstring As Boolean = False) As Long
Dim w2 As Long, w1 As Long, nbstack As basetask
PushStage bstack, False
w2 = var2used
If forstring Then
w1 = globalvarGroup("A_" + CStr(w2) + "$", 0#)
 Set var(w1) = bstack.lastobj
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "$()", "", , , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "$()", "", , , w1
    End If
 Set nbstack = New basetask
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
 CallLambdaASAP = GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "$()", a$, R)

Else
w1 = globalvarGroup("A_" + CStr(w2), 0#)
 Set var(w1) = bstack.lastobj
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w1
    End If
     Set nbstack = New basetask
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
 CallLambdaASAP = GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "()", a$, R)
End If


                 
PopStage bstack
End Function

Function ProcText(basestack As basetask, isHtml As Boolean, rest$) As Boolean
Dim x1 As Long, frm$, pa$, s$
ProcText = True
If IsSymbol(rest$, "UTF-8", 5) Then
x1 = 2
ElseIf IsSymbol(rest$, "UTF-16", 6) Then
x1 = 0 ' only little endian (but if something convert it to big we can read...)
Else
x1 = 3
End If

s$ = vbNullString
If Not IsStrExp(basestack, rest$, s$) Then
If Not FastPureLabel(rest$, s$) = 1 Then
    ProcText = False
    Exit Function
End If
End If
FastSymbol rest$, ","
If s$ <> "" Then

If FastSymbol(rest$, "+") Then pa$ = vbNullString Else pa$ = "new"
If FastSymbol(rest$, "{") Then frm$ = NLTrim2$(blockString(rest$, 125))
If frm$ <> "" Then
If isHtml Then
If ExtractType(s$) = vbNullString Then s$ = s$ & ".html"
End If
 textPUT basestack, mylcasefILE(s$), frm$, pa$, x1
Else
 textDel (mylcasefILE(s$))
 ProcText = True
 Exit Function
End If
ProcText = FastSymbol(rest$, "}")
End If
Exit Function

End Function
Private Function textPUT(bstack As basetask, ByVal ThisFile As String, THISBODY As String, c$, mode2save As Long) As Boolean
Dim chk As String, b$, j As Long, PREPARE$, VR$, s$, v As Double, buf$, i As Long
ThisFile = strTemp + ThisFile
chk = GetDosPath(ThisFile)
If chk <> "" And c$ = "new" Then KillFile GetDosPath(chk)
On Error GoTo HM
textPUT = True
Do
j = InStr(THISBODY, "##")
If j = 0 Then PREPARE$ = PREPARE$ & THISBODY: Exit Do
If j > 1 Then PREPARE$ = PREPARE$ & Mid$(THISBODY, 1, InStr(THISBODY, "##") - 1)
THISBODY = Mid$(THISBODY, j + 2)
j = InStr(THISBODY, "##")
If j = 0 Then PREPARE$ = PREPARE$ & THISBODY: Exit Do
If j > 1 Then VR$ = Mid$(THISBODY, 1, InStr(THISBODY, "##") - 1)
THISBODY = Mid$(THISBODY, j + 2)
'
If IsExp(bstack, VR$, v, , True) Then
buf$ = Trim$(Str$(v))
ElseIf IsStrExp(bstack, VR$, s$, Len(bstack.tmpstr) = 0) Then
buf$ = s$
Else
buf$ = VR$
End If
PREPARE$ = PREPARE$ & buf$
Loop
           If Not WeCanWrite(ThisFile) Then GoTo HM

textPUT = SaveUnicode(ThisFile, PREPARE$, mode2save, Not (c$ = "new"))
Exit Function
HM:
textPUT = False
End Function
Private Function textDel(ByVal ThisFile As String) As Boolean
Dim chk As String
ThisFile = strTemp + ThisFile
chk = CFname(ThisFile)
textDel = (chk <> "")
If chk <> "" Then KillFile chk
End Function
Function MyPset(bstack As basetask, rest$) As Boolean
Dim prive As Long, x As Double, p As Variant, y As Double, Col As Long
Dim Scr As Object, ss$
Set Scr = bstack.Owner
prive = GetCode(Scr)
With players(prive)
    Col = players(prive).mypen
    If IsExp(bstack, rest$, p, , True) Then Col = mycolor(p)
    If FastSymbol(rest$, ",") Then
        If IsExp(bstack, rest$, x, , True) Then
            If FastSymbol(rest$, ",") Then
                If IsExp(bstack, rest$, y, , True) Then
                    If TypeOf Scr Is MetaDc Then
                     
                        Scr.Line2 x, y, x, y, Col, False, False
                    Else
                        Scr.PSet (x, y), Col
                    End If
                     MyPset = True
                Else
                    MissPar
                End If
            End If
        Else
            MissPar
        End If
    Else
        If TypeOf Scr Is MetaDc Then
            Scr.Line2 x, y, x, y, Col, False, False
        Else
            Scr.PSet (.XGRAPH, .YGRAPH), Col
        End If
        MyPset = True
    End If
End With
MyDoEvents1 Scr
Set Scr = Nothing
End Function
Function Matrix(bstack As basetask, a$, Arr As Variant, res As Variant) As Boolean
Dim Pad$, cut As Long, pppp As mArray, pppp1 As mArray, st1 As mStiva, anything As Object, w3 As Long, usehandler As mHandler, R As Variant, p As Variant
Dim cur As Long, w2 As Long, w4 As Long, retresonly As Boolean, s$
Dim multi As Boolean, original As Long, bhas As Integer
Set anything = Arr
If TypeOf anything Is mArray Then
Set usehandler = New mHandler
usehandler.t1 = 3
Set usehandler.objref = Arr
Set anything = usehandler
Else
If Not CheckLastHandlerOrIterator(anything, w3) Then Exit Function
End If
Pad$ = myUcase(Left$(a$, 20))  ' 20??
cut = InStr(Pad$, "(")

If cut <= 1 Then Exit Function
Mid$(a$, 1, cut) = space$(cut)
Set usehandler = anything
If TypeOf usehandler.objref Is mArray Then
Set pppp = usehandler.objref

If Left$(Pad$, 1) = Chr$(1) Then LSet Pad$ = Mid$(Pad$, 2): cut = cut - 1
Do
multi = False
Select Case Left$(Pad$, cut - 1)
Case "SUM", "ΑΘΡ"
res = 0
For w3 = 0 To pppp.count - 1
If pppp.MyIsNumeric(pppp.item(w3)) Then res = res + pppp.item(w3)
Next w3
Case "MIN", "ΜΙΚ"
res = 0
w4 = -1
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 1
If pppp.MyIsNumeric(w3) Then res = pppp.itemnumeric(w3): w4 = w3: Exit For
Next w3

For w3 = w3 To pppp.count - 1
If pppp.MyIsNumeric(pppp.item(w3)) Then If pppp.item(w3) < res Then res = pppp.item(w3): w4 = w3
Next w3
End If
If Not FastSymbol(a$, ")") Then
    bstack.soros.PushVal w4
    If Not getone(bstack, a$) Then Exit Function
    Else
    Matrix = True
    Exit Function
End If
Case "MIN$", "ΜΙΚ$"
res = vbNullString
w4 = -1
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 1
If pppp.IsStringItem(w3) Then res = pppp.item(w3): w4 = w3: Exit For
Next w3

For w3 = w3 To pppp.count - 1
If pppp.IsStringItem(w3) Then If pppp.item(w3) < res Then res = pppp.item(w3): w4 = w3
Next w3
End If
If Not FastSymbol(a$, ")") Then
    bstack.soros.PushVal w4
    If Not getone(bstack, a$) Then Exit Function
Else
    Matrix = True
    Exit Function
End If

Case "MAX$", "ΜΕΓ$"
res = vbNullString
w4 = -1
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 1
If pppp.IsStringItem(w3) Then res = pppp.item(w3): w4 = w3: Exit For
Next w3

For w3 = w3 To pppp.count - 1
If pppp.IsStringItem(w3) Then If pppp.item(w3) > res Then res = pppp.item(w3): w4 = w3
Next w3
End If
If Not FastSymbol(a$, ")") Then
    bstack.soros.PushVal w4
    If Not getone(bstack, a$) Then Exit Function
Else
    Matrix = True
    Exit Function
End If
Case "MAX", "ΜΕΓ"
res = 0
w4 = -1
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 1
If pppp.MyIsNumeric(pppp.item(w3)) Then res = pppp.itemnumeric(w3): w4 = w3: Exit For
Next w3

For w3 = w3 To pppp.count - 1
If pppp.MyIsNumeric(pppp.item(w3)) Then If pppp.item(w3) > res Then res = pppp.item(w3): w4 = w3
Next w3
End If
If Not FastSymbol(a$, ")") Then
    
    bstack.soros.PushVal w4
    If Not getone(bstack, a$) Then Exit Function
    
Else
    Matrix = True
    Exit Function
End If
Case "EVAL", "EVAL$", "ΕΚΦΡ", "ΕΚΦΡ$"
If IsExp(bstack, a$, p, , True) Then
    w2 = CLng(p)
Else
    w2 = 0
End If
If w2 < 0 Or w2 >= pppp.count Then
MyEr "offset out of limits", "Δείκτης εκτός ορίων"
Matrix = False
Exit Function
Else
If pppp.MyIsObject(pppp.item(w2)) Then
Set bstack.lastobj = pppp.item(w2)
res = 0
lookagain:
    If Not bstack.lastobj Is Nothing Then

        If TypeOf bstack.lastobj Is mHandler Then
            Set usehandler = bstack.lastobj
            With usehandler
                If .t1 = 3 Then
                    If TypeOf .objref Is mArray Then
                        If FastSymbol(a$, ",") Then
                            If IsExp(bstack, a$, p, , True) Then
                                w4 = CLng(Int(p))
                                If w4 < 0 Or w4 >= .objref.count Then
                                    indexout a$
                                    Exit Function
                                End If
                                Set pppp = .objref
                                pppp.Index = w4
                                If pppp.IsObj Then
                                Set res = pppp.Value
                                If lookOne(a$, ",") Then w2 = w4: Set bstack.lastobj = res: GoTo lookagain
                                Set anything = res
                                If Not anything Is Nothing Then
                                    If TypeOf anything Is mHandler Then
                                        Set usehandler = anything
                                        If usehandler.t1 = 3 Then
                                            Set bstack.lastobj = usehandler
                                            Set pppp = usehandler.objref
                                            res = 0
                                            multi = True
                                            GoTo conthere11
                                        End If
                                    End If
                                End If
                                Else
                                If Mid$(Pad$, cut - 1, 1) = "$" Then
                                res = CStr(MyVal(pppp.Value))

                                Else
                                res = MyVal(pppp.Value)
                                End If
                                End If
                            Set bstack.lastobj = Nothing
                            Matrix = FastSymbol(a$, ")")
                            Exit Function
                            End If
                        Else
                            Set pppp = .objref
                            multi = True
                        End If
                        
                    ElseIf TypeOf .objref Is mStiva Then
  
                    If FastSymbol(a$, ",") Then
                            If IsExp(bstack, a$, p, , True) Then
                                w4 = CLng(Int(p))
                                If w4 < 0 Or w4 >= .objref.count Then
                                    indexout a$
                                    Exit Function
                                End If
                             
                                .objref.Index = w4
                                If .objref.IsObj Then
                                Set res = .objref.Value
                                If lookOne(a$, ",") Then w2 = w4: Set bstack.lastobj = res: GoTo lookagain
                                Set anything = res
                                If Not anything Is Nothing Then
                                    If TypeOf anything Is mHandler Then
                                        Set usehandler = anything
                                        If usehandler.t1 = 3 Then
                                            Set bstack.lastobj = usehandler
                                            Set pppp = usehandler.objref
                                            res = 0
                                            multi = True
                                            GoTo conthere11
                                        End If
                                    End If
                                End If
                                
                                
                                
                                
                                
                                Else
                                If Mid$(Pad$, cut - 1, 1) = "$" Then
                                res = CStr(MyVal(.objref.Value))
                                Else
                                res = MyVal(.objref.Value)
                                End If
                                End If
                                
                            Set bstack.lastobj = Nothing
                            Matrix = FastSymbol(a$, ")")
                            Exit Function
                            End If
                       End If
                    End If
                ElseIf usehandler.t1 = 1 And FastSymbol(a$, ",") Then
                    If IsExp(bstack, a$, p, , True) Then
                            w4 = CLng(Int(p))
                            If w4 < 0 Or w4 >= .objref.count Then
                            indexout a$
                            Exit Function
                            End If
                        End If
                    

                        .objref.Index = w4
                        .objref.Done = True

                    If .objref.IsObj Then
                        res = usehandler.objref.ValueObj
                                If lookOne(a$, ",") Then w2 = w4: Set bstack.lastobj = res: GoTo lookagain
                                Set anything = res
                                If Not anything Is Nothing Then
                                    If TypeOf anything Is mHandler Then
                                        Set usehandler = anything
                                        If usehandler.t1 = 3 Then
                                            Set bstack.lastobj = usehandler
                                            Set pppp = usehandler.objref
                                            res = 0
                                            multi = True
                                            GoTo conthere11
                                        End If
                                    End If
                                End If
                        
                        
                        
                    Else
                                If Mid$(Pad$, cut - 1, 1) = "$" Then
                                res = CStr(MyVal(.objref.Value))
                                Else
                                res = MyVal(.objref.Value)
                                End If
                        Set bstack.lastobj = Nothing
                    End If
                    .objref.Done = False
            End If
            End With
        ElseIf TypeOf bstack.lastobj Is lambda Then
                            PushStage bstack, False

                            If Mid$(Pad$, cut - 1, 1) = "$" Then
                                w3 = globalvarGroup("A_" + CStr(Abs(w2)) + "$", 0#)
                                Set var(w3) = bstack.lastobj
                                If here$ = vbNullString Then
                                    GlobalSub "A_" + CStr(Abs(w2)) + "$()", "", , , w3
                                Else
                                    GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "$()", "", , , w3
                                End If
                                If Not FastSymbol(a$, ")(", , 2) Then FastSymbol a$, ","
                                bstack.tmpstr = "A_" + CStr(Abs(w2)) + "$(" + Left$(a$, 1)
                                BackPort a$
                               Matrix = IsStr1(bstack, a$, s$)
                               res = s$
                            Else
                                w3 = globalvarGroup("A_" + CStr(Abs(w2)), 0#)
                                Set var(w3) = bstack.lastobj
                                If here$ = vbNullString Then
                                    GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , , w3
                                Else
                                    GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w3
                                End If
                                If Not FastSymbol(a$, ")(", , 2) Then FastSymbol a$, ","
                                bstack.tmpstr = "A_" + CStr(Abs(w2)) + "(" + Left$(a$, 1)
                                BackPort a$
                                Matrix = IsNumberNew(bstack, a$, res, (1), False)
                            End If
                            PopStage bstack
                            If Not bstack.lastobj Is Nothing Then
                            If TypeOf bstack.lastobj Is mHandler Then
                                Set usehandler = bstack.lastobj
                                If usehandler.t1 = 3 Then
                                    If TypeOf usehandler.objref Is mArray Then
                                        Set pppp = usehandler.objref
                                        multi = True
                                        GoTo conthere11
                                    End If
                                End If
                            End If
                                                      
                            End If
                            Exit Function
conthere11:
        ElseIf TypeOf bstack.lastobj Is Group Then
                                    Set pppp = BoxGroupVar(bstack.lastobj)
                                    Set bstack.lastpointer = Nothing
                                    Set bstack.lastobj = Nothing
                                w2 = 0
                
                a$ = NLtrim(a$)
            If Left$(a$, 1) = "," Then
                If Mid$(Pad$, cut - 1, 1) = "$" Then
                    Mid$(a$, 1, 1) = " "
                    Matrix = SpeedGroup(bstack, pppp, "VAL$", pppp.CodeName + "$(", a$, w2) = 1
                
                Else
                    Mid$(a$, 1, 1) = "("
                    Matrix = SpeedGroup(bstack, pppp, "VAL", pppp.CodeName, a$, w2) = 1
                    res = bstack.LastValue
                End If
            Else
                If Mid$(Pad$, cut - 1, 1) = "$" Then
                    Matrix = SpeedGroup(bstack, pppp, "VAL$", pppp.CodeName + "$(", a$, w2) = 1
                
                Else
                    Matrix = SpeedGroup(bstack, pppp, "VAL", pppp.CodeName, a$, w2) = 1
                    res = bstack.LastValue
                End If
                    FastSymbol a$, ")", True
            End If
            
            res = bstack.LastValue
            
            Exit Function
        Else
            Set bstack.lastobj = Nothing
            res = rValue(bstack, pppp.item(w2))
            If Not bstack.lastobj Is Nothing Then
                If TypeOf bstack.lastobj Is mHandler Then
                    Set usehandler = bstack.lastobj
                    If usehandler.t1 = 3 Then
                        If TypeOf usehandler.objref Is mArray Then
                        res = 0
                        Set pppp = CopyArray(usehandler.objref)
                        multi = True
                        End If
                    End If
                End If
            End If
        End If
    End If
Else
    If Mid$(Pad$, cut - 1, 1) = "$" Then
        If Not pppp.IsStringItem(w2) Then
            On Error Resume Next
            res = CStr(pppp.item(w2))
                If Err Then res = vbNullString
            Else
                res = pppp.item(w2)
            End If
        Else
            res = pppp.item(w2)
        End If
    End If

End If

Case "VAL", "ΤΙΜΗ", "VAL$", "ΤΙΜΗ$"
If IsExp(bstack, a$, p, , True) Then
    w2 = CLng(p)
Else
    w2 = 0
End If
If w2 < 0 Or w2 >= pppp.count Then
MyEr "offset out of limits", "Δείκτης εκτός ορίων"
Matrix = False
Exit Function
Else
If pppp.MyIsObject(pppp.item(w2)) Then
Set bstack.lastobj = pppp.item(w2)
res = 0
If Not bstack.lastobj Is Nothing Then
If TypeOf bstack.lastobj Is mHandler Then
Set usehandler = bstack.lastobj
With usehandler
If .t1 = 3 Then
    If TypeOf usehandler.objref Is mArray Then
    Set pppp = usehandler.objref
    multi = True
    End If
    ElseIf .t1 = 1 And FastSymbol(a$, ",") Then
        If IsExp(bstack, a$, p, , True) Then
                w4 = CLng(Int(p))
                If w4 < 0 Or w4 >= .objref.count Then
                indexout a$
                Exit Function
                End If
            End If
        

            .objref.Index = w4
            .objref.Done = True

        If .objref.IsObj Then
            res = rValue(bstack, usehandler.objref.ValueObj)
        Else
                    If Mid$(Pad$, cut - 1, 1) = "$" Then
                    res = CStr(MyVal(.objref.Value))
                    Else
                    res = MyVal(.objref.Value)
                    End If
            Set bstack.lastobj = Nothing
        End If
        .objref.Done = False
End If
End With
ElseIf TypeOf bstack.lastobj Is Group Then
    If bstack.lastpointer Is Nothing Then
    If bstack.lastobj.IamApointer Then
        Set bstack.lastpointer = bstack.lastobj
    
    End If
    End If
ElseIf TypeOf bstack.lastobj Is mArray Then

Set pppp = bstack.lastobj
multi = True
End If
End If
Else
    If Mid$(Pad$, cut - 1, 1) = "$" Then
        If Not pppp.IsStringItem(w2) Then
            On Error Resume Next
            res = CStr(pppp.item(w2))
                If Err Then res = vbNullString
            Else
                res = pppp.item(w2)
            End If
        Else
            res = pppp.item(w2)
        End If
    End If
End If
Case "SLICE", "ΜΕΡΟΣ"
If IsExp(bstack, a$, p, , True) Then
If p < 0 Or p >= pppp.count Then
    MyEr "start offset out of limits", "Δείκτης αρχής εκτός ορίων"
    Matrix = False
    Exit Function
End If
Else
p = 0
End If
If FastSymbol(a$, ",") Then
If IsExp(bstack, a$, R, , True) Then
    If R >= pppp.count Or R < p Then
    MyEr "end offset out of limits", "Δείκτης τέλους εκτός ορίων"
    Matrix = False
    Exit Function
    End If
Else
R = pppp.count - 1
End If
Else
R = pppp.count - 1
End If
If original > 0 Then
pppp.CopyArraySliceFast pppp1, CLng(p), CLng(R)
Else
pppp.CopyArraySlice pppp1, CLng(p), CLng(R)
End If
original = original + 1
Set pppp = pppp1
Set pppp1 = Nothing
multi = True
Matrix = True
Set usehandler = New mHandler
usehandler.t1 = 3
Set usehandler.objref = pppp
Set bstack.lastobj = usehandler
Case "SORT", "ΤΑΞΙΝΟΜΗΣΗ"
w2 = 0
w3 = -1
w4 = -1
If IsExp(bstack, a$, p, , True) Then
w2 = CLng(p)
End If
If FastSymbol(a$, ",") Then
If IsExp(bstack, a$, p, , True) Then
w3 = CLng(p)
End If
End If
If FastSymbol(a$, ",") Then
If IsExp(bstack, a$, p, , True) Then
w4 = CLng(p)
End If
End If
If original > 0 Then
    Set pppp1 = pppp
Else
    Set pppp1 = New mArray
    pppp.CopyArray pppp1
End If
original = original + 1

If w2 Then
pppp1.SortDesTuple w3, w4
Else
pppp1.SortTuple w3, w4
End If
Set pppp = pppp1
Set pppp1 = Nothing
multi = True
Matrix = True
Set usehandler = New mHandler
usehandler.t1 = 3
Set usehandler.objref = pppp
Set bstack.lastobj = usehandler
Case "STR$", "ΓΡΑΦΗ$"
If Not IsStrExp(bstack, a$, Pad$) Then Pad$ = " "
If FastSymbol(a$, ",") Then
    If Not IsStrExp(bstack, a$, s$) Then
        s$ = "": If Pad$ = " " Then Pad$ = ","
    End If
End If
res = ""
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 2
If pppp.IsStringItem(w3) Then
    If Len(s$) > 0 Then
        res = res + """" + StringToEscapeStr(pppp.item(w3)) + """" + Pad$
    Else
    res = res + pppp.item(w3) + Pad$
    End If
Else
    If Len(s$) > 0 Then
    res = res + Num2Str(pppp.itemnumeric(w3), s$) + Pad$
    Else
    res = res + Trim$(Str$(pppp.itemnumeric(w3))) + Pad$
    End If
End If
Next
If pppp.IsStringItem(w3) Then
    If Len(s$) > 0 Then
        res = res + """" + StringToEscapeStr(pppp.item(w3)) + """"
    Else
        res = res + pppp.item(w3)
    End If
Else
If Len(s$) > 0 Then
    res = res + Num2Str(pppp.itemnumeric(w3), s$)
    Else
    res = res + Trim$(Str$(pppp.itemnumeric(w3)))
    End If
End If
End If
Case "FOLD", "ΠΑΚ", "FOLD$", "ΠΑΚ$"
If IsExp(bstack, a$, p) Then
    If Not bstack.lastobj Is Nothing Then
        Set anything = bstack.lastobj
        If FastSymbol(a$, ",") Then
            If IsExp(bstack, a$, p, , True) Then
                 res = p
            ElseIf IsStrExp(bstack, a$, Pad$, Len(bstack.tmpstr) = 0) Then
                res = Pad$
            Else
               MissParam a$
                Matrix = False
                Exit Function
            End If
            
        End If
        Set bstack.lastobj = Nothing
        CallLambdaArrayFold bstack, pppp, anything, res
    Else
        MyEr "missing a lambda function", "λείπει μια λάμδα συνάρτηση"
        Matrix = False
        Exit Function
    End If
Else
    MyEr "missing a lambda function", "λείπει μια λάμδα συνάρτηση"
    Matrix = False
    Exit Function
End If
Case "REV", "ΑΝΑΠ"
Set pppp1 = New mArray
If original > 0 Then
    pppp.CopyArrayRevFast pppp1
Else
    pppp.CopyArrayRev pppp1
End If
original = original + 1
Set pppp = pppp1
Set pppp1 = Nothing
res = 0
multi = True
Matrix = True
Set usehandler = New mHandler
usehandler.t1 = 3
Set usehandler.objref = pppp
Set bstack.lastobj = usehandler

Case "MAP", "ΑΝΤ"
againmap:
If IsExp(bstack, a$, p) Then
    If Not bstack.lastobj Is Nothing Then
    CallLambdaArrayMap bstack, pppp, bstack.lastobj
    If FastSymbol(a$, ",") Then GoTo againmap
    Else
    MyEr "missing a lambda function", "λείπει μια λάμδα συνάρτηση"
    Matrix = False
    Exit Function
End If
Else
Set pppp1 = New mArray
pppp.CopyArray pppp1
Set pppp = pppp1
original = original + 1
End If
res = 0
If FastSymbol(a$, ",") Then
    If pppp.count = 0 Then
        If IsExp(bstack, a$, p, , True) Then
             res = p: retresonly = True
        ElseIf IsStrExp(bstack, a$, Pad$, Len(bstack.tmpstr) = 0) Then
            res = Pad$: retresonly = True
        Else
           MyEr "No value", "Χωρίς τιμή"
            Matrix = False
            Exit Function
        End If
    Else
        w2 = 1
        aheadstatus a$, , w2
        If w2 > 1 Then Mid$(a$, 1, w2 - 1) = space$(w2)
    End If
End If
Matrix = True
If Not retresonly Then
Set usehandler = New mHandler
usehandler.t1 = 3
Set usehandler.objref = pppp
Set bstack.lastobj = usehandler
End If
multi = True

Case "FILTER", "ΦΙΛΤΡΟ"
again:
If IsExp(bstack, a$, p) Then
    If Not bstack.lastobj Is Nothing Then
    CallLambdaArray bstack, pppp, bstack.lastobj
    Else
    MyEr "missing a lambda function", "λείπει μια λάμδα συνάρτηση"
    Matrix = False
    Exit Function
End If
Else
Set pppp1 = New mArray
pppp.CopyArray pppp1
Set pppp = pppp1
original = original + 1
End If
res = 0
If FastSymbol(a$, ",") Then
    If pppp.count = 0 Then
        If IsExp(bstack, a$, p) Then
            If Not bstack.lastobj Is Nothing Then
            If TypeOf bstack.lastobj Is mHandler Then
            Set usehandler = bstack.lastobj
            If usehandler.t1 = 3 Then
            If TypeOf usehandler.objref Is mArray Then
                Set pppp = usehandler.objref
                GoTo again
           ' End If
            End If
            End If
            End If
            End If
             res = p: retresonly = True
        ElseIf IsStrExp(bstack, a$, Pad$, Len(bstack.tmpstr) = 0) Then
            res = Pad$: retresonly = True
        Else
           MyEr "No value", "Χωρίς τιμή"
            Matrix = False
            Exit Function
        End If
    Else
        w2 = 1
        aheadstatus a$, , w2
        If w2 > 1 Then Mid$(a$, 1, w2 - 1) = space$(w2)
    End If
End If
Matrix = True
If Not retresonly Then
Set usehandler = New mHandler
usehandler.t1 = 3
Set usehandler.objref = pppp
Set bstack.lastobj = usehandler
End If
multi = True
Case "NOTHAVE", "ΔΕΝΕΧΕΙ"
bhas = 2
GoTo jpos
Case "HAVE", "ΕΧΕΙ"
bhas = 1
GoTo jpos
Case "POS", "ΘΕΣΗ"
jpos:
    res = -1
    cur = 0
    Dim st() As String, sn() As Variant
    If IsExp(bstack, a$, R, , True) Then
        p = Int(R)
        If p < 0 Then p = 0
again1:
        If FastSymbol(a$, "->", , 2) Then
            If IsExp(bstack, a$, R) Then
                If bstack.lastobj Is Nothing Then
                    ReDim sn(0 To 4) As Variant
                    sn(0) = R
                Else
dothis:
                    Set anything = bstack.lastobj
                    Set bstack.lastobj = Nothing
                    If Not CheckLastHandlerOrIterator(anything, w3) Then Exit Function
                    Set usehandler = anything
                    If Not TypeOf usehandler.objref Is mArray Then
                        If usehandler.t1 = 4 Then Set usehandler = Nothing: Set anything = Nothing: GoTo again1
                        Exit Function
                    End If
                    Set pppp1 = usehandler.objref
            
                    sn() = pppp1.GetCopy()
                    If pppp1.count > 0 Then
                        ReDim Preserve sn(0 To pppp1.count - 1)
                    End If
                    cur = pppp1.count - 1
                    w3 = p
                End If
            ElseIf IsStrExp(bstack, a$, Pad$, Len(bstack.tmpstr) = 0) Then
                GoTo there
            Else
                MissParam a$: Exit Function
            End If
        Else
            If bstack.lastobj Is Nothing Then
'                r = p
                p = 0
                ReDim sn(0 To 4) As Variant
                sn(0) = R
            Else
                GoTo dothis
            End If
        End If
        
        If pppp.count > 0 Then
            res = -1
            If Not pppp1 Is Nothing Then
                While res = -1 And cur >= 0 And w3 < pppp.count - cur - 1
                    For w3 = w3 To pppp.count - cur - 1
                        If pppp.MyIsObject(pppp.item(w3)) Then
                             If pppp.MyIsObject(sn(0)) Then GoTo inside
                        Else
                            If pppp.MyIsObject(sn(0)) Then GoTo inside
                            If pppp.item(w3) = sn(0) Then
inside:
                                res = w3
                                w4 = w3 + 1
                                For w2 = 1 To cur
                                    If w4 < pppp.count Then
                                        If pppp.MyIsObject(pppp.item(w4)) Then
                                            If Not pppp.MyIsObject(sn(w2)) Then
                                                res = -1
                                                Exit For
                                            End If
                                        Else
                                            If pppp.MyIsObject(sn(w2)) Then
                                                res = -1
                                                Exit For
                                            Else
                                                If pppp.item(w4) <> sn(w2) Then res = -1: Exit For
                                            End If
                                        End If
                                    End If
                                    w4 = w4 + 1
                                Next w2
                                If w2 > cur Then Exit For
                            End If
                        End If
                    Next w3
                Wend
            Else
                For w3 = p To pppp.count - 1
                    If pppp.MyIsNumeric(pppp.item(w3)) Then
                        If pppp.item(w3) = R Then
                        res = w3: Exit For
                        End If
                    Else
                        If pppp.ItemType(w3) = "mHandler" Then
                            Set usehandler = pppp.item(w3)
                            If usehandler.t1 = 4 Then
                                If usehandler.index_cursor * usehandler.sign = R Then res = w3: Exit For
                            End If
                        End If
                    End If
                Next w3
                w2 = 1
                Do While FastSymbol(a$, ",")
                    If cur = UBound(sn()) Then ReDim Preserve sn(0 To cur * 2 - 1) As Variant
                    cur = cur + 1
                    If IsExp(bstack, a$, sn(cur), , True) Then
                        If res > -1 Then
                            w3 = w3 + 1
                            If w3 < pppp.count Then
                                If pppp.MyIsNumeric(pppp.item(w3)) Then
                                    If pppp.item(w3) <> sn(cur) Then w2 = -1
                                ElseIf pppp.ItemType(w3) = "mHandler" Then
                                    Set usehandler = pppp.item(w3)
                                    If usehandler.t1 = 4 Then
                                        If usehandler.index_cursor * usehandler.sign <> sn(cur) Then w2 = -1
                                    Else
                                        w2 = -1
                                    End If
                                Else
                                    w2 = -1
                                End If
                            Else
                                w2 = -1
                            End If
                        End If
                    End If
                Loop
                If w2 = -1 Then
                    w3 = res + 1
                    res = -1
                    While res = -1 And cur > 0 And w3 < pppp.count - cur - 1
                        For w3 = w3 To pppp.count - cur - 1
                            If pppp.MyIsNumeric(pppp.item(w3)) Then
                                If pppp.item(w3) = sn(0) Then
                                    res = w3
                                    w4 = w3 + 1
                                    For w2 = 1 To cur
                                        If w4 < pppp.count Then
                                            If pppp.MyIsNumeric(pppp.itemnumeric(w4)) Then
                                                If pppp.item(w4) <> sn(w2) Then res = -1: Exit For
                                            ElseIf pppp.ItemType(w4) = "mHandler" Then
                                                Set usehandler = pppp.item(w4)
                                                If usehandler.t1 = 4 Then
                                                    If usehandler.index_cursor * usehandler.sign <> sn(w2) Then res = -1
                                                Else
                                                    res = -1
                                                End If
                                            Else
                                                res = -1
                                            End If
                                        End If
                                        w4 = w4 + 1
                                    Next w2
                                    If w2 > cur Then Exit For
                                End If
                            End If
                        Next w3
                    Wend
                End If
            End If
        End If
    ElseIf IsStrExp(bstack, a$, Pad$, Len(bstack.tmpstr) = 0) Then
        
there:
        ReDim st(0 To 4) As String
        st(0) = Pad$
        If pppp.count > 0 Then
            res = -1
            For w3 = p To pppp.count - 1
                If pppp.IsStringItem(w3) Then
                    If pppp.item(w3) = st(0) Then res = w3: Exit For
                End If
            Next w3
            w2 = 1
            Do While FastSymbol(a$, ",")
                If cur = UBound(st()) Then ReDim Preserve st(0 To cur * 2 - 1) As String
                cur = cur + 1
                If IsStrExp(bstack, a$, st(cur)) Then
                    If res > -1 Then
                        w3 = w3 + 1
                        If w3 < pppp.count Then
                            If pppp.IsStringItem(w3) Then
                                If pppp.item(w3) <> st(cur) Then w2 = -1
                            Else
                                w2 = -1
                            End If
                        Else
                            w2 = -1
                        End If
                    End If
                Else
                    w2 = -1
                End If
            Loop
            If w2 = -1 Then
                w3 = res + 1
                res = -1
                While res = -1 And cur > 0 And w3 < pppp.count - cur - 1
                    For w3 = w3 To pppp.count - cur - 1
                        If pppp.IsStringItem(w3) Then
                            If pppp.item(w3) = st(0) Then
                                res = w3
                                w4 = w3 + 1
                                For w2 = 1 To cur
                                    If w4 < pppp.count Then
                                        If pppp.IsStringItem(w4) Then
                                            If pppp.item(w4) <> st(w2) Then res = -1: Exit For
                                        Else
                                            res = -1
                                        End If
                                    End If
                                    w4 = w4 + 1
                                Next w2
                                If w2 > cur Then Exit For
                            End If
                        End If
                    Next w3
                Wend
            End If
        End If
    Else
        MissParam a$
        Exit Function
    End If
    If bhas > 0 Then
        If bhas = 1 Then
            res = res <> -1
        Else
            res = res = -1
        End If
    End If
End Select
Matrix = FastSymbol(a$, ")")
If Not multi Then Exit Do
If Matrix = False Then Exit Do
If Not IsOperator(a$, "#") Then Exit Do
If pppp.count = 0 Then
        Do
        w2 = 1
        aheadstatus a$, , w2
        If w2 > 1 Then Mid$(a$, 1, w2 - 1) = space$(w2)
        If Not FastSymbol(a$, ")") Then Matrix = False: Exit Function
        Loop Until Not IsOperator(a$, "#")
Exit Function
End If
Pad$ = myUcase(Left$(a$, 20))
cut = InStr(Pad$, "(")
If cut <= 1 Then Exit Do
Mid$(a$, 1, cut) = space$(cut)
Set bstack.lastobj = Nothing

Loop
Else
WrongObject
End If
End Function
Function getone(bstack As basetask, rest$) As Boolean
Dim what$, ss$, x1 As Long
getone = True
FastSymbol rest$, "&"
   x1 = Abs(IsLabelBig(bstack, rest$, what$))
    
    If x1 <> 0 Then
            If x1 > 4 Then
                    ss$ = BlockParam(rest$)
                    what$ = what$ + ss$ + ")"
                    'rest$ = Mid$(rest$, Len(ss$) + 2)
                    Mid$(rest$, 1, Len(ss$) + 1) = space(Len(ss$) + 1)
                    Do While IsSymbol(rest$, ".")
                    x1 = IsLabel(bstack, rest$, ss$)
                    If x1 > 0 Then what$ = what$ + "." + ss$ Else Exit Do
                            If x1 > 4 Then
                            ss$ = BlockParam(rest$)
                            what$ = what$ + ss$ + ")"
                            'rest$ = Mid$(rest$, Len(ss$) + 2)
                            Mid$(rest$, 1, Len(ss$) + 1) = space(Len(ss$) + 1)
                            End If
                    Loop
            End If
    
            



              
              getone = MyRead(6, bstack, (what$), 1, what$, x1, True)

             Else
             MissParamref rest$
             Exit Function
             End If
  



End Function
Sub CallLambdaArray(bstack As basetask, ByRef pppp As mArray, mylambda As lambda)
Dim w2 As Long, w1 As Long, nbstack As basetask
PushStage bstack, False
w2 = var2used
w1 = globalvarGroup("A_" + CStr(w2), 0#)
 Set var(w1) = mylambda
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w1
    End If
     Set nbstack = New basetask
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
Dim aa As Object, oldsoros As mStiva, tempsoros As New mStiva, finalpppp As New mArray
finalpppp.StartResize: finalpppp.PushDim pppp.count: finalpppp.PushEnd
Set oldsoros = bstack.soros
Set bstack.Sorosref = tempsoros
Dim R, what As Long, where As Long
For w1 = 0 To pppp.count - 1
 
  If pppp.IsStringItem(w1) Then
    tempsoros.PushStrVariant pppp.item(w1)
    what = 1
  ElseIf pppp.MyIsObject(pppp.item(w1)) Then
    tempsoros.PushObj pppp.item(w1)
    what = 2
  Else
      tempsoros.PushVal pppp.item(w1)
      what = 3
  End If
  If Not GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "()", vbNullString, R, w2, , , True) Then Exit For
  If CBool(R) Then
  If what <> 2 Then
  finalpppp.item(where) = pppp.item(w1)
 
  Else
   Set finalpppp.item(where) = pppp.item(w1)
  End If
  where = where + 1

  End If
  tempsoros.Flush
Next w1
Set bstack.Sorosref = oldsoros
PopStage bstack
finalpppp.StartResize: finalpppp.PushDim where: finalpppp.PushEnd
Set pppp = finalpppp

End Sub

Sub CallLambdaArrayMap(bstack As basetask, ByRef pppp As mArray, mylambda As lambda)
Dim w2 As Long, w1 As Long, nbstack As basetask
PushStage bstack, False
w2 = var2used
w1 = globalvarGroup("A_" + CStr(w2), 0#)
 Set var(w1) = mylambda
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w1
    End If
     Set nbstack = New basetask
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
Dim aa As Object, oldsoros As mStiva, tempsoros As New mStiva, finalpppp As New mArray
finalpppp.StartResize: finalpppp.PushDim pppp.count: finalpppp.PushEnd
Set oldsoros = bstack.soros
Set bstack.Sorosref = tempsoros
Dim R, what As Long, where As Long
For w1 = 0 To pppp.count - 1
 
  If pppp.IsStringItem(w1) Then
    tempsoros.PushStrVariant pppp.item(w1)
  ElseIf pppp.MyIsObject(pppp.item(w1)) Then
    tempsoros.PushObj pppp.item(w1)
  Else
      tempsoros.PushVal pppp.item(w1)
  End If
  If Not GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "()", vbNullString, R, w2, , , True) Then Exit For
  If tempsoros.count > 0 Then
  If tempsoros.StackItemTypeIsObject(1) Then
  Set finalpppp.item(w1) = tempsoros.PopObj
  Else
   finalpppp.item(w1) = tempsoros.PopAnyNoObject
  End If
    End If
  
  tempsoros.Flush
Next w1
Set bstack.Sorosref = oldsoros
PopStage bstack
Set pppp = finalpppp

End Sub
Sub CallLambdaArrayFold(bstack As basetask, pppp As mArray, mylambda As lambda, res As Variant)
Dim w2 As Long, w1 As Long, nbstack As basetask
PushStage bstack, False
w2 = var2used
w1 = globalvarGroup("A_" + CStr(w2), 0#)
 Set var(w1) = mylambda
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w1
    End If
     Set nbstack = New basetask
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
Dim aa As Object, oldsoros As mStiva, tempsoros As New mStiva
Set oldsoros = bstack.soros
Set bstack.Sorosref = tempsoros
Dim R, what As Long, where As Long
If pppp.MyIsNumeric(res) Then
tempsoros.PushVal res
ElseIf pppp.MyIsObject(res) Then
Set aa = res
tempsoros.PushObj aa
Else
tempsoros.PushStrVariant res
End If

For w1 = 0 To pppp.count - 1

  If pppp.IsStringItem(w1) Then
    tempsoros.PushStrVariant pppp.item(w1)
  ElseIf pppp.MyIsObject(pppp.item(w1)) Then
    tempsoros.PushObj pppp.item(w1)
  Else
      tempsoros.PushVal pppp.item(w1)
  End If
  If Not GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "()", vbNullString, R, w2, , , True) Then Exit For
  
Next w1
  If tempsoros.count > 0 Then
  If tempsoros.StackItemTypeIsObject(1) Then
        Set bstack.lastobj = tempsoros.PopObj
        res = 0
  Else
        res = tempsoros.PopAnyNoObject
  End If
    End If
Set bstack.Sorosref = oldsoros
PopStage bstack
End Sub

Function ChangeValues(bstack As basetask, rest$) As Boolean
Dim aa As mHandler, bb As FastCollection, ah As String, p As Variant, s$, lastindex As Long
Set aa = bstack.lastobj
Set bstack.lastobj = Nothing
Set bb = aa.objref
If bb.StructLen > 0 Then
MyEr "Structure members are ReadOnly", "Τα μέλη της δομής είναι μόνο για ανάγνωση"
Exit Function
End If
If bb.Done And FastSymbol(rest$, ":=", , 2) Then
        ' change one value
        ah = aheadstatus(rest$, False) + " "
        If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
            If Not IsExp(bstack, rest$, p) Then
                ChangeValues = False
                GoTo there
            End If
            ChangeValues = True
            If Not bstack.lastobj Is Nothing Then
                Set bb.ValueObj = bstack.lastobj
                Set bstack.lastobj = Nothing
            Else
                bb.Value = p
            End If
            
        ElseIf Left$(ah, 1) = "S" Then
            If Not IsStrExp(bstack, rest$, s$) Then
                ChangeValues = False
                GoTo there
            End If
            ChangeValues = True
            If Not bstack.lastobj Is Nothing Then
                Set bb.ValueObj = bstack.lastobj
                Set bstack.lastobj = Nothing
            Else
                bb.Value = s$
            End If
        Else
                MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
                ChangeValues = False
        End If
        GoTo there

ElseIf lookOne(rest$, ",") Then
    Do While FastSymbol(rest$, ",")
        ChangeValues = True
        ah = aheadstatus(rest$, False) + " "
        If InStr(ah, "l") Then
                MyEr "Found logical expression", "Βρήκα λογική έκφραση"
                ChangeValues = False
        Else
                If Left$(ah, 1) = "N" Then
                    If Not IsExp(bstack, rest$, p) Then
                        ChangeValues = False
                        GoTo there
                    End If
                        If VarType(p) = vbBoolean Then p = CLng(p)
                    If Not bstack.lastobj Is Nothing Then
                        MyEr "No Object Allowed for Key", "Δεν επιτρέπεται αντικείμενο για κλειδί"
                        ChangeValues = False
                        GoTo there
                    End If
                    If Not bb.Find(p) Then
                         MyEr "No Key found", "Δεν βρέθηκε κλειδί"
                        ChangeValues = False
                        GoTo there
                    End If
                    
                ElseIf Left$(ah, 1) = "S" Then
                    If Not IsStrExp(bstack, rest$, s$) Then
                        ChangeValues = False
                        GoTo there
                    End If
                    If Not bstack.lastobj Is Nothing Then
                        MyEr "No Object Allowed for Key", "Δεν επιτρέπεται αντικείμενο για κλειδί"
                        ChangeValues = False
                        GoTo there
                    End If
                    If Not bb.Find(s$) Then
                          MyEr "No Key found", "Δεν βρέθηκε κλειδί"
                        ChangeValues = False
                        GoTo there
                    End If
                Else
                        MyEr "No Key found", "Δεν βρέθηκε κλειδί"
                        ChangeValues = False
                        GoTo there
                
                End If
lastindex = bb.Index
                
                If FastSymbol(rest$, ":=", , 2) Then
                    ah = aheadstatus(rest$, False) + " "
                    If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
                        If Not IsExp(bstack, rest$, p) Then
                            ChangeValues = False
                            GoTo there
                        End If
                        ChangeValues = True
                        bb.Index = lastindex
                        If Not bstack.lastobj Is Nothing Then
                            Set bb.ValueObj = bstack.lastobj
                            Set bstack.lastobj = Nothing
                        Else
                            bb.Value = p
                        End If
                ElseIf Left$(ah, 1) = "S" Then
                        If Not IsStrExp(bstack, rest$, s$) Then
                            ChangeValues = False
                            GoTo there
                        End If
                        ChangeValues = True
                        bb.Index = lastindex
                        If Not bstack.lastobj Is Nothing Then
                            Set bb.ValueObj = bstack.lastobj
                            Set bstack.lastobj = Nothing
                        Else
                            bb.Value = s$
                        End If
                Else
                        MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
                        ChangeValues = False
                End If
                End If
        End If
    Loop
ElseIf bb.Done Then
If Not bstack.soros.IsEmpty Then
With bstack.soros
If .StackItemTypeIsObject(1) Then
Set bb.ValueObj = .PopObj

Else
bb.Value = .StackItem(1)
.drop 1
End If
End With
ChangeValues = True
Else

End If
Set bstack.lastobj = Nothing
End If


there:
Set bb = Nothing
Set aa = Nothing
End Function
Function ChangeValuesArray(bstack As basetask, rest$) As Boolean
Dim aa As mHandler, p As Variant, pppp As mArray, w As Long, s$, ah As String, stiva As mStiva
Dim bs As Long
Set aa = bstack.lastobj
Dim anything As Object
Set anything = aa
If CheckIsmArray(anything) Then
    Set pppp = anything
    Set anything = Nothing
    If Not pppp.Arr Then
        NotArray
        Exit Function
    End If
    bs = pppp.myarrbase
    FastSymbol rest$, ","
    Do
    If IsExp(bstack, rest$, p, , True) Then
    On Error Resume Next
    w = CLng(Fix(p)) + bs
    If Err Then
        Err.Clear
        OutOfLimit
        Exit Function
    End If
    On Error GoTo 0
    If w < 0 Then w = pppp.count - w + bs
    If w > (pppp.count + bs) Then GoTo outlimit
    If w < 0 Then
outlimit:
    MyEr "Index out of limits", "Ο δείκτης είναι εκτός ορίων"
    Exit Function
    End If
    
    If FastSymbol(rest$, ":=", , 2) Then
            
            ah = aheadstatus(rest$, False) + " "
            If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
                If Not IsExp(bstack, rest$, p) Then
                    ChangeValuesArray = False
                    GoTo there
                End If
                ChangeValuesArray = True
                If Not bstack.lastobj Is Nothing Then
                    Set pppp.item(w) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                Else
                    pppp.item(w) = p
                End If
                
            ElseIf Left$(ah, 1) = "S" Then
                If Not IsStrExp(bstack, rest$, s$) Then
                    ChangeValuesArray = False
                    GoTo there
                End If
                ChangeValuesArray = True
                If Not bstack.lastobj Is Nothing Then
                    Set pppp.item(w) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                Else
                    pppp.item(w) = s$
                End If
            Else
                    MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
                    ChangeValuesArray = False
            End If
    End If
    End If
    Loop Until Not FastSymbol(rest$, ",")
Else
Set bstack.lastobj = aa
If CheckStackObj(bstack, anything) Then
    Set stiva = anything
    Set anything = Nothing
    FastSymbol rest$, ","
    Do
    If IsExp(bstack, rest$, p, , True) Then
    On Error Resume Next
    w = CLng(Fix(p))
    If Err Then
        Err.Clear
        OutOfLimit
        Exit Function
    End If
    On Error GoTo 0
    If w < 0 Then w = stiva.count - w + 1
    If w > stiva.count Then GoTo outlimit
    If w < 0 Then GoTo outlimit
   
    If FastSymbol(rest$, ":=", , 2) Then
            
            ah = aheadstatus(rest$, False) + " "
            If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
                If Not IsExp(bstack, rest$, p) Then
                    ChangeValuesArray = False
                    GoTo there
                End If
                ChangeValuesArray = True
                If Not bstack.lastobj Is Nothing Then
                    stiva.MakeTopItem w
                    stiva.drop 1
                    stiva.PushObj bstack.lastobj
                    stiva.MakeTopItemBack w
                    Set bstack.lastobj = Nothing
                Else
                    stiva.MakeTopItem w
                    stiva.drop 1
                    stiva.PushVal p
                    stiva.MakeTopItemBack w
                End If
                
            ElseIf Left$(ah, 1) = "S" Then
                If Not IsStrExp(bstack, rest$, s$) Then
                    ChangeValuesArray = False
                    GoTo there
                End If
                ChangeValuesArray = True
                If Not bstack.lastobj Is Nothing Then
                    stiva.MakeTopItem w
                    stiva.drop 1
                    stiva.PushObj bstack.lastobj
                    stiva.MakeTopItemBack w
                    Set bstack.lastobj = Nothing
                Else
                    stiva.MakeTopItem w
                    stiva.drop 1
                    stiva.PushStr s$
                    stiva.MakeTopItemBack w
                End If
            Else
                    MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
                    ChangeValuesArray = False
            End If
    End If
    End If
    Loop Until Not FastSymbol(rest$, ",")
End If
End If
there:
Set anything = Nothing
Set aa = Nothing
End Function
Sub FeedCopyInOut(bstack As basetask, var$, where As Long, Arr$)
Dim a  As Variant
a = Array(var$, Arr$, where)
If bstack.CopyInOutCol Is Nothing Then Set bstack.CopyInOutCol = New Collection
bstack.CopyInOutCol.Add a
End Sub
Sub CopyBack(bstack As basetask)
Dim a As Variant, aa As Object, x1 As Long, what$, rest$, oldhere$, w2 As Long, ii As Long
Dim pppp As mArray, s As String
Dim actualvar As String, ArrArg As String, LocalVar As String
oldhere$ = here$
here$ = vbNullString
again:
For Each a In bstack.CopyInOutCol
actualvar = a(0)
ArrArg = a(1)
LocalVar = a(2)
If Len(ArrArg) > 0 Then
x1 = rinstr(actualvar, ArrArg) + Len(ArrArg) - 1
JUMPHERE:
    If neoGetArray(bstack, Left$(actualvar, x1), pppp) Then
        If Not pppp.Arr Then GoTo cont123
        If Not NeoGetArrayItem(pppp, bstack, Left$(actualvar, x1), w2, Mid$(actualvar, x1 + 1)) Then GoTo cont123
        If MyIsObject(var(LocalVar)) Then
        If TypeOf pppp.itemObject(w2) Is Group Then
        If pppp.item(w2).IamApointer Then
            Set pppp.item(w2) = var(LocalVar)
        Else
        Set pppp.item(w2) = CopyGroupObj(var(LocalVar), Not pppp.GroupRef Is Nothing)
        End If
        Else
            Set pppp.item(w2) = var(LocalVar)
            End If
        Else
            pppp.item(w2) = var(LocalVar)
        End If
    End If
Else
    If bstack.ExistVar2(actualvar) Then
        
        If MyIsObject(var(LocalVar)) Then
            bstack.SetVarobJ actualvar, var(LocalVar)
        Else
            bstack.SetVar actualvar, var(LocalVar)
        End If
    Else
    ii = 1
    If FastPureLabel(actualvar, what$, ii, , , , False) > 4 Then
    x1 = ii - 1 'rinstr(actualvar, what$) + Len(what$) - 1
    GoTo JUMPHERE
    
    End If
    
    End If
End If
cont123:
Next
here$ = oldhere$
Set bstack.CopyInOutCol = Nothing

End Sub
Function GetOneAsString(bstack As basetask, rest$, what$, x1 As Long) As Boolean
Dim ss$
            If x1 > 4 Then
                    ss$ = BlockParam(rest$)
                    If Mid$(rest$, Len(ss$) + 1, 1) <> ")" Then Exit Function
                    what$ = what$ + ss$ + ")"
                   rest$ = Mid$(rest$, Len(ss$) + 2)
                    GetOneAsString = True
            End If
    
End Function
Function NewVarItem() As VarItem
    If TrushCount = 0 Then
    Set NewVarItem = New VarItem
      Exit Function
    End If
    Set NewVarItem = Trush(TrushCount)
    Set Trush(TrushCount) = Nothing
    TrushCount = TrushCount - 1
End Function
Function ExpMatrix(bstack As basetask, a$, R) As Boolean
Dim usehandler As mHandler
 If Not bstack.lastobj Is Nothing Then
                                If Typename(bstack.lastobj) = "mHandler" Then
                                    Set usehandler = bstack.lastobj
                                    Set bstack.lastobj = Nothing
                                    ExpMatrix = Matrix(bstack, a$, usehandler, R)
                                    If MyIsObject(R) Then R = CDbl(0)
                                   ' If SG < 0 Then r = -r
                                    Exit Function
                                ElseIf Typename(bstack.lastobj) = "mArray" Then
                                Set usehandler = New mHandler
                                usehandler.t1 = 3
                                Set usehandler.objref = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                    ExpMatrix = Matrix(bstack, a$, usehandler, R)
                                    If MyIsObject(R) Then R = CDbl(0)
                                   ' If SG < 0 Then r = -r
                                    Exit Function
                                End If
                            End If
                                SyntaxError
                                ExpMatrix = False
                                Exit Function
End Function
Sub targetsMyExec(MyExec As Long, b$, bb$, v As Long, di As Object, w$, bstack As basetask, VarStat As Boolean, temphere$)
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, SBB$, nd&, p As Variant
Dim notglobal As Boolean, GuiForm As GuiM2000
notglobal = TypeOf bstack.Owner Is GuiM2000

     If Abs(IsLabel(bstack, b$, w$)) = 1 Then
         If Not GetVar(bstack, w$, v) Then
             v = globalvar(w$, 0#, , VarStat, temphere$)
         Else
             If var(v) >= 330000 And Not notglobal Then
                MyEr "wrong target handler", "λάθος χειριστής στόχου"
                MyExec = 0
                Exit Sub
             End If
         End If
     Else
         MyExec = 0
         Exit Sub
     End If

     If Not FastSymbol(b$, ",") Then
         MyExec = 0
         Exit Sub
     ElseIf IsStrExp(bstack, b$, bb$) Then
     If NocharsInLine(bb$) Then MyExec = 0: Exit Sub
     With players(GetCode(di))
        x1 = 1
        y1 = 1
        x2 = &H81000000
        y2 = &H81000000
        nd& = 0
        SBB$ = vbNullString
        On Error GoTo err123
        If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then x1 = Abs(p) Mod (.mx + 1)
        If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then y1 = Abs(p) Mod (.My + 1)
        If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then x2 = CLng(Fix(p))
        If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then y2 = CLng(Fix(p))
        If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then nd& = Abs(p)
        If FastSymbol(b$, ",") Then If Not IsStrExp(bstack, b$, SBB$) Then MyExec = 0: Exit Sub
err123:
        If Err.Number = 6 Then
            Overflow
            MyExec = 0
            Exit Sub
        End If
    
        If Not notglobal Then
            Targets = False
            MyDoEvents1 Form1
            ReDim Preserve q(UBound(q()) + 1)
            q(UBound(q()) - 1) = BoxTarget(bstack, x1, y1, x2, y2, SBB$, nd&, bb$, .Xt, .Yt, .uMineLineSpace)
            var(v) = UBound(q()) - 1
            Targets = True
        Else
            Set GuiForm = di
            var(v) = GuiForm.AddTarget(BoxTarget(bstack, x1, y1, x2, y2, SBB$, nd&, bb$, .Xt, .Yt, .uMineLineSpace))
        End If
    End With
ElseIf IsExp(bstack, b$, p, , True) Then
    If Not notglobal Then
    If var(v) < 320000 Then
    
        q(var(v)).Enable = Not (p = 0)
        RTarget bstack, q(var(v))
        End If
    Else
    Set GuiForm = di
    GuiForm.EnableTarget bstack, var(v), p
    End If
Else
    MyExec = 0
    Exit Sub
End If
End Sub

Function ProcUSE(basestack As basetask, rest$, Lang As Long) As Boolean
Dim ss$, ML As Long, x As Double, pa$, s$, stac1$, p As Variant, frm$, i As Long, w$, pppp As mArray
Dim it As Long, what$
If IsStrExp(basestack, rest$, ss$) Then   'gsb
ElseIf Not Abs(IsLabel(basestack, rest$, ss$)) = 1 Then ' WITHOUT " .gsb"
SyntaxError
Exit Function
End If
ML = 0
If UCase(ss$) = "PIPE" Or UCase(ss$) = "ΑΥΛΟΥ" Then
ML = 1
End If

stac1$ = vbNullString
If FastSymbol(rest$, "!") And ML <> 1 Then
If VALIDATEpart(rest$, s$) Then
Do While s$ <> ""
    If ISSTRINGA(s$, pa$) Then
        basestack.soros.DataStr pa$
    ElseIf IsNumberD2(s$, x) Then
        basestack.soros.DataVal x
        x = vbEmpty
    Else
        Exit Do
    End If
Loop
Else
SyntaxError
Exit Function
End If
Else
If ML <> 1 Then
    Do
        If IsExp(basestack, rest$, p) Then
        stac1$ = stac1$ & Str$(p)
        ElseIf IsStrExp(basestack, rest$, s$) Then
         stac1$ = stac1$ & Sput(s$)
        Else
        Exit Do
        End If
        If Not FastSymbol(rest$, ",") Then Exit Do
    Loop
    pa$ = ExtractPath(ss$)
    para$ = RTrim$(".gsb " & Mid$(ss$, Len(ExtractPath(ss$) + ExtractName(ss$)) + 1))
    If pa$ = vbNullString Then pa$ = mcd
    frm$ = ExtractNameOnly(ss$)
    End If
End If

If Not IsLabelSymbolNew(rest$, "ΣΤΟ", "TO", Lang) Then

w$ = "S" & CStr(Int(Rnd(12) * 100000))

Else

Select Case Abs(IsLabel(basestack, rest$, ss$))

Case 3
    If GetVar(basestack, ss$, i) Then
       w$ = "V" & CStr(i)
       s$ = frm$
       frm$ = var(i)
       var(i) = vbNullString
      Else
     i = globalvar(ss$, "")
             If i <> 0 Then
              w$ = "V" & CStr(i)
              
            var(i) = vbNullString
            End If
                        
     End If
Case 6
   
     If neoGetArray(basestack, ss$, pppp) Then
            If Not NeoGetArrayItem(pppp, basestack, ss$, it, rest$) Then
        MyEr "Not such index for array", "Περίμενα σωστούς δείκτες για πίνακα"
        
        Exit Function
        End If
     Else
     MyEr "Not such array, need to DIM fisrt", "Περίμενα πίνακα, πρέπει να ορίσεις έναν"
: Exit Function
     End If
    
    w$ = "A" & CopyArrayItems(basestack, ss$) + Str(it) ''''''''''εδω για τον νεο πίνακα πρέπει να δώσω το mArray???
    s$ = frm$
    frm$ = pppp.item(it)
    If pppp.ItemType(it) = doc Then
    Set pppp.item(it) = New Document
    Else
     pppp.item(it) = vbNullString
    End If
   
    Case Else
    SyntaxError
: Exit Function
   End Select
   If Left$(w$, 1) <> "S" Then

p = GetTaskId + 10000 ' starts from 10000
If Not IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then
's$ = validpipename(ss$)

If frm$ <> "" Then

ss$ = frm$
Else

ss$ = "M" & CStr(p)
End If

Thing w$, validpipename(ss$)
sThread CLng(p), 0, ss$, w$
TaskMaster.Message CLng(p), 3, CLng(100)
Exit Function
Else
Select Case Abs(IsLabel(basestack, rest$, what$))
Case 0 ' TAKE A NUMBER
If IsNumberLabel(rest$, what$) Then
frm$ = "S" + Right$("0000" + what$, 5)
p = val(what$)
s$ = frm$
If Left$(w$, 1) = "V" Then var(val(Mid$(w$, 2))) = validpipename(frm$)
Else
MyEr "No number found (5 digits)", "Δεν βρήκα αριθμό (5 ψηφία)"
Exit Function
End If
Case 1
    If GetVar(basestack, what, i) Then
    If var(i) < 10000 Then var(i) = p Else p = var(i)
      Else
      globalvar what, p
                             
     End If
Case 5, 7
   
     If neoGetArray(basestack, what, pppp) Then
        If Not NeoGetArrayItem(pppp, basestack, ss$, it, rest$) Then
        MyEr "Not such index for array", "Περίμενα σωστούς δείκτες για πίνακα"
      
        Exit Function
        End If
     Else
     MyEr "Not such array, need to DIM fisrt", "Περίμενα πίνακα, πρέπει να ορίσεις έναν"
     
      Exit Function
     End If
     If pppp.item(it) < 10000 Then pppp.item(it) = p Else p = pppp.item(it)
    Case Else
    MyEr "Wrong parameter", "Λάθος παράμετρος"
     Exit Function
   End Select
   End If
End If
'ss$ = validpipename("M" & CStr(p))
'stac1$ = Sput(ss$) + stac1$
If frm$ <> "" Then
ss$ = frm$
Else
ss$ = "M" & CStr(p)
End If
frm$ = s$
sThread CLng(p), 0, ss$, w$
TaskMaster.Message CLng(p), 3, CLng(100)
ss$ = validpipename(ss$)
stac1$ = Sput(ss$) + stac1$
ss$ = "M" & CStr(p)
End If
If ML <> 1 Then
If stac1$ = vbNullString And Left$(s$, 1) = "S" Then
s$ = App.Path
AddDirSep s$
s$ = s$ & "M2000.EXE "
If Shell(s$ & Chr(34) + pa$ & frm$ & ".gsb" & para$ & Chr(34), vbNormalFocus) > 0 Then
End If
End If
If Left$(w$, 1) = "V" Then
ss$ = GetTag$ & ".gsb"
Else
ss$ = w$ & ".gsb"
End If
i = FreeFile
On Error Resume Next
 If Not NeoUnicodeFile(strTemp + ss$) Then
 MyEr "can't save " + strTemp + ss$, "δεν μπορώ να σώσω " + strTemp + ss$
what$ = vbNullString
 Exit Function
End If

Open GetDosPath(strTemp + ss$) For Output As i
If Err.Number > 0 Then
InternalEror
what$ = vbNullString
Exit Function
End If
If stac1$ <> "" Then

' look for unicode...
Print #i, "STACK !" & stac1$ & ": DIR " & Chr(34) + pa$ & Chr(34) & " : LOAD " & Chr(34) + frm$ & para$ & Chr(34)

Else
Print #i, "DIR " & Chr(34) + pa$ & Chr(34) & " : LOAD " & Chr(34) + frm$ & para$ & Chr(34)

End If
Close i
tempList2delete = Sput(strTemp + ss$) + tempList2delete
s$ = App.Path
AddDirSep s$
s$ = s$ & "M2000.EXE "
LastUse = MyShell(s$ & Chr(34) + strTemp + ss$ & Chr(34), vbNormalFocus - 4 * (ML <> 0 Or IsSymbol(rest$, ";")))
Sleep 1
If LastUse <> 0 Then

If ML = 0 Then
If IsSymbol(rest$, ";") Then
Else
'AppActivate LastUse
End If
End If
'killfile strTemp + ss$
End If
End If
ProcUSE = True
End Function

Function AddInventory(bstack As basetask, rest$, Optional ret2logical As Boolean = False) As Boolean
Dim p As Variant, s$, pppp As mArray, lastindex As Long, usehandler As mHandler
If Not bstack.lastobj Is Nothing Then
If Typename(bstack.lastobj) = "mHandler" Then
Dim aa As mHandler
Set aa = bstack.lastobj
Set bstack.lastobj = Nothing
If Not aa.objref Is Nothing Then
If TypeOf aa.objref Is FastCollection Then
Dim bb As FastCollection
Set bb = aa.objref
If bb.StructLen > 0 Then
MyEr "Structure members are ReadOnly", "Τα μέλη της δομής είναι μόνο για ανάγνωση"
Exit Function
End If

Dim ah As String
FastSymbol rest$, ","
again:
AddInventory = True
ah = aheadstatus(rest$, False) + " "
If InStr(ah, "l") Then
If ret2logical Then ret2logical = False: Exit Function
MyEr "No logical expression", "Όχι λογική έκφραση"
AddInventory = False
Else
If Left$(ah, 1) = "N" Then
    If Not IsExp(bstack, rest$, p) Then
        AddInventory = False
        GoTo there
    End If
    If VarType(p) = vbBoolean Then p = CLng(p)
    If Not bstack.lastobj Is Nothing Then
        If TypeOf bstack.lastobj Is mHandler Then
        Set usehandler = bstack.lastobj
        If usehandler.t1 = 4 Then
        Set bstack.lastobj = Nothing
        GoTo noenum
        End If
        End If
        MyEr "No Object Allowed for Key", "Δεν επιτρέπεται αντικείμενο για κλειδί"
        AddInventory = False
        GoTo there
    End If
noenum:
    If bb.ExistKey0(p) Then
        MyEr "Key exist, must be unique", "Το κλειδί υπάρχει, πρέπει να είναι μοναδικό"
        AddInventory = False
        GoTo there
    End If
    bb.AddKey p
ElseIf Left$(ah, 1) = "S" Then
    If Not IsStrExp(bstack, rest$, s$) Then
        AddInventory = False
        GoTo there
    End If
    If Not bstack.lastobj Is Nothing Then
        MyEr "No Object Allowed for Key", "Δεν επιτρέπεται αντικείμενο για κλειδί"
        AddInventory = False
        GoTo there
    End If
    If bb.ExistKey0(s$) Then
        MyEr "Key exist, must be unique", "Το κλειδί υπάρχει, πρέπει να είναι μοναδικό"
        AddInventory = False
        GoTo there
    End If
    bb.AddKey s$

Else
        MyEr "No Key found", "Δεν βρέθηκε κλειδί"
        AddInventory = False
        GoTo there

End If
lastindex = bb.Index
If FastSymbol(rest$, ":=", , 2) Then
ah = aheadstatus(rest$, False) + " "
If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
    If Not IsExp(bstack, rest$, p) Then
        AddInventory = False
        GoTo there
    End If
    bb.Index = lastindex
    If Not bstack.lastobj Is Nothing Then
    If TypeOf bstack.lastobj Is mArray Then
        Set pppp = New mArray
        bstack.lastobj.CopyArray pppp
        Set bb.ValueObj = pppp
        Set pppp = Nothing
        
    Else
      Set bb.ValueObj = bstack.lastobj
    End If
        Set bstack.lastobj = Nothing
        If TypeOf bb.ValueObj Is Group Then
        bb.ValueObj.ToDelete = False
        End If
    Else
        bb.Value = p
    End If
    
ElseIf Left$(ah, 1) = "S" Then
    If Not IsStrExp(bstack, rest$, s$) Then
        AddInventory = False
        GoTo there
    End If
    bb.Index = lastindex
    If Not bstack.lastobj Is Nothing Then
    If TypeOf bstack.lastobj Is mArray Then
        Set pppp = New mArray
        bstack.lastobj.CopyArray pppp
        Set bb.ValueObj = pppp
        Set pppp = Nothing
    Else
    
        Set bb.ValueObj = bstack.lastobj
        
    End If
        Set bstack.lastobj = Nothing
             If TypeOf bb.ValueObj Is Group Then
        bb.ValueObj.ToDelete = False
        End If
    Else
        bb.Value = s$
    End If

Else
        MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
        AddInventory = False
        GoTo there

End If


End If


End If
If FastSymbol(rest$, ",") Then GoTo again
there:
Set bb = Nothing
Set aa = Nothing

Exit Function
ElseIf TypeOf aa.objref Is mArray Then
While IsSymbol(rest$, ",")
If Not IsExp(bstack, rest$, p) Then
    MyEr "Expected Array", "Περίμενα Πίνακα"
    Set aa = Nothing
    Set bstack.lastobj = Nothing
    Exit Function
End If
Dim myobject As Object
Set myobject = bstack.lastobj
Set bstack.lastobj = Nothing
If CheckIsmArray(myobject) Then
Set pppp = myobject
pppp.AppendArray aa.objref
Else
    MyEr "Expected Array", "Περίμενα Πίνακα"
    Set aa = Nothing
    Set bstack.lastobj = Nothing
    Exit Function
End If
Wend
Set aa = Nothing
AddInventory = True
Exit Function
End If
End If
End If
MyEr "Wrong type of object (not Inventory or pointer to Array)", "Λάθος τύπος αντικειμένου (όχι Κατάσταση ή δείκτης σε Πίνακα)"
Set aa = Nothing
End If
Set bstack.lastobj = Nothing
End Function
Sub OnlyForInventory()
MyEr "Only for Inventory object", "Μόνο για αντικείμενο Κατάσταση"
End Sub
Function NewInventory(bstack As basetask, rest$, R, Queue As Boolean) As Boolean
            Dim serr As Boolean, usehandler As mHandler
            
                    MakeitObjectInventory R, Queue
                    Set usehandler = R
                    If Queue Then usehandler.objref.AllowAnyKey
                    Set bstack.lastobj = usehandler
                    If FastSymbol(rest$, ":=", , 2) Then
                    If AddInventory(bstack, rest$, serr) Then
                            Set bstack.lastobj = R
                        R = 0
                        NewInventory = True
                    End If
                    Else
                        Set bstack.lastobj = R
                        R = 0
                        NewInventory = True
                    End If
                    
End Function
Function IsCdate(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim pp As Variant, par As Boolean, r2 As Variant, r3 As Variant, r4 As Variant
   If IsExp(bstack, a$, R, , True) Then
    pp = Abs(R) - Fix(Abs(R))
    R = Abs(R Mod 2958466)
    par = True
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r2, , True)
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r3, , True) And par
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r4, , True) And par
    
    End If
    End If
    End If
    
    If Not par Then
     MissParam a$
     Exit Function
                End If
                On Error Resume Next
    r3 = r3 + (r2 - Int(r2)) * 365
    r2 = Int(r2)
    r4 = r4 + (r3 - Int(r3)) * 30
    r3 = Int(r3)
     R = CDbl(DateSerial(Year(R) + r2, Month(R) + r3, Day(R) + r4) + pp)
     If SG < 0 Then R = -R
              If Err.Number > 0 Then
    WrongArgument a$
    Err.Clear
    Exit Function
    End If
    On Error GoTo 0
 IsCdate = FastSymbol(a$, ")", True)
   Else
   
     MissParam a$
    
    End If
    
End Function
Function IsTimeVal(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim s$
    If IsStrExp(bstack, a$, s$) Then
    On Error Resume Next
    If UCase(s$) = "UTC" Then
    R = CDbl(GetUTCTime)
    R = R - Int(R)
    Else
    R = CDbl(CDate(TimeValue(s$)))
    End If
    If SG < 0 Then R = -R
         If Err.Number > 0 Then
    
    WrongArgument a$
    Err.Clear
    Exit Function
    End If
        On Error GoTo 0
    
    
    Else
     Dim usehandler As mHandler
     Set usehandler = New mHandler
     usehandler.t1 = 1
     usehandler.ReadOnly = True
     Set usehandler.objref = zones
        Set bstack.lastobj = usehandler
     R = R - R
    End If
IsTimeVal = FastSymbol(a$, ")", True)
End Function
Function IsDataVal(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
 Dim s$, p
    If IsStrExp(bstack, a$, s$) Then
    If FastSymbol(a$, ",") Then
    If Not IsExp(bstack, a$, p, , True) Then
        p = Clid
    End If
    On Error Resume Next
    R = CDbl(DateFromString(s$, p))
    If SG < 0 Then R = -R
     If Err.Number > 0 Then
    
    WrongArgument a$
    Err.Clear
    Exit Function
    End If
    Else
    On Error Resume Next
    If s$ = "UTC" Then
    R = CDbl(Int(GetUTCDate))
    Else
    R = CDbl(DateValue(s$))
    End If
    If SG < 0 Then R = -R
     If Err.Number > 0 Then
    
    WrongArgument a$
    Err.Clear
    Exit Function
    End If
    End If
    On Error GoTo 0
    
    IsDataVal = FastSymbol(a$, ")", True)
      Else
     
                MissParam a$
    End If

End Function
Function IsSymbolNoSpace(a$, c$, Optional l As Long = 1) As Boolean
' not for greek identifiers. see isStr1()
    Dim j As Long
    j = Len(a$)
    If j = 0 Then Exit Function
    If UCase(Mid$(a$, 1, l)) = c$ Then
        a$ = NLtrim$(Mid$(a$, l + 1))
        
        IsSymbolNoSpace = True
    End If
End Function
Function FindItem(bstackstr As basetask, v As Variant, a$, R$, w2 As Long, Optional ByVal wasarr As Boolean = False) As Boolean
Dim usehandler As mHandler, fastcol As FastCollection, pppp As mArray, w1 As Long, p As Variant, s$
'Dim prev As Variant
'Set prev = v
FindItem = True
againtype:
        R$ = Typename(v)
        If R$ = "mHandler" Then
            Set usehandler = v
            Select Case usehandler.t1
            Case 1
                Set fastcol = usehandler.objref
                If FastSymbol(a$, ")(", , 2) Or True Then
                    If IsExp(bstackstr, a$, p) Then
                        If Not fastcol.Find(p) Then GoTo keynotexist
                        If fastcol.IsObj Then
                            w2 = fastcol.Index
                            Set v = fastcol.ValueObj
                            GoTo againtype
                        Else
                            wasarr = True
                            GoTo checkit
                        End If
                    ElseIf IsStrExp(bstackstr, a$, s$, Len(bstackstr.tmpstr) = 0) Then
                        If fastcol.IsObj Then
                            w2 = fastcol.Index
                           ' Set prev = v
                            Set v = fastcol.ValueObj
                            GoTo againtype
                        Else
                        If fastcol.StructLen > 0 Then GoTo checkit
                            R$ = Typename(fastcol.Value)
                        End If
                    Else
                        FindItem = False
                        Exit Function
keynotexist:
                        indexout a$
                        FindItem = False
                        Exit Function
                    End If
                Else
                    ' new
checkit:
                    If fastcol.StructLen > 0 Then
                    FindItem = False
                    Exit Function
                    ElseIf wasarr Then
                    'r$ = Typename(fastcol.Value)
                    FindItem = False
                    Exit Function
                    Else
                    FindItem = False
                    Exit Function
                    End If
                End If
            Case 2
                'r$ = "Buffer"
                    FindItem = False
                    Exit Function
                
            Case 3
                w1 = usehandler.indirect
                If w1 > -1 And w1 <= var2used Then
                                R$ = Typename(var(w1))
                                If R$ = "mHandler" Then Set v = var(w1): GoTo againtype
                    Else
                            R$ = Typename(usehandler.objref)
                                       If FastSymbol(a$, ")(", , 2) Or True Then
                                        If R$ = "mArray" Then
                                            Set pppp = usehandler.objref
                                                If IsExp(bstackstr, a$, p) Then
                                                   pppp.Index = p
                                                    If MyIsObject(pppp.Value) Then
                                                    w2 = p
                                                   ' Set prev = v
                                                         Set v = pppp.Value
                                                         wasarr = False
                                                         GoTo againtype
                                                    Else
                                                       ' r$ = Typename(pppp.Value)
                                                    End If
                                                Else

                                                FindItem = False
                                                Exit Function
                                            End If
                                        Else
                                                FindItem = False
                                                Exit Function
                                        End If
                                        
                                        End If
                                        
                                        End If
                                    
            Case 4
                    FindItem = False
                    Exit Function
            Case Else
                R$ = Typename(usehandler.objref)
                
            End Select
        ElseIf Typename(v) = "PropReference" Then
                    FindItem = False
                    Exit Function
        End If
        Set bstackstr.lastobj = Nothing
        Set bstackstr.lastpointer = Nothing
        FindItem = FastSymbol(a$, ")")
        If usehandler Is Nothing Then FindItem = False: Exit Function
        If TypeOf usehandler.objref Is mArray Then
            Set v = usehandler.objref
        Else
         Set pppp = New mArray
            Set pppp.GroupRef = usehandler
            pppp.Arr = False
            w2 = -101
            Set v = pppp
        End If
End Function


Function HaveMark(bstack As basetask, a As Long, b As Boolean) As Boolean
Dim s As mStiva2
Set s = bstack.RetStack
If s.Total >= 3 Then
HaveMark = s.LookTopVal = -3
a = s.StackItem(2)
b = s.StackItem(3)
End If
End Function
Function HaveMark2(bstack As basetask) As Boolean
Dim s As mStiva2
Set s = bstack.RetStack
If s.Total >= 3 Then
If s.LookTopVal = -3 Then s.drop 3: HaveMark2 = True
End If
End Function
Sub DropMark(bstack As basetask)
Dim s As mStiva2
Set s = bstack.RetStack
If s.Total >= 3 Then
If s.LookTopVal = -3 Then s.drop 3
End If
End Sub
Function interpret(bstack As basetask, b$, Optional ByPass As Boolean) As Boolean
Dim di As Object, myobject As Object, i As Long, x1 As Long, ok As Boolean, sp As Variant
Dim usehandler As mHandler, usehandler2 As mHandler
Set di = bstack.Owner
Dim prive As basket
'b$ = Trim$(b$)
Dim w$, ww#, LLL As Long, sss As Long, v As Long, p As Variant, ss$, sw$, ohere$
Dim pppp As mArray, i1 As Long, Lang As Long
Dim R1 As Long, r2 As Long
' uink$ = VbNullString
di.FontTransparent = True
SetBkMode di.Hdc, 1
ohere$ = here$
If Not ByPass Then here$ = vbNullString
bstack.LoadOnly = ByPass
sss = Len(b$)
Do While Len(b$) <> LLL
If LastErNum <> 0 Then Exit Do
LLL = Len(b$)


If FastSymbol(b$, "{") Then
If Not interpret(bstack, block(b$)) Then interpret = False: here$ = ohere$: GoTo there1
If FastSymbol(b$, "}") Then
sss = Len(b$)
GoTo loopcontinue1
'LLL = Len(b$)


Else
interpret = False: here$ = ohere$: GoTo there1
End If
End If
jumpforCR1:
If FastSymbol(b$, vbCrLf, , 2) Then
        While FastSymbol(b$, vbCrLf, , 2)
        Wend
     ''   UINK$ = VbNullString
        sss = LLL
        End If

While MaybeIsSymbol(b$, "\'/")
 SetNextLine b$
    sss = Len(b$)
    LLL = sss
Wend
While FastSymbol(b$, ":")
sss = LLL

Wend
If NOEXECUTION Then interpret = False: here$ = ohere$: GoTo there1

If NocharsInLine(b$) Then interpret = True: here$ = ohere$: GoTo there1
If IsSymbol(b$, "@") Then
i1 = IsLabelAnew("", b$, w$, Lang)  '' NO FORM AA@BBB ALLOWED HERE
w$ = "@" + w$
GoTo PROCESSCOMMAND   'IS A COMMAND
Else
i1 = IsLabelAnew("", b$, w$, Lang) '' NO FORM AA@BBB ALLOWED HERE
End If
  If trace And (bstack.Process Is Nothing) And Not bypasstrace Then
  If bstack.IamLambda Then
  If pagio$ = "GREEK" Then
  Form2.label1(0) = "ΛΑΜΔΑ()"
  Else
  Form2.label1(0) = "LAMBDA()"
  End If
  Else
    Form2.label1(0) = here$
    End If
    Form2.label1(1) = w$
    Form2.label1(2) = GetStrUntil(vbCrLf, b$ & vbCrLf, False)
 TestShowSub = vbNullString
 TestShowStart = 0
    Set Form2.Process = bstack
    stackshow bstack
    If Not Form1.Visible Then
    Form1.Show , Form5   'OK
    End If

    If STbyST And bstack.IamChild Then
        STbyST = False
        If Not STEXIT Then
        If Not STq Then
        Form2.gList4.ListIndex = 0
        End If
        End If
        Do
        If di.Visible Then di.Refresh
        ProcTask2 bstack
        Loop Until STbyST Or STq Or STEXIT Or NOEXECUTION Or myexit(bstack)
            If Not STEXIT Then
        If Not STq Then
        Form2.gList4.ListIndex = 0
        End If
        End If
        STq = False
        If STEXIT Then
        NOEXECUTION = True
        trace = False
        STEXIT = False
        GoTo there1
        End If
    End If
'Sleep 5
   '' SleepWaitNO 5
    If STEXIT Then
    
    trace = False
    STEXIT = False
    GoTo there1
    Else
    
    End If
End If
Select Case i1
Case 1234
GoTo jumpforCR1
Case 2
NoRef2
interpret = False
GoTo there1
Case 1

    If sss = LLL Then
  If comhash.Find2(w$, i, v) Then
  If v <> 0 Then GoTo PROCESSCOMMAND
  End If
    ss$ = vbNullString
    If MaybeIsSymbol(b$, "/*-+=~^|<>") Then
        If FastOperator(b$, "<=", i, 2, False) Then
        ' LOOK GLOBAL
        If GetVar(bstack, w$, v, True) Then
        w$ = varhash.lastkey
            Mid$(b$, i, 2) = "  "
            
            GoTo assignvalue
        ElseIf GetlocalVar(w$, v) Then
            w$ = varhash.lastkey
            Mid$(b$, i, 2) = "  "
            GoTo assignvalue
        Else
            ' NO SUCH VARIABLE
            interpret = False
            GoTo there1
        End If
        ' do something here
        ElseIf varhash.Find(myUcase(w$), v) Then
        ' CHECK VAR
            If FastOperator(b$, "=", i) Then
assignvalue:
                If MyIsNumeric(var(v)) Then
assignvalue2:
                    If IsExp(bstack, b$, p) Then
assignvalue3:
                        If bstack.lastobj Is Nothing Then
                        If VarType(var(v)) = vbLong Then
                        On Error Resume Next
                            var(v) = CLng(Int(p))
                            If Err.Number > 0 Then OverflowLong: interpret = 0: GoTo there1
                            On Error GoTo 0
                        Else
                            var(v) = p
                        End If
                        Else
checkobject:
                        Set myobject = bstack.lastobj
                            If TypeOf bstack.lastobj Is Group Then ' oh is a group
                                Set bstack.lastobj = Nothing
                                UnFloatGroup bstack, w$, v, myobject, True ' global??
                                Set myobject = Nothing
                            ElseIf CheckIsmArray(myobject) Then
                                    Set usehandler = New mHandler
                                    usehandler.t1 = 3
                                    Set usehandler.objref = myobject
                                    Set var(v) = usehandler
                                    If TypeOf bstack.lastobj Is mHandler Then
                                        Set usehandler2 = bstack.lastobj
                                        With usehandler2
                                            If .UseIterator Then
                                                usehandler.UseIterator = True
                                                usehandler.index_start = .index_start
                                                usehandler.index_End = .index_End
                                                usehandler.index_cursor = .index_cursor
                                            End If
                                        End With
                                    End If
                                    Set usehandler2 = Nothing
                                    Set usehandler = Nothing
                            ElseIf TypeOf myobject Is mHandler Then
                                Set usehandler2 = myobject
                                If usehandler2.indirect > -1 Then
                                    Set var(v) = var(usehandler2.indirect)
                                Else
                                    Set var(v) = usehandler2
                                End If
                                Set usehandler2 = bstack.lastobj
                                Set usehandler = var(v)
                                With usehandler2
                                    If .UseIterator Then
                                        usehandler.UseIterator = True
                                        usehandler.index_start = .index_start
                                        usehandler.index_End = .index_End
                                        usehandler.index_cursor = .index_cursor
                                    End If
                                End With
                                Set usehandler2 = Nothing
                                Set usehandler = Nothing
                                Set bstack.lastobj = Nothing
                            ElseIf TypeOf myobject Is lambda Then
                                GlobalSub w$ + "()", "", , , v
                                Set var(v) = myobject
                                Set bstack.lastobj = Nothing
                            ElseIf TypeOf myobject Is mEvent Then
                             Set var(v) = myobject
                            CopyEvent var(v), bstack
                            Set var(v) = bstack.lastobj
                            ElseIf TypeOf myobject Is VarItem Then
                                
                                var(v) = myobject.ItemVariant
                            Else
                                Set myobject = Nothing
                                Set bstack.lastobj = Nothing
                                If VarType(var(v)) = vbLong Then
                                    NoObjectpAssignTolong
                                Else
                                    NoObjectAssign
                                End If
                                interpret = False: GoTo there1
                            End If
                            Set bstack.lastobj = Nothing
                            Set myobject = Nothing
                        End If

                    ElseIf IsStrExp(bstack, b$, ss$, Len(bstack.tmpstr) = 0) Then
                    If bstack.lastobj Is Nothing Then
                    If ss$ = vbNullString Then
                    var(v) = 0#
                    Else
                    If IsNumberCheck(ss$, p) Then
                    var(v) = p
                    End If
                    End If
                    Else
                    GoTo checkobject
                    End If
                    Else
                    ' if is string then what???
                    If Typename(bstack.lastobj) = "mHandler" Then
                    GoTo checkobject
                    End If
                        NoValueForVar w$
                        interpret = False
                        GoTo there1
                    End If
                    GoTo loopcontinue1
                    
                Else
                    If Not MyIsObject(var(v)) Then
                        If IsStrExp(bstack, b$, ss$) Then
                            If ss$ = vbNullString Then
                            var(v) = 0#
                        Else
                            If IsNumberCheck(ss$, p) Then
                                var(v) = p
                            End If
                        End If
                        GoTo loopcontinue1
                    Else
                        MyEr "Expected String expression", "Περίμενα έκφραση Αλφαριθμητική"
                        Exit Function
                    End If
                ElseIf var(v) Is Nothing Then
                    AssigntoNothing  ' Use Declare
                    interpret = False
                    GoTo there1
                ElseIf TypeOf var(v) Is Group Then
                    If IsExp(bstack, b$, p) Then
                        If var(v).HasSet Then
                            If bstack.lastobj Is Nothing Then
                                bstack.soros.PushVal p
                            Else
                                bstack.soros.PushObj bstack.lastobj
                                Set bstack.lastobj = Nothing
                            End If
                            NeoCall2 bstack, w$ + "." + ChrW(&H1FFF) + ":=()", ok
                    ElseIf bstack.lastobj Is Nothing Then
                        NeedAGroupInRightExpression
                        interpret = False
                        GoTo there1
                    ElseIf TypeOf bstack.lastobj Is Group Then
                        Set myobject = bstack.lastobj
                        Set bstack.lastobj = Nothing
                        ss$ = bstack.GroupName
                        If var(v).HasValue Or var(v).HasSet Then
                            PropCantChange
                            interpret = 0
                            GoTo there1
                        Else
                            If Len(var(v).GroupName) > Len(w$) Then
                                UnFloatGroupReWriteVars bstack, w$, v, myobject
                            Else
                                bstack.GroupName = Left$(w$, Len(w$) - Len(var(v).GroupName) + 1)
                                If Len(var(v).GroupName) > 0 Then
                                    w$ = Left$(var(v).GroupName, Len(var(v).GroupName) - 1)
                                    UnFloatGroupReWriteVars bstack, w$, v, myobject
                                Else
                                    GroupWrongUse
                                    interpret = 0
                                    GoTo there1
                                End If
                            End If
                        End If
                        Set myobject = Nothing
                        bstack.GroupName = ss$
                    Else
                        WrongObject
                        interpret = False
                        GoTo there1
                    End If
                    GoTo loopcontinue1
                Else
noexpression:
                Set myobject = Nothing
                Set bstack.lastobj = Nothing
                MissNumExpr
                interpret = False
                GoTo there1
            End If
        ElseIf TypeOf var(v) Is PropReference Then
            If IsExp(bstack, b$, p) Then
                If FastSymbol(b$, "@") Then
                    If IsExp(bstack, b$, sp) Then
                        var(v).Index = p: sp = 0
                    ElseIf IsStrExp(bstack, b$, ss$) Then
                        var(v).Index = ss$: ss$ = vbNullString
                    End If
                    var(v).UseIndex = True
                End If
                var(v).Value = p
            Else
                GoTo noexpression
            End If
            GoTo loopcontinue1
        ElseIf TypeOf var(v) Is Constant Then
            CantAssignValue
            interpret = False
            GoTo there1
        ElseIf TypeOf var(v) Is lambda Then
            ' exist and take something else
            If IsExp(bstack, b$, p) Then
                If bstack.lastobj Is Nothing Then
                    Expected "lambda", "λάμδα"
                ElseIf TypeOf bstack.lastobj Is lambda Then
                    Set var(v) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                    GoTo loopcontinue1
                Else
                    Expected "lambda", "λάμδα"
                End If
                interpret = False
                GoTo there1
            Else
                MissNumExpr
                interpret = False
                GoTo there1
            End If
        ElseIf TypeOf var(v) Is mHandler Then  ' CHECK IF IT IS A HANDLER
            Set usehandler = var(v)
            If IsExp(bstack, b$, p) Then
                If usehandler.ReadOnly Then
                    ReadOnly
                    interpret = False: GoTo there1
                End If
                If bstack.lastobj Is Nothing Then
                    MissingObjReturn
                    interpret = False: GoTo there1
                ElseIf Typename(bstack.lastobj) = "mHandler" Then
                    Set usehandler = bstack.lastobj
                    Set myobject = New mHandler
                    usehandler.CopyTo myobject
                    Set var(v) = myobject
                ElseIf Typename(bstack.lastobj) = myArray Then
                    Set usehandler = New mHandler
                    usehandler.t1 = 3
                    Set usehandler.objref = bstack.lastobj
                    Set var(v) = usehandler
                    Set usehandler = Nothing
                Else
                    Set usehandler = var(v)
                    usehandler.t1 = 0
                    Set usehandler.objref = bstack.lastobj
                    Set usehandler = Nothing
                End If
                Set myobject = Nothing
            Else
                MissNumExpr
                interpret = False
                GoTo there1
            End If
            Set bstack.lastobj = Nothing
            Set myobject = Nothing
        ElseIf TypeOf var(v) Is mEvent Then
            If IsExp(bstack, b$, p) Then
                Set var(v) = bstack.lastobj
                CopyEvent var(v), bstack
                Set var(v) = bstack.lastobj
                Set bstack.lastobj = Nothing
            Else
                MissNumExpr
                interpret = 0
                GoTo there1
            End If
        Else
            i = 1
            GoTo somethingelse
        End If
    End If
Else
    ' or do something else
somethingelse:
    If InStr("/*-+=~^&|<>", Mid$(b$, i, 1)) > 0 Then
        If InStr("/*-+=~^&|<>!", Mid$(b$, i + 1, 1)) > 0 Then
            ss$ = Mid$(b$, i, 2)
            Mid$(b$, i, 2) = "  "
        Else
            ss$ = Mid$(b$, i, 1)
            Mid$(b$, i, 1) = " "
        End If
    Else
        GoTo PROCESSCOMMAND
    End If
    On Error GoTo err123456
    If MyIsNumeric(var(v)) Then
        If VarType(var(v)) = vbLong Then
            On Error GoTo forlong
                Select Case ss$
                    Case "="
                        v = globalvar(w$, CLng(Int(p)), , True)
                        GoTo assignvalue2
                    Case "+="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(Int(p) + var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "-="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(-Int(p) + var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "*="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(Int(p) * var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "/="
                        If IsExp(bstack, b$, p) Then
                            If Int(p) = 0 Then
                                DevZero
                                interpret = False
                                GoTo there1
                            End If
                            var(v) = CLng(var(v) / Int(p))
                        Else
                            GoTo noexpression
                        End If
                    Case "-!"
                        var(v) = CLng(-var(v))
                    Case "++"
                        var(v) = CLng(1 + var(v))
                    Case "--"
                        var(v) = CLng(var(v) - 1)
                    Case "~"
                        var(v) = CLng(-1 - (var(v) <> 0))
                    Case Else
                        GoTo PROCESSCOMMAND
                End Select
                On Error GoTo 0
            Else
                Select Case ss$
                    Case "="
                        v = globalvar(w$, p, , True)
                        GoTo assignvalue2
                    Case "+="
                        If IsExp(bstack, b$, p) Then
                            var(v) = p + var(v)
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "-="
                        If IsExp(bstack, b$, p) Then
                            var(v) = -p + var(v)
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "*="
                        If IsExp(bstack, b$, p) Then
                            var(v) = p * var(v)
                             If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "/="
                        If IsExp(bstack, b$, p) Then
                            If p = 0 Then
                                DevZero
                                interpret = False
                                GoTo there1
                            End If
                            var(v) = var(v) / p
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "-!"
                        var(v) = -var(v)
                    Case "++"
                        var(v) = 1 + var(v)
                    Case "--"
                        var(v) = var(v) - 1
                    Case "~"
                     
                         Select Case VarType(var(v))
                        Case vbBoolean
                            var(v) = Not CBool(var(v))
                        Case vbCurrency
                            var(v) = CCur(Not CBool(var(v)))
                        Case vbDecimal
                            var(v) = CDec(Not CBool(var(v)))
                        Case Else
                            var(v) = CDbl(Not CBool(var(v)))
                        End Select
                        
                        
                    Case Else
                        GoTo PROCESSCOMMAND
                    End Select
                    On Error Resume Next
                End If
            ElseIf TypeOf var(v) Is Group Then
                If IsExp(bstack, b$, p) Then
                    If bstack.lastobj Is Nothing Then
                        bstack.soros.PushVal p
                    Else
                        bstack.soros.PushObj bstack.lastobj
                        Set bstack.lastobj = Nothing
                    End If
                End If
                NeoCall2 bstack, w$ + "." + ChrW(&H1FFF) + ss$ + "()", ok
                If Not ok Then
                    If LastErNum = 0 Then
                        MisOperatror (ss$)
                    End If
                    interpret = False
                    GoTo there1
                End If
            Else
                Set myobject = var(v)
                    If CheckIsmArray(myobject) Then
                        If IsExp(bstack, b$, p) Then
                            If Not bstack.lastobj Is Nothing Then
                                If TypeOf bstack.lastobj Is mArray Then
                                    Set usehandler = New mHandler
                                    Set var(v) = usehandler
                                    usehandler.t1 = 3
                                    Set usehandler.objref = bstack.lastobj
                                    Set usehandler = Nothing
                                Else
                                    Set myobject = bstack.lastobj
                                    If CheckIsmArray(myobject) Then
                                        Set usehandler = New mHandler
                                        Set var(v) = usehandler
                                        usehandler.t1 = 3
                                        Set usehandler.objref = myobject
                                        Set usehandler = Nothing
                                    Else
                                        NotArray
                                        interpret = False
                                        GoTo there1
                                    End If
                                End If
                            Else
                                myobject.Compute2 p, ss$
                            End If
                            Set myobject = Nothing
                            Set bstack.lastobj = Nothing
                        Else
                            myobject.Compute3 ss$
                            Set myobject = Nothing
                            Set bstack.lastobj = Nothing
                        End If
                    ElseIf TypeOf myobject Is mHandler Then
                    Set usehandler = myobject
                    If usehandler.t1 = 4 Then
                        If usehandler.ReadOnly Then
                                ReadOnly
                             interpret = False
                                GoTo there1
                        ElseIf ss$ = "++" Then
                        If usehandler.index_start < usehandler.objref.count - 1 Then
                            usehandler.index_start = usehandler.index_start + 1
                            usehandler.objref.Index = usehandler.index_start
                            usehandler.index_cursor = usehandler.objref.Value
                        End If
                        ElseIf ss$ = "--" Then
                    If usehandler.index_start > 0 Then
                            usehandler.index_start = usehandler.index_start - 1
                            usehandler.objref.Index = usehandler.index_start
                            usehandler.index_cursor = usehandler.objref.Value
                        End If
                        ElseIf ss$ = "-!" Then
                        usehandler.sign = -usehandler.sign
                        Else
                        NoOperatorForThatObject ss$
                         interpret = False
                            GoTo there1
                        End If
                        End If
                    Else
                    MyEr "Object not support operator " + ss$, "Το αντικείμενο δεν υποστηρίζει το τελεστή " + ss$
                    interpret = False
                    GoTo there1
                    End If
                End If
            End If
            Set usehandler = Nothing
            Set myobject = Nothing
            GoTo loopcontinue1
        ElseIf Not bstack.StaticCollection Is Nothing Then
            If bstack.ExistVar(w$) Then
                If FastOperator(b$, "=", i) Then
                    If IsExp(bstack, b$, p) Then
checkobject1:
                        Set myobject = bstack.lastobj
                        If CheckIsmArray(myobject) Then
                            Set usehandler = New mHandler
                            Set bstack.lastobj = usehandler
                            usehandler.t1 = 3
                            bstack.SetVarobJ w$, myobject
                            Set usehandler = Nothing
                        Else
                            bstack.SetVar w$, p
                        End If
                        Set myobject = Nothing
                        Set usehandler = Nothing
                        Set bstack.lastobj = Nothing
                        GoTo loopcontinue1
                    ElseIf IsStrExp(bstack, b$, ss$) Then
                        If bstack.lastobj Is Nothing Then
                            If ss$ = vbNullString Then
                                p = 0
                            Else
                                p = val(ss$)
                            End If
                        End If
                        GoTo checkobject1
                    Else
                        GoTo aproblem1
                    End If
                Else
                    If InStr("/*-+~", Mid$(b$, i, 1)) > 0 Then
                        If InStr("=+-!", Mid$(b$, i + 1, 1)) > 0 Then
                            ss$ = Mid$(b$, i, 2)
                            Mid$(b$, i, 2) = "  "
                        Else
                            ss$ = Mid$(b$, i, 1)
                            Mid$(b$, i, 1) = " "
                        End If
                    End If
                    If Not bstack.AlterVar(w$, p, ss$, False) Then interpret = False: GoTo there1
                    GoTo loopcontinue1
                End If
            End If
            If FastOperator(b$, "=", i) Then ' MAKE A NEW ONE IF FOUND =
                v = globalvar(w$, p, , True)
                GoTo assignvalue
            ElseIf GetVar(bstack, w$, v, True) Then
                GoTo somethingelse
            End If
        ElseIf FastOperator(b$, "=", i) Then ' MAKE A NEW ONE IF FOUND =
jumpiflocal:
            v = globalvar(w$, p, , True)
            GoTo assignvalue
        ElseIf GetVar(bstack, w$, v, True) Then
        ' CHECK FOR GLOBAL
            GoTo somethingelse
            Else
        GoTo PROCESSCOMMAND
        End If
    Else
          '**********************************************************
PROCESSCOMMAND:
        Dim y1 As Long
        Dim x2 As Long, y2 As Long, SBR$, nd&
        If Trim$(w$) <> "" Then
            Select Case w$
                Case "CALL", "ΚΑΛΕΣΕ"
                    ' CHECK FOR NUMBER...
                    If bstack.NoRun Then
                        bstack.callx1 = 0
                        bstack.callohere = vbNullString
                        b$ = NLtrim(b$)
                        SetNextLineNL b$
                    Else
                        If lckfrm > 0 Then lckfrm = sb2used + 1
                        NeoCall ObjPtr(bstack), b$, Lang, ok
                        If Not ok Then
                            interpret = 0
                            GoTo there1
                        End If
                    End If
                Case " ", ChrW(160), vbTab
                Case "SLOW", "ΑΡΓΑ"
                    extreme = False
                    SLOW = True
                    interpret = True
                    here$ = ohere$
                    GoTo there1
                Case "FAST", "ΓΡΗΓΟΡΑ"
                    If FastSymbol(b$, "!") Then extreme = True Else extreme = False
                    SLOW = False
                    interpret = True
                    here$ = ohere$
                    GoTo there1
                Case "GLOBAL", "ΓΕΝΙΚΟ", "ΓΕΝΙΚΗ", "ΓΕΝΙΚΕΣ", "LOCAL", "ΤΟΠΙΚΑ", "ΤΟΠΙΚΗ", "ΤΟΠΙΚΕΣ"
                    b$ = w$ + " " + b$
                    interpret = Execute(bstack, b$, True) = 1
                    GoTo there1
                Case "USER", "ΧΡΗΣΤΗΣ"
                    ss$ = PurifyPath(GetStrUntil("\", Trim$(GetNextLine(b$) + "\")))
                    If ss$ <> "" Then
                        dset
                        userfiles = GetSpecialfolder(CLng(26)) & "\M2000_USER\"
                        If Not isdir(userfiles) Then MkDir userfiles
                        ss$ = AddBackslash(userfiles + ss$)
                        If PathMakeDirs(ss$) Or isdir(ss$) Then
                            userfiles = ss$
                            mcd = userfiles
                            original bstack, "CLS"
                        Else
                            PlainBaSket di, players(GetCode(di)), "Bad User Name"
                        End If
                    Else
                        ss$ = UCase(userfiles)
                        DropLeft "\M2000_USER\", ss$
                        If Len(ss$) > 0 Then
                            PlainBaSket di, players(GetCode(di)), GetStrUntil("\", Tcase(ss$))
                        End If
                    End If
                    interpret = True
                    GoTo there1
                Case "TARGET", "ΣΤΟΧΟΣ"
                    If Abs(IsLabel(bstack, b$, w$)) = 1 Then
                        If Not GetVar(bstack, w$, v) Then
                            v = globalvar(w$, 0#, , True)
                        End If
                    Else
                        interpret = False
                        here$ = ohere$: GoTo there1
                    End If
                    If Not FastSymbol(b$, ",") Then
                        interpret = False
                        Exit Do
                    ElseIf IsExp(bstack, b$, p) Then
                        q(var(v)).Enable = Not (p = 0)
                        RTarget bstack, q(var(v))
                    ElseIf IsStrExp(bstack, b$, ss$, Len(bstack.tmpstr) = 0) Then
                        If ss$ = vbNullString Then interpret = False: here$ = ohere$: GoTo there1
                        x1 = 1
                        y1 = 1
                        x2 = -1
                        y2 = -1
                        nd& = 0
                        SBR$ = vbNullString
                        On Error GoTo err123456
                        With players(GetCode(di))
                            If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then x1 = Abs(p) Mod (.mx + 1)
                            If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then y1 = Abs(p) Mod (.My + 1)
                            If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then x2 = CLng(p)
                            If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then y2 = CLng(p)
                            If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then nd& = Abs(p)
                            If FastSymbol(b$, ",") Then If Not IsStrExp(bstack, b$, SBR$) Then interpret = False: here$ = ohere$: GoTo there1
                
err123456:
                            If Err.Number = 6 Then
                                OverflowLong
                                interpret = False: here$ = ohere$: GoTo there1
                            End If
                            Targets = False
                            ReDim Preserve q(UBound(q()) + 1)
                            q(UBound(q()) - 1) = BoxTarget(bstack, x1, y1, x2, y2, SBR$, nd&, ss$, .Xt, .Yt, .uMineLineSpace)
                        End With
                        var(v) = UBound(q()) - 1
                        Targets = True
                    Else
                        interpret = False
                        here$ = ohere$:             GoTo there1
                    End If
                Case "ΔΙΑΚΟΠΤΕΣ", "SWITCHES"
                    If IsStrExp(bstack, b$, ss$) Then
                    Switches ss$, bstack.IamChild Or bstack.IamAnEvent  ' NON LOCAL FROM cli OR using SET SWITCHES
                End If
                Case "MONITOR", "ΕΛΕΓΧΟΣ"
                    If IsSupervisor Then
                    prive = players(GetCode(di))
                    
                    monitor bstack, prive, Lang
                    players(GetCode(di)) = prive
                    Else
                    BadCommand
                    End If
                Case "ΣΕΝΑΡΙΟ", "SCRIPT"
                If IsLabelOnly(b$, ss$) Then
                 If GetSub(myUcase(ss$, True), nd&) Then
                           b$ = vbCrLf + sbf(nd&).sb & b$
                   Else
                   b$ = ss$ + " " + b$
                   If IsStrExp(bstack, b$, w$) Then
                           b$ = vbCrLf + w$ + b$
                   Else
                   ' skip
                   End If
                   End If
                ElseIf IsStrExp(bstack, b$, w$) Then
                           b$ = vbCrLf + w$ + b$
                   End If
Case "RETURN", "ΕΠΙΣΤΡΟΦΗ"
    LastErNum = 0
       If IsExp(bstack, b$, p) Then
                If bstack.lastobj Is Nothing Then
                ElseIf Typename(bstack.lastobj) = "mHandler" Then
                        Select Case bstack.lastobj.t1
                           Case 1
                                  If ChangeValues(bstack, b$) Then GoTo loopcontinue1
                                  
                           Case 2
                                    If ChangeValuesMem(bstack, b$, Lang) Then GoTo loopcontinue1
                           Case 3
                                    If ChangeValuesArray(bstack, b$) Then GoTo loopcontinue1
                           End Select
                End If
            ElseIf IsStrExp(bstack, b$, ss$, Len(bstack.tmpstr) = 0) Then
                    append_table bstack, ss$, b$, True, Lang
                GoTo loopcontinue1
                 End If
  BadUseofReturn
       interpret = False
       GoTo there1
            Case "CONTINUE", "ΣΥΝΕΧΙΣΕ"
            If HaltLevel > 0 Then
                     If NORUN1 Then NORUN1 = False: interpret = True: b$ = vbNullString: GoTo there1   ' send environment....to hell
                    If bstack.IamChild Or bstack.IamAnEvent Then NERR = True: NOEXECUTION = True
                    ExTarget = True: INK$ = Chr(27): UKEY$ = Chr$(27)  ': UINK$ = Chr(27)    ' send escape...for any good reason...
            Else
            GoTo contnoproper
            End If
            Case "CONST", "ΣΤΑΘΕΡΗ", "ΣΤΑΘΕΡΕΣ"
            ConstNew bstack, b$, w$, True, Lang
                    If LastErNum = -1 Then
                    interpret = False
                    GoTo there1
                    End If
                Case "ΤΕΛΟΣ", "END"
                    
                    If NORUN1 Then NORUN1 = False: interpret = True: b$ = vbNullString: GoTo there1   ' send environment....to hell
                    If bstack.IamChild Or bstack.IamAnEvent Then
                    NERR = True: NOEXECUTION = True
                    ElseIf Not bstack.IamChild And Not bstack.IamAnEvent And Not HaltLevel > 0 Then
                    If Check2Save Then
                        GoTo loopcontinue1
                    Else
                        Check2SaveModules = False
                    End If
                    End If
                    ExTarget = True: INK$ = Chr(27): UKEY$ = Chr$(27)
                Case Else
                    LastErNum = 0 ' LastErNum1 = 0
                    LastErName = vbNullString   ' every command from Query call identifier
                    LastErNameGR = vbNullString  ' interpret is like execute without if for repeat while select structures
                    If comhash.Find2(w$, i, v) Then
                        If v <> 0 Then
                            If v = 32 Then
                                If Not Identifier(bstack, w$, b$, True, Lang) Then
                                    If NOEXECUTION Then
                                            MyEr "", ""
                                            interpret = False
                                    End If
                                    here$ = ohere$: GoTo there1
                              Else
                              If bstack.callx1 > 0 Then
                              If bstack.NoRun Then
                              bstack.callx1 = 0
                              bstack.callohere = vbNullString
                              b$ = NLtrim(b$)
                              SetNextLineNL b$
                              ElseIf Not ProcModuleEntry(bstack, "", 0, b$, Lang) Then
                                    If MOUT And b$ = vbNullString Then
                                    Else
                                        MyErMacro b$, "unknown identifier " & w$, "’γνωστο αναγνωριστικό " & w$
                                    End If
                                End If
                                bstack.RemoveOptionals
                                End If
                              GoTo loopcontinue1
                              End If
                           
                         '' ElseIf v = 2000 Then
                          
                          Else
contnoproper:
                            MyEr "No proper command for command line interpreter", "Δεν είναι η κατάλληλη εντολή για τον διερμηνευτή γραμμής"
                            interpret = False
                          
                            here$ = ohere$: GoTo there1
                          End If
                     End If
                    
                    If i <> 0 Then
                     If IsBadCodePtr(i) = 0 Then
                        If Not CallByPtr(i, bstack, b$, Lang) Then
                               If NOEXECUTION Then
                                    MyEr "", ""
                                    interpret = False
                                    End If
                                    here$ = ohere$: GoTo there1
                        End If
                        End If
                    Else
                            If Not Identifier(bstack, w$, b$, Not comhash.Find(w$, i1), Lang) Then
                            
                                    If NOEXECUTION Then
                                    MyEr "", ""
                                    interpret = False
                                    End If
                                    here$ = ohere$: GoTo there1
                            End If
                    End If
                    
                    ElseIf Not Identifier(bstack, w$, b$, Not comhash.Find(w$, i1), Lang) Then
                    
                        If NOEXECUTION Then
                            MyEr "", ""
                            interpret = False
                        End If
                        here$ = ohere$: GoTo there1
                    ElseIf bstack.callx1 > 0 Then
                        If lckfrm > 0 Then lckfrm = sb2used + 1
                            If bstack.NoRun Then
                                bstack.callx1 = 0
                                bstack.callohere = vbNullString
                                b$ = NLtrim(b$)
                                SetNextLineNL b$
                              ElseIf Not ProcModuleEntry(bstack, "", 0, b$, Lang) Then
                                If MOUT And b$ = vbNullString Then
                                Else
                                    MyErMacro b$, "unknown identifier " & w$, "’γνωστο αναγνωριστικό " & w$
                                End If
                            
                            
                             End If
                            
                            If bstack.Parent Is Nothing Then
                                If NOEXECUTION Then
                                    NOEXECUTION = False
                                    MyEr "", ""
                                    Set Basestack1.Sorosref = New mStiva
                                    b$ = vbNullString
                                     ClearState
                                End If
                               
                            End If
                    End If
                    
                End Select
                End If
            End If
        Else
        If w$ <> "" Then
       '' b$ = w$ & " " & b$
        If Abs(IsLabel(bstack, (w$), w$)) Then
        
         If FindNameForGroup(bstack, w$) Then
 MyEr "Unknown Property " & w$, "’γνωστη ιδιότητα " & w$
 Else
MyEr "Unknown Variable " & w$, "’γνωστη μεταβλητή " & w$
End If
b$ = w$ & " " & b$
        
        Else

       SyntaxError
        End If
        b$ = vbNullString
        interpret = False
        GoTo there1
        End If
    End If
Case 3

ss$ = vbNullString
        i = 1
        If Len(b$) > 1 Then
        If InStr("/*-+=~^&|<>", Mid$(b$, i, 1)) > 0 Then
        
                    If InStr("/*-+=~^&|<>!", Mid$(b$, i + 1, 1)) > 0 Then
                        ss$ = Mid$(b$, i, 2)
                        Mid$(b$, i, 2) = "  "
                        If ss$ = "<=" Then ss$ = "g"
                    Else
                        ss$ = Mid$(b$, i, 1)
                        Mid$(b$, i, 1) = " "
                    End If
         End If
       End If

If ss$ <> "" Then
            If ss$ = "=" Then
                If GetVar(bstack, w$, v) Then
                sw$ = ss$
                    If IsStrExp(bstack, b$, ss$) Then
                    If Typename$(bstack.lastobj) = "lambda" Then
                                  GlobalSub w$ + "()", "", , , v
                                               Set var(v) = bstack.lastobj
                                                Set bstack.lastobj = Nothing
                    ElseIf Typename$(var(v)) = "Group" Then
                    
                    If sw$ = "g" Then
                           sw$ = ":="
                           If Not var(v).HasSet Then GroupCantSetValue: interpret = False: GoTo there1
                           End If
                           If bstack.lastobj Is Nothing Then
                                bstack.soros.PushStr ss$
                            Else
                                bstack.soros.PushObj bstack.lastobj
                                Set bstack.lastobj = Nothing
                            End If
                            NeoCall2 bstack, Left$(w$, Len(w$) - 1) + "." + ChrW(&H1FFF) + sw$ + "()", ok
                    If Not ok Then
                        If LastErNum = 0 Then
                            MisOperatror (ss$)
                        End If
                        interpret = False
                        GoTo there1
                    End If
                    ElseIf Typename(var(v)) = "PropReference" Then
                        If FastSymbol(b$, "@") Then
                            If IsExp(bstack, b$, sp) Then
                            var(v).Index = p: sp = 0
                            ElseIf IsStrExp(bstack, b$, sw$) Then
                            var(v).Index = sw$: sw$ = vbNullString
                            End If
                             var(v).UseIndex = True
                        End If
                        var(v).Value = ss$
                  ElseIf TypeOf var(v) Is Constant Then
                    CantAssignValue
                    interpret = False
                    GoTo there1
                    Else
                         
                         If CheckVarOnlyNo(var(v), ss$) Then
                           ExpectedObj Typename(var(v))
                           GoTo there1
                         End If
                        End If
                    Else
aproblem1:
                       NoValueForVar w$
                    Exit Do  '???
                    End If
                ElseIf IsStrExp(bstack, b$, ss$) Then
                    
                                If bstack.lastobj Is Nothing Then
              globalvar w$, ss$, , True
            Else
            If Typename$(bstack.lastobj) = "lambda" Then
                       If Not GetVar(bstack, w$, x1, True) Then x1 = globalvar(w$, p, , True)
                             GlobalSub w$ + "()", "", , , x1
                                        Set myobject = bstack.lastobj
                                        Set bstack.lastobj = Nothing
                                        If x1 <> 0 Then
                                        
                                          Set var(x1) = myobject
                                                Set myobject = Nothing
                                           
                                            
                                        End If
            End If
            End If
                ElseIf LastErNum = 0 Then
                                    
                    SyntaxError
                    interpret = False
                    GoTo there1
                    Else
                   Exit Do  '???
                End If
          
            ElseIf ss$ = "+=" Then
                            If GetVar(bstack, w$, v) Then
                                If IsStrExp(bstack, b$, ss$) Then
                                    If MyIsObject(var(v)) Then

                                            NoOperatorForThatObject "+="
                                            
                                            interpret = False
                                            GoTo there1

                                    Else
                                var(v) = CStr(var(v)) + ss$
                                    End If
                                Else
                                    MissStringExpr
                                End If
                            Else
                                ExpectedVariable
                            End If
            Else
            ' one now option
                If GetVar(bstack, w$, v) Then
                        If IsStrExp(bstack, b$, ss$) Then
                             CheckVar var(v), ss$
                        Else
                            NoValueForVar w$
                        Exit Do
                        End If
                Else
                    Nosuchvariable w$
                End If
        End If
End If
          
Case 4
If FastSymbol(b$, "=") Then '................................
           
            If GetVar(bstack, w$, v) Then
                If IsExp(bstack, b$, p) Then
                
                
                If Not bstack.lastobj Is Nothing Then
                        If TypeOf bstack.lastobj Is lambda Then
                        If Typename(var(v)) = "lambda" Then
                                                Set var(v) = bstack.lastobj

                                                Else
                                    GlobalSub w$ + "()", "", , , v
                                               Set var(v) = bstack.lastobj
                                                
                                        End If
                           Set bstack.lastobj = Nothing
                         Else
                       SyntaxError
                        End If
                        ElseIf MyIsObject(var(v)) Then
                        If TypeOf var(v) Is Constant Then
                            CantAssignValue
                            interpret = False
                            GoTo there1
                        Else
                           ExpectedObj Typename(var(v))
                           GoTo there1
                           End If
                        Else
                        var(v) = MyRound(p)
                        End If
                Else
                  MissNumExpr
                Exit Do
                End If
            ElseIf IsExp(bstack, b$, p) Then
             If Not bstack.lastobj Is Nothing Then
                
                If Typename$(bstack.lastobj) = "lambda" Then
                    
                       If Not GetVar(bstack, w$, x1, True) Then x1 = globalvar(w$, p, , True)
                             GlobalSub w$ + "()", "", , , x1
                                        Set myobject = bstack.lastobj
                                        Set bstack.lastobj = Nothing
                                        If x1 <> 0 Then
                                        
                                          Set var(x1) = myobject
                                                Set myobject = Nothing
                                           
                                            
                                        End If
                                        Else
                                SyntaxError
            End If
            Else
            globalvar w$, p, , True
            End If
                ElseIf LastErNum = 0 Then
                                
                SyntaxError
                interpret = False
                GoTo there1
                Else
               Exit Do
            End If
 Else
    If FastSymbol(b$, "+=", , 2) Then
    ss$ = "+"
    ElseIf FastSymbol(b$, "/=", , 2) Then
    ss$ = "/"
    ElseIf FastSymbol(b$, "-=", , 2) Then
    ss$ = "-"
    ElseIf FastSymbol(b$, "*=", , 2) Then
    ss$ = "*"
    ElseIf IsOperator0(b$, "++", 2) Then
    ss$ = "++"
    ElseIf IsOperator0(b$, "--", 2) Then
    ss$ = "--"
    ElseIf IsOperator0(b$, "-!", 2) Then
    ss$ = "-!"
         ElseIf IsOperator0(b$, "~") Then
        ss$ = "!!"
    ElseIf FastSymbol(b$, "<=", , 2) Then
    ss$ = "="
    End If
        If ss$ = vbNullString Then
                    NoValueForVar w$
                    interpret = False
                     GoTo there1
    End If
    If GetVar(bstack, w$, v) Then
        If Len(ss$) = 1 Then
                    If IsExp(bstack, b$, p) Then
                            On Error Resume Next
                            Select Case ss$
                            Case "="
                            var(v) = MyRound(p)
                                Case "+"
                                var(v) = MyRound(p) + MyRound(var(v))
                                Case "*"
                                 var(v) = MyRound(MyRound(p) * MyRound(var(v)))
                                Case "-"
                                var(v) = MyRound(var(v)) - MyRound(p)
                                Case "/"
                                If MyRound(p) = 0 Then
                                   interpret = False
                                 GoTo there1
                                End If
                                 var(v) = MyRound(MyRound(var(v) / MyRound(p)))
                                 Case "!"
                                 var(v) = -1 - (var(v) <> 0)
                            End Select
                            If Err.Number = 6 Then
                            interpret = False
                            GoTo there1
                            End If
                
                    Else
                                   interpret = False
                                 GoTo there1
                    End If
        Else
        If ss$ = "++" Then
        var(v) = 1 + var(v)
        ElseIf ss$ = "--" Then
        var(v) = var(v) - 1
        ElseIf ss$ = "-!" Then
        var(v) = -var(v)
        Else

                      var(v) = -1 - (var(v) <> 0)
        End If
        End If
    Else
                   interpret = False
        GoTo there1
    End If
End If
Case 5

If neoGetArray(bstack, w$, pppp) Then
againarray22:
    If FastSymbol(b$, ")") Then
    'need to found an expression
        If FastSymbol(b$, "=") Then
            If IsExp(bstack, b$, p) Then
                If Not bstack.lastobj Is Nothing Then
                    bstack.lastobj.CopyArray pppp
                    pppp.Final = False
                    Set bstack.lastobj = Nothing
                    GoTo loopcontinue1
                End If
            Else
                SyntaxError
            End If
            interpret = False
            GoTo there1
        End If
        End If
If Not NeoGetArrayItem(pppp, bstack, w$, v, b$) Then interpret = False: here$ = ohere$: GoTo there1
On Error Resume Next
If MaybeIsSymbol(b$, ":+-*/~") Then
With pppp
        If IsOperator0(b$, "++", 2) Then
            .item(v) = .itemnumeric(v) + 1
            GoTo loopcontinue1
        ElseIf IsOperator0(b$, "--", 2) Then
            .item(v) = .itemnumeric(v) - 1
            GoTo loopcontinue1
        ElseIf IsOperator(b$, "+=", 2) Then
            If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
            .item(v) = .itemnumeric(v) + p
        ElseIf IsOperator(b$, "-=", 2) Then
            If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
            .item(v) = .itemnumeric(v) - p
        ElseIf IsOperator(b$, "*=", 2) Then
            If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
            .item(v) = .itemnumeric(v) * p
        ElseIf IsOperator(b$, "/=", 2) Then
            If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
            If p = 0 Then
             DevZero
             Else
             .item(v) = pppp.itemnumeric(v) / p
            End If
        ElseIf IsOperator0(b$, "-!", 2) Then
            .item(v) = -.itemnumeric(v)
            GoTo loopcontinue1
        ElseIf IsOperator0(b$, "~") Then
            Select Case VarType(.itemnumeric(v))
            Case vbBoolean
                .item(v) = Not CBool(.itemnumeric(v))
            Case vbInteger
                .item(v) = CInt(Not CBool(.itemnumeric(v)))
            Case vbLong
                .item(v) = CLng(Not CBool(.itemnumeric(v)))
            Case vbCurrency
                .item(v) = CCur(Not CBool(.itemnumeric(v)))
            Case vbDecimal
                .item(v) = CDec(Not CBool(.itemnumeric(v)))
            Case Else
                .item(v) = CDbl(Not CBool(.itemnumeric(v)))
            End Select
            GoTo loopcontinue1
      ElseIf FastSymbol(b$, ":=", , 2) Then

    If IsExp(bstack, b$, p) Then
        .item(v) = p
    ElseIf IsStrExp(bstack, b$, ss$, Len(bstack.tmpstr) = 0) Then
      If Not MyIsObject(.item(v)) Then
          .item(v) = ss$
          Else
        CheckVar .item(v), ss$
        
        End If

    Else
        Exit Do
    End If
    If FastSymbol(b$, ",") Then v = v + 1: GoTo contarr1
    GoTo loopcontinue1
        End If
.item(v) = MyRound(.itemnumeric(v), 13)
GoTo loopcontinue1
End With
End If


If IsOperator0(b$, ".") Then

If pppp.ItemType(v) = "Group" Then
interpret = SpeedGroup(bstack, pppp, "", w$, b$, v)
Set pppp = Nothing
GoTo loopcontinue1
End If
ElseIf IsOperator(b$, "(") Then
If pppp.ItemType(v) = myArray Then
Set pppp = pppp.item(v)
GoTo againarray22
End If
ElseIf Not FastSymbol(b$, "=") Then
here$ = ohere$: GoTo there1
End If

If Not IsExp(bstack, b$, p) Then here$ = ohere$: GoTo there1

 If Not bstack.lastobj Is Nothing Then
     Set myobject = pppp.GroupRef
     If pppp.IHaveClass Then

            Set pppp.item(v) = bstack.lastobj
            Set pppp.item(v).LinkRef = myobject
            With pppp.item(v)
                 .HasStrValue = myobject.HasStrValue
                .HasValue = myobject.HasValue
                .HasSet = myobject.HasSet
                .HasParameters = myobject.HasParameters
                .HasParametersSet = myobject.HasParametersSet
                
                Set .SuperClassList = myobject.SuperClassList
                Set .Events = myobject.Events
                .highpriorityoper = myobject.highpriorityoper
                .HasUnary = myobject.HasUnary
            End With
     Else
            If Typename(bstack.lastobj) = "mHandler" Then
                               Set pppp.item(v) = bstack.lastobj
     
            Else
                   If Not bstack.lastobj Is Nothing Then
                          If TypeOf bstack.lastobj Is mArray Then
                                 If bstack.lastobj.Arr Then
                                         Set pppp.item(v) = CopyArray(bstack.lastobj)

                                 Else
  
   
                                            Set pppp.item(v) = bstack.lastobj
                                            If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                                 End If
                          Else
                          
                                  Set pppp.item(v) = bstack.lastobj
                                  If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                          End If
                   Else
                  
                          Set pppp.item(v) = bstack.lastobj
                          If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                   End If
            End If
        End If
     
     Set bstack.lastobj = Nothing
     Else
     If pppp.Arr Then
     pppp.item(v) = p
     ElseIf Typename(pppp.GroupRef) = "PropReference" Then
    
     pppp.GroupRef.Value = p
     End If
    End If
Do While FastSymbol(b$, ",")
If pppp.UpperMonoLimit > v Then
v = v + 1
If Not IsExp(bstack, b$, p) Then here$ = ohere$: GoTo there1
If Not bstack.lastobj Is Nothing Then
     Set myobject = pppp.GroupRef
     If pppp.IHaveClass Then
         Set pppp.item(v) = bstack.lastobj
            
            With pppp.item(v)
                 .HasStrValue = myobject.HasStrValue
                .HasValue = myobject.HasValue
                .HasSet = myobject.HasSet
                .HasParameters = myobject.HasParameters
                .HasParametersSet = myobject.HasParametersSet
                 Set .SuperClassList = myobject.SuperClassList
                Set .Events = myobject.Events
                .highpriorityoper = myobject.highpriorityoper
                .HasUnary = myobject.HasUnary
            End With
        
        
     Else
        Set pppp.item(v) = bstack.lastobj
    End If
    Set pppp.item(v).LinkRef = myobject
    Set bstack.lastobj = Nothing
     Else
pppp.item(v) = p
End If
Else
Exit Do
End If
Loop
Else
interpret = False: here$ = ohere$: GoTo there1
End If
Case 6
If neoGetArray(bstack, w$, pppp) Then
    If FastSymbol(b$, ")") Then
    'need to found an expression
        If FastSymbol(b$, "=") Then
            If IsStrExp(bstack, b$, ss$) Then
                If Not bstack.lastobj Is Nothing Then
                    If TypeOf bstack.lastobj Is mHandler Then
                        Set usehandler = bstack.lastobj
                        If usehandler.t1 = 3 Then
                            If pppp.Arr Then
                                usehandler.objref.CopyArray pppp
                                pppp.Final = False
                            Else
                                NotArray
                            End If
                        Else
                            NotArray
                        End If
                        Set usehandler = Nothing
                    Else
                        bstack.lastobj.CopyArray pppp
                    End If
                    Set bstack.lastobj = Nothing
                    GoTo loopcontinue1
                End If
            Else
                SyntaxError
            End If
               interpret = False
            GoTo there1
        End If
    End If
     
againstrarr22:
If Not NeoGetArrayItem(pppp, bstack, w$, v, b$) Then interpret = False: here$ = ohere$: GoTo there1
On Error Resume Next
If pppp.ItemType(v) = myArray And pppp.Arr Then
If FastSymbol(b$, "(") Then
Set pppp = pppp.item(v)
GoTo againstrarr22
End If
End If
If Not FastSymbol(b$, "=") Then
    If FastSymbol(b$, ":=", , 2) Then
contarr1:
    ss$ = Left$(aheadstatus(b$), 1)
        If ss$ = "S" Then
        If Not IsStrExp(bstack, b$, ss$) Then interpret = False: here$ = ohere$: GoTo there1
        Else
        If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
        ss$ = Trim$(Str$(p))
        End If
             If Not MyIsObject(pppp.item(v)) Then
          pppp.item(v) = ss$
          Else
        CheckVar pppp.item(v), ss$
        
        End If
        Do While FastSymbol(b$, ",")
        If pppp.UpperMonoLimit > v Then
        v = v + 1
          ss$ = Left$(aheadstatus(b$), 1)
                        If ss$ = "S" Then
        If Not IsStrExp(bstack, b$, ss$) Then interpret = False: here$ = ohere$: GoTo there1
        Else
        If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
        ss$ = Trim$(Str$(p))
        End If
        
                If Not MyIsObject(pppp.item(v)) Then
                  pppp.item(v) = ss$
                  Else
                CheckVar pppp.item(v), ss$
                
                End If
        Else
        Exit Do
        End If
        Loop
   ElseIf IsOperator(b$, "+=", 2) Then
    If pppp.IsStringItem(v) Then
    If Not IsStrExp(bstack, b$, ss$) Then GoTo st1222
    If bstack.lastobj Is Nothing Then
        pppp.ItemStr(v) = pppp.item(v) + ss$
    Else
st1222:
        MyEr "Need a string", "Χρειάζομαι ένα αλφαριθμητικό"
        interpret = False: here$ = ohere$: GoTo there1
    End If
    Else
    GoTo st1222
    End If
        
    Else
    interpret = False: here$ = ohere$: GoTo there1
    End If
Else
        If Not IsStrExp(bstack, b$, ss$) Then interpret = False: here$ = ohere$: GoTo there1
        
        
    If Not MyIsObject(pppp.item(v)) Then
    If pppp.Arr Then
    If bstack.lastobj Is Nothing Then
        pppp.item(v) = ss$
    
    Else
    If Typename(bstack.lastobj) = myArray Then
    If bstack.lastobj.Arr Then
        Set pppp.item(v) = CopyArray(bstack.lastobj)
    Else
         Set pppp.item(v) = bstack.lastobj.GroupRef
    End If
    Else
        Set pppp.item(v) = bstack.lastobj
        End If
        Set bstack.lastobj = Nothing
        End If
        Else
        pppp.GroupRef.Value = ss$
        End If
    Else
        CheckVar pppp.item(v), ss$
    End If
        Do While FastSymbol(b$, ",")
        If pppp.UpperMonoLimit > v Then
        v = v + 1
                If Not IsStrExp(bstack, b$, ss$) Then here$ = ohere$: GoTo there1
        
                If Not MyIsObject(pppp.item(v)) Then
                  pppp.item(v) = ss$
                  Else
                CheckVar pppp.item(v), ss$
                
                End If
        Else
        Exit Do
        End If
        Loop
End If
Else
interpret = 0: here$ = ohere$: GoTo there1
End If
Case 7
If neoGetArray(bstack, w$, pppp) Then
    If FastSymbol(b$, ")") Then
    'need to found an expression
        If FastSymbol(b$, "=") Then
            If IsStrExp(bstack, b$, ss$) Then
                If Not bstack.lastobj Is Nothing Then
                    bstack.lastobj.CopyArray pppp
                    Set bstack.lastobj = Nothing
                    GoTo loopcontinue1
                End If
            Else
                SyntaxError
            End If
              interpret = False
            GoTo there1
        End If
        End If
againintarr7:
If Not NeoGetArrayItem(pppp, bstack, w$, v, b$) Then interpret = False: here$ = ohere$: GoTo there1
On Error Resume Next
If pppp.ItemType(v) = myArray And pppp.Arr Then
If FastSymbol(b$, "(") Then
Set pppp = pppp.item(v)
GoTo againintarr7
End If
End If
If MaybeIsSymbol(b$, "+-*/~") Then
If IsOperator0(b$, "++", 2) Then
pppp.item(v) = pppp.itemnumeric(v) + 1
ElseIf IsOperator0(b$, "--", 2) Then
pppp.item(v) = pppp.itemnumeric(v) - 1
ElseIf IsOperator(b$, "+=", 2) Then
If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
pppp.item(v) = pppp.itemnumeric(v) + MyRound(p)
ElseIf IsOperator(b$, "-=", 2) Then
If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
pppp.item(v) = pppp.itemnumeric(v) - MyRound(p)
ElseIf IsOperator(b$, "*=", 2) Then
If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
pppp.item(v) = MyRound(pppp.itemnumeric(v) * MyRound(p))
ElseIf IsOperator(b$, "/=", 2) Then
If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
If MyRound(p) = 0 Then
 DevZero
 Else
 pppp.item(v) = MyRound(pppp.itemnumeric(v) / MyRound(p))
End If
ElseIf IsOperator0(b$, "-!", 2) Then
pppp.item(v) = -pppp.itemnumeric(v)
ElseIf IsOperator0(b$, "~") Then
        With pppp
        Select Case VarType(.itemnumeric(v))
            Case vbBoolean
                .item(v) = Not CBool(.itemnumeric(v))
            Case vbInteger
                .item(v) = CInt(Not CBool(.itemnumeric(v)))
            Case vbLong
                .item(v) = CLng(Not CBool(.itemnumeric(v)))
            Case vbCurrency
                .item(v) = CCur(Not CBool(.itemnumeric(v)))
            Case vbDecimal
                .item(v) = CDec(Not CBool(.itemnumeric(v)))
            Case Else
                .item(v) = CDbl(Not CBool(.itemnumeric(v)))
        End Select
        End With
End If

GoTo loopcontinue1
End If
If Not FastSymbol(b$, "=") Then here$ = ohere$: GoTo there1
If Not IsExp(bstack, b$, p) Then here$ = ohere$: GoTo there1
If Not bstack.lastobj Is Nothing Then
    If TypeOf bstack.lastobj Is mArray Then
                                 If bstack.lastobj.Arr Then
                                         Set pppp.item(v) = CopyArray(bstack.lastobj)

                                 Else
  
   
                                            Set pppp.item(v) = bstack.lastobj
                                            If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                                 End If
                          Else
                          
                                  Set pppp.item(v) = bstack.lastobj
                                  If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                          End If
Else
p = MyRound(p)

If Err.Number > 0 Then interpret = False: here$ = ohere$: GoTo there1
pppp.item(v) = p
End If
Do While FastSymbol(b$, ",")

If pppp.UpperMonoLimit > v Then
v = v + 1
If Not IsExp(bstack, b$, p) Then here$ = ohere$: GoTo there1
pppp.item(v) = MyRound(p)
Else
Exit Do
End If
Loop
Else
interpret = False: here$ = ohere$: GoTo there1
End If
Case Else
If FastSymbol(b$, "(") Then
            i = 1
            x1 = 0
            While Len(aheadstatus(b$, False, i)) > 0
                x1 = i - 1
                i = i + 1
            Wend
            ss$ = Left$(b$, x1)
            If x1 > 0 And MyTrim(ss$) <> vbNullString Then
                Mid$(b$, 1, x1) = space$(x1)
                If FastSymbol(b$, ")", True) Then
                    If FastSymbol(b$, "=") Then
                        If IsExp(bstack, b$, p) Then
                            If Not bstack.lastobj Is Nothing Then
                                If TypeOf bstack.lastobj Is mArray Then
                                Set pppp = bstack.lastobj
                                GoTo wehavearray
                                ElseIf TypeOf bstack.lastobj Is mHandler Then
                                    If CheckIsmArray(bstack.lastobj) Then
                                        Set usehandler = bstack.lastobj
                                        Set pppp = usehandler.objref
                                        Set usehandler = Nothing
wehavearray:
                                        Set bstack.lastobj = Nothing
                                        Set myobject = bstack.soros
                                        Set bstack.Sorosref = New mStiva
                                        bstack.soros.MergeBottomCopyArray pppp
                                        If Not MyRead(1, bstack, ss$, 1) Then
                                            Set bstack.lastobj = Nothing
                                            Set bstack.Sorosref = myobject
                                            interpret = False
                                            Exit Function
                                        End If
                                        Set bstack.lastobj = Nothing
                                        Set bstack.Sorosref = myobject
                                        Set myobject = Nothing
                                        GoTo loopcontinue1
                                    Else
a123321:                                    NotArray
                                            interpret = False
                                            Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        GoTo a123321
        End If



If MaybeIsSymbol(b$, ",-+*/_!@()[];<>|~`'\") Then
SyntaxError
End If
End Select


loopcontinue1:
If MaybeIsSymbol(b$, "'\/") Then
ElseIf Not NocharsInLine(b$) Then
    If Not MaybeIsSymbol(b$, vbCr) Then
        If Not MaybeIsSymbol(b$, ":") Then Exit Do
    End If
End If
Loop
here$ = ohere$
If LastErNum = -2 Then
sss = CLng(Execute(bstack, b$, True))
b$ = vbNullString
interpret = False

GoTo there1
forlong:
OverflowLong
interpret = False

GoTo there1


ElseIf LastErNum <> 0 Then
b$ = " "
End If
interpret = b$ = vbNullString
If Not interpret Then
If LastErNum = 0 Then SyntaxError
End If
there1:
bstack.LoadOnly = False
End Function
Public Sub PushErrStage(basestack As basetask)
        With basestack.RetStack
                        .PushVal subHash.count
                        .PushVal varhash.count
                        .PushVal sb2used
                        .PushVal basestack.SubLevel
                        .PushVal var2used

                        .PushVal -4
                         basestack.ErrVars = var2used
        End With
       
End Sub
Function PrepareLambda(basestask As basetask, myl As lambda, ByVal v As Long, frm$, c As Constant) As Boolean
On Error GoTo 1234
If Typename(var(v)) = "Constant" Then
    Set c = var(v)
    If Not c.flag Then
    InternalError
    PrepareLambda = False
    Exit Function
    End If
    Set myl = c.Value
Else
    Set myl = var(v)
End If
         myl.Name = here$
            
            myl.CopyToVar basestask, here$ = vbNullString, var()
            basestask.OriginalCode = -v
            basestask.FuncRec = subHash.LastKnown

            frm$ = myl.code$
PrepareLambda = True
Exit Function
1234
InternalError
PrepareLambda = False

End Function

Sub BackPort(a$)
If Len(a$) = 0 Then a$ = Chr(8) Else Mid$(a$, 1, 1) = Chr(8)
End Sub
Function ExistNum(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim p As Variant, dd As Long, dn As Long, x As Variant, anything As Object, s$, usehandler As mHandler
ExistNum = False
    If IsExp(bstack, a$, p) Then
    If Typename(bstack.lastobj) = "mHandler" Then
    Set anything = bstack.lastobj
    
    Set bstack.lastobj = Nothing
       If Not CheckLastHandler(anything) Then
        InternalError
        ExistNum = False
        Exit Function
        End If
    Set usehandler = anything
    With usehandler
        If TypeOf .objref Is FastCollection Then
            If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p, , True) Then
                    If FastSymbol(a$, ",") Then
                        If IsExp(bstack, a$, x, , True) Then
                        x = Int(x)
                        If x = 0 Then
                                R = .objref.FindOne(p, x)
                                R = SG * x
                        ElseIf x > 0 Then
                            dn = x
                            x = 0
                            R = .objref.FindOne(p, x)
                            x = -x + dn - 1
                            R = SG * .objref.FindOne(p, x)
                            Else
                                R = SG * .objref.FindOne(p, x)
                            End If
                        Else
                            MissParam a$
                            Set anything = Nothing
                            Exit Function
                        End If
                    Else
                    R = SG * .objref.Find(p)
                    End If
                ElseIf IsStrExp(bstack, a$, s$) Then
                    Set bstack.lastobj = Nothing
                    If FastSymbol(a$, ",") Then
                        If IsExp(bstack, a$, x, , True) Then
                            x = Int(x)
                            If x = 0 Then
                                R = .objref.FindOne(s$, x)
                                R = SG * x
                        ElseIf x > 0 Then
                            dn = x
                            x = 0
                            R = .objref.FindOne(s$, x)
                            x = -x + dn - 1
                            R = SG * .objref.FindOne(s$, x)
                            Else
                                R = SG * .objref.FindOne(s$, x)
                            End If
                        Else
                            MissParam a$
                            Exit Function
                        End If
                    Else
                        R = SG * .objref.Find(s$)
                        
                    End If
                End If
                
                
                ExistNum = FastSymbol(a$, ")", True)
                
                Exit Function
            End If
        End If
    End With
      
    End If

    MissParam a$
    Set bstack.lastobj = Nothing
    ElseIf IsStrExp(bstack, a$, s$) Then
    s$ = CFname(s$)
    If s$ <> "" Then
      R = SG * (InStr(s$, "*") = 0 And InStr(s$, "?") = 0)
      Else
  
    R = 0
    End If
   
    
    
    ExistNum = FastSymbol(a$, ")", True)
    Else
        MissParam a$
    End If

    
End Function
Function MySwap(bstack As basetask, rest$, Lang As Long) As Boolean
Dim s$, ss$, F As Long, Col As Long, x1 As Long, i As Long, pppp As mArray, pppp1 As mArray
    F = Abs(IsLabel(bstack, rest$, s$))
    MySwap = True
    If F = 1 Or F = 4 Then Col = 1
    If F = 5 Or F = 7 Then Col = 2
    If F = 0 Then MissingnumVar:  Exit Function
    If (F = 3 Or F = 6) And Col > 0 Then SyntaxError: MySwap = False:    Exit Function
    If Col = 1 Then
        If GetVar(bstack, s$, F) Then
                If Not FastSymbol(rest$, ",") Then MissingnumVar:  Exit Function
                i = Abs(IsLabel(bstack, rest$, ss$))
              If i = 1 Or i = 4 Then
                If GetVar(bstack, ss$, x1) Then
         If MyIsObject(var(F)) Then
            If var(F) Is Nothing Then
            ElseIf TypeOf var(F) Is Constant Then
            CantAssignValue
            MySwap = False: Exit Function
            End If
        End If
        If MyIsObject(var(x1)) Then
            If TypeOf var(x1) Is Constant Then
            CantAssignValue
            MySwap = False: Exit Function
            End If
        End If
                    SwapVariant var(F), var(x1)
                    
                    
                Exit Function
                Else
                    Nosuchvariable ss$
                    MySwap = False
                    Exit Function
                End If
            ElseIf i = 5 Or i = 7 Then
                If neoGetArray(bstack, ss$, pppp) Then
                If Not pppp.Arr Then NotArray: Exit Function
                    If Not NeoGetArrayItem(pppp, bstack, ss$, x1, rest$, True) Then Exit Function
                        If MyIsObject(var(F)) Then
                            If TypeOf var(F) Is Constant Then
                            CantAssignValue
                            MySwap = False: Exit Function
                            End If
                        End If
                    SwapVariant2 var(F), pppp, x1
                    
                    
                Else
                
                    NoSwap ss$
                    MySwap = False
                    Exit Function
                End If
            Else
                MissingnumVar
                MySwap = False
                Exit Function
            End If
        Else
            Nosuchvariable s$
            
            Exit Function
        End If
    ElseIf Col = 2 Then
        If neoGetArray(bstack, s$, pppp) Then
        If Not pppp.Arr Then NotArray: Exit Function
        If Not NeoGetArrayItem(pppp, bstack, s$, F, rest$) Then Exit Function
            If Not FastSymbol(rest$, ",") Then MissingnumVar:  Exit Function
                i = Abs(IsLabel(bstack, rest$, ss$))
                  
            If i = 1 Or i = 4 Then
                    If GetVar(bstack, ss$, x1) Then
                    If pppp.IHaveClass Then
                            NoSwap ""
                    Else
                             If MyIsObject(var(x1)) Then
                                If TypeOf var(x1) Is Constant Then
                                CantAssignValue
                                MySwap = False: Exit Function
                                End If
                            End If
                    
                    
                         SwapVariant2 var(x1), pppp, F
                     End If
                        
                    Else
                        MissingnumVar
                        MySwap = False
                        Exit Function
                    End If
            ElseIf i = 5 Or i = 7 Then
                    If neoGetArray(bstack, ss$, pppp1) Then
                    If Not pppp1.Arr Then NotArray: Exit Function
                        If Not NeoGetArrayItem(pppp1, bstack, ss$, x1, rest$) Then Exit Function
                   If pppp.IHaveClass Xor Not pppp1.IHaveClass Then
                            
                        SwapVariant3 pppp, F, pppp1, x1
                        If pppp.IHaveClass Then
                            Set pppp.item(F).LinkRef = pppp1.GroupRef
                            Set pppp1.item(x1).LinkRef = pppp.GroupRef
                            End If
                        Else
                        NoSwap ""
                        Exit Function
                        End If
                        
                    Else
                        MissingnumVar
                        
                        Exit Function
                    End If
            Else
                MissingnumVar
                
                Exit Function
            End If
        Else
            MissingnumVar
            
            Exit Function
        End If
    ElseIf F = 3 Then
            If GetVar(bstack, s$, F) Then
            If Not FastSymbol(rest$, ",") Then MissingnumVar:  Exit Function
                i = Abs(IsLabel(bstack, rest$, ss$))
                 If i = 6 Then
                    If Not neoGetArray(bstack, ss$, pppp) Then MissingStrVar:  Exit Function
                    If Not pppp.Arr Then NotArray: Exit Function
                    If Not NeoGetArrayItem(pppp, bstack, ss$, x1, rest$) Then Exit Function
                     If MyIsObject(var(F)) Then
                        If TypeOf var(F) Is Constant Then
                        CantAssignValue
                        MySwap = False: Exit Function
                        End If
                    End If
                    SwapVariant2 var(F), pppp, x1

                ElseIf i = 3 Then
                    If Not GetVar(bstack, ss$, x1) Then: Exit Function
                     If MyIsObject(var(F)) Then
                        If TypeOf var(F) Is Constant Then
                        CantAssignValue
                        MySwap = False: Exit Function
                        End If
                    End If
                   SwapVariant var(F), var(x1)
                Else
                MissFuncParameterStringVar
                MySwap = False
                End If
                
                
            Else
                    
                    MissFuncParameterStringVar
                    MySwap = False
            End If
    ElseIf F = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
            If Not pppp.Arr Then NotArray: Exit Function
                If Not NeoGetArrayItem(pppp, bstack, s$, x1, rest$) Then Exit Function
                If Not FastSymbol(rest$, ",") Then MissingnumVar:  Exit Function
                i = Abs(IsLabel(bstack, rest$, ss$))
     
                If i = 6 Then
                    If Not neoGetArray(bstack, ss$, pppp1) Then MissingStrVar:  Exit Function
                    If Not pppp.Arr Then NotArray: Exit Function
                    If Not NeoGetArrayItem(pppp1, bstack, ss$, i, rest$) Then Exit Function

                   SwapVariant3 pppp, x1, pppp1, i
 
                ElseIf i = 3 Then
                    If Not GetVar(bstack, ss$, i) Then: Exit Function
                    If MyIsObject(var(i)) Then
                        If TypeOf var(i) Is Constant Then
                        CantAssignValue
                        MySwap = False: Exit Function
                        End If
                    End If


                  SwapVariant2 var(i), pppp, x1
                    Else
                MissFuncParameterStringVar
                MySwap = False
                End If
                
                
            Else
                
                MissPar
                MySwap = False
                
            End If
    Else
                 
                MissPar
                MySwap = False
    End If
    Exit Function

End Function
Public Function TraceThis(bstack As basetask, di As Object, b$, w$, SBB$) As Boolean
    TraceThis = True
    PrepareLabel bstack
    Form2.label1(1) = w$
    Form2.label1(2) = GetStrUntil(vbCrLf, b$ & vbCrLf, False)
    If Len(b$) = 0 Then
    WaitShow = 0
    bypassST = False
    Set Form2.Process = bstack
    Exit Function
    Else
    If TestShowBypass Then
    
        ElseIf WaitShow = 0 Or Len(b$) < WaitShow Then
            WaitShow = 0
            If bstack.OriginalCode < 0 Then
            lasttracecode = -bstack.OriginalCode
                SBB$ = GetNextLine((var(-bstack.OriginalCode).code$))
            Else
            lasttracecode = bstack.OriginalCode
                SBB$ = GetNextLine((sbf(Abs(bstack.OriginalCode)).sb))
            End If
            If Left$(SBB$, 10) = "'11001EDIT" Then
                TestShowSub = Mid$(sbf(Abs(bstack.OriginalCode)).sb, Len(SBB$) + 3)
                If TestShowSub = vbNullString Then
                    TestShowSub = Mid$(sbf(FindPrevOriginal(bstack)).sb, Len(SBB$) + 3)
                End If
                If InStr(TestShowSub, b$) = 0 Then
                    WaitShow = Len(b$)
                End If
            Else
                If bstack.OriginalCode <> 0 Then
                    If bstack.OriginalCode < 0 Then
                        TestShowSub = var(-bstack.OriginalCode).code$
                    Else
                        TestShowSub = sbf(Abs(bstack.OriginalCode)).sb
                    End If
                Else
                    If bstack.IamThread Then
                        If bstack.Process Is Nothing Then
                        Else
                            TestShowSub = bstack.Process.CodeData
                        End If
                    Else
                        TestShowSub = b$
                    End If
                End If
            End If
        End If
        If bstack.addlen Then
            If Len(TestShowSub) - bstack.addlen - Len(b$) > 0 Then
                TestShowStart = Len(TestShowSub) - bstack.addlen - Len(b$) + 1
            Else
                TestShowStart = 1
            End If
        Else
            TestShowStart = Len(TestShowSub) - Len(b$) + 1 ' rinstr(TestShowSub, b$)
        End If
        If TestShowStart <= 0 Then
            TestShowStart = rinstr(TestShowSub, Mid$(b$, 2)) - 1
        End If
     bypassST = False
          
    Set Form2.Process = bstack
    stackshow bstack
        
    End If
    
    If Not Form1.Visible Then
        Form1.Show , Form5   'OK
    End If

    If STbyST Then
        STbyST = False
        If Not STEXIT Then
            If Not STq Then
                Form2.gList4.ListIndex = 0
            End If
        End If
        If Not TaskMaster Is Nothing Then
            If TaskMaster.QueueCount > 0 And TaskMaster.Processing Then TaskMaster.StopProcess
        End If
      
        Do
            BLOCKkey = False
            If Not IsWine Then If di.Visible Then di.Refresh
            ProcTask2 bstack
        Loop Until STbyST Or STq Or STEXIT Or bypassST Or NOEXECUTION Or myexit(bstack) Or Not Form2.Visible

        If Not TaskMaster Is Nothing Then
           If TaskMaster.QueueCount > 0 And Not TaskMaster.Processing Then TaskMaster.StartProcess
        End If
        If Not STEXIT Then
            If Not STq Then
                Form2.gList4.ListIndex = 0
            End If
        End If
        STq = False
        If STEXIT Then
            NOEXECUTION = True
            trace = False
            STEXIT = False
            TraceThis = False
            Exit Function
        End If
    Else
If tracecounter > 0 Then If Not IsWine Then MyDoEvents1 Form2, True

    End If
    If STEXIT Then
        trace = False
        STEXIT = False
        TraceThis = False
        Exit Function
    End If
End Function


Function DriveSerial1(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
    Dim s$
    If IsStrExp(bstack, a$, s$) Then
    R = SG * DriveSerial(Left$(s$, 3))
  
    
    
    DriveSerial1 = FastSymbol(a$, ")", True)
    Else
         MissParam a$
    End If
End Function
Function MakeForm(basestack As basetask, rest$) As Boolean
On Error Resume Next
MakeForm = True
Dim Scr As Object, xx As Single, p As Variant, x1 As Long, y1 As Long, x As Double, y As Double
Dim w3 As Long, w4 As Long, sX As Double, adjustlinespace As Boolean, SZ As Single, reduce As Single
reduce = 1
Set Scr = basestack.Owner

Dim basketcode As Long, mAddTwipsTop As Long


If Left$(Typename(Scr), 3) = "Gui" Then
If Typename(Scr) = "GuiM2000" Then FastSymbol rest$, "!": GoTo there1
ElseIf Scr.Name = "Form1" Then

Else
If FastSymbol(rest$, "!") Then
If Not TypeOf Scr Is MetaDc Then reduce = 0.9
End If
there1:

basketcode = GetCode(Scr)
With players(basketcode)
SetNormal Scr
mAddTwipsTop = .uMineLineSpace  ' the basic
Dim wishX As Long, wishY As Long
If IsExp(basestack, rest$, p) Then
    If p < 10 Then p = 10
    x = 4
    xx = 4
    If Scr.Name = "DIS" Then
    Do
    y = CDbl(xx)
    xx = CSng(x)
    nForm basestack, xx, w3, w4, mAddTwipsTop   'using line spacing
    If xx > CSng(x) Then x = CDbl(xx)
    
    If Form1.Width * reduce < w3 * p Then Exit Do
    x = x + 0.25
    Loop
 
    
    Else
    Do
    
    y = CDbl(xx)
    xx = CSng(x)

    nForm basestack, xx, w3, w4, mAddTwipsTop  'using line spacing
    If xx > CSng(x) Then x = CDbl(xx)
    
    If Scr.Width * reduce < w3 * p Then Exit Do
    
    x = x + 0.4
    Loop
    End If
    x = y
    sX = 0
    wishX = p
    If FastSymbol(rest$, ",") Then
        If IsExp(basestack, rest$, sX) Then
        '' ok
        wishY = sX
       mAddTwipsTop = 0  ' find a new one
       players(basketcode).MineLineSpace = 0
       players(basketcode).uMineLineSpace = 0
        adjustlinespace = True
    ''    mmx = scr.Width
''mmy = scr.Height
        Else
        MakeForm = False
        MissNumExpr
        Set Scr = Nothing
        Exit Function
        End If
   
End If
If FastSymbol(rest$, ";") And Scr.Name = "DIS" Then
adjustlinespace = False
If IsWine Then
    Form1.move ScrInfo(Console).Left, ScrInfo(Console).top, ScrInfo(Console).Width - 1, ScrInfo(Console).Height - 1
Else
    Form1.move ScrInfo(Console).Left, ScrInfo(Console).top, ScrInfo(Console).Width, ScrInfo(Console).Height
End If
    Form1.backcolor = players(-1).Paper
    
Sleep 1
End If
nForm basestack, CSng(x), w3, w4, 0
Dim mmx As Long, mmy As Long
If sX = 0 Then
SZ = CSng(x)
mmx = Scr.Width * reduce
 If Scr.Name = "DIS" Then
 mmy = CLng(mmx * Form1.Height / Form1.Width) ' WHY 3/4 ??
 Else
 mmy = Scr.Width * reduce
 End If
 players(basketcode).MineLineSpace = mAddTwipsTop
 players(basketcode).uMineLineSpace = mAddTwipsTop
FrameText Scr, SZ, CLng(w3 * p), mmy, players(basketcode).Paper
Else
If Scr.Name = "DIS" Then
If (sX * w4) > Form1.Height * reduce Then
y = Form1.Height * reduce
Dim safety As Long

While sX * w4 > Form1.Height * reduce And Not safety = sX * w4
safety = sX * w4

xx = y / (dv20 * sX)

nForm basestack, xx, w3, w4, 0  'using no spacing so we put a lot of lines
x = CDbl(xx)
y = y * 0.9
Wend


End If
Else
If sX * w4 > Scr.Height * reduce Then
y = Scr.Height * reduce
Do While sX * w4 > Scr.Height * reduce

xx = y / (dv20 * sX)
nForm basestack, xx, w3, w4, 0  'using no spacing so we put a lot of lines
If x = CDbl(xx) Then Exit Do
x = CDbl(xx)
y = y * 0.9
Loop


End If

End If

If Scr.Name = "DIS" Then
If Not adjustlinespace Then If Scr.Height * reduce >= Form1.Height * reduce - dv15 Then mAddTwipsTop = dv15 * (((Scr.Height * reduce - sX * w4) / sX / 2) \ dv15)
End If
nForm basestack, (x), w3, w4, mAddTwipsTop
SZ = CSng(x)
'If mmx < scr.Width Then
mmx = Scr.Width * reduce


'If mmx < scr.Width Then
mmy = Scr.Height * reduce
If adjustlinespace Then
If Scr.Name = "DIS" Then
mAddTwipsTop = dv15 * (((Form1.Height * reduce - sX * w4) / sX / 2) \ dv15)

Else
mAddTwipsTop = dv15 * (((Scr.Height * reduce - sX * w4) / sX / 2) \ dv15)
End If
sX = CLng(sX * (w4 + mAddTwipsTop * 2))
Else
sX = CLng(sX * w4)
End If
players(basketcode).MineLineSpace = mAddTwipsTop
players(basketcode).uMineLineSpace = mAddTwipsTop
FrameText Scr, SZ, CLng(w3 * p), CLng(sX), players(basketcode).Paper, Not (Scr.Name = "DIS")

End If


ElseIf FastSymbol(rest$, ";") And Scr.Name = "DIS" Then



If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top
If IsWine Then
         Form1.Width = ScrInfo(Console).Width - 1
         Form1.Height = ScrInfo(Console).Height
         Form1.move ScrInfo(Console).Left, ScrInfo(Console).top
Else
        Form1.move ScrInfo(Console).Left, ScrInfo(Console).top, ScrInfo(Console).Width, ScrInfo(Console).Height
    
End If
Form1.backcolor = players(-1).Paper
Form1.Cls
With players(-1)
        .mysplit = 0
        .MAXXGRAPH = Form1.Width
        .MAXYGRAPH = Form2.Height
        SetText Form1
        End With
MyMode Scr
ElseIf Scr.Name = "DIS" Then

w3 = Form1.Left + Scr.Left
w4 = Form1.top + Scr.top
If IsWine And Form1.Width = ScrInfo(Console).Width Then Form1.Width = ScrInfo(Console).Width - dv15
If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top: w4 = Form1.top + Scr.top
scrMove00 Scr
If IsWine Then
    If Scr.Width = ScrInfo(Console).Width Then
       Form1.Width = Scr.Width - 1
    Else
        Form1.Width = Scr.Width
    End If
    Form1.Height = Scr.Height
    Form1.move w3, w4
Else
    Form1.move w3, w4, Scr.Width, Scr.Height
End If
Form1.Cls
        With players(-1)
        .mysplit = 0
        .MAXXGRAPH = Form1.Width
        .MAXYGRAPH = Form2.Height
        SetText Form1
        End With
SetText Scr

Set Scr = Nothing
Exit Function
Else
'' CROP LAYER
If basketcode > 0 Then
With players(basketcode)
.MAXXGRAPH = .mx * .Xt
.MAXYGRAPH = .My * .Yt
End With
With Form1.dSprite(basestack.tolayer)
.move .Left, .top, players(basketcode).MAXXGRAPH, players(basketcode).MAXYGRAPH
End With

End If
End If

players(basketcode).MineLineSpace = mAddTwipsTop
players(basketcode).uMineLineSpace = mAddTwipsTop
MakeForm = True
.curpos = 0
.currow = 0

End With
End If
SetText Scr

If TypeOf Scr Is MetaDc Then
x1 = (players(basketcode).mx \ 20) * 20
players(basketcode).mx = (players(basketcode).mx \ (players(basketcode).Column + 1)) * (players(basketcode).Column + 1)
If x1 >= wishX Then
    x1 = (wishX \ 20) * 20
    If x1 < wishX Then
        x1 = (wishX \ 8) * 8
        If x1 < wishX Then x1 = x1 + 16
        players(basketcode).mx = x1: players(basketcode).Column = 7
    Else
        players(basketcode).mx = x1: players(basketcode).Column = 9
    End If
End If
If players(basketcode).mx = 0 Then SetText Scr
players(basketcode).MAXXGRAPH = players(basketcode).Xt * players(basketcode).mx
Scr.Width = players(basketcode).MAXXGRAPH
If wishY > 0 Then
players(basketcode).MAXYGRAPH = wishY * players(basketcode).Yt
If players(basketcode).MAXYGRAPH > Scr.Height Then
SetText Scr, 0, True
players(basketcode).MAXYGRAPH = wishY * (players(basketcode).Yt) + dv15
players(basketcode).MineLineSpace = 0
players(basketcode).uMineLineSpace = 0
mAddTwipsTop = 0
End If
players(basketcode).My = wishY
Scr.Height = players(basketcode).MAXYGRAPH


End If
SetText Scr, mAddTwipsTop, True
End If
End Function

Sub ClearLoadedForms()
Dim i As Long, j As Long, Start As Long
j = Forms.count
Debug.Print ""
While j > 0
For i = Start To Forms.count - 1
If TypeOf Forms(i) Is GuiM2000 Then Unload Forms(i): Start = i: Exit For
Next i
j = j - 1
Wend
End Sub
Function getSafeFormList() As LongHash
Dim i As Long, mycol As safeforms
If Not varhash.Find(ChrW(&HFFBF) + here$, i) Then
i = AllocVar()
varhash.ItemCreator ChrW(&HFFBF) + here$, i
Set mycol = New safeforms
Set var(i) = mycol
Else
Set mycol = var(i)
End If
Set getSafeFormList = mycol.mylist
End Function
Function ProcBrowser(bstack As basetask, rest$, Lang As Long) As Boolean
Dim s$, w$, x As Double
ProcBrowser = True
If Not IsStrExp(bstack, rest$, s$) Then

    If Not Abs(IsLabelFileName(bstack, rest$, s$, , w$)) = 1 Then
         If NOEDIT Then
                If Form1.view1.Visible Then
                    Form1.KeyPreview = True
                    ProcTask2 bstack
                    Form1.view1.SetFocus: Form1.KeyPreview = False
                Else
                    Form1.KeyPreview = True
                End If
        End If
            Exit Function
    Else
     s$ = w$ '' low case
    End If
End If
            If FastSymbol(rest$, ",") Then
                    If IsExp(bstack, rest$, x) Then IEX = CLng(x): IESizeX = Form1.Scalewidth - IEX Else MissNumExpr: ProcBrowser = False: Exit Function
                If FastSymbol(rest$, ",") Then
                    If IsExp(bstack, rest$, x) Then IEY = CLng(x): IESizeY = Form1.Scaleheight - IEY Else MissNumExpr: ProcBrowser = False: Exit Function
                                If FastSymbol(rest$, ",") Then
                    If IsExp(bstack, rest$, x) Then IESizeX = CLng(x) Else MissNumExpr: ProcBrowser = False: Exit Function
                                    If FastSymbol(rest$, ",") Then
                    If IsExp(bstack, rest$, x) Then IESizeY = CLng(x) Else MissNumExpr: ProcBrowser = False: Exit Function
                 End If
                End If
             End If
           End If
           If IESizeX = 0 Or IESizeY = 0 Then
           IEX = Form1.Scalewidth / 8
           IEY = Form1.Scaleheight / 8
           IESizeX = Form1.Scalewidth * 6 / 8
           IESizeY = Form1.Scaleheight * 6 / 8
           End If

If myLcase(Left$(s$, 8)) = "https://" Or myLcase(Left$(s$, 7)) = "http://" Or myLcase(Left$(s$, 4)) = "www." Or myLcase(Left$(s$, 6)) = "about:" Then
Form1.IEUP s$
ElseIf s$ <> "" Then
Form1.IEUP "file:" & strTemp + s$
Else
Form1.IEUP ""
Form1.KeyPreview = True
End If
ProcTask2 bstack

End Function

Function MyScore(bstack As basetask, rest$) As Boolean
Dim s$, sX As Double, p As Variant
MyScore = False
If IsExp(bstack, rest$, p) Then
If p >= 1 And p <= 16 Then
If FastSymbol(rest$, ",") Then
If IsExp(bstack, rest$, sX) Then
If FastSymbol(rest$, ",") Then
If IsStrExp(bstack, rest$, s$) Then
voices(p - 1) = s$
BEATS(p - 1) = sX
MyScore = True
End If
End If
End If
End If
End If
End If
End Function

Function MyPlayScore(bstack As basetask, rest$) As Boolean
Dim task As TaskInterface, sX As Double, p As Variant

MyPlayScore = True
If IsExp(bstack, rest$, p) Then
    If p = 0 Then
    TaskMaster.MusicTaskNum = 0
    TaskMaster.OnlyMusic = True
    Do
    TaskMaster.TimerTickNow
    Loop Until TaskMaster.PlayMusic = False
    TaskMaster.OnlyMusic = False   '' forget it in revision 130
   mute = True
    Else
    mute = False
    If FastSymbol(rest$, ",") Then
        If IsExp(bstack, rest$, sX) Then
          If sX < 1 Then
          sX = 0
          Do While TaskMaster.ThrowOne(CLng(p))
          sX = sX - 1
          If sX < -100 Then Exit Do
          Loop
          Else
          Set task = New MusicBox
          Set task.Owner = Form1.DIS
         
          task.Parameters CLng(p), CLng(sX)
          TaskMaster.MusicTaskNum = TaskMaster.MusicTaskNum + 1
          TaskMaster.AddTask task
          End If
          Do While FastSymbol(rest$, ",")
           MyPlayScore = False
        If IsExp(bstack, rest$, p) Then
             If FastSymbol(rest$, ",") Then
                If IsExp(bstack, rest$, sX) Then
                If sX < 1 Then
                        sX = 0
                        Do While TaskMaster.ThrowOne(CLng(p))
                        sX = sX - 1
                        If sX < -100 Then Exit Do
                        Loop
                  Else
                    Set task = New MusicBox
                    Set task.Owner = Form1.DIS
                    task.Parameters CLng(p), CLng(sX)
                    TaskMaster.MusicTaskNum = TaskMaster.MusicTaskNum + 1
                     TaskMaster.AddTask task
              End If
                MyPlayScore = True
                 End If
            End If
        End If
        If MyPlayScore = False Then
          mute = True
        Exit Do
        End If
          Loop
        End If
    End If
    End If
Else

MyPlayScore = False
End If
End Function

Function IdPara(basestack As basetask, rest$, Lang As Long) As Boolean
Dim x1 As Long, y1 As Long, i As Long, it As Long, vvl As Variant
Dim x As Double, y As Double, s$, what$, w3 As Long, w4 As Long, z As Double
Dim xa As Long, ya As Long
Dim pppp As mArray


IdPara = True
If IsLabelSymbolNew(rest$, "ΣΤΟ", "TO", Lang) Then
        If Not IsExp(basestack, rest$, y) Then
            MissNumExpr
            IdPara = False
            Exit Function
        Else
        
          y = y - 1
                     If y < 0 Then y = -1
         If FastSymbol(rest$, ",") Then
                    If IsExp(basestack, rest$, x) Then
                        x = Int(x)
                        If x < 1 Then
                        MyErMacro rest$, "the index base must be >=1", "η βάση δείκτη πρέπει να είναι >=1"
                        
                        Exit Function
                        End If
                   
                    End If
                    If FastSymbol(rest$, ",") Then
                     If IsExp(basestack, rest$, z) Then
                         z = Int(z)
                         If z < 1 Then
                         MyErMacro rest$, "the lenght base must be >=1", "το μήκος πρέπει να είναι >=1"
                         
                         Exit Function
                         End If
                    
                     End If
                    Else
                    z = 0
                    End If
            Else
                x = 0
            End If
        
            x1 = Abs(IsLabel(basestack, rest$, what$))
            If x1 = 3 Then
                    If GetVar(basestack, what$, i) Then
                        If Typename(var(i)) = doc Then
                                If Not FastSymbol(rest$, "=") Then
                                    MissSymbolMyEr "="
                                    IdPara = False
                                    Exit Function
                                Else
                                    If Not IsStrExp(basestack, rest$, s$) Then
                                        MissStringExpr
                                        IdPara = False
                                        Exit Function
                                    Else
                                    If y = -1 Then
                                    y = var(i).DocParagraphs
                                    End If
                                   If var(i).ParagraphFromOrder(y + 1) = -1 Then
                                   CheckVar var(i), s$
                                    ElseIf y < 1 Then
                                     w3 = var(i).ParagraphFromOrder(1)
                                     w4 = x
                                       If z > 0 Then
                                    var(i).BackSpaceNchars w3, w4 + CLng(z), CLng(z)
                                    End If
                                    If w3 < 1 Then w3 = 1
                                    If Len(s$) > 0 Then var(i).InsertDoc w3, w4, s$
                                    Else
                                    w3 = var(i).ParagraphFromOrder(y + 1)
                                    w4 = x
                                       If z > 0 Then
                                    var(i).BackSpaceNchars w3, w4 + CLng(z), CLng(z)
                                    End If
                                    If w3 < 1 Then w3 = 1
                                    If Len(s$) > 0 Then var(i).InsertDoc w3, w4, s$
                                    End If
                                    End If
                                End If
                        Else
                             MissingDoc   ' only doc not string var
                             IdPara = False
                            Exit Function
                        End If
                    Else
                        Nosuchvariable what$
                        IdPara = False
                        Exit Function
                    End If
            ElseIf x1 = 6 Then
                    If neoGetArray(basestack, what$, pppp) Then
                        If Not NeoGetArrayItem(pppp, basestack, what$, it, rest$) Then IdPara = False: Exit Function
                        If pppp.ItemType(it) = doc Then
                                    If Not FastSymbol(rest$, "=") Then
                                            MissSymbolMyEr "="
                                            IdPara = False
                                            Exit Function
                                            Else
                                If IsStrExp(basestack, rest$, s$) Then
                                
                                
                                    If pppp.item(it).ParagraphFromOrder(y + 1) = -1 Then
                                       CheckVar pppp.item(it), s$
                                        ElseIf y < 1 Then
                                   w3 = pppp.item(it).ParagraphFromOrder(1)
                                     w4 = x
                                       If z > 0 Then
                                    pppp.item(it).BackSpaceNchars w3, w4 + CLng(z), CLng(z)
                                    End If
                                    If w3 < 1 Then w3 = 1
                                    If Len(s$) > 0 Then pppp.item(it).InsertDoc w3, w4, s$
                                   
                                        Else
                                        w3 = pppp.item(it).ParagraphFromOrder(y + 1)
                                    w4 = x
                                       If z > 0 Then
                                    pppp.item(it).BackSpaceNchars w3, w4 + CLng(z), CLng(z)
                                    End If
                                    If w3 < 1 Then w3 = 1
                                    If Len(s$) > 0 Then pppp.item(it).InsertDoc w3, w4, s$
                                    
                                        End If
                                
                                
                                
                                Else
                                    MissStringExpr
                                    IdPara = False
                                    Exit Function
                                
                                End If

                            End If
                        Else
                             MissingDoc   ' only doc not string var
                             IdPara = False
                            Exit Function
                        End If
                    End If
            Else
                MissingDoc   ' only doc not string var
                IdPara = False
                Exit Function
            End If
        End If
 ElseIf IsExp(basestack, rest$, x) Then
    x = Int(x)
    If x < 1 Then
    MyErMacro rest$, "the index base must be >=1", "η βάση δείκτη πρέπει να είναι >=1"
    ' not needed to change idpara must be true because macro embed an ERROR command
    Exit Function
    End If
    If FastSymbol(rest$, ",") Then
        If Not IsExp(basestack, rest$, y) Then
        MissNumExpr
        IdPara = False
        Exit Function
        End If
        y = Int(y)
        If y < 0 Then
            MyErMacro rest$, "number to delete chars must positive or zero", "ο αριθμός για να διαγράψω πρέπει να είναι θετικός ή μηδέν"
            Exit Function
        End If
    Else
    y = 0  ' only insert
    End If

     x1 = Abs(IsLabel(basestack, rest$, what$))
        If x1 = 3 Then
            If GetVar(basestack, what$, i) Then
        
                If Typename(var(i)) = doc Then
                    If Not FastSymbol(rest$, "=") Then
                    MissSymbolMyEr "="
                    IdPara = False
                    Exit Function
                    Else
                            If Not IsStrExp(basestack, rest$, s$) Then
                                MissStringExpr
                                IdPara = False
                                Exit Function
                            Else
                                    If y = 0 Then
                                           var(i).FindPos 1, 0, CLng(x), x1, y1, w3, w4
                                           If w4 = 0 Then
                                          ' ' merge to previous
                                           End If
       
                                    Else
                                             var(i).FindPos 1, 0, x + y, x1, y1, w3, w4
                                            ' so now we now the paragraph w3 and the position w4
                                            var(i).BackSpaceNchars w3, w4, y
                                    End If
                                    If s$ <> "" Then var(i).InsertDoc w3, w4, s$
                            End If
                     End If
                ElseIf Typename(var(i)) = "Constant" Then
                CantAssignValue
                    IdPara = False
            Exit Function
                
                Else
                    If Not FastSymbol(rest$, "=") Then
                    MissSymbolMyEr "="
                    IdPara = False
                    Exit Function
                    Else
                    If Not IsStrExp(basestack, rest$, s$) Then
                                MissStringExpr
                                IdPara = False
                                Exit Function
                            Else
                                    If y = 0 Then
                                        var(i) = Left$(var(i), x - 1) & s$ & Mid$(var(i), x)
                                    Else
                                        If s$ = vbNullString Then
                                        var(i) = Left$(var(i), x - 1) & Mid$(var(i), x + y)
                                        Else
                                        If Len(s$) = y Then
                                        Mid$(var(i), x, y) = s$
                                        ElseIf Len(s$) < y Then
                                        Mid$(var(i), x, y) = s$ + space$(y - Len(s$))
                                        Else
                                        var(i) = Left$(var(i), x - 1) & s$ & Mid$(var(i), x + y)
                                        End If
                                        End If
                                    End If
                            End If
                    End If
                
                End If
            Else
            Nosuchvariable what$
            IdPara = False
            Exit Function
            
            End If
        ElseIf x1 = 6 Then
        
        
        If neoGetArray(basestack, what$, pppp) Then
                If Not NeoGetArrayItem(pppp, basestack, what$, it, rest$) Then IdPara = False: Exit Function
                If pppp.ItemType(it) = doc Then
                    If FastSymbol(rest$, "=") Then
                        If IsStrExp(basestack, rest$, s$) Then
                      If y = 0 Then
                                     pppp.item(it).FindPos 1, 0, CLng(x), xa, ya, w3, w4
                                           If w4 = 0 Then
                                          ' ' merge to previous
                                           End If

                      Else
                                     pppp.item(it).FindPos 1, 0, x + y, xa, ya, w3, w4
                                            ' so now we now the paragraph w3 and the position w4
                                            pppp.item(it).BackSpaceNchars w3, w4, y
                      End If
                       If s$ <> "" Then pppp.item(it).InsertDoc w3, w4, s$
                        Else
                            MissStringExpr
                            IdPara = False
                        End If
                    End If
                Else
                If FastSymbol(rest$, "=") Then
                If IsStrExp(basestack, rest$, s$) Then
                If y = 0 Then
                    pppp.item(it) = Left$(pppp.item(it), x - 1) & s$ & Mid$(var(i), x)
                Else
                                                        If s$ = vbNullString Then
                                        pppp.item(it) = Left$(pppp.item(it), x - 1) & Mid$(pppp.item(it), x + y)
                                        Else
                                      vvl = pppp.item(it)
                                       If Len(vvl) = y Then
                                      
                                        Mid$(vvl, x, y) = s$
                                        ElseIf Len(s$) < y Then
                                            Mid$(vvl, x, y) = s$ + space$(y - Len(s$))
                                        Else
                                        vvl = Left$(vvl, x - 1) & s$ & Mid$(vvl, x + y)
                                        End If
                                        pppp.item(it) = vvl
                                        End If
                End If
                Else
                     MissStringExpr
                            IdPara = False
                End If
                End If
                End If
        Else
            IdPara = True
        End If
        
        
        
        Else
        MissingStrVar
        IdPara = False
        ' wrong parameter
        End If


 
 
End If

End Function
Sub stackshow(b As basetask)
Static OldPagio$
Dim p As Variant, R$, al$, s$, dl$, dl2$
Static once As Boolean, ok As Boolean
If once Then Exit Sub
once = True


If TestShowCode Then
With Form2.testpad
.enabled = True
.SelectionColor = rgb(255, 64, 128)
.nowrap = True
.Text = TestShowSub
If Len(Form2.label1(1)) > 0 Then
If AscW(Form2.label1(1)) = 8191 Then
.SelStartSilent = TestShowStart - 1
.SelLength = Len(Mid$(Form2.label1(1), 7))
Else
.SelStartSilent = TestShowStart - Len(Form2.label1(1)) - 1
.SelLength = Len(Form2.label1(1))
End If


.enabled = False
If Len(Form2.label1(1)) > 0 Then
If .SelLength > 1 And Not AscW(Form2.label1(1)) = 8191 Then
If Not myUcase(.SelText, True) = Form2.label1(1) Then
End If
End If
End If
Else
.enabled = False
End If

End With

once = False
Exit Sub
Else
Form2.testpad.nowrap = False
End If

If pagio$ <> OldPagio$ Then
Form2.FillAgainLabels
OldPagio$ = pagio$
End If


Dim stack As mStiva
Set stack = b.soros

If Form2.Compute <> "" Then
If Form2.Compute.Prompt = "? " Then dl$ = Form2.Compute
With Form2.testpad
.enabled = True
.ResetSelColors
''
.nowrap = False
''
End With
Do
dl2 = dl$
ok = False
stackshowonly = True
If FastSymbol(dl$, ")") Then
ok = True
ElseIf IsExp(b, dl$, p) Then
    If al$ = vbNullString Then
        If pagio$ = "GREEK" Then
        al$ = "? " & Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & MyCStr(p)
        Else
        al$ = "? " & Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & MyCStr(p)
        End If
            
    Else
        al$ = al$ & "," & Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & MyCStr(p)
    End If
    ok = True
    ElseIf IsStrExp(b, dl$, s$) Then
    If Len(dl2$) - Len(dl$) >= 0 Then
    
    
    If al$ = vbNullString Then
        al$ = Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & Chr(34) + s$ & Chr(34)
    Else
        al$ = al$ + "," + Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & Chr(34) + s$ & Chr(34)
    End If
    ok = True
    End If
    ElseIf InStr(dl$, ",") > 0 Then
       If InStr(dl$, Chr(2)) > 0 Then
     R$ = GetStrUntil(Chr(2), dl$, False)
     s$ = "<"
If ISSTRINGA(dl$, R$) Then If pagio$ <> "GREEK" Then s$ = s$ & R$
If ISSTRINGA(dl$, R$) Then If pagio$ = "GREEK" Then s$ = s$ & R$
al$ = s$ & ">" & al$
ok = True
Else
al$ = al$ & " " & GetStrUntil(",", dl$)
    
     dl$ = vbNullString
  
End If
    
    ok = True
    ElseIf dl$ <> "" Then
      If InStr(dl$, Chr(2)) > 0 Then
     R$ = GetStrUntil(Chr(2), dl$, False)
     s$ = "<"
If ISSTRINGA(dl$, R$) Then If pagio$ <> "GREEK" Then s$ = s$ & R$
If ISSTRINGA(dl$, R$) Then If pagio$ = "GREEK" Then s$ = s$ & R$
al$ = s$ & ">" & al$
ok = True
Else
     al$ = al$ & " " & dl$
     dl$ = vbNullString
  
End If

    End If
    
DropLeft ",", dl$

Loop Until Not ok
End If
stackshowonly = False
If al$ <> "" Then al$ = al$ & vbCrLf
    If pagio$ = "GREEK" Then
    al$ = al$ & "Σωρός "
    Else
    al$ = al$ & "Stack "
    End If
If stack.Total = 0 Then
    If pagio$ = "GREEK" Then
    al$ = al$ & "Αδειος"
    Else
    al$ = al$ & "Empty"
    End If
Else
    If pagio$ = "GREEK" Then
    al$ = al$ & "Κορυφή "
    Else
    al$ = al$ & "Top "
    End If

End If
Dim i As Long

Do
i = i + 1
If stack.Total < i Or Len(al$) > 400 Then Exit Do

If stack.StackItemType(i) = "N" Or stack.StackItemType(i) = "L" Then
al$ = al$ & MyCStr(stack.StackItem(i)) & " "
ElseIf stack.StackItemType(i) = "S" Then
R$ = stack.StackItem(i)
    If Len(R$) > 78 Then
    al$ = al$ & Chr(34) + Left$(R$, 75) & "..." & Chr(34)
    Else
    al$ = al$ & Chr(34) + R$ & Chr(34)
    End If
 ElseIf stack.StackItemType(i) = ">" Then
   If pagio$ = "LATIN" Then
    al$ = al$ & "[Optional] "
    Else
    al$ = al$ & "[Προαιρετικό] "
    End If
ElseIf stack.StackItemType(i) = "*" Then

al$ = al$ & stack.StackItemTypeObjectType(i) & " "
Else  '??
al$ = al$ & stack.StackItemTypeObjectType(i) & " "
End If

Loop
With Form2
    .gList1.backcolor = &H3B3B3B
        .label1(2) = .label1(2)
    
        .testpad.enabled = True
        .testpad.Text = al$
        .testpad.SetRowColumn 1, 1
        .testpad.enabled = False
End With
once = False
End Sub

Function MyCStr(p) As String
    Select Case VarType(p)
    Case vbBoolean
        MyCStr = Format$(p, ";\T\r\u\e;\F\a\l\s\e")
    Case vbLong
        MyCStr = LTrim$(Str(p)) & "&"
    Case vbInteger
        MyCStr = LTrim$(Str(p)) & "%"
    Case vbDecimal
        MyCStr = LTrim$(Str(p)) & "@"
    Case vbSingle
        MyCStr = LTrim$(Str(p)) & "~"
    Case vbCurrency
        MyCStr = LTrim$(Str(p)) & "#"
    Case Else
        MyCStr = LTrim$(Str(p))
    End Select
End Function
Sub mylist(bstack As basetask, Optional tofile As Long = -1, Optional Lang As Long)
Dim Scr As Object, prive As Long
Set Scr = bstack.Owner
prive = GetCode(Scr)
Dim p As Variant, i As Long, s$, pn&, x As Double, y As Double, it As Long, F As Long, pa$
Dim x1 As Long, y1 As Long, frm$, par As Boolean, ohere$, ss$, w$, sX As Double, sY As Double, modname$
Dim pppp As mArray, hlp$, H&, all$, myobject As Object, usehandler As mHandler, usegroup As Group
Dim w1 As Long, w2 As Long, w3 As Long, dum As Boolean, virtualtop As Long
pn& = 0
virtualtop = varhash.count - 1
'GoTo ByPass  '********************************
For pn& = virtualtop To 0 Step -1
varhash.ReadVar pn&, s$, H&
If InStr(s$, ChrW(&H1FFF)) = 0 And InStr(s$, ChrW(&HFFBF)) = 0 Then Exit For
virtualtop = virtualtop - 1
Next pn&
ByPass:
pn& = 0
Dim a() As String
With players(prive)
Do While pn& < varhash.count
varhash.ReadVar pn&, s$, H&
If SecureNames Then
a() = Split(s$, "].")
If UBound(a()) = 1 Then
a(0) = GetName(a(0))
s$ = Join(a(), ".")
End If
End If
's$ = Replace(s$, ChrW(&HFFBF), "")
's$ = Replace(s$, ChrW(&H1FFF), "")
If H& = -1 Then
Else
'If InStr(s$, ChrW(&HFFBF)) > 0 And False Then   '*******************
If InStr(s$, ChrW(&HFFBF)) > 0 Then
GoTo LOOPNEXT
'ElseIf InStr(s$, ChrW(&H1FFF)) > 0 And False Then '******************
ElseIf InStr(s$, ChrW(&H1FFF)) > 0 Then '******************
' DO NOTHING
GoTo LOOPNEXT
ElseIf Right$(s$, 1) = "(" Then
    If MyIsObject(var(H&)) Then
        Set myobject = var(H&)
        If Not CheckIsmArray(myobject) Then
        If Not myobject Is Nothing Then
        If TypeOf myobject Is mArray Then
        If Typename(myobject.GroupRef) = "PropReference" Then
        s$ = s$ + ") [Object Property]"
        GoTo conthere
        End If
        End If
        End If
        GoTo LOOPNEXT
        End If
        Set pppp = myobject
            Set myobject = Nothing
       F = pppp.bDnum
       w1 = 0
        pppp.GetDnum w1, w2, w3
        w1 = w1 + 1

        If F > 1 Then
            If tofile < 0 Then
                If tofile = -1 Then
                            If .mx - .curpos < Len(s$ & Trim$(Str$(w2)) & ",") Then crNew bstack, players(prive)
                            PlainBaSket Scr, players(prive), s$ & Trim$(Str$(w2)) & ","
                Else
                            all = all + " " + s$ + Trim$(Str$(w2)) + ","
                            s$ = vbNullString
                End If
            Else
                If uni(tofile) Then
                    putUniString tofile, s$ & Trim$(Str$(w2)) & ","
                Else
                    putANSIString tofile, s$ & Trim$(Str$(w2)) & ","
                End If
            End If
        Else
        If pn& < virtualtop Then
        If tofile < 0 Then
            If tofile = -1 Then
            If .mx - .curpos < Len(s$ & Trim$(Str$(w2)) & "), ") Then crNew bstack, players(prive)
            PlainBaSket Scr, players(prive), s$ & Trim$(Str$(w2)) & "), "
            Else
            'prop
                all = all + " " + s$ + Trim$(Str$(w2)) + "),"
            End If
            Else
                   If uni(tofile) Then
        putUniString tofile, s$ & Trim$(Str$(w2)) & "), "
                Else
                putANSIString tofile, s$ & Trim$(Str$(w2)) & "), "
            
            End If
            End If
       Else
        If tofile < 0 Then
                If tofile = -1 Then
          If .mx - .curpos < Len(s$ & Trim$(Str$(w2)) & ")") Then crNew bstack, players(prive)
         PlainBaSket Scr, players(prive), s$ & Trim$(Str$(w2)) & ")"
         Else
           all = all + " " + s$ + Trim$(Str$(w2)) + ")"
         End If
         Else
                   If uni(tofile) Then
        putUniString tofile, s$ & Trim$(Str$(w2)) & ")"
                Else
putANSIString tofile, s$ & Trim$(Str$(w2)) & ")"

         End If
         End If
     End If
    End If
x = F - 1

While x > 0
x = x - 1
pppp.GetDnum w1, w2, w3
w1 = w1 + 1
If x > 0 Then
If tofile < 0 Then
    If tofile = -1 Then
    
    If .mx - .curpos < Len(Trim$(Str$(w2)) & ",") Then crNew bstack, players(prive)
    PlainBaSket Scr, players(prive), Trim$(Str$(w2)) & ","
    Else
    ' prop
        all = all + " " + s$ + Trim$(Str$(w2)) + ","
    End If
       Else
        If uni(tofile) Then
        putUniString tofile, Trim$(Str$(w2)) & ","
                Else
         putANSIString tofile, Trim$(Str$(w2)) & ","
        ' Print #tofile, Trim$(str$(w2)) & ",";
         End If
         End If
    
Else
        If pn& < virtualtop Then
         If tofile < 0 Then
            If tofile = -1 Then
            If .mx - .curpos < Len(Trim$(Str$(w2)) & "), ") Then crNew bstack, players(prive)
            PlainBaSket Scr, players(prive), Trim$(Str$(w2)) & "), "
            Else
            'prop
                all = all + " " + s$ + Trim$(Str$(w2)) + "),"
            End If
            Else
     If uni(tofile) Then
        putUniString tofile, Trim$(Str$(w2)) & "), "
                Else
                putANSIString tofile, Trim$(Str$(w2)) & "), "
            'Print #tofile, Trim$(str$(w2)) & "), ";
            End If
            End If
        Else
        If tofile < 0 Then
    If tofile = -1 Then
            If .mx - .curpos < Len(Trim$(Str$(w2)) & ")") Then crNew bstack, players(prive)
            PlainBaSket Scr, players(prive), Trim$(Str$(w2)) & ")"
            Else
            ' prop
                                all = all + " " + s$ + Trim$(Str$(w2)) + ")"
            End If
            Else
                               If uni(tofile) Then
        putUniString tofile, Trim$(Str$(w2)) & ")"
                Else
                putANSIString tofile, Trim$(Str$(w2)) & ")"
            'Print #tofile, scr, Trim$(str$(w2)) & ")";
            End If
            End If
        End If
    End If
Wend

    End If
    GoTo LOOPNEXT
 ElseIf Right$(s$, 1) = "$" Or Right$(s$, 3) = "$()" Then  ' WHY "$()"
    ' h& = Val(Mid$(VarName$, pn&))
        If Typename(var(H&)) = doc Then
            If var(H&).IsEmpty Then
                hlp$ = " [Empty Document]"
            Else
                hlp$ = " [Document " + CStr(var(H&).SizeCRLF) & " chars]"
            End If
        ElseIf Typename(var(H&)) = "PropReference" Then
        hlp$ = " [Object Property]"
        Else
        If MyIsObject(var(H&)) Then
        If TypeOf var(H&) Is lambda Then
            hlp$ = "[lambda$]"
            Else
            
            If TypeOf var(H&) Is Constant Then
            If var(H&).flag Then
            hlp$ = "[" + Typename(var(H&).Value) + "$]"
            Else
                 If Len(CStr(var(H&))) > 3 * .mx Then
                    hlp$ = " = [" + Left$(CStr(var(H&)), 4) & "...]"
                Else
                    hlp$ = " = [" + CStr(var(H&)) + "]"
                End If
            End If
            Else
            hlp$ = "[" + Typename(var(H&)) + "]"
            End If
            End If
        Else
        
            If Len(var(H&)) > 3 * .mx Then
                hlp$ = " = " & Chr(34) + Left$(CStr(var(H&)), 4) & "..." & Chr(34)
            Else
                hlp$ = " = " & Chr(34) + CStr(var(H&)) + Chr(34)
            End If
            
        End If
        End If
    s$ = s$ & hlp$
    
Else

If MyIsObject(var(H&)) Then
If var(H&) Is Nothing Then
s$ = s$ + "*[Nothing]"
ElseIf TypeOf var(H&) Is mHandler Then
Set usehandler = var(H&)
Select Case usehandler.t1
Case 1
If usehandler.ReadOnly Then
    If usehandler.objref.StructLen > 0 Then
    s$ = s$ + "*[Structure]"
    Else
    s$ = s$ + "*[Inventory/ReadOnly]"
    End If
Else
    s$ = s$ + "*[Inventory]"
    End If
Case 2
    s$ = s$ + "*[Buffer]"
Case 4
    s$ = s$ + "*[" + usehandler.objref.EnumName + "]"
Case Else
If Not usehandler.objref Is Nothing Then
    If TypeOf usehandler.objref Is mHandler Then
        If usehandler.objref.t1 = 4 Then
        s$ = s$ + "*[" + usehandler.objref.objref.EnumName + "]"
        Else
        s$ = s$ + "*[" + Typename(usehandler.objref) + "]"
        End If
    Else
    s$ = s$ + "*[" + Typename(usehandler.objref) + "]"
    End If
ElseIf usehandler.indirect > 0 Then
    s$ = s$ + "*[" + Typename(var(usehandler.indirect)) + "]"
End If
End Select
Set usehandler = Nothing
Else
    If TypeOf var(H&) Is Constant Then
        On Error Resume Next
        s$ = s$ & " = [" & LTrim$(Str(var(H&))) + "]"
        If Err Then
            s$ = s$ & " = [" & var(H&) + "]"
            Err.Clear
        End If
    Else
        If TypeOf var(H&) Is Group Then
            Set usegroup = var(H&)
            If usegroup.IamApointer Then
                s$ = s$ & "*[Group]"
            Else
                s$ = s$ & "[Group]"
            End If
            Set usegroup = Nothing
        ElseIf TypeOf var(H&) Is PropReference Then
            s$ = s$ & " [Object Property]"
        Else
            s$ = s$ & "[" & Typename(var(H&)) & "]"
        End If
    End If
End If
Else
On Error Resume Next
s$ = s$ & " = " & LTrim$(Str(var(H&)))
If Err Then
s$ = s$ & " = " & Chr(34) & var(H&) & Chr(34)
Err.Clear
End If
Select Case VarType(var(H&))
        Case vbLong
        s$ = s$ & "&"
        Case vbDecimal
        s$ = s$ & "@"
        Case vbSingle
        s$ = s$ & "~"
        Case vbCurrency
        s$ = s$ & "#"
        Case vbInteger
        s$ = s$ & "%"
        End Select
End If
End If
conthere:
If pn& < virtualtop Then s$ = s$ & ", "
If tofile < 0 Then
   If tofile = -1 Then
   If .mx - .curpos < Len(s$) Then crNew bstack, players(prive)
    PlainBaSket Scr, players(prive), s$
    End If
    ' proportional
    all = all + " " + s$
   Else
   If uni(tofile) Then
putUniString tofile, s$
   Else
    putANSIString tofile, s$

    End If
    End If
 End If
LOOPNEXT:
pn& = pn& + 1
Loop
s$ = vbNullString
If Not bstack.StaticCollection Is Nothing Then
Dim st1 As Long, mList As FastCollection
Set mList = bstack.StaticCollection
For st1 = 1 To mList.count
mList.Index = st1 - 1
If Left$(mList.KeyToString, 2) <> "%_" Then
    If s$ <> "" Then s$ = s$ + ", "
    If mList.IsObj Then
        If TypeOf mList.ValueObj Is mHandler Then
            Set usehandler = mList.ValueObj
            Select Case usehandler.t1
            Case 1
                If usehandler.ReadOnly Then
                    If usehandler.objref.StructLen > 0 Then
                    s$ = s$ + mList.KeyToString + "*[Structure]"
                    Else
                    s$ = s$ + mList.KeyToString + "*[Inventory/ReadOnly]"
                    End If
                Else
                    s$ = s$ + mList.KeyToString + "*[Inventory]"
                End If
            Case 2
                s$ = s$ + mList.KeyToString + "*[Buffer]"
            Case 3
            
            
                s$ = s$ + mList.KeyToString + "*[" + Typename(usehandler.objref) + "]"
            End Select
        Else
            s$ = s$ + mList.KeyToString + " [" + Typename(mList.ValueObj) + "]"
        End If
    Else
If IsNumeric(mList.Value) Then
s$ = s$ + mList.KeyToString + " = " + LTrim$(Str$(mList.Value))
Else
If Len(var(H&)) > 3 * .mx Then
 s$ = s$ + mList.KeyToString + " = " & Chr(34) + Left$(mList.Value, 4) & "..." & Chr(34)
 Else
s$ = s$ + mList.KeyToString + " = " & Chr(34) + mList.Value + Chr(34)
End If
End If
End If
End If
Next st1
If s$ <> "" Then
If Lang = 1 Then

s$ = " Static Variables: " + s$
Else
s$ = " Στατικές Μεταβλητές: " + s$
End If
If tofile >= 0 Then
 putUniString tofile, vbCrLf
If uni(tofile) Then
        putUniString tofile, s$
                Else
                putANSIString tofile, s$

    End If
End If
End If
End If
    If tofile < -1 Then
        If Scr.currentX <> 0 Then crNew bstack, players(prive)
        If s$ <> "" Then s$ = vbCrLf + s$
        wwPlain2 bstack, players(prive), all$ + s$, Scr.Width, 1000, True, , 3
    ElseIf tofile = -1 Then
        If s$ <> "" Then
         crNew bstack, players(prive)
        
        PlainBaSket Scr, players(prive), s$
        End If
    End If
    
    If tofile < 0 Then crNew bstack, players(prive)
      End With
End Sub
Function PathFromApp(ByVal nap$) As String
Dim ap$
nap$ = nap$ & " "
ap$ = GetStrUntil(" ", nap$)
If ExtractType(ap$) = vbNullString Then ap$ = ap$ & ".exe"
Dim cc As New cRegistry
cc.ClassKey = HKEY_CURRENT_USER
cc.SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & ap$
cc.ValueKey = vbNullString
cc.ValueType = REG_SZ
On Error GoTo 1111
If IsEmpty(cc.Value) Then
    cc.ClassKey = HKEY_LOCAL_MACHINE
    If IsEmpty(cc.Value) Then
        PathFromApp = vbNullString
    Else
        PathFromApp = Trim$(mylcasefILE(cc.Value & " " & nap$))
    End If
Else
    PathFromApp = Trim$(mylcasefILE(cc.Value & " " & nap$))
End If
Exit Function
1111:
PathFromApp = vbNullString
End Function


Public Function PCall(ByVal sFile As String, Optional param As String) As String
Dim s2 As String, i As Long, bsfile As String, rfile As String, MYNULL$
bsfile = mylcasefILE(sFile)
   s2 = String(MAX_FILENAME_LEN, 32)
   'Retrieve the name and handle of the executable, associated with this file
   i = FindExecutable(StrPtr(sFile), StrPtr(MYNULL$), StrPtr(s2))
   If i > 32 Then
   rfile = mylcasefILE(Left$(s2, InStr(s2, Chr$(0)) - 1))
   If ExtractName(bsfile, True) = ExtractName(rfile, True) Then
   ' it is an executable
   PCall = mylcasefILE(bsfile)
   Else
   If param <> "" Then
   PCall = rfile & " " & param & " " & Chr(34) + bsfile + Chr(34)
   Else
      PCall = rfile & " " & Chr(34) + bsfile + Chr(34)
      End If
      End If
      Else
      PCall = vbNullString
End If
End Function

Sub makegroup(bstack As basetask, what$, i As Long)
Dim it As Long
it = globalvar(what$, it)
    MakeitObject2 var(it)
    If var(i).IamApointer Then
        If var(i).link.IamFloatGroup Then
           Set var(it).LinkRef = var(i).link
            var(it).IamApointer = True
            var(it).isref = True
        Else
            With var(i).link
            
                var(it).edittag = .edittag
                var(it).FuncList = .FuncList
                var(it).GroupName = myUcase(what$) + "."
                Set var(it).Sorosref = .soros.Copy
                var(it).HasValue = .HasValue
                var(it).HasSet = .HasSet
                var(it).HasStrValue = .HasStrValue
                var(it).HasParameters = .HasParameters
                var(it).HasParametersSet = .HasParametersSet
            
                        Set var(it).Events = .Events
            
                var(it).highpriorityoper = .highpriorityoper
                var(it).HasUnary = .HasUnary
            End With
        End If
    
    Else
        With var(i)
            var(it).edittag = .edittag
            var(it).FuncList = .FuncList
            var(it).GroupName = myUcase(what$) + "."
            Set var(it).Sorosref = .soros.Copy
            var(it).HasValue = .HasValue
            var(it).HasSet = .HasSet
            var(it).HasStrValue = .HasStrValue
            var(it).HasParameters = .HasParameters
            var(it).HasParametersSet = .HasParametersSet
            Set var(it).Events = .Events
            var(it).highpriorityoper = .highpriorityoper
            var(it).HasUnary = .HasUnary
        End With
        var(it).IamRef = Len(bstack.UseGroupname) > 0
    End If
    If var(i).HasStrValue Then
        globalvar what$ + "$", it, True
    End If
            
        
End Sub
Function ExecCode(basestack As basetask, rest$) As Boolean ' experimental
' ver .001
Dim p As Variant, mm As MemBlock, w2 As Long, usehandler As mHandler
    If IsExp(basestack, rest$, p) Then
        If Not basestack.lastobj Is Nothing Then
          If Not TypeOf basestack.lastobj Is mHandler Then
            Set basestack.lastobj = Nothing
            Exit Function
            End If
            Set usehandler = basestack.lastobj
            With usehandler
                  If Not TypeOf .objref Is MemBlock Then
                      Set basestack.lastobj = Nothing
                      Exit Function
                  ElseIf .objref.NoRun Then
                       Set basestack.lastobj = Nothing
                       Exit Function
                  End If
            End With
            Set mm = usehandler.objref
            Set usehandler = Nothing
            If mm.status = 0 Then
            w2 = mm.GetPtr(0)
            If FastSymbol(rest$, ",") Then
            If Not IsExp(basestack, rest$, p) Then
                Set basestack.lastobj = Nothing
                Set mm = Nothing
                MissPar
                Exit Function
            End If
            If p < 0 Or p >= mm.SizeByte Then
                Set basestack.lastobj = Nothing
                Set mm = Nothing
                MyEr "Offset out of buffer", "Διεύθυνση εκτός διάρθρωσης"
                Exit Function
            End If

            SetUpForExecution w2, mm.SizeByte
            w2 = cUlng(uintnew(w2) + p)
            End If
            Set basestack.lastobj = Nothing
            Dim what As Long
            what = CallWindowProc(w2, 0&, 0&, 0&, 0&)
            If what <> 0 Then MyEr "Error " & what, "Λάθος " & what
                ReleaseExecution w2, mm.SizeByte
                ExecCode = what = 0
                Set mm = Nothing
            End If
        End If
    End If
    Set basestack.lastobj = Nothing
End Function
Sub MyMode(Scr As Object)
Dim x1 As Long, y1 As Long
On Error Resume Next
With players(GetCode(Scr))
    x1 = Scr.Width
    y1 = Scr.Height
    If Left$(Typename(Scr), 3) = "Gui" Then
    Else
    If Scr.Name = "Form1" Then
    DisableTargets q(), -1
    
    ElseIf Scr.Name = "DIS" Then
    DisableTargets q(), 0
    
    ElseIf Scr.Name = "dSprite" Then
        DisableTargets q(), val(Scr.Index)
    ElseIf TypeOf Scr Is GuiM2000 Then
        Scr.DisAllTargets
    End If
    End If
    If .SZ < 4 Then .SZ = 4
        Err.Clear
        Scr.Font.Size = .SZ
        If Err.Number > 0 Then
                MyFont = "ARIAL"
                Scr.Font.Name = MyFont
                Scr.Font.charset = .charset
                Scr.Font.Name = MyFont
                Scr.Font.charset = .charset
        End If
        .uMineLineSpace = .MineLineSpace
        FrameText Scr, .SZ, x1, y1, .Paper
    .currow = 0
    .curpos = 0
    .XGRAPH = 0
    .YGRAPH = 0
End With
End Sub
Function ProcSave(basestack As basetask, rest$, Lang As Long) As Boolean
Dim pa$, w$, s$, Col As Long, prg$, x1 As Long, par As Boolean, i As Long, noUse As Long, lcl As Boolean
Dim askme As Boolean, k As Long, m As Long
If lckfrm <> 0 Then MyEr "Save is locked", "Η αποθήκευση είναι κλειδωμένη": rest$ = vbNullString: Exit Function
lcl = IsLabelSymbolNew(rest$, "ΤΟΠΙΚΑ", "LOCAL", Lang) Or basestack.IamChild Or basestack.IamAnEvent
x1 = Abs(IsLabelFileName(basestack, rest, pa$, , s$))

If x1 <> 1 Then
rest$ = pa$ + rest$: x1 = IsStrExp(basestack, rest$, pa$)
Else
pa$ = s$: s$ = vbNullString
End If

If x1 <> 0 Then
        If subHash.count = 0 Or pa$ = vbNullString Then MyEr "Nothing to save", "Δεν υπάρχει κάτι να σώσω":              Exit Function
        If ExtractType(pa$) = "gsb1" Then
            MyEr "Recovery Mode (file is a gsb1 type): use another name or set type to gsb", "Κατάσταση Ανάκτησης (o τύπος αρχείου είναι gsb1): χρησιμοποίησε άλλο όνομα χωρίς τύπο, ή τύπο gsb"
            Exit Function
        End If

        If ExtractType(pa$) = "gsb" Then pa$ = ExtractPath(pa$) + ExtractNameOnly(pa$, True)
        If ExtractPath(pa$, True) <> "" Then
                If InStr(ExtractPath(pa$, True), mcd) <> 1 Then pa$ = pa$ & ".gsb" Else pa$ = pa$ & ".gsb"
        Else
                pa$ = mcd + pa$ & ".gsb"
        End If
        If Not WeCanWrite(pa$) Then Exit Function
        
      
           For i = subHash.count - 1 To 0 Step -1
       subHash.ReadVar i, s$, Col
       If Not Len(sbf(Col).sbgroup) > 0 Then
       
                If Right$(s$, 2) = "()" Then
                If Not InStr(s$, ChrW(&H1FFF)) > 0 Then
                s$ = Left$(s$, Len(s$) - 2)
                
                If Right$(sbf(Col).sb, 2) <> vbCrLf Then sbf(Col).sb = sbf(Col).sb + vbCrLf
                If Lang Then
                
                                If Not blockCheck(sbf(Col).sb, DialogLang, noUse, "Function " & s$ + "()" + vbCrLf) Then Exit Function
                                If sbf(Col).IamAClass Then

                                k = InStr(sbf(Col).sb, vbCrLf)
                                m = rinstr(sbf(Col).sb, vbCrLf + vbCrLf)
                                k = InStr(k + 1, sbf(Col).sb, "{")
                                prg$ = "CLASS " + Left$(sbf(Col).goodname, Len(sbf(Col).goodname) - 2) + " {" + Mid$(sbf(Col).sb, k + 3, m - k - 3) + vbCrLf + prg$
                               
                               
                                Else
                                prg$ = s$ & " {" & sbf(Col).sb & "}" & vbCrLf + prg$
                                If lcl Then
                                    prg$ = "FUNCTION " + prg$
                                Else
                                    prg$ = "FUNCTION GLOBAL " + prg$
                                End If
                                End If
                        Else
                                If Not blockCheck(sbf(Col).sb, DialogLang, noUse, "Συνάρτηση " & s$ + "()" + vbCrLf) Then Exit Function
                                If sbf(Col).IamAClass Then
                                k = InStr(sbf(Col).sb, vbCrLf)
                                m = rinstr(sbf(Col).sb, vbCrLf + vbCrLf)
                                k = InStr(k + 1, sbf(Col).sb, "{")
                                prg$ = "ΚΛΑΣΗ " + Left$(sbf(Col).goodname, Len(sbf(Col).goodname) - 2) + " {" + Mid$(sbf(Col).sb, k + 3, m - k - 3) + vbCrLf + prg$
                                Else
                                prg$ = s$ & " {" & sbf(Col).sb & "}" & vbCrLf + prg$
                                If lcl Then
                                    prg$ = "ΣΥΝΑΡΤΗΣΗ " + prg$
                                Else
                                    prg$ = "ΣΥΝΑΡΤΗΣΗ ΓΕΝΙΚΗ " + prg$
                                End If
                                End If
                        End If
                End If
                Else
                        If Right$(sbf(Col).sb, 2) <> vbCrLf Then sbf(Col).sb = sbf(Col).sb + vbCrLf
                        If Lang Then
                                If Not blockCheck(sbf(Col).sb, DialogLang, noUse, "Module " & s$ + vbCrLf) Then Exit Function
                                prg$ = s$ & " {" & sbf(Col).sb & "}" & vbCrLf + prg$
                                If lcl Then
                                    prg$ = "MODULE " + prg$
                                Else
                                    prg$ = "MODULE GLOBAL " + prg$
                                End If
                        Else
                                If Not blockCheck(sbf(Col).sb, DialogLang, noUse, "Τμήμα " & s$ + vbCrLf) Then Exit Function
                                prg$ = s$ & " {" & sbf(Col).sb & "}" & vbCrLf + prg$
                                If lcl Then
                                    prg$ = "ΤΜΗΜΑ " + prg$
                                Else
                                    prg$ = "ΤΜΗΜΑ ΓΕΝΙΚΟ " + prg$
                                End If
                        End If
                End If
        End If
        Next i
        w$ = vbNullString
        If FastSymbol(rest$, "@@", , 2) Then
            ' default password  - one space only - coder use default internal password
                If Not IsStrExp(basestack, rest$, w$) Then w$ = " "
        ElseIf FastSymbol(rest$, "@") Then
                ' One space only
                w$ = " "
        End If
        par = False
        If FastSymbol(rest$, ",") Then
                If Abs(IsLabel(basestack, rest$, s$)) = 1 Then
                        prg$ = prg$ & s$
                ElseIf FastSymbol(rest$, "{") Then
                        prg$ = prg$ & block(rest$)
                        If Not FastSymbol(rest$, "}") Then Exit Function
                End If
        End If
        ' reuse s$, col$
        If Len(w$) > 1 Then  'scrable col by George
                s$ = vbNullString: For Col = 1 To Int((33 * Rnd) + 1): s$ = s$ & Chr(65 + Int((23 * Rnd) + 1)): Next Col
                ' insert a variable length label......to make a variable length file
                prg$ = s$ & ":" & vbCrLf + prg$
                prg$ = mycoder.encryptline(prg$, w$, Len(prg$) Mod 33)
                par = True
        ElseIf Len(w$) = 1 Then   ' I have to check that...
                s$ = vbNullString:   For Col = 1 To Int((33 * Rnd) + 1): s$ = s$ & Chr(65 + Int((23 * Rnd) + 1)): Next Col
                prg$ = s$ & ":" & vbCrLf + prg$
                prg$ = mycoder.must1(prg$)
                par = True
        End If
        s$ = vbNullString

        If Not WeCanWrite(pa$) Then Exit Function
       ' If CFname(pa$) <> "" Then
        If par Then
        If Not SaveUnicode(pa$ + "1", prg$, 0) Then BadFilename: Exit Function
        Else
        If Not SaveUnicode(pa$ + "1", prg$, 2) Then BadFilename: Exit Function
        End If
        ProcTask2 Basestack1
        If CFname(ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".gsb") <> "" Then
            If CFname(ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck1") <> "" Then
                KillFile ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck1"
            End If
            RenameFile2 pa$, ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck1"
            askme = True
        End If
     ProcTask2 Basestack1
     
        If askme Then
                
                If Lang = 1 Then
                        If MsgBoxN("Replace " + ExtractNameOnly(pa$, True), vbOKCancel, MesTitle$) <> vbOK Then
nogood:
                        If CFname(pa$ + "1") <> "" Then
                            KillFile pa$ + "1"
                        End If
                        RenameFile2 ExtractPath(pa$) & ExtractNameOnly(pa$, True) + ".bck1", ExtractPath(pa$) & ExtractNameOnly(pa$, True) + ".gsb"
                        MyEr "File not saved -1005", "Δεν σώθηκε το αρχείο -1005"
                        ProcSave = True
                        Exit Function
                        Else
                        Check2SaveModules = False
                        End If
                Else
                        If MsgBoxN("Αλλαγή " + ExtractNameOnly(pa$, True), vbOKCancel, MesTitle$) <> vbOK Then
                        GoTo nogood
                        Else
                        Check2SaveModules = False
                        End If
                End If
                s$ = "*"
        End If
        If par Then
                If s$ = "*" Then
                        MakeACopy ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck1", ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck"
                        KillFile ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck1"
                End If
                RenameFile2 pa$ + "1", pa$
                Check2SaveModules = False
        Else
                If s$ <> "" Then
                        MakeACopy ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck1", ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck"
                        KillFile ExtractPath(pa$) & ExtractNameOnly(pa$, True) & ".bck1"
                End If
                RenameFile2 pa$ + "1", pa$
                If here$ = vbNullString Then LASTPROG$ = pa$
                Check2SaveModules = False
        End If
 ProcSave = True
Else
MyEr "A name please or use Ctrl+A to perform SAVE COMMAND$  (the last loading)", "Ένα όνομα παρακαλώ, ή πάτα το ctrl+Α για να αποθηκεύσεις με το όνομα του προγράμματος που φορτώθηκε τελευταία"
End If

End Function
Sub ProcChooseFont(bstack As basetask, Lang As Long)
If Form4Loaded Then
If Form4.Visible Then
Form4.Visible = False
    If Form1.TEXT1.Visible Then
        Form1.TEXT1.SetFocus
    Else
        Form1.SetFocus
    End If
End If
End If
DialogSetupLang Lang
With bstack.Owner
    ReturnFontName = .Font.Name
    ReturnBold = .Font.bold
    ReturnItalic = .Font.Italic
    ReturnSize = CSng(.Font.Size)
    ReturnCharset = .Font.charset
End With
FeedFont2Stack bstack, OpenFont(bstack, Form1)

End Sub


Function ProcMode(bstack As basetask, rest$) As Boolean
Dim Scr As Object, p As Variant, x1 As Long, y1 As Long
Dim prive As Long
ProcMode = True
'' Kform is global
Kform = True
On Error Resume Next
Set Scr = bstack.Owner
With players(GetCode(Scr))
If .double Then SetNormal Scr

x1 = Scr.Width
y1 = Scr.Height
If Scr.Name = "GuiM2000" Then
    Else
If Scr.Name = "Form1" Then
DisableTargets q(), -1

ElseIf Scr.Name = "DIS" Then
DisableTargets q(), 0

ElseIf Scr.Name = "dSprite" Then
DisableTargets q(), val(Scr.Index)
End If
End If
If IsExp(bstack, rest$, p) Then
If bstack.toprinter Then
If prFactor = 0 Then
prFactor = 1
End If
.SZ = CSng(p) * szFactor
Else
.SZ = CSng(p)
End If
If .SZ < 4 Then .SZ = 4
If Not bstack.toprinter Then
If FastSymbol(rest$, ",") Then
    If IsExp(bstack, rest$, p) Then x1 = CLng(p): y1 = CLng(x1 * ScrInfo(Console).Height / ScrInfo(Console).Width)
    If FastSymbol(rest$, ",") Then
            If IsExp(bstack, rest$, p) Then y1 = CLng(p)
        
    End If
ElseIf FastSymbol(rest$, ";") Then
prive = GetCode(bstack.Owner)
.mysplit = 0
Scr.Font.Size = .SZ
       SetText Scr
        GetXYb Scr, players(prive), .curpos, .currow

Set Scr = Nothing
Exit Function
End If
Else
'.SZ = .SZ * 3
End If
Err.Clear
Scr.Font.Size = .SZ
If Err.Number > 0 Then

MyFont = "ARIAL"
Scr.Font.Name = MyFont
Scr.Font.charset = bstack.myCharSet
Scr.Font.Name = MyFont
Scr.Font.charset = bstack.myCharSet
End If
.SZ = Scr.Font.Size
     .uMineLineSpace = .MineLineSpace
    
 FrameText Scr, .SZ, x1, y1, .Paper
 
    Else
    ProcMode = False
    Exit Function
    End If
    .currow = 0
    .curpos = 0
    .XGRAPH = 0
    .YGRAPH = 0
End With
Set Scr = Nothing

End Function



Function Infinity() As Double
PutMem1 VarPtr(Infinity) + 7, &H7F
PutMem1 VarPtr(Infinity) + 6, &HF0
End Function
Function ProcAbout(basestack As basetask, rest$, Lang As Long) As Boolean

Dim par As Boolean, s$, ss$, x As Double, y As Double, i As Long
Dim kk As New Document
Dim UAddPixelsTop As Long  ' just not used
If IsLabelSymbolNewExp(rest$, "ΔΕΙΞΕ", "SHOW", Lang, ss$) Then
If lastAboutHTitle <> "" Then abt = True: vH_title$ = vbNullString
If IsStrExp(basestack, rest$, ss$) Then
    feedback$ = ss$
feednow$ = FeedbackExec$
CallGlobal feednow$
Else
    vHelp
 End If
ProcAbout = True
Exit Function
ElseIf FastSymbol(rest$, "!") Then
'*******
vH_title$ = vbNullString
par = True
par = par And IsStrExp(basestack, rest$, s$)

If par Then
If s$ = vbNullString Then
mHelp = False
abt = False
lastAboutHTitle = vbNullString
sHelp "", "", 0, 0
GoTo conthere
Else
par = par And FastSymbol(rest$, ",")
par = par And IsExp(basestack, rest$, x)
par = par And FastSymbol(rest$, ",")
par = par And IsExp(basestack, rest$, y)
par = par And FastSymbol(rest$, ",")
par = par And IsStrExp(basestack, rest$, ss$)
If par Then
abt = True
kk.EmptyDoc
kk.textDoc = ReplaceSpace(s$)
s$ = kk.textFormat(vbCrLf)
kk.EmptyDoc
kk.textDoc = ReplaceSpace(s$)
s$ = kk.TextParagraph(1)
kk.EmptyDoc
kk.textDoc = ReplaceSpace(ss$)
' save to
lastAboutHTitle = s$
LastAboutText = kk.textFormat(vbCrLf)
sHelp lastAboutHTitle, LastAboutText, CLng(x), CLng(y)
End If
End If
End If
'*******
ElseIf IsLabelSymbolNew(rest$, "ΚΑΛΕΣΕ", "CALL", Lang) Then
mHelp = True
abt = True
If IsStrExp(basestack, rest$, ss$) Then

kk.textDoc = ReplaceSpace(ss$)
FeedbackExec$ = kk.textFormat(vbCrLf)
End If
Else
If IsStrExp(basestack, rest$, s$) Then
mHelp = True
If s$ = vbNullString Then
If Lang = 0 Then
lastAboutHTitle = "Βοήθεια Εφαρμογής"
LastAboutText = vbNullString
Else
lastAboutHTitle = "Application Help"
LastAboutText = vbNullString
End If
GoTo conthere
End If
kk.EmptyDoc
kk.textDoc = ReplaceSpace(s$)
s$ = kk.textFormat(vbCrLf)
kk.EmptyDoc
kk.textDoc = ReplaceSpace(s$)
s$ = kk.TextParagraph(1)

        i = 0
       
        x = (ScrInfo(Console).Width - 1) * 2 / 5
        y = (ScrInfo(Console).Height - 1) / 7
        vH_title$ = s$
        If FastSymbol(rest$, ",") Then
                par = True
                    If Not IsExp(basestack, rest$, x) Then
                    x = (ScrInfo(Console).Width - 1) * 2 / 5: par = False
                    Else
                    i = 1
                    End If
                    If FastSymbol(rest$, ",") Then
                        par = True
                        If Not IsExp(basestack, rest$, y) Then
                        y = (ScrInfo(Console).Height - 1) / 7: par = False
                        Else
                        i = 2
                        End If
                    End If
        End If

        If Not Form4.Visible Then
        Helplastfactor = 1
       helpSizeDialog = 1
           vH_x = CLng(x * Helplastfactor)
           vH_y = CLng(y * Helplastfactor)
           
            If Screen.ActiveForm Is Nothing Then
                        Form4.Show
            Else
                If Screen.ActiveForm.Name = "MyPopUp" Then
                    Form4.Show , Form1
                Else
                    Form4.Show , Screen.ActiveForm
                End If
       
          End If
               
           
                myform Form4, ScrInfo(Console).Width - CLng(x * Helplastfactor), ScrInfo(Console).Height - CLng(y * Helplastfactor), CLng(x * Helplastfactor), CLng(y * Helplastfactor), True, 1  'Helplastfactor
                  MoveFormToOtherMonitorOnly Form4, True
                HelpLastWidth = x
                GoTo there:
        ElseIf i <> 0 Then
            '    Form4.Show , Form1
            If Screen.ActiveForm Is Nothing Then
                        Form4.Show
                        
            ElseIf Form4 Is Screen.ActiveForm Then
            Form4.Show
            Else
                            Form4.Show , Screen.ActiveForm
       
            End If
                myform Form4, Form4.Left, Form4.top, CLng(x * Helplastfactor), CLng(y * Helplastfactor), True, Helplastfactor
                MoveFormToOtherMonitorOnly Form4
            GoTo there:
        End If
        Form4.Hide
            If Screen.ActiveForm Is Nothing Then
                        Form4.Show
            ElseIf Form4 Is Screen.ActiveForm Then
            Form4.Show
            Else
                            Form4.Show , Screen.ActiveForm
       
            End If
there:
        Form4.Line (0, 0)-(Form4.Scalewidth - dv15, Form4.Scaleheight - dv15), Form4.backcolor, BF
        Form4.moveMe
        If FastSymbol(rest$, ",") Or Not par Then
        If IsStrExp(basestack, rest$, ss$) Then
        kk.EmptyDoc
        kk.textDoc = ReplaceSpace(ss$)
        Form4.label1.Text = kk.textFormat(vbCrLf)
        End If
        End If
        
With Form4.label1

.EditDoc = False
.NoMark = True
.enabled = True
.NewTitle s$, 4 + UAddPixelsTop
.glistN.DragEnabled = False
.glistN.WordCharLeft = "["
.glistN.WordCharRight = "]"
.glistN.WordCharRightButIncluded = vbNullString
End With
Else
conthere:
'vH_title$ = vbNullString
If Not (basestack.IamChild Or basestack.IamAnEvent) Then
abt = False
End If
If Form4Loaded Then
If Form4.Visible Then
Form4.Visible = False
If Form1.Visible Then
    If Form1.TEXT1.Visible Then
        Form1.TEXT1.SetFocus
    Else
        Form1.SetFocus
    End If
    End If
End If
Helplastfactor = 1
helpSizeDialog = 1
Unload Form4
End If
End If
End If
Exit Function
End Function
Function ProcRemove(basestack As basetask, rest$, Lang As Long) As Boolean
Dim ss$
ProcRemove = True
If IsLabelSymbolNew(rest$, "ΑΔΕΙΑΣ", "LICENSE", Lang) Then
If IsStrExp(basestack, rest$, ss$) Then
Licenses.Remove ss$
Else
MissStringExpr
End If
Else
If IsStrExp(basestack, rest$, ss$) Then
RemoveDll ss$
ElseIf sb2used <= basestack.OriginalCode And basestack.OriginalCode <> 0 Then
MyEr "Can't Remove Last Module/Function", "Δεν μπορώ να διαγράψω το τελευταίο τμήμα/συνάρτηση"
ProcRemove = False
Else
If basestack.IamChild Or basestack.IamAnEvent Or basestack.IamThread Or HaltLevel > 0 Then Exit Function
If sb2used > 0 Then
If MsgBoxN(IIf(pagio$ <> "GREEK", "Remove last module/function", "Να διαγράψω το τελευταίο τμήμα") & vbCrLf & sbf(subHash.count).goodname, 1) = 1 Then


If subHash.count > 0 Then subHash.ReduceHash subHash.count - 1, sbf()
    If UBound(sbf()) / 2 > sb2used And UBound(sbf()) > 19 Then
       ReDim Preserve sbf(UBound(sbf()) / 2 + 1) As modfun
         
    End If
    sb2used = sb2used - 1

If lckfrm <> 0 Then
If lckfrm > sb2used Then lckfrm = 0

End If
End If
End If

End If
End If
End Function
Function BreakMes() As String
If Check2SaveModules Then
If pagio$ = "GREEK" Then
BreakMes = "Διακοπή και Επανεκκίνηση" & vbCrLf & "Προσοχή υπάρχουν μη αποθηκευμένες αλλαγές που θα χαθούν"
Else
BreakMes = "Break Key - Hard Reset" & vbCrLf & "Warning, changes lost"
End If
Else
If pagio$ = "GREEK" Then
BreakMes = "Διακοπή και Επανεκκίνηση"
Else
BreakMes = "Break Key - Hard Reset"
End If
End If
End Function
Function Check2Save() As Boolean
If Check2SaveModules Then
Check2Save = MsgBoxN(IIf(pagio$ = "GREEK", "Προσοχή υπάρχουν μη αποθηκευμένες αλλαγές που θα χαθούν", "Warning, changes lost"), 1, IIf(pagio$ = "GREEK", "Τερματισμός Διερμηνευτή", "Quit Interpreter")) <> 1
Else
Check2Save = False
End If

End Function
Function IsCtime(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim par As Boolean, r2 As Variant, r3 As Variant, r4 As Variant
If IsExp(bstack, a$, R, , True) Then
  
    R = Abs(R)
    par = True
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r2, , True)
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r3, , True) And par
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r4, , True) And par
    End If
    End If
    End If
    
       If Not par Then
     MissParam a$
     Exit Function
                End If
                On Error Resume Next
    r3 = r3 + (r2 - Int(r2)) * 60
    r2 = Int(r2)
    r4 = r4 + (r3 - Int(r3)) * 60
    r3 = Int(r3)
    R = CDbl(TimeSerial(Hour(CDate(R)) + r2, Minute(CDate(R)) + r3, Second(CDate(R)) + r4) + Int(Abs(R)))
    If SG < 0 Then R = -R
                If Err.Number > 0 Then
    WrongArgument a$
    Err.Clear
    Exit Function
    End If
    On Error GoTo 0
     IsCtime = FastSymbol(a$, ")", True)
      Else
   
     MissParam a$
    End If
End Function
Function IsRecords(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim VR As Long
    FastSymbol a$, "#"
    If IsExp(bstack, a$, R, , True) Then
        VR = R Mod 512
        If FLEN(VR) = 0 Then
            wrongfilenumber a$
            
        Else
            R = SG * LOF(VR) / FLEN(VR)
            
            IsRecords = FastSymbol(a$, ")", True)
        End If
    Else
        
        MissParam a$
    End If
End Function
Function ExpandGui(bstack As basetask, what$, rest$, ifier As Boolean, Lang As Long, oName As String)
Dim pppp As mArray, aVar As Variant, H$
 Dim i As Long
'' add new GuiItems but not visible
what$ = Left$(what$, Len(what$) - 1)
 If neoGetArray(bstack, oName$, pppp, here$ <> "") Then
 
 If Not (pppp.IHaveGui Or pppp.comevents) Then
 MyEr "Only for Gui controls or Com objects with events", "Μόνο για στοιχεία διεπαφής ή για αντικείμενα COM με γεγονότα"
 ifier = False: Exit Function
  End If
 i = pppp.UpperMonoLimit + 1
 pppp.SerialItem 0, i + 2, 9
  

  On Error Resume Next
  Select Case Typename$(pppp.item(0))
    Case "GuiM2000"
        CreateFormObject aVar, 1
        Set pppp.item(i) = aVar
        Dim aaa As GuiM2000
        Set aaa = aVar
        With pppp.item(0)
        Set aaa.EventObj = .EventObj
        H$ = .modulename
        End With
        
        With aaa
        .MyName = what$
        .modulename = H$
        .TempTitle = what$ + "(" + LTrim$(Str$(i)) + ")"
        .Index = i
        End With
        Set aaa = Nothing
    Case "GuiButton"
        CreateFormObject aVar, 2
        
        Set pppp.item(i) = aVar
        Dim aaa1 As GuiButton
        Set aaa1 = pppp.item(0)
        With aVar
        .ConstructArray aaa1.GetCallBack, what$, i
        .move 0, 2000, 6000, 600
        .Caption = what$ + "(" + LTrim$(Str$(i)) + ")"
        .SetUp
        End With
    Case "GuiTextBox"
        CreateFormObject aVar, 3
        Set pppp.item(i) = aVar
        Dim aaa2 As GuiTextBox
        Set aaa2 = pppp.item(0)
        With aVar
        .ConstructArray aaa2.GetCallBack, what$, i
        .move 0, 2000, 6000, 600
        .SetUp
        .Text = what$ + "(" + LTrim$(Str$(i)) + ")"
        End With
        
    Case "GuiCheckBox"
        CreateFormObject aVar, 4
        Set pppp.item(i) = aVar
        Dim aaa3 As GuiCheckBox
        Set aaa3 = pppp.item(0)
        With aVar
        .ConstructArray aaa3.GetCallBack, what$, i
        .move 0, 2000, 6000, 600
        .Caption = what$ + "(" + LTrim$(Str$(i)) + ")"
        .SetUp
        End With
    
    Case "GuiEditBox"
        CreateFormObject aVar, 5
        Set pppp.item(i) = aVar
        Dim aaa4 As GuiEditBox
        Set aaa4 = pppp.item(0)
        Set aaa = aaa4.GetCallBack
        With aVar
        .ConstructArray aaa, what$, i
        .move 0, 2000, 6000, 1200
        
        If aaa.prive <> 0 Then
        .Linespace = players(aaa.prive).uMineLineSpace
        Else
        .Linespace = players(0).uMineLineSpace
        End If
        .SetUp
        .Text = what$ + "(" + LTrim$(Str$(i)) + ")"
        End With
        Set aaa = Nothing
    
    Case "GuiListBox"
        CreateFormObject aVar, 6
        Set pppp.item(i) = aVar
        Dim aaa5 As GuiListBox
        Set aaa5 = pppp.item(0)
        Set aaa = aaa5.GetCallBack
        With aVar
        .ConstructArray aaa5.GetCallBack, what$, i
        .move 0, 2000, 6000, 600
        If aaa.prive <> 0 Then
        .Linespace = players(aaa.prive).uMineLineSpace
        Else
        .Linespace = players(0).uMineLineSpace
        End If
        .ListText = what$ + "(" + LTrim$(Str$(i)) + ")"
        .SetUp
        End With
    Case "GuiDropDown"
        CreateFormObject aVar, 7
        Set pppp.item(i) = aVar
        Dim aaa6 As GuiDropDown
        Set aaa6 = pppp.item(0)
        Set aaa = aaa6.GetCallBack
        With aVar
        .ConstructArray aaa5.GetCallBack, what$, i
        .move 0, 2000, 6000, 600
        .SetUp
        End With
    Case "Socket"
    Set aVar = New Socket
    Set pppp.item(pppp.count - 1) = aVar
    Case "SerialPort"
    Set aVar = New SerialPort
    Set pppp.item(pppp.count - 1) = aVar
    Case "ShellPipe"
    Set aVar = New ShellPipe
    Set pppp.item(pppp.count - 1) = aVar
    Case "cHttpDownload"
    Set aVar = New cHttpDownload
    Set pppp.item(pppp.count - 1) = aVar
    Case "CallBack2"
    Dim Aaaa As CallBack2
    
    Set Aaaa = pppp.item(pppp.count - 1)
    Set aVar = Aaaa.CopyOfMe
    Set pppp.item(pppp.count - 1) = aVar
    Case Else
    ' any other com object
    
    End Select
    If pppp.comevents Then
    Dim pppp1 As mArray
    If neoGetArray(bstack, ChrW(&HFFBF) + "_" + what$ + "(", pppp1, True) Then
     i = pppp1.UpperMonoLimit + 1
    pppp1.SerialItem 0, i + 2, 9
    GetShinkArr2 pppp, pppp1, what$, (pppp1.count - 1), (pppp1.count)
    End If
    End If
ifier = True
End If
End Function

Function MyDeclare(bstack As basetask, rest$, Lang As Long, Optional groupok As Boolean) As Boolean
Dim p As Variant, i As Long, s$, pa$
Dim x1 As Long, y1 As Long, par As Boolean, ss$, w$, what$, Y3 As Boolean, Y4 As Boolean
Dim declobj As Object, ML As Long
Dim dum As Boolean
Dim pppp As mArray
Dim ii As Long, ev As ComShinkEvent
MyDeclare = True
ML = -1
    If Not groupok Then y1 = IsLabelSymbolNew(rest$, "ΓΕΝΙΚΟ", "GLOBAL", Lang)
If Not y1 Then

    If Not groupok Then y1 = IsLabelSymbolNew(rest$, "ΤΟΠΙΚΟ", "LOCAL", Lang)
    If y1 Then y1 = y1 * -100
End If
If Not groupok Then Y3 = IsLabelSymbolNew(rest$, "ΜΕΓΕΓΟΝΟΤΑ", "WITHEVENTS", Lang)
x1 = Abs(innerIsLabel(bstack, rest$, what$, , True, True))
If Not groupok Then
w$ = myUcase(what$, True)
Else
w$ = bstack.GroupName + myUcase(what$, True)
End If
gohere:
    If x1 = 1 Or x1 = 3 Then
        If x1 = 1 Then
            If GetVar(bstack, w$, i) And y1 = 0 Then
            
            If Y3 Then
            If TypeOf var(i) Is GuiM2000 Then
                    MyEr "Can't define com events here", "Δεν μπορώ να χειριστώ γεγονότα COM εδώ"
                            MyDeclare = False
            Exit Function
            End If
                        If GetShink(ev, i, w$) Then
                            If Not GetVar(bstack, ChrW(&HFFBF) + "_" + w$, ii) Then
                            
                            ii = globalvarGroup(ChrW(&HFFBF) + "_" + w$, s$, , y1 = True)
                            End If
                            
                            
                            Set var(ii) = ev
                                        
                            Else
                            MyEr "Can't handle events here", "Δεν μπορώ να χειριστώ γεγονότα"
                            MyDeclare = False
                         End If
                        Exit Function
            End If
            ss$ = vbNullString

               If IsLabelSymbolNewExp(rest$, "ΤΙΠΟΤΑ", "NOTHING", Lang, ss$) Then
                   If ML >= 0 Then
goNothing:
              If neoGetArray(bstack, w$, pppp) Then
              ' no check for real array Why ?
              
              For i = 0 To pppp.UpperMonoLimit
                If MyIsObject(pppp.item(i)) Then
                    If Typename$(pppp.item(i)) = "GuiM2000" Then
                    Set declobj = pppp.item(i)
                    Unload declobj
                    End If
                 Set pppp.item(i) = Nothing  ' special use
                Else
                pppp.item(i) = Empty
                End If
              Next i
              Else
              MyDeclare = False
                Exit Function
              End If
                   
                   ElseIf Not MyIsObject(var(i)) Then
                        If groupok Then
                        Set var(i) = Nothing
                        
                        Else
                        
                        BadObjectDecl
                        MyDeclare = False
                        End If
                   Else
                   If var(i) Is Nothing Then Exit Function
                      If TypeOf var(i) Is GuiM2000 Then
                      var(i).CloseNow
                      Unload var(i)
                      Else
                      If GetVar(bstack, ChrW(&HFFBF) + "_" + w$, ii) Then
                        Set var(ii) = Nothing
                      End If
                      End If
                       Set var(i) = Nothing
                   End If
                   Exit Function
               
               ElseIf IsLabelSymbolNewExp(rest$, "ΝΕΟ", "NEW", Lang, ss$) Then
               Set var(i) = Nothing
               GoTo THEREnew
               ElseIf IsLabelSymbolNewExp(rest$, "ΝΕΑ", "NEW", Lang, ss$) Then
               Set var(i) = Nothing
               GoTo THEREnew
               Else
                  BadObjectDecl
                    MyDeclare = False
                  Exit Function
               End If
                     
            End If
   
         End If
         ss$ = vbNullString
         If IsLabelSymbolNewExp(rest$, "ΒΑΣΗ", "BASE", Lang, ss$) Then
         If Not IsStrExp(bstack, rest$, pa$) Then
                   BadObjectDecl
                    MyDeclare = False
                  Exit Function
         End If

         If GetVar(bstack, w$, i) Then
                      BadObjectDecl
                    MyDeclare = False
                  Exit Function
         
         End If

         If Not getone2(pa$, p) Then
         Set p = New Mk2Base
         PushOne pa$, p
         End If
         globalvar bstack.GroupName & w$, p, y1 = True
          MyDeclare = True
                  Exit Function
         ElseIf IsLabelSymbolNewExp(rest$, "ΑΠΟ", "LIB", Lang, ss$) Then
               If y1 = 100 Then y1 = 0
              par = Fast2Label(rest$, "C", 1, "", 0, 1)
              
            If IsStrExp(bstack, rest$, pa$) Then
                       ' we get the lib
                If Not IsStrExp(bstack, rest$, ss$) Then ss$ = vbNullString
                ' clear vbcr
                
                Do
                If FastSymbol(ss$, vbCrLf, , 2) Then
                s$ = vbNullString
                Else
                s$ = GetNextLine(ss$)
                If Not MaybeIsSymbol(s$, "'\") Then Exit Do
                End If
                Loop
                ss$ = s$ + CleanStr(ss$, vbCrLf)
                s$ = vbNullString
                i = AllocVar()
                Set declobj = New stdCallFunction
                If IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then
                If Not IsExp(bstack, rest$, p) Then
                BadObjectDecl
                MyDeclare = False
                Set declobj = Nothing
                Exit Function
                Else
            declobj.RetType = CLng(p)
                End If
                End If
                declobj.CallThis pa$, ss$, Lang
                If par Then declobj.CallType = 1
            
                Set var(i) = declobj
                
                s$ = w$
                If x1 = 3 Then
                    If here$ = vbNullString Or y1 Then
                        GlobalSub s$ + "()", "CALL EXTERN " & Str(i) & " : = LETTER$"
                    Else
                        GlobalSub here$ & "." & bstack.GroupName & s$ + "()", "CALL EXTERN " & Str(i) & " : = LETTER$"
                    End If
                Else
                    If here$ = vbNullString Or y1 Then
                        GlobalSub s$ + "()", "CALL EXTERN " & Str(i) & " : = NUMBER"
                    Else
                        GlobalSub here$ & "." & bstack.GroupName & s$ + "()", "CALL EXTERN " & Str(i) & " : = NUMBER"
                    End If
                End If
               Set declobj = Nothing
                Exit Function
            Else
                BadObjectDecl
                 MyDeclare = False
                 Set declobj = Nothing
                Exit Function
            End If
        ElseIf IsLabelSymbolNewExp(rest$, "ΜΕ", "USE", Lang, ss$) Then
                If Y3 <> 0 Then
                SyntaxError
                Exit Function
                End If


                    x1 = Abs(IsLabel(bstack, rest$, pa$))
                    IsSymbol3 rest$, ","
                    If x1 = 1 Then
                            If GetVar(bstack, pa$, x1) Then
                                    If MyIsObject(var(x1)) Then
                                                i = globalvar(w$, s$, , y1 = True)
                                                If IsStrExp(bstack, rest$, ss$) Then
                                                Set var(i) = MakeObjectFromString(var(x1), ss$)
                                         
                                                
                                                
                                                
                                                End If
                                    End If
                            End If
                    Else
                    MissPar
                    MyDeclare = False
                    End If
                    Exit Function
        End If
        If ML >= 0 Then
         If neoGetArray(bstack, w$, pppp, True) Then
         MyEr "Array exist", "Υπάρχει ήδη αυτός ο πίνακας"
                MyDeclare = False
                Exit Function
         Else
         x1 = -1
         GlobalArr bstack, w$ + "(", LTrim$(Str(ML)) + ")", i, x1, , , CBool(y1)
          If DeclareGUI(bstack, what$, rest$, MyDeclare, Lang, (i), ML, w$, y1, Y3, x1) Then
            If Y3 And y1 = 0 Then
            ii = -1
            GlobalArr bstack, ChrW(&HFFBF) + "_" + w$ + "(", LTrim$(Str(ML)) + ")", i, ii, , , CBool(y1)
              If Not GetShinkArr(x1, ii, w$, 0, i) Then
              MyEr "Events for this type of object", "Δεν υπάρχουν γεγονότα για αυτό το αντικείμενο"
              MyDeclare = False
              Else
                    var(x1).comevents = True
                End If
            End If

        
        
        Exit Function
        End If
        End If
        Else
        i = globalvar(w$, s$, , y1 = True)
        End If
THEREnew:
MyDeclare = False
        Y4 = IsLabelSymbolNew(rest$, "ΠΑΡΕ", "GET", Lang)
        If Y4 And MaybeIsSymbol(rest$, ",") Then
            FastSymbol rest$, ","
            If IsStrExp(bstack, rest$, pa$) Then
                GetitObject var(i), , CVar(pa$)
            End If
            GoTo jump1
        ElseIf IsStrExp(bstack, rest$, s$) Then
            If FastSymbol(rest$, ",") Then
                If IsStrExp(bstack, rest$, pa$) Then
                    If Y4 Then
                        On Error Resume Next
                        GetitObject var(i), CVar(s$), CVar(pa$)
                        If Err.Number > 0 Then Exit Function
                    ElseIf FastSymbol(rest$, ",") Then
                        If IsStrExp(bstack, rest$, w$) Then
                            Err.Clear
                            On Error Resume Next
                            If w$ = vbNullString Then
                                Licenses.Add s$
                            Else
                                Licenses.Add s$, w$
                            End If
                            If Err.Number > 0 And Err.Number <> 732 Then
                                MissLicense
                                Err.Clear
                            Else
                                Err.Clear
                                CreateitObject var(i), s$, CVar(pa$)
                                If Err.Number > 0 Then
                                    Err.Clear
                                    MissLicense
                                End If
                            End If
                            Licenses.Remove s$
                        Else
                            MissStringExpr
                        End If
                    Else
                        Err.Clear
                        On Error Resume Next
                        CreateitObject var(i), s$, CVar(pa$)

                        If Err.Number > 0 Then
                            Err.Clear
                            MissLicense
                        End If
                    End If
                    
                    
                Else
                 If FastSymbol(rest$, ",") Then
                 If IsStrExp(bstack, rest$, pa$) Then
                    Err.Clear
                        On Error Resume Next
                        If pa$ = vbNullString Then
                        Licenses.Add s$
                        Else
                 Licenses.Add s$, pa$
                 End If
                     If Err.Number > 0 And Err.Number <> 732 Then
                        MissLicense
                        Err.Clear
                        Else
                        Err.Clear
                 CreateitObject var(i), s$
                 If Err.Number > 0 Then
                        Err.Clear
                        MissLicense
                        End If
                        End If
                 Licenses.Remove s$
                 Else
                    MissStringExpr
                 End If
                Else
                    MissStringExpr
                    End If
                End If
            Else
               Err.Clear
                        On Error Resume Next
                        If Y4 Then
                            GetitObject var(i), CVar(s$)
                        Else
                            CreateitObject var(i), s$
                        End If
jump1:
                 If Err.Number > 0 Then
                        Err.Clear
                        MissLicense
                        ElseIf Y3 <> 0 Then
                        
                         
                         If GetShink(ev, i, w$) Then
                         ' private
                         ' no need for a name, we do not address it in any way.
                         ' ev knows how to call handler
                            If Not GetVar(bstack, ChrW(&HFFBF) + "_" + w$, ii) Then
                            
                            ii = globalvarGroup(ChrW(&HFFBF) + "_" + w$, s$, , y1 = True)
                            End If
                            
                            
                            Set var(ii) = ev

                         Else
                         MyEr "Can't handle events here", "Δεν μπορώ να χειριστώ γεγονότα"
                            MyDeclare = False
            
                         End If
                         
                        End If
            End If
        ElseIf DeclareGUI(bstack, what$, rest$, MyDeclare, Lang, i, , , , Y3) Then
        If Y3 Then
                If GetShink(ev, i, w$) Then
                    If Not GetVar(bstack, ChrW(&HFFBF) + "_" + w$, ii) Then
                    
                    ii = globalvarGroup(ChrW(&HFFBF) + "_" + w$, s$, , y1 = True)
                    End If
                    Set var(ii) = ev
                End If
        
        
        
        End If
         
        End If
           Err.Clear
                        On Error Resume Next
            
        If Not MyIsObject(var(i)) Then
        If groupok Then
        
        Set var(i) = Nothing
        'OptVariant var(i)
        Else
        BadObjectDecl
        End If
        End If
       
    ElseIf x1 = 5 Then
If FastSymbol(rest$, ")") Then
 'w$ = Left$(w$, Len(w$) - 1)
'w$ = w$ + ")"
 If IsLabelSymbolNewExp(rest$, "ΤΙΠΟΤΑ", "NOTHING", Lang, ss$) Then GoTo goNothing
 If IsLabelSymbolNewExp(rest$, "ΠΑΝΩ", "OVER", Lang, ss$) Then
    ExpandGui bstack, what$, rest$, MyDeclare, Lang, w$

Else

 
 BadObjectDecl
 End If
ElseIf IsExp(bstack, rest$, p) Then
ML = CLng(p)
If FastSymbol(rest$, ")") Then
x1 = 1
 w$ = Left$(w$, Len(w$) - 1)
   GoTo gohere
Else
BadObjectDecl
End If
Else
BadObjectDecl
End If
    Else
    
    BadObjectDecl
    
    End If
     MyDeclare = True
     Set declobj = Nothing
    

End Function

Private Function GetShinkArr(v As Long, where As Long, objname$, from As Long, count As Long) As Boolean
Dim aa As Object, ok As Boolean, i As Long, pppp As mArray, pppp1 As mArray, ev As ComShinkEvent
Set pppp = var(v)
Set pppp1 = var(where)
ok = True
For i = from To count + from - 1
    Set ev = New ComShinkEvent
    
    Set aa = pppp.item(i)
    
                         ev.VarIndex = v
                         ev.ItemIndex = i
                         If here$ <> "" Then
                         ev.modulename = here$ + "." + objname$
                         ev.modulenameonly = here$
                         Else
                         ev.modulename = objname$
                         ev.modulenameonly = vbNullString
                         End If
                         ev.Attach aa
   Set aa = Nothing
   If Not ev.Attached() Then ok = False: Exit For
   Set pppp1.item(i) = ev
   Next i
   GetShinkArr = ok
End Function
Private Function GetShinkArr2(pppp As mArray, pppp1 As mArray, objname$, from As Long, count As Long) As Boolean
Dim aa As Object, ok As Boolean, i As Long, ev As ComShinkEvent
ok = True
Set ev = pppp1.item(0)
Dim v As Long
v = ev.VarIndex
For i = from To count - 1
    Set ev = New ComShinkEvent
    
    Set aa = pppp.item(i)
    
                         ev.VarIndex = v
                         ev.ItemIndex = i
                         If here$ <> "" Then
                         ev.modulename = here$ + "." + objname$
                         ev.modulenameonly = here$
                         Else
                         ev.modulename = objname$
                         ev.modulenameonly = vbNullString
                         End If
                         ev.Attach aa
   Set aa = Nothing
   If Not ev.Attached() Then ok = False: Exit For
   Set pppp1.item(i) = ev
   Next i
   GetShinkArr2 = ok
End Function
Public Function GetShink(ev As ComShinkEvent, v As Long, objname$) As Boolean
Dim aa As Object
    Set ev = New ComShinkEvent

    Set aa = var(v)
    
                         ev.VarIndex = v
                         ev.ItemIndex = -1
                         If here$ <> "" Then
                         ev.modulename = here$ + "." + objname$
                         ev.modulenameonly = here$
                         Else
                         ev.modulename = objname$
                         ev.modulenameonly = vbNullString
                         End If
                         ev.Attach aa
   Set aa = Nothing
   GetShink = ev.Attached()
End Function
Function DeclareGUI(bstack As basetask, what$, rest$, ifier As Boolean, Lang As Long, i As Long, Optional ar As Long = 0, Optional oName As String = vbNullString, Optional glob As Long = 0, Optional Y3 As Boolean, Optional ArrPos As Long = 0)
DeclareGUI = True
Dim w$, x1 As Long, y1 As Long, s$, useold As Boolean, bp As Long
Dim alfa As GuiM2000, mm As CallBack2
Dim aVar As Variant, p As Variant
Dim pppp As mArray, mmmm As mEvent
' for these events no use of noHere property because no copyevent/upgrade happens
'         see .upgrade in UnFloatGroupReWriteVars and UnFloatGroup

 If IsLabelSymbolNew(rest$, "ΦΟΡΜΑ", "FORM", Lang) Then
    Y3 = False ' no withevents in declare
    If IsLabelSymbolNew(rest$, "ΓΕΓΟΝΟΣ", "EVENT", Lang) Then
        x1 = Abs(IsLabel(bstack, rest$, w$))
        If x1 <> 1 Then
                BadObjectDecl
        Else
           If GetlocalVar(bstack.GroupName & w$, y1) Then
               useold = True
           ElseIf GetVar(bstack, bstack.GroupName & w$, y1) Then
               useold = True
           Else
            y1 = globalvar(bstack.GroupName & w$, s$)
            MakeitObjectEvent var(y1)
           End If
           If Typename(var(y1)) <> "mEvent" Then
               ifier = False
               DeclareGUI = False: Exit Function
           End If
        End If
        If ar = 0 Then
            If Not useold Then ProcEvent bstack, "{Read msg$, &obj}", 1, y1
                CreateFormObject var(i), 1
                Set alfa = var(i)
                Set alfa.safe = getSafeFormList()
                Set alfa.EventObj = var(y1)
                alfa.Index = -1
                alfa.MyName = what$
                alfa.modulename = here$
                alfa.TempTitle = what$
                Set alfa = Nothing
            Else
                ProcEvent bstack, "{Read index, msg$, &obj}", 1, y1
                Set mmmm = var(y1)
                GoTo contEvArray
            End If
        Else
            bp = True
            If ar = 0 Then
                CreateFormObject var(i), 1
                Set alfa = var(i)
                Set mmmm = New mEvent
                Set alfa.safe = getSafeFormList()
                Set alfa.EventObj = mmmm
                With mmmm
                    .BypassInit 10
                    .VarIndex = i * 1000 + varhash.count
                    .enabled = True
                    .ParamBlock "Read msg$, &obj", 2
                    .GenItemCreator LTrim$(Str(i * 456)), "{ Module " + here$ + vbCrLf + "try { Call local " + here$ + "." + bstack.GroupName + what$ + "() } }" + here$ + "." + bstack.GroupName
                End With
                alfa.MyName = what$
                alfa.Index = -1
                alfa.modulename = here$
                alfa.ByPass = bp
                alfa.TempTitle = what$
                Set mmmm = Nothing
                Set alfa = Nothing
            Else
                Set mmmm = New mEvent
contEvArray:
                'If neoGetArray(bstack, oName$ + "(", pppp, , CBool(glob)) Then
                    Set pppp = var(ArrPos)
                    what$ = Left$(what$, Len(oName$))
                    If y1 = 0 Then
                        With mmmm
                            .BypassInit 10
                            .VarIndex = i * 1000 + varhash.count
                            .enabled = True
                            .ParamBlock "Read Index, msg$, &obj", 3
                            .GenItemCreator LTrim$(Str(i * 3456)), "{ Module " + here$ + vbCrLf + "try { call local " + here$ + "." + bstack.GroupName + what$ + "() } }" + here$ + "." + bstack.GroupName
                        End With
                    End If
                    For i = 0 To ar - 1
                        CreateFormObject aVar, 1
                        Set pppp.item(i) = aVar
                        Dim aaa As GuiM2000, safe As LongHash
                        Set safe = getSafeFormList()
                        Set aaa = aVar
                        With aaa
                            Set .safe = safe
                            Set .EventObj = mmmm
                            .MyName = what$
                            .modulename = here$
                            .ByPass = bp
                            .TempTitle = what$ + "(" + LTrim$(Str$(i)) + ")"
                            .Index = i
                        End With
                    Next i
                    Set aaa = Nothing
                'End If
                pppp.IHaveGui = True
                Set mmmm = Nothing
            End If
        End If
 ElseIf IsLabelSymbolNew(rest$, "ΠΛΗΚΤΡΟ", "BUTTON", Lang) Then
     Y3 = False ' no withevents in declare
    If IsLabelSymbolNew(rest$, "ΦΟΡΜΑ", "FORM", Lang) Then
        If Not IsExp(bstack, rest$, p) Then
            BadObjectDecl
        Else
            If bstack.lastobj Is Nothing Then
                BadObjectDecl
                Exit Function
            End If
            If ar = 0 Then
                CreateFormObject var(i), 2
                Set alfa = bstack.lastobjIndirect(var())
                Set bstack.lastobj = Nothing
                With var(i)
                    .Construct alfa, what$
                    .move 0, 2000, 6000, 600
                    .Caption = what$
                    .SetUp
                End With
                Set alfa = Nothing
            Else
                If neoGetArray(bstack, oName$ + "(", pppp, , CBool(glob)) Then
                    what$ = Left$(what$, Len(oName$))
                    Set alfa = bstack.lastobjIndirect(var())
                    Set bstack.lastobj = Nothing
                    For i = 0 To ar - 1
                        CreateFormObject aVar, 2
                        Set pppp.item(i) = aVar
                        With aVar
                            .ConstructArray alfa, what$, i
                            .move 0, 2000, 6000, 600
                            .Caption = what$ + "(" + LTrim$(Str$(i)) + ")"
                            .SetUp
                        End With
                    Next i
                    pppp.IHaveGui = True
                    Set alfa = Nothing
                End If
            End If
        End If
    End If
ElseIf IsLabelSymbolNew(rest$, "ΕΙΣΑΓΩΓΗ", "TEXTBOX", Lang) Then
    Y3 = False ' no withevents in declare
    If IsLabelSymbolNew(rest$, "ΦΟΡΜΑ", "FORM", Lang) Then
        If Not IsExp(bstack, rest$, p) Then
            BadObjectDecl
        Else
            If bstack.lastobj Is Nothing Then
                BadObjectDecl
                Exit Function
            End If
            If ar = 0 Then
                CreateFormObject var(i), 3
                Set alfa = bstack.lastobjIndirect(var())
                Set bstack.lastobj = Nothing
                With var(i)
                    .Construct alfa, what$
                    .move 0, 2000, 6000, 600
                    .SetUp
                    .Text = what$
                End With
                Set alfa = Nothing
            Else
                If neoGetArray(bstack, oName$ + "(", pppp, , CBool(glob)) Then
                    what$ = Left$(what$, Len(oName$))
                    Set alfa = bstack.lastobjIndirect(var())
                    Set bstack.lastobj = Nothing
                    For i = 0 To ar - 1
                        CreateFormObject aVar, 3
                        Set pppp.item(i) = aVar
                        With aVar
                            .ConstructArray alfa, what$, i
                            .move 0, 2000, 6000, 600
                            .SetUp
                            .Text = what$ + "(" + LTrim$(Str$(i)) + ")"
                        End With
                    Next i
                    pppp.IHaveGui = True
                    Set alfa = Nothing
                End If
            End If
        End If
    End If
ElseIf IsLabelSymbolNew(rest$, "ΕΠΙΛΟΓΗ", "CHECKBOX", Lang) Then
    Y3 = False ' no withevents in declare
     If IsLabelSymbolNew(rest$, "ΦΟΡΜΑ", "FORM", Lang) Then
        If Not IsExp(bstack, rest$, p) Then
            BadObjectDecl
        Else
            If bstack.lastobj Is Nothing Then
                BadObjectDecl
                Exit Function
            End If
            If ar = 0 Then
                CreateFormObject var(i), 4
                Set alfa = bstack.lastobjIndirect(var())
                Set bstack.lastobj = Nothing
                With var(i)
                    .Construct alfa, what$
                    .move 0, 2000, 6000, 600
                    .Caption = what$
                    .SetUp
                End With
                Set alfa = Nothing
            Else
                If neoGetArray(bstack, oName$ + "(", pppp, , CBool(glob)) Then
                    what$ = Left$(what$, Len(oName$))
                    Set alfa = bstack.lastobjIndirect(var())
                    Set bstack.lastobj = Nothing
                    For i = 0 To ar - 1
                        CreateFormObject aVar, 4
                        Set pppp.item(i) = aVar
                        With aVar
                            .ConstructArray alfa, what$, i
                            .move 0, 2000, 6000, 600
                            .Caption = what$ + "(" + LTrim$(Str$(i)) + ")"
                            .SetUp
                        End With
                    Next i
                    pppp.IHaveGui = True
                    Set alfa = Nothing
                End If
            End If
        End If
    End If
ElseIf IsLabelSymbolNew(rest$, "ΚΕΙΜΕΝΟ", "EDITBOX", Lang) Then
    Y3 = False ' no withevents in declare
    If IsLabelSymbolNew(rest$, "ΦΟΡΜΑ", "FORM", Lang) Then
        If Not IsExp(bstack, rest$, p) Then
            BadObjectDecl
        Else
            If bstack.lastobj Is Nothing Then
                BadObjectDecl
                Exit Function
            End If
            If ar = 0 Then
                CreateFormObject var(i), 5
                Set alfa = bstack.lastobjIndirect(var())
                Set bstack.lastobj = Nothing
                With var(i)
                    .Construct alfa, what$
                    .move 0, 2000, 6000, 1200
                    If alfa.prive <> 0 Then
                        .Linespace = players(alfa.prive).uMineLineSpace
                    Else
                        .Linespace = players(0).uMineLineSpace
                    End If
                    .SetUp
                    .Text = what$
                End With
                Set alfa = Nothing
            Else
                If neoGetArray(bstack, oName$ + "(", pppp, , CBool(glob)) Then
                    what$ = Left$(what$, Len(oName$))
                    Set alfa = bstack.lastobjIndirect(var())
                    Set bstack.lastobj = Nothing
                    For i = 0 To ar - 1
                        CreateFormObject aVar, 5
                        Set pppp.item(i) = aVar
                        With aVar
                            .ConstructArray alfa, what$, i
                            .move 0, 2000, 6000, 1200
                            If alfa.prive <> 0 Then
                                .Linespace = players(alfa.prive).uMineLineSpace
                            Else
                                .Linespace = players(0).uMineLineSpace
                            End If
                            .SetUp
                            .Text = what$ + "(" + LTrim$(Str$(i)) + ")"
                        End With
                    Next i
                    pppp.IHaveGui = True
                    Set alfa = Nothing
                End If
            End If
        End If
    End If
ElseIf IsLabelSymbolNew(rest$, "ΛΙΣΤΑ", "LISTBOX", Lang) Then
    Y3 = False ' no withevents in declare
    If IsLabelSymbolNew(rest$, "ΦΟΡΜΑ", "FORM", Lang) Then
        If Not IsExp(bstack, rest$, p) Then
            BadObjectDecl
        Else
            If bstack.lastobj Is Nothing Then
                BadObjectDecl
                Exit Function
            End If
            If ar = 0 Then
                CreateFormObject var(i), 6
                Set alfa = bstack.lastobjIndirect(var())
                Set bstack.lastobj = Nothing
                With var(i)
                    .Construct alfa, what$
                    .move 0, 2000, 6000, 600
                            If alfa.prive <> 0 Then
                                .Linespace = players(alfa.prive).uMineLineSpace
                            Else
                                .Linespace = players(0).uMineLineSpace
                            End If
                    .ListText = what$
                    .SetUp
                End With
                Set alfa = Nothing
            Else
                If neoGetArray(bstack, oName$ + "(", pppp, , CBool(glob)) Then
                    what$ = Left$(what$, Len(oName$))
                    Set alfa = bstack.lastobjIndirect(var())
                    Set bstack.lastobj = Nothing
                    For i = 0 To ar - 1
                        CreateFormObject aVar, 6
                        Set pppp.item(i) = aVar
                        With aVar
                            .ConstructArray alfa, what$, i
                            .move 0, 2000, 6000, 600
                            If alfa.prive <> 0 Then
                                .Linespace = players(alfa.prive).uMineLineSpace
                            Else
                                .Linespace = players(0).uMineLineSpace
                            End If
                            .ListText = what$ + "(" + LTrim$(Str$(i)) + ")"
                            .SetUp
                        End With
                    Next i
                    pppp.IHaveGui = True
                    Set alfa = Nothing
                End If
            End If
        End If
    End If
ElseIf IsLabelSymbolNew(rest$, "ΛΙΣΤΑ.ΕΙΣΑΓΩΓΗΣ", "COMBOBOX", Lang) Then
    Y3 = False ' no withevents in declare
    If IsLabelSymbolNew(rest$, "ΦΟΡΜΑ", "FORM", Lang) Then
        If Not IsExp(bstack, rest$, p) Then
            BadObjectDecl
        Else
            If bstack.lastobj Is Nothing Then
                BadObjectDecl
                Exit Function
            End If
            If ar = 0 Then
                CreateFormObject var(i), 7
                Set alfa = bstack.lastobjIndirect(var())
                Set bstack.lastobj = Nothing
                With var(i)
                    .Construct alfa, what$
                    .move 0, 2000, 6000, 600
                    .SetUp
                End With
                Set alfa = Nothing
            Else
                If neoGetArray(bstack, oName$ + "(", pppp, , CBool(glob)) Then
                    what$ = Left$(what$, Len(oName$))
                    Set alfa = bstack.lastobjIndirect(var())
                    Set bstack.lastobj = Nothing
                    For i = 0 To ar - 1
                        CreateFormObject aVar, 7
                        Set pppp.item(i) = aVar
                        With aVar
                            .ConstructArray alfa, what$, i
                            .move 0, 2000, 6000, 600
                            .SetUp
                        End With
                    Next i
                    pppp.IHaveGui = True
                    Set alfa = Nothing
                End If
            End If
        End If
    End If
ElseIf IsLabelSymbolNew(rest$, "ΠΛΗΡΟΦΟΡΙΕΣ", "INFORMATION", Lang) Then
    
    Y3 = False ' no withevents in declare
    If ar > 0 Then
    MyEr "not for array", "όχι για πίνακα"
    DeclareGUI = False
    End If
    Set var(i) = OsInfo
ElseIf IsLabelSymbolNew(rest$, "ΣΥΜΠΙΕΣΤΗΣ", "COMPRESSOR", Lang) Then
    If ar = 0 Then
        Set var(i) = New ZipTool
    Else
        what$ = Left$(what$, Len(oName$))
        Set pppp = var(ArrPos)
        For i = 0 To ar - 1
            Set pppp.item(i) = New ZipTool
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΜΑΘΗΜΑΤΙΚΑ", "MATH", Lang) Then
    Y3 = False ' no withevents in declare
    Set var(i) = New Math
ElseIf IsLabelSymbolNew(rest$, "ΣΤΑΘΜΟΣ", "SOCKET", Lang) Then
    If ar = 0 Then
        Set var(i) = New Socket
    Else
        what$ = Left$(what$, Len(oName$))
        Set pppp = var(ArrPos)
        For i = 0 To ar - 1
            Set pppp.item(i) = New Socket
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΑΥΛΟΣ", "SHELLPIPE", Lang) Then
    If ar = 0 Then
        Set var(i) = New ShellPipe
    Else
        what$ = Left$(what$, Len(oName$))
        Set pppp = var(ArrPos)
        For i = 0 To ar - 1
            Set pppp.item(i) = New ShellPipe
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΣΕΙΡΙΑΚΗ", "SERIALPORT", Lang) Then
    If ar = 0 Then
        Set var(i) = New SerialPort
    Else
        what$ = Left$(what$, Len(oName$))
        Set pppp = var(ArrPos)
        For i = 0 To ar - 1
            Set pppp.item(i) = New SerialPort
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΚΑΤΕΒΑΣΜΑ", "DOWNLOAD", Lang) Then
    If ar = 0 Then
        Set var(i) = New cHttpDownload
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set pppp.item(i) = New cHttpDownload
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΠΕΛΑΤΗΣ", "CLIENT", Lang) Then
    Y3 = False ' no withevents in declare
    If ar = 0 Then
        Set var(i) = New cTlsClient
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set pppp.item(i) = New cTlsClient
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΣΥΛΛΟΓΗ", "COLLECTION", Lang) Then
    Y3 = False ' no withevents in declare
    If ar = 0 Then
        Set var(i) = New Collection
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set pppp.item(i) = New Collection
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΣΤΟΙΧΕΙΑXML", "XMLDATA", Lang) Then
    Y3 = False ' no withevents in declare
    If ar = 0 Then
        Set var(i) = xmlMonoNew
        
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set pppp.item(i) = xmlMonoNew
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΠΙΝΑΚΑΣJSON", "JSONARRAY", Lang) Then
    Y3 = False
    If ar = 0 Then
        Set var(i) = New JsonArray
        
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set pppp.item(i) = New JsonArray
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΛΙΣΤΑJSON", "JSONOBJECT", Lang) Then
    Y3 = False
    If ar = 0 Then
        Set var(i) = New JsonObject
        
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set pppp.item(i) = New JsonObject
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΜΟΝΑΔΙΚΟ", "MUTEX", Lang) Then
    Y3 = False ' no withevents in declare
    If ar = 0 Then
        Set var(i) = New Mutex
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set pppp.item(i) = New Mutex
        Next i
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΕΦΑΡΜΟΓΗ", "APPLICATION", Lang) Then
    If IsLabelSymbolNew(rest$, "ΦΟΡΜΑ", "FORM", Lang) Then
        Set var(i) = Form1
    Else
        If ar = 0 Then
            Set mm = New CallBack2
            mm.NoPublic bstack, ""
            Set var(i) = mm
            Set mm = Nothing
        Else
            Set pppp = var(ArrPos)
            what$ = Left$(what$, Len(oName$))
            For i = 0 To ar - 1
                Set mm = New CallBack2
                mm.NoPublic bstack, ""
                Set pppp.item(i) = mm
                Set mm = Nothing
            Next i
        End If
    End If
    Exit Function
ElseIf IsLabelSymbolNew(rest$, "ΤΜΗΜΑ", "MODULE", Lang) Then
    If ar = 0 Then
            Set mm = New CallBack2
            mm.NoPublic bstack, here$
            Set var(i) = mm
            Set mm = Nothing
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set mm = New CallBack2
            mm.NoPublic bstack, here$
            Set pppp.item(i) = mm
            Set mm = Nothing
        Next i
    End If
ElseIf IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then
    If IsStrExp(bstack, rest$, s$) Then GoTo ccc
ElseIf IsStrExp(bstack, rest$, s$) Then
ccc:
    Y3 = False ' no withevents in declare
    If ar = 0 Then
        CreateitObject var(i), s$
    Else
        Set pppp = var(ArrPos)
        what$ = Left$(what$, Len(oName$))
        For i = 0 To ar - 1
            Set p = Nothing
            CreateitObject p, s$
            Set pppp.item(i) = p
        Next i
    End If
    Exit Function


Stop
End If
End Function

Function ReArrangePara(what$) As String
ReArrangePara = what$
Exit Function
Dim a1() As Integer, A2() As Integer, WHAT1$, R As Long, mark1 As Long, ii As Long, mark2 As Long
Dim wr$
If Len(what$) = 0 Then Exit Function
    ReDim a1(Len(what$) + 10)
    ReDim A2(Len(what$) + 10)
    Dim skip As Boolean
    skip = GetStringTypeExW(&HB, 4, StrPtr(what$), Len(what$), a1(0)) = 0
    skip = GetStringTypeExW(&HB, 2, StrPtr(what$), Len(what$), A2(0)) = 0 Or skip
    
    If Not skip Then
    For R = 0 To Len(what$) - 1
    If (A2(R) And 254) = 2 And (a1(R) And &H8000) <> 0 Then
    If mark1 = 0 Then WHAT1$ = WHAT1$ + Left$(what$, R)
    mark1 = R + 1
    mark2 = 0
    For ii = mark1 + 1 To Len(what$) - 1
    If (A2(ii) And 254) > 2 And (a1(ii) And 7) = 0 Then mark2 = ii: Exit For
    Next ii
    If mark2 = 0 Then
    wr$ = Mid$(what$, mark1) + wr$
    R = ii
    Else
    wr$ = Mid$(what$, mark1, mark2 - mark1 + 1) + wr$

    R = ii - 1
    End If
    ElseIf (a1(R) And &HFFF8) <> 0 And (a1(ii) And 7) = 0 Then
    mark1 = 0
    For ii = R To Len(what$) - 1
  
    If mark2 > 0 Then
    If (a1(R) And &H8000) <> 0 Then
    If (A2(ii) And 2) = 2 Then mark1 = ii: Exit For
    Else
    If A2(ii) > 3 And mark1 > 0 Then
    
   ' If (A2(ii + 1) And 15) <> 3 Then
   If A2(ii) > 9 Then ii = ii - 1: Exit For
    ElseIf (A2(ii) And 3) = 2 Then
        If A2(ii) > 3 And A2(ii + 1) = 1 Then
                WHAT1$ = WHAT1$ + wr$
                wr$ = ""
                mark2 = 0
         Else
            mark1 = ii: Exit For
        End If
    ElseIf A2(ii) > 3 Then

        If (A2(ii + 1) And 3) > 1 Then mark1 = ii
        Exit For

    End If
        mark1 = ii
    
    End If
    Else
    If Len(wr$) > 0 Then
    If A2(ii) > 3 Then
            If A2(ii) > 9 And A2(ii + 1) > 9 Then
        mark1 = 0
        
        Else
    If (A2(ii + 1) And 1) = 1 And A2(ii + 1) <> 3 Then
          WHAT1$ = WHAT1$ + wr$
    wr$ = ""
    mark1 = ii
    Else
    
    mark2 = ii
    mark1 = ii
    
    If A2(ii + 1) < 4 Then Exit For
    End If
    End If
    ElseIf A2(ii + 1) > 3 Then

    If A2(R) = 3 Then
    mark2 = ii
    mark1 = ii
    
    Exit For
    End If
    End If
    End If
    If (A2(ii) And 254) = 2 And (a1(ii) And &H8000) <> 0 Then mark1 = ii: Exit For
    End If
    
    Next ii

    If mark2 = 0 Then
    WHAT1$ = WHAT1$ + Mid$(what$, R + 1, ii - R)
    R = ii - 1
    ElseIf mark1 = 0 Then
    If Len(wr$) > 0 Then wr$ = wr$ + Mid$(what$, R + 1, ii - R + 1)
    R = ii
    Else
   
    If Len(wr$) > 0 Then wr$ = Mid$(what$, R + 1, ii - R + 1) + wr$
    R = ii
    End If
    If mark2 > 0 Then
    If (A2(ii) And 3) = 3 Then

        mark2 = 0
    ElseIf (A2(ii) And 3) = 0 Then
      WHAT1$ = WHAT1$ + wr$
    wr$ = ""
    mark2 = 0
    End If
    
    End If
    
    
    Else
    
    End If
    Next R
    
    
    
    End If
    If Len(wr$) = 0 And Len(WHAT1$) = 0 Then
    ReArrangePara = what$
    Else
    ReArrangePara = WHAT1$ + wr$
    End If
End Function
Function mydata2(bstack As basetask, rest$, RetStack As mStiva) As Boolean
Dim s$, p As Variant, usehandler As mHandler ', vvl As Variant, photo As Object
mydata2 = True
Do
    If FastSymbol(rest$, "!") Then
                If IsExp(bstack, rest$, p) Then
                    If bstack.lastobj Is Nothing Then
                        RetStack.DataValLong p
                    ElseIf TypeOf bstack.lastobj Is mHandler Then
                        Set usehandler = bstack.lastobj
                        If TypeOf usehandler.objref Is mStiva Then
                            RetStack.MergeBottom usehandler.objref
                        ElseIf TypeOf usehandler.objref Is mArray Then
                            RetStack.MergeBottomCopyArray usehandler.objref
                        Else
                            mydata2 = False
                            MyEr "Expected Stack Object or Array after !", "Περίμενα αντικείμενο Σωρό ή πίνακα μετά το !"
                            Set bstack.lastobj = Nothing
                            Exit Function
                        End If
                        Set usehandler = Nothing
                        Set bstack.lastobj = Nothing
                    ElseIf TypeOf bstack.lastobj Is mArray Then
                        RetStack.MergeBottomCopyArray bstack.lastobj
                        Set bstack.lastobj = Nothing
                    End If
                        
                     End If
                ElseIf IsExp(bstack, rest$, p) Then
                  If bstack.lastobj Is Nothing Then
                      RetStack.DataVal p
                 Else
                   If TypeOf bstack.lastobj Is mStiva Then
                   Set bstack.Sorosref = bstack.lastobj
                   ElseIf TypeOf bstack.lastobj Is VarItem Then
                    RetStack.DataObjVaritem bstack.lastobj
                      Else
                      RetStack.DataObj bstack.lastobj
                    Set bstack.lastpointer = Nothing
                    End If
                      Set bstack.lastobj = Nothing
                End If
        ElseIf IsStrExp(bstack, rest$, s$) Then
                If bstack.lastobj Is Nothing Then
                        RetStack.DataStr s$
                Else
                        RetStack.DataObj bstack.lastobj
                        Set bstack.lastobj = Nothing
                          Set bstack.lastpointer = Nothing
                End If
        Else
                mydata2 = LastErNum1 = 0
                Exit Do
        End If
        If Not FastSymbol(rest$, ",") Then Exit Do
        
Loop
End Function
Function IsLabelDot2(bstack As basetask, a$, R$) As Long    'ok
' for left side...no &

Dim rr&, one As Boolean, c$, firstdot$, gr As Boolean
R$ = vbNullString
If a$ = vbNullString Then IsLabelDot2 = 0: Exit Function

a$ = NLtrim$(a$)
    Do While Len(a$) > 0
    c$ = Left$(a$, 1)
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 64  '"@"
           
              IsLabelDot2 = 0: a$ = firstdot$ + a$: Exit Function

        Case 63 '"?"
        If R$ = vbNullString And firstdot$ = vbNullString Then
        R$ = "?"
        a$ = Mid$(a$, 2)
        IsLabelDot2 = 1
        Exit Function
    
        ElseIf firstdot$ = vbNullString Then
        IsLabelDot2 = 1
        Exit Function
        Else
        IsLabelDot2 = 0
        Exit Function
        End If
        Case 46 '"."
            If one Then
            Exit Do
            Exit Do
            ElseIf R$ <> "" And Len(a$) > 1 Then
            If Mid$(a$, 2, 2) = ". " Or Mid$(a$, 2, 1) = " " Then Exit Do
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1
            Else
            firstdot$ = firstdot$ + "."
            a$ = Mid$(a$, 2)
            End If
       Case 92, 94, 123 To 126, 160 '"\","^", "{" To "~"
        Exit Do

        Case 48 To 57, 95 '"0" To "9", "_"
           If one Then
            If firstdot$ <> "" Then a$ = firstdot$ + a$
            Exit Do
            ElseIf R$ <> "" Then
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit Do
            End If
        Case Is < 0, Is > 64 ' >=A and negative
            If one Then
            Exit Do
            Else
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        Case 36 ' "$"
            If one Then Exit Do
            If R$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If R$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 40 ' "("
            If R$ <> "" Then
                            If Mid$(a$, 2, 2) = ")@" Then
                                    R$ = R$ & "()."
                                  
                                 a$ = Mid$(a$, 4)
                               Else
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit Do
                                       End Select
                                       R$ = R$ & Left$(a$, 1)
                                       a$ = Mid$(a$, 2)
                                   Exit Do
                            
                          End If
               Else
                        Exit Do
            
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
            gr = True
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        End If

    Loop
    If Len(firstdot$) > 0 Then
     R$ = myUcase(R$, gr)
    rr& = bstack.GetDotNew(R$, Len(firstdot$)) * rr&
    Else
       R$ = myUcase(R$, gr)
       End If
    IsLabelDot2 = rr&
   

End Function
Function ProcDrawWidth(bstack As basetask, rest$) As Boolean
Dim x As Double, p As Variant, it As Long, ss$, i As Long, x1 As Long, nd&, once As Boolean
ProcDrawWidth = True
Dim Scr As Object
Set Scr = bstack.Owner
If IsExp(bstack, rest$, p, , True) Then
    i = Scr.DrawWidth
    x1 = Scr.DrawStyle
    If Int(p) < 1 Then p = 1
    Scr.DrawWidth = p
   
        If FastSymbol(rest$, ",") Then
            If IsExp(bstack, rest$, x, , True) Then
                On Error Resume Next
                x = Int(x)
                If x >= 0 Or x <= 6 Then
                    Scr.DrawStyle = x
                    If Err Then x = 0: Scr.DrawStyle = Int(x)
                    Scr.DrawWidth = p
                End If
            End If
        End If
   
    If FastSymbol(rest$, "{") Then
        ss$ = block(rest$)
         TraceStore bstack, nd&, rest$, 0
        If FastSymbol(rest$, "}") Then
            Call executeblock(it, bstack, ss$, False, once, , True)
        End If
        bstack.addlen = nd&
    Else
        MissingBlockCode
    End If
Else
MissNumExpr
End If

If it = 2 Then
If ss$ = "" Then
If once Then rest$ = ": Break": If trace Then WaitShow = 2: TestShowSub = vbNullString
Else
rest$ = ": Goto " + ss$
If trace Then WaitShow = 2: TestShowSub = rest$
End If

it = 1
End If
If it <> 1 Then ProcDrawWidth = False: rest$ = ss$ + rest$
If i <= 0 Then i = 1
Scr.DrawWidth = i
Scr.DrawStyle = x1
Scr.DrawWidth = i
Scr.DrawStyle = x1
Set Scr = Nothing
End Function
Function ProcCurve(bstack As basetask, rest$, Lang As Long) As Boolean
Dim par As Boolean, sX As Double, sY As Double, x As Double, y As Double, x1 As Integer, p As Variant, F As Long
Dim Scr As Object

Set Scr = bstack.Owner
ProcCurve = True
With players(GetCode(Scr))
If IsLabelSymbolNew(rest$, "ΓΩΝΙΑ", "ANGLE", Lang) Then par = True
F = 32
ReDim PLG(F)
x1 = 1
PLG(0).x = Scr.ScaleX(.XGRAPH, 1, 3)
PLG(0).y = Scr.ScaleY(.YGRAPH, 1, 3)
Do
If x1 >= F Then F = F * 2: ReDim Preserve PLG(F)
If IsExp(bstack, rest$, p) Then
x = p

If Not FastSymbol(rest$, ",") Then ProcCurve = False: MissNumExpr: Exit Function
If IsExp(bstack, rest$, p) Then
If par Then
sX = x / PI2
sX = (sX - Fix(sX)) * PI2
.XGRAPH = .XGRAPH + Cos(sX) * p
.YGRAPH = .YGRAPH - Sin(sX) * p
Else
.XGRAPH = .XGRAPH + CLng(x)
.YGRAPH = .YGRAPH + CLng(p)
End If
PLG(x1).x = Scr.ScaleX(.XGRAPH, 1, 3)
PLG(x1).y = Scr.ScaleY(.YGRAPH, 1, 3)

Else
 ProcCurve = False: MissNumExpr: Exit Function
End If
Else
 ProcCurve = False: MissNumExpr: Exit Function
End If

x1 = x1 + 1
Loop Until Not FastSymbol(rest$, ",")
x1 = x1 - 1
Dim mGDILines As Boolean
'If GDILines Then
'    mGDILines = Not (TypeOf scr Is MetaDc And scr.DrawWidth = 1) And Not .NoGDI
'ElseIf scr.DrawWidth > 1 And Not .NoGDI Then
'    mGDILines = Not TypeOf scr Is MetaDc
'End If
Dim trans As Long
trans = .mypentrans
If .IamEmf Then

   mGDILines = Not .NoGDI And Not ((Scr.DrawStyle > 0 And Scr.DrawWidth = 1) And Not Scr.DrawStyle)
ElseIf GDILines Or (trans < 255) Then
    mGDILines = Not .NoGDI
Else
    mGDILines = Scr.DrawWidth > 1

End If
Dim pencol As Long, bstyle As Long, Col As Long
If .pathfillstyle = 1 And .IamEmf Then mGDILines = False
If mGDILines Then
    pencol = .mypen
    If .pathgdi > 0 Then
        Col = .pathcolor: bstyle = .pathfillstyle
        M2000Pen trans, Col
      
         If trans < 255 And bstyle = 5 Then M2000Pen trans, pencol Else M2000Pen 255, pencol
        If x1 + 4 >= F Then F = F + 4: ReDim Preserve PLG(F)
        PLG(x1 + 1) = PLG(0)
        PLG(x1 + 2) = PLG(0)
        PLG(x1 + 3) = PLG(0)
        PLG(x1 + 4) = PLG(0)
        DrawBezierGdi Scr.Hdc, pencol, Col, .pathfillstyle, Scr.DrawWidth, Scr.DrawStyle, PLG(), CLng(x1 + 4)
    Else
        DrawBezierGdi Scr.Hdc, pencol, 0, 1, Scr.DrawWidth, Scr.DrawStyle, PLG(), CLng(x1 + 1)
    End If
Else
If PolyBezier(Scr.Hdc, PLG(0), x1 + 1) = 0 Then
BadGraphic
 Exit Function
End If
End If
Scr.fillstyle = vbSolid
End With
MyDoEvents1 Scr


End Function
Function ProcSort(basestack As basetask, rest$, Lang As Long) As Boolean
Dim i As Long, s$, sX As Double, sY As Double, pppp As mArray
Dim x1 As Long, y1 As Long, p As Variant, ML As Long, desc As Boolean, numb As Boolean, useclid As Boolean
ProcSort = False
Dim mm As mArray, uHandler As mHandler
desc = IsLabelSymbolNew(rest$, "ΦΘΙΝΟΥΣΑ", "DESCENDING", Lang)
If Not desc Then
    useclid = IsLabelSymbolNew(rest$, "ΑΥΞΟΥΣΑ", "ASCENDING", Lang)
Else
    useclid = True
End If
    y1 = Abs(IsLabel(basestack, rest$, s$))
    If y1 = 1 Then
         If GetVar(basestack, s$, i) Then
                If Typename(var(i)) = "mHandler" Then
                Set uHandler = var(i)
usehandler:
                        If uHandler.ReadOnly Then
                        ReadOnly
                        Exit Function
                        End If
                            If uHandler.t1 = 1 Then
                                If IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then
                                    numb = IsLabelSymbolNew(rest$, "ΑΡΙΘΜΟΣ", "NUMBER", Lang, , , , False)
                                    If Not numb Then
                                    If Not IsLabelSymbolNew(rest$, "ΚΕΙΜΕΝΟ", "TEXT", Lang, , , , False) Then
                                    MyEr "Expected Text or Number", "Περίμενα Κείμενο ή Αριθμός"
                                    ProcSort = False
                                    Exit Function
                                    End If
                                    End If
                                    uHandler.objref.NumericSort = numb
                                End If
                                If FastSymbol(rest$, ",") Then
                                    If IsExp(basestack, rest$, p) Then
                                        If FastSymbol(rest$, ",") Then
                                            Set mm = New mArray
                                            mm.PushDim 20, (0)
                                            mm.PushEnd
                                            If Not useclid Then
                                            mm.item(0) = p
                                            y1 = 1
                                            Else
                                            x1 = p
                                            y1 = 0
                                            End If
                                            Do
                                            If IsExp(basestack, rest$, p) Then
                                                mm.item(y1) = p
                                            Else
                                                mm.item(y1) = 0
                                            End If
                                            y1 = y1 + 1
                                            Loop While FastSymbol(rest$, ",")
                                            mm.StartResize
                                            mm.PushDim y1, (0)
                                            mm.PushEnd
                                            mm.ExportArrayNow
                                            With uHandler.objref
                                                .FeedSCol2 mm.refArray
                                                If useclid Then
                                                    .SetBinaryCompare
                                                    .useclid = x1
                                                    If desc Then
                                                        .SortDes
                                                    Else
                                                        .Sort
                                                    End If
                                                    .SetTextCompare
                                                Else
                                                    .useclid = 0
                                                    .Sort
                                                End If
                                                .DisposeCol
                                            End With
                                        Else
                                            With uHandler.objref
                                                If useclid Then
                                                    .SetBinaryCompare
                                                    .useclid = CLng(p)
                                                
                                                        If desc Then
                                                            .SortDes
                                                        Else
                                                            .Sort
                                                        End If
                                                    .SetTextCompare
                                                ElseIf p <> 0 Then
                                                    
                                                    .useclid = 0
                                                    .Sort
                                                Else
                                                
                                                    .useclid = 0
                                                    .SortDes
                                                End If
                                             End With
                                        End If
                                    Else
                                        MissPar
                                        Exit Function
                                    End If

                            Else
                                With uHandler.objref
                                    If desc Then
                                     .useclid = 0
                                     .SortDes
                                    Else
                                    .useclid = 0
                                        .Sort
                                    End If
                                End With
                            End If
                            
                            ProcSort = True
                            ElseIf uHandler.t1 = 3 Then
                            Set pppp = uHandler.objref
useArray:
                            If FastSymbol(rest$, ",") Then
                                  If pppp.bDnum = 2 Then
                                            pppp.GetDnum 0, ML, i
                                            pppp.SerialItem x1, 1, 6
                                            
                                            Set mm = New mArray
                                            
                                            mm.PushDim x1 * 2 + 2, (0)
                                            mm.PushEnd
                                            y1 = 0
                                            Do
                                            If IsExp(basestack, rest$, p) Then
                                            If y1 > 1 And (y1 And 1) = 1 And desc Then
                                                mm.item(y1) = (CLng(p) And 2) + 1 - (CLng(p) And 1) = 1
                                            Else
                                                mm.item(y1) = p
                                                End If
                                            Else
                                                mm.item(y1) = 0
                                            End If
                                            y1 = y1 + 1
                                            Loop While FastSymbol(rest$, ",")
                                            If y1 Mod 2 = 1 Then
                                            y1 = y1 + 1
                                            End If
                                            mm.item(0) = mm.item(0) + i
                                            mm.item(1) = mm.item(1) + i
                                            If y1 = 2 Then
                                            If desc Then
                                            mm.item(2) = 0
                                            mm.item(3) = 1
                                            y1 = 4
                                            End If
                                            
                                            End If
                                            mm.StartResize
                                            mm.PushDim y1, (0)
                                            mm.PushEnd
                                            mm.ExportArrayNow
                                            pppp.SortColumns mm.refArray
                                  Else
                                            If IsExp(basestack, rest$, p) Then
                                                x1 = p
                                            Else
                                                x1 = -1
                                            End If
                                            y1 = -1
                                            If FastSymbol(rest$, ",") Then
                                            If IsExp(basestack, rest$, p) Then
                                                y1 = p
                                            End If
                                        If desc Then
                                           pppp.SortDesTuple x1, y1
                                        Else
                                           pppp.SortTuple x1, y1
                                        End If
                                  End If
                                  End If
                                Else
                                        If desc Then
                                           pppp.SortDesTuple
                                        Else
                                           pppp.SortTuple
                                        End If
                                End If
                            ProcSort = True
                            Else
                            MyEr "Expected Inventory", "Περίμενα Κατάσταση"
                            End If
                Else
                   MyEr "Expected Inventory", "Περίμενα Κατάσταση"
                End If
            Else
                   MissFuncParameterStringVar
            End If
    Exit Function
    ElseIf y1 = 5 Then
            If neoGetArray(basestack, s$, pppp) Then
                If pppp.Arr Then
                If FastSymbol(rest$, ")") Then
                GoTo useArray
                End If
                End If
                 If Not NeoGetArrayItem(pppp, basestack, s$, i, rest$) Then Exit Function
                 If pppp.IsStringItem(i) Or pppp.itemnumeric(i) Then
                 MissingArrayOrInventory
                 Exit Function
                 End If
                 If pppp.item(i) Is Nothing Then
                 MissingArrayOrInventory
                 Exit Function
                 End If
                 If TypeOf pppp.itemObject(i) Is mHandler Then
                    Set uHandler = pppp.item(i)
                    GoTo usehandler
                 ElseIf TypeOf pppp.itemObject(i) Is mArray Then
                    Set pppp = pppp.item(i)
                    GoTo useArray
                 End If
                End If
                MissingArrayOrInventory
                Exit Function
                
    ElseIf y1 = 6 Then
                If neoGetArray(basestack, s$, pppp) Then
                If pppp.Arr Then
                If FastSymbol(rest$, ")") Then
                GoTo useArray
                End If
                End If
                
                 If Not NeoGetArrayItem(pppp, basestack, s$, i, rest$) Then Exit Function
                 If pppp.IsStringItem(i) Or pppp.itemnumeric(i) Then
                 MissingDocOrArrayOrInventory
                 Exit Function
                 End If
                 If pppp.item(i) Is Nothing Then
                 MissingDocOrArrayOrInventory
                 Exit Function
                 End If
                 If TypeOf pppp.itemObject(i) Is mHandler Then
                    Set uHandler = pppp.item(i)
                    GoTo usehandler
                 ElseIf TypeOf pppp.itemObject(i) Is mArray Then
                    Set pppp = pppp.item(i)
                    GoTo useArray
                 End If
                Else
                MissingDoc
                Exit Function
                End If
    End If
    If FastSymbol(rest$, ",") Then
        If Not IsExp(basestack, rest$, sX) Then    ' FROM
            If FastSymbol(rest$, ",") Then
            sX = 1
                GoTo sort2
            Else
                MissNumExpr
                Exit Function
            End If
        End If
    Else
 sX = 1
    End If
    
     If FastSymbol(rest$, ",") Then
sort2:
        If Not IsExp(basestack, rest$, p) Then   ' TO
            If FastSymbol(rest$, ",") Then
            x1 = 0
                GoTo sort3
            Else
                MissNumExpr
                Exit Function
            End If
        End If
        x1 = CLng(p)
     Else
        x1 = 0    ' TO THE LAST
    End If
         If FastSymbol(rest$, ",") Then
sort3:
        If Not IsExp(basestack, rest$, sY) Then   ' TO
                        MissNumExpr
                        Exit Function
        End If
        ML = CLng(sY)
     Else
        ML = 1   ' KEYSTART
    End If
        If y1 = 3 Then
            If GetVar(basestack, s$, i) Then
                If Typename(var(i)) = doc Then
                            If desc Then
                            var(i).SortDocDes ML, CLng(sX), x1
                            Else
                            var(i).SortDoc ML, CLng(sX), x1
                            End If
                            ProcSort = True '*******************************************
                Else
                   MissingDoc
                End If
            Else
                   MissFuncParameterStringVar
            End If
        ElseIf y1 = 6 Then
                    If pppp.ItemType(i) = doc Then
                                If desc Then
                                pppp.item(i).SortDocDes ML, CLng(sX), x1
                                Else
                                pppp.item(i).SortDoc ML, CLng(sX), x1
                                End If
                                ProcSort = True  '*****************************************
                      Else
                                MissingDoc
                      End If
        Else
                    MissPar
        End If

End Function

Function GetData(bstack As basetask, rest$, obj As Object) As Boolean
Dim s$, p As Variant, usehandler As mHandler ', vvl As Variant, photo As Object
Set obj = New mStiva
Dim soros As mStiva
Set soros = obj
GetData = True
Do
    If FastSymbol(rest$, "!") Then
                If IsExp(bstack, rest$, p) Then
                    If bstack.lastobj Is Nothing Then
                        soros.DataValLong p
                    ElseIf TypeOf bstack.lastobj Is mHandler Then
                        Set usehandler = bstack.lastobj
                        If TypeOf usehandler.objref Is mStiva Then
                            soros.MergeBottom usehandler.objref
                        ElseIf TypeOf usehandler.objref Is mArray Then
                            soros.MergeBottomCopyArray usehandler.objref
                        Else
                            GetData = False
                            MyEr "Expected Stack Object or Array after !", "Περίμενα αντικείμενο Σωρό ή πίνακα μετά το !"
                            Set bstack.lastobj = Nothing
                            Exit Function
                        End If
                        Set usehandler = Nothing
                        Set bstack.lastobj = Nothing
                    ElseIf TypeOf bstack.lastobj Is mArray Then
                        soros.MergeBottomCopyArray bstack.lastobj
                        Set bstack.lastobj = Nothing
                    End If
                        
                     End If
                ElseIf IsExp(bstack, rest$, p) Then
                  If bstack.lastobj Is Nothing Then
                      soros.DataVal p
                 Else
                   If TypeOf bstack.lastobj Is mStiva Then
                   Set bstack.Sorosref = bstack.lastobj
                   ElseIf TypeOf bstack.lastobj Is VarItem Then
                    soros.DataObjVaritem bstack.lastobj
                      Else
                      If TypeOf bstack.lastobj Is Group Then
                      bstack.lastobj.ToDelete = False
                      End If
                      soros.DataObj bstack.lastobj
                      Set bstack.lastpointer = Nothing
                      
                    End If
                      Set bstack.lastobj = Nothing
                End If
        ElseIf IsStrExp(bstack, rest$, s$, Len(bstack.tmpstr) = 0) Then
                If bstack.lastobj Is Nothing Then
                        soros.DataStr s$
                Else
                        soros.DataObj bstack.lastobj
                        Set bstack.lastobj = Nothing
                          Set bstack.lastpointer = Nothing
                End If
        Else
                GetData = LastErNum1 = 0
                Exit Do
        End If
        If Not FastSymbol(rest$, ",") Then Exit Do
        
Loop
End Function
Function MyWith(bstack As basetask, rest$, Lang As Long) As Boolean
Dim i As Long, ss$, s$, pppp As mArray, pa$, x1 As Long, id, p
MyWith = True
x1 = Abs(IsLabel(bstack, rest$, s$))
If x1 = 1 Or x1 = 3 Then
    If GetVar(bstack, s$, i) Then
            If Not MyIsObject(var(i)) Then BadObjectDecl:  Exit Function
            If Not var(i) Is Nothing Then  ''???
                   Do While FastSymbol(rest$, ",")
                    If Not ss$ = vbNullString Then ss$ = vbNullString
                    If IsExp(bstack, rest$, id) Then
                                        
                    ProcProperty bstack, var(), i, ss$, rest$, Lang, , CLng(id)
                    If LastErNum <> 0 Then MyWith = False: Exit Do
                    ElseIf IsStrExp(bstack, rest$, ss$, Len(bstack.tmpstr) = 0) Then
                    On Error Resume Next

                      ProcProperty bstack, var(), i, ss$, rest$, Lang
                      If Err.Number > 0 Then
                      MyEr "Property " + ss$ + " problem", "Πρόβλημα με ιδιότητα " + ss$
                      Err.Clear
                      MyWith = False
                    Exit Do
                      End If
                      MyWith = Err = 0
                      Err.Clear
                    Else
                    MissStringNumber
                    MyWith = False
                    Exit Do
                    End If
                    Loop
                    Exit Function
            Else
                    BadObjectDecl
                    Exit Function
            End If
    Else
    
     Nosuchvariable s$
    End If
ElseIf x1 = 5 Or x1 = 6 Then
  If neoGetArray(bstack, s$, pppp) Then
    If NeoGetArrayItem(pppp, bstack, s$, i, rest$) Then
      Do While FastSymbol(rest$, ",")
                    If Not ss$ = vbNullString Then ss$ = vbNullString
                    If IsExp(bstack, rest$, id) Then
                    ProcPropertyArray bstack, pppp, i, ss$, rest$, Lang, MyWith
                    If LastErNum <> 0 Then MyWith = False: Exit Do
                    
                    ElseIf IsStrExp(bstack, rest$, ss$, Len(bstack.tmpstr) = 0) Then
                    If TypeOf pppp.itemObject(i) Is GuiM2000 Then
                      If UCase(ss$) = "VISIBLE" Then ss$ = "TrueVisible"
                      End If
                  ProcPropertyArray bstack, pppp, i, ss$, rest$, Lang, MyWith
                  If LastErNum1 = -1 Then
                  Exit Do
                  End If
                  If Err.Number > 0 Then
                      MyEr "Property " + ss$ + " problem", "Πρόβλημα με ιδιότητα " + ss$
                      Err.Clear
                      MyWith = False
                    Exit Do
                      End If
                      MyWith = Err = 0
                      Err.Clear
                    Else
                    MissStringNumber
                    MyWith = False
                    Exit Do
                    End If
                    Loop
                    Exit Function
           
            
     End If
  Else
      MissingObj
  End If
Else
MissingObj
End If
End Function

Function MyWrite(basestack As basetask, rest$, Lang As Long) As Boolean
Dim p As Variant, s$, it As Long, par As Boolean, i As Long, skip As Boolean
If IsLabelSymbolNew(rest$, "ΜΕ", "WITH", Lang) Then
    If IsStrExp(basestack, rest$, s$) Then
        csvsep$ = Left$(s$, 1)
        If FastSymbol(rest$, ",") Then If IsStrExp(basestack, rest$, s$) Then csvDec$ = Left$(s$, 1): MyWrite = True
        csvuseescape = False
        If FastSymbol(rest$, ",") Then If IsExp(basestack, rest$, p, , True) Then csvuseescape = CBool(p): MyWrite = True
        If FastSymbol(rest$, ",") Then MyWrite = False: If IsExp(basestack, rest$, p, , True) Then cleanstrings = CBool(p): MyWrite = True
    End If
    Exit Function
End If
MyWrite = True
If csvsep$ = vbNullString Then csvsep$ = ","
If IsLabelSymbolNew(rest$, "ΔΕΚΑΕΞ", "HEX", Lang) Then it = 1
If FastSymbol(rest$, "#") Then

    MyWrite = False
    If IsExp(basestack, rest$, p, , True) Then
    skip = p < 0
        On Error Resume Next
    If skip Then
    Dim Scr As Object, prive As basket, basketcode As Long
    Set Scr = basestack.Owner
    basketcode = GetCode(Scr)
    prive = players(GetCode(Scr))
    Else
    i = CLng(MyMod(p, 512))
    If FKIND(i) = FnoUse Or FKIND(i) = Finput Or FKIND(i) = Frandom Then MyEr "Wrong File Handler", "Λάθος Χειριστής Αρχείου": MyWrite = False: Exit Function
    End If
        
        par = False
        Do While FastSymbol(rest$, ",")

            If IsExp(basestack, rest$, p, , True) Then
            s$ = LTrim$(Str$(p))
            If it Then
            Else
             If Left$(s$, 1) = "." Then
                s$ = "0" + s$
                ElseIf Left$(s$, 2) = "-." Then s$ = "-0" + Mid$(s$, 2)
                End If
            End If
                If par Then
                    If skip Then
                        PlainBaSket Scr, prive, csvsep$
                    ElseIf uni(i) Then
                            putUniString i, csvsep$
                    Else
                            putANSIString i, csvsep$
                           ' Print #i, ",";
                    End If
                End If
                If skip Then
                        If it Then
                            PlainBaSket Scr, prive, PACKLNGUnsign$(p)
                        Else
                            If Len(csvDec$) = 0 Then
                                PlainBaSket Scr, prive, s$
                            Else
                                PlainBaSket Scr, prive, Replace(Str$(p), ".", csvDec$)
                            End If
                        End If
                ElseIf uni(i) Then
                        If it Then
                            putUniString i, PACKLNGUnsign$(p)
                        ElseIf Len(csvDec$) = 0 Then
                            putUniString i, s$
                        Else
                            putUniString i, Replace(s$, ".", csvDec$)
                        End If
                Else
                        If Len(csvDec$) = 0 Then
                            putANSIString i, s$
                        Else
                            putANSIString i, Replace(s$, ".", csvDec$)
                        End If
                        
                        If Err.Number > 0 Then Exit Function
                End If
            ElseIf IsStrExp(basestack, rest$, s$, Len(basestack) = 0) Then
            If csvuseescape Then s$ = StringToEscapeStr(s$, False)
                If par Then
                    If skip Then
                        PlainBaSket Scr, prive, csvsep$
                    ElseIf uni(i) Then
                        putUniString i, csvsep$
                    Else
                        putANSIString i, csvsep$

                    End If
                End If
                If Not cleanstrings Then s$ = Replace$(s$, Chr(34), Chr(34) + Chr(34))
                If skip Then
                    PlainBaSket Scr, prive, Chr(34) + s$ + Chr(34)
                ElseIf uni(i) Then
                    If cleanstrings Then
                        putUniString i, s$
                    Else
                        putUniString i, Chr(34) + s$ + Chr(34)
                    End If
                Else
                    If cleanstrings Then
                        putANSIString i, s$
                    Else
                        putANSIString i, Chr(34) + s$ + Chr(34)
                    End If
                    End If
                    If Err.Number > 0 Then Exit Function
            Else
                
                    Exit Function
            End If
            par = True
            If skip Then players(basketcode) = prive
        Loop
        If skip Then
            crNew basestack, prive
        ElseIf uni(i) Then
            putUniString i, vbCrLf
        Else
            putANSIString i, vbCrLf
        End If
        MyWrite = True

    End If
End If
If skip Then players(basketcode) = prive

End Function
Function IsPureLabel(a$, R$) As Long 'ok
Dim rr&, one As Boolean, c$
R$ = vbNullString
If a$ = vbNullString Then IsPureLabel = 0: Exit Function

a$ = NLtrim$(a$)
    Do While Len(a$) > 0
    c$ = myUcase(Left$(a$, 1))
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 46 '"."
            If one Then
            Exit Do
            Exit Do
            ElseIf R$ <> "" Then
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1
            Else
            IsPureLabel = 0
            Exit Function
            End If
        Case 65 To 90, 913 To 937, 902, 904, 906, 908, 905, 911, 910, 962 '"A" To "Z", "Α" To "Ω", "’", "Έ", "Ί", "Ό", "Ή", "Ώ", "Ύ", "ς"
            If one Then
            Exit Do
            Else
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        Case 48 To 57, 95 '"0" To "9", "_"
           If one Then
            Exit Do
            ElseIf R$ <> "" Then
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit Do
            End If
            
        Case 36 ' "$"
            If one Then Exit Do
            If R$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If R$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 40 ' "("
            If R$ <> "" Then
                            If Mid$(a$, 2, 2) = ")@" Then
                                    R$ = R$ & "()."
                                  
                                 a$ = Mid$(a$, 4)
                               Else
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit Do
                                       End Select
                                       R$ = R$ & Left$(a$, 1)
                                       a$ = Mid$(a$, 2)
                                   Exit Do
                            
                          End If
               Else
                        Exit Do
            
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
            R$ = R$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        End If

    Loop
    IsPureLabel = rr&
   a$ = NLtrim$(a$)

End Function
Function GetRealRow(dq As Object) As Long
Dim mybasket As basket
mybasket = players(GetCode(dq))
If mybasket.lastprint Then
GetXYb dq, mybasket, mybasket.curpos, mybasket.currow
End If

 GetRealRow = mybasket.currow
End Function

Function GetRealPos(dq As Object) As Long
Dim mybasket As basket, oldx
mybasket = players(GetCode(dq))
If mybasket.lastprint Then
oldx = dq.currentX
dq.currentX = dq.currentX + mybasket.Xt - dv15
GetXYb dq, mybasket, mybasket.curpos, mybasket.currow
dq.currentX = oldx
End If

 GetRealPos = mybasket.curpos
End Function

Sub ProcWindow(bstack As basetask, rest$, Scr As Object, ifier As Boolean)
Dim x1 As Long, y1 As Long, p As Variant, useScreen As Long

If Scr.Name = "GuiM2000" Then
    Else
If Scr.Name = "Form1" Then
    DisableTargets q(), -1
ElseIf Scr.Name = "DIS" Then
    DisableTargets q(), 0
ElseIf Scr.Name = "dSprite" Then
    DisableTargets q(), val(Scr.Index)
End If
End If
With players(GetCode(Scr))
If .double Then SetNormal Scr
        If IsExp(bstack, rest$, p) Then
            .SZ = p
            If .SZ < 4 Then .SZ = 4
            If FastSymbol(rest$, ",") Then
                If IsExp(bstack, rest$, p) Then
                    x1 = CLng(p)
again:
                    If x1 >= 0 And x1 <= DisplayMonitorCount - 1 And Scr.Name = "DIS" Then
                    Console = x1
                    
                    If Not Form1.WindowState = 0 Then Form1.WindowState = 0: Form1.Refresh
                    
                    If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top
                    If IsWine Then
                        Form1.move ScrInfo(Console).Left, ScrInfo(Console).top
                        If Form1.Width = ScrInfo(Console).Width Then
                        Form1.Width = ScrInfo(Console).Width - 1
                        Else
                        Form1.Width = ScrInfo(Console).Width
                        End If
                        Form1.Height = ScrInfo(Console).Height
                        Form1.move ScrInfo(Console).Left, ScrInfo(Console).top
                    
                    Else
                        Form1.move ScrInfo(Console).Left, ScrInfo(Console).top, ScrInfo(Console).Width, ScrInfo(Console).Height
                    End If
                    FrameText Scr, .SZ, CLng(Form1.Width), CLng(Form1.Height), .Paper
                    players(-1).MAXXGRAPH = Form1.Width ' .MAXXGRAPH
                    players(-1).MAXYGRAPH = Form1.Height '.MAXYGRAPH
                    'If Scr.Name = "Form1" Then SetText Scr
                    Exit Sub
                    
                ElseIf x1 > 3000 Then
                    y1 = CLng(x1 * ScrInfo(Console).Height / ScrInfo(Console).Width)
                ElseIf Scr.Name <> "DIS" Then
                    
                Else
                    x1 = 0
                    GoTo again
                End If
        End If
        If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then y1 = CLng(p)
    End If
    If Scr.Name = "GuiM2000" Then
        useScreen = FindMonitorFromMouse
        Set Scr.Picture = LoadPicture("")
               If FastSymbol(rest$, ";") Then 'CENTER
                    FrameText Scr, .SZ, x1, y1, .Paper, True
                    Scr.move (ScrInfo(useScreen).Width - .MAXXGRAPH) / 2, (ScrInfo(useScreen).Height - .MAXYGRAPH) / 2
                Else
                    If x1 = 0 Then x1 = 14000
                    If y1 = 0 Then y1 = 6000
                    Scr.move Scr.Left, Scr.top, x1, y1
                    FrameText Scr, .SZ, 0, 0, .Paper, True
                            
            End If
            SetTextSZ Scr, .SZ
    ElseIf Scr.Name = "dSprite" Then
            RsetRegion Scr
            Set Scr.Picture = LoadPicture("")
            If FastSymbol(rest$, ";") Then 'CENTER
                        FrameText Scr, .SZ, x1, y1, .Paper, True
                        Scr.move (ScrInfo(Console).Width - .MAXXGRAPH) / 2, (ScrInfo(Console).Height - .MAXYGRAPH) / 2
            Else
                        Scr.move Scr.Left, Scr.top, x1, y1

                        FrameText Scr, .SZ, 0, 0, .Paper, True
                        
            End If
            SetTextSZ Scr, .SZ
    ElseIf Scr.Name = "Emf" Then
        FastSymbol rest$, ";"
        If x1 <= 0 Then x1 = Scr.Width
        If y1 <= 0 Then y1 = Scr.Height
        Scr.move Scr.Left, Scr.top, x1, y1
        FrameText Scr, .SZ, 0, 0, .Paper, True
        
        Exit Sub
    Else
        If FastSymbol(rest$, ";") Then 'CENTER
            Form1.WindowState = 0
            If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top
            If Scr.Name = "Form1" Then
            If x1 = 0 Then x1 = ScrInfo(Console).Width
            If y1 = 0 Then y1 = ScrInfo(Console).Height
            
            .MAXXGRAPH = x1
            .MAXYGRAPH = y1
            If Scr.Visible Then
                ClearScr Scr, .Paper
            Else
                Scr.backcolor = .Paper
            End If
            Else
            FrameText Scr, .SZ, x1, y1, .Paper
            End If
                    If IsWine Then
                        If .MAXXGRAPH = ScrInfo(Console).Width Then
                        Form1.Width = .MAXXGRAPH - 1
                        Else
                        Form1.Width = .MAXXGRAPH
                        End If
                        Form1.Height = .MAXYGRAPH
                        Form1.move ((ScrInfo(Console).Width - 1) - Form1.Width) / 2 + ScrInfo(Console).Left, ((ScrInfo(Console).Height - 1) - Form1.Height) / 2 + ScrInfo(Console).top
                    Else

            
            Form1.move ScrInfo(Console).Left + (ScrInfo(Console).Width - .MAXXGRAPH) / 2, ScrInfo(Console).top + (ScrInfo(Console).Height - .MAXYGRAPH) / 2, .MAXXGRAPH, .MAXYGRAPH
 
            End If
            Form1.Up
            If Scr.Name <> "Form1" Then
            scrMove00 Scr
            Else
            Scr.currentX = 0
            Scr.currentY = 0
            .curpos = 0
            .currow = 0
            SetText Scr

            End If
            players(-1).MAXXGRAPH = Form1.Width ' .MAXXGRAPH
            players(-1).MAXYGRAPH = Form1.Height '.MAXYGRAPH
            If Form1.Visible Then
                 If form5iamloaded Then
                     Form1.Show , Form5
                 Else
                     Form1.Show
                 End If
             End If
        Else
            Form1.WindowState = 0
            If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top
            If Scr.Name = "Form1" Then
            If x1 = 0 Then x1 = ScrInfo(Console).Width
            If y1 = 0 Then y1 = ScrInfo(Console).Height
            .MAXXGRAPH = x1
            .MAXYGRAPH = y1
            If Scr.Visible Then
                ClearScr Scr, .Paper
            Else
                Scr.backcolor = .Paper
            End If
            Else
            FrameText Scr, .SZ, x1, y1, .Paper
            End If
            If IsWine Then
                If .MAXXGRAPH = ScrInfo(Console).Width Then
                Form1.Width = .MAXXGRAPH - 1
                Else
                Form1.Width = .MAXXGRAPH
                End If
                Form1.Height = .MAXYGRAPH
                Form1.move ScrInfo(Console).Left, ScrInfo(Console).top
            Else
                Form1.move ScrInfo(Console).Left, ScrInfo(Console).top, .MAXXGRAPH, .MAXYGRAPH
            End If
            players(-1).MAXXGRAPH = Form1.Width ' .MAXXGRAPH
            players(-1).MAXYGRAPH = Form1.Height '.MAXYGRAPH
            Form1.Up
            If Scr.Name = "Form1" Then
            Scr.currentX = 0
            Scr.currentY = 0
            .curpos = 0
            .currow = 0
            SetText Scr
            End If
            scrMove00 Scr
            If Form1.Visible Then
                If form5iamloaded Then
                    Form1.Show , Form5
                Else
                    Form1.Show
                End If
            End If
    End If
End If
Else
ifier = False
Exit Sub
End If
''SleepWait 4
End With
End Sub



Sub scrMove00(Scr As Object, Optional ByVal w As Variant, Optional ByVal H As Variant)
If TypeOf Scr Is Form Then
If IsMissing(w) And IsMissing(H) Then
Scr.move ScrInfo(Console).Left, ScrInfo(Console).top
Else
Scr.move ScrInfo(Console).Left, ScrInfo(Console).top, w, H
End If
Else
 If IsMissing(w) And IsMissing(H) Then
Scr.move 0, 0
Else
Scr.move 0, 0, w, H
End If
End If
End Sub
Function MySounds(bstack As basetask, rest$, Lang As Long) As Boolean
MySounds = True
Dim mDir As recDir, s$, ss$, frm$
Set mDir = New recDir
mDir.IncludedFolders = False
mDir.Nofiles = False
mDir.TopFolder = mcd
mDir.SortType = Abs(FastSymbol(rest$, "!"))
frm$ = ExtractNameOnly(mDir.Dir2$(mcd, "WAV", False), True)
s$ = vbNullString
ss$ = vbNullString
Do While frm$ <> ""
s$ = frm$
If ss$ <> "" Then ss$ = ss$ & ", " & s$ Else ss$ = s$
s$ = mDir.Dir2
If s$ <> "" Then frm$ = ExtractNameOnly(s$, True) Else frm$ = vbNullString
Loop
If Lang Then
ss$ = "Sounds: " & ss$
Else
ss$ = "Ήχοι: " & ss$
End If
Set mDir = Nothing
RepPlain bstack, bstack.Owner, ss$
End Function
Function ProcCat(bstack As basetask, rest$, Lang As Long) As Boolean
Dim aDir As New recDir, ss$, s$, pa$, frm$, par As Boolean, i As Long, Col As Long
aDir.IncludedFolders = True
aDir.Nofiles = True
aDir.TopFolder = mcd
aDir.LevelStop = 1
aDir.SortType = Abs(FastSymbol(rest$, "!")) + Abs(FastSymbol(rest$, "!"))
ProcCat = True
s$ = vbNullString
pa$ = vbNullString

par = Lang = 1
i = FastSymbol(rest$, "+")
If FastSymbol(rest$, "*") Then
ss$ = "*"
ElseIf Not IsStrExp(bstack, rest$, ss$) Then
ss$ = vbNullString
Else
ss$ = myUcase(ss$)
End If
''stac1 = VbNullString

If InStr(ss$, "?") > 0 Or InStr(ss$, "*") > 0 Then
aDir.Pattern = ss$
frm$ = mylcasefILE$(aDir.Dir2$(mcd, "", False))

Else
frm$ = mylcasefILE$(aDir.Dir2$(mcd, ss$, False))

End If

If i = False Then
    If frm$ <> "" Then
        If par Then
            pa$ = "Folders " & ss$ & ": "
        Else
            pa$ = "Κατάλογοι " & ss$ & ": "
        End If
    End If
End If
ss$ = vbNullString
Col = Len(mcd)
Do While frm$ <> ""
MyDoEvents
s$ = Mid$(frm$, Col + 2)
If s$ <> "" Then
If i Then
Form1.List1.additem s$
Else
If InStr(s$, " ") > 0 Then s$ = Chr(34) + s$ + Chr(34)
If ss$ <> "" Then ss$ = ss$ & ", " & s$ Else ss$ = s$
End If
End If
s$ = aDir.Dir2
If s$ <> "" Then frm$ = s$ Else frm$ = vbNullString
Loop
If i = False Then RepPlain bstack, bstack.Owner, pa$ + ss$

End Function
Function ProcFiles(bstack As basetask, rest$, Lang As Long) As Boolean
Dim aDir As New recDir, ss$, pa$, s$, par As Boolean, stac1 As String, frm$, i As Long
Dim addext As Boolean, addpath As Boolean, p As Variant
ProcFiles = True
aDir.IncludedFolders = False
aDir.Nofiles = False
aDir.TopFolder = mcd
aDir.SortType = Abs(FastSymbol(rest$, "!")) + Abs(FastSymbol(rest$, "!"))

s$ = vbNullString
pa$ = vbNullString

par = Lang = 1
i = FastSymbol(rest$, "+")
If FastSymbol(rest$, "*") Then
ss$ = "*"
ElseIf Not IsStrExp(bstack, rest$, ss$) Then
ss$ = "TXT"
Else
ss$ = myUcase(ss$)
End If
stac1 = vbNullString
If FastSymbol(rest$, ",") Then
' SEARCH INSIDE FILE
If Not IsStrExp(bstack, rest$, stac1$) Then
If Not IsExp(bstack, rest$, p) Then
SyntaxError: Exit Function
Else
Select Case p
Case 1
addext = True
Case 2
addext = True
addpath = True
Case Else
MyEr "Not defined yet", "δεν έχει οριστεί ακόμα"
ProcFiles = False
Exit Function
End Select
End If
End If
End If
If i <> 0 Then addpath = False
If InStr(ss$, "?") > 0 Or InStr(ss$, "*") > 0 Then
aDir.Pattern = ss$
frm$ = mylcasefILE$(aDir.Dir2$(mcd, "", False))
'frm$ = mylcasefILE$(Dir(mcd + mylcasefILE(sS$)))
Else
frm$ = mylcasefILE$(aDir.Dir2$(mcd, ss$, False))
'frm$ = mylcasefILE$(Dir(mcd & "*." & sS$))
End If

If i = False Then
    If frm$ <> "" Then
        If par Then
            If stac1 <> "" Then
            pa$ = "Files " & Replace(ss$, "|", ", ") & " with text " & Chr(34) + stac1 + Chr(34) & " :  "
            Else
            pa$ = "Files " & Replace(ss$, "|", ", ") & ": "
            End If
        Else
            If stac1 <> "" Then
            pa$ = "Αρχεία " & Replace(ss$, "|", ", ") & " με κείμενο " & Chr(34) + stac1 + Chr(34) & " :  "
            
            Else
            pa$ = "Αρχεία " & Replace(ss$, "|", ", ") & ": "
            End If
        End If
    End If
Else
   ' Form1.List2.clear
End If
If InStr(ss$, "|") > 0 Then addext = True
ss$ = vbNullString
Do While frm$ <> ""
MyDoEvents
If NOEXECUTION Then Exit Do
If Right$(" " & aDir.Pattern, 1) <> "*" And InStr(" " & aDir.Pattern, "|") = 0 Then
    If addext Then
        If addpath Then
            s$ = Included(mcd + frm$, stac1)
            If LastErNum1 <> 0 Then Exit Function
            If s$ <> "" Then s$ = mcd + s$
        Else
            s$ = ExtractName(Included(mcd + frm$, stac1), True)
            If LastErNum1 <> 0 Then Exit Function
        End If
    Else
        s$ = ExtractNameOnly(Included(mcd + frm$, stac1), True)
    If LastErNum1 <> 0 Then Exit Function
    End If
Else
If addpath Then
            s$ = Included(mcd + frm$, stac1)
            If LastErNum1 <> 0 Then Exit Function
            If s$ <> "" Then s$ = mcd + s$
Else
   s$ = Included(mcd + frm$, stac1)
   If LastErNum1 <> 0 Then Exit Function
   End If
End If
If s$ <> "" Then
If i Then
If Right$(" " & aDir.Pattern, 1) <> "*" Then
Form1.List1.additem s$
Else
Form1.List1.additem s$
End If
Else

If ss$ <> "" Then ss$ = ss$ & ", " & s$ Else ss$ = s$
End If
End If
s$ = aDir.Dir2
If s$ <> "" Then frm$ = s$ Else frm$ = vbNullString
Loop
If i = False Then
RepPlain bstack, bstack.Owner, pa$ + ss$
End If
End Function
Sub RepPlain(bstack As basetask, Scr As Object, txt$)
Dim prive As Long
prive = GetCode(Scr)
If players(prive).curpos > 0 Then crNew bstack, players(prive)
wwPlain2 bstack, players(prive), txt$, Scr.Width, 100000, True
If players(prive).curpos > 0 Then crNew bstack, players(prive)
End Sub

Function MyModules(bstack As basetask, rest$, Lang As Long) As Boolean
Dim frm$, s$, pa$, ss$, mDir As recDir, showlocal As Boolean, i As Long, Filter$, filter2$
showlocal = FastSymbol(rest$, "?")
LastErNum = 0
If IsStrExp(bstack, rest$, Filter$) Then
Filter$ = myUcase(Filter$, True)
If Left$(Filter$, 1) <> "*" Then Filter$ = Filter$ + "*"
If FastSymbol(rest$, ",") Then
If IsStrExp(bstack, rest$, filter2$) Then
If filter2$ <> vbNullString Then filter2$ = myUcase(filter2$, True)
ElseIf LastErNum Then
    Exit Function
End If
End If
ElseIf LastErNum Then
    Exit Function
End If
MyModules = True
frm$ = subHash.ShowRev    ' Mid$(SubName$, 2)
s$ = vbNullString
pa$ = vbNullString

Do While ISSTRINGA(frm$, s$)

If pa$ <> "" Then
    If InStrRev(s$, ") ") > 0 Then
    ss$ = FixName(Left$(s$, InStrRev(s$, ")")))
    If Len(Filter$) > 0 Then
    If ss$ Like Filter$ Then
        If Len(filter2$) > 0 Then
            If InStr(myUcase(sbf(val(Mid$(s$, InStrRev(s$, ")") + 1))).sb, True), filter2$) > 0 Then
                pa$ = pa$ + ", " + ss$
            End If
        Else
            pa$ = pa$ + ", " + ss$
        End If
    End If
    Else
    pa$ = pa$ + ", " + ss$
    End If
    ElseIf InStrRev(s$, " ") > 0 Then
    ss$ = FixName(Left$(s$, InStrRev(s$, " ") - 1))
    If Len(Filter$) > 0 Then
        If ss$ Like Filter$ Then
            If Len(filter2$) > 0 Then
            If InStr(myUcase(sbf(val(Mid$(s$, InStrRev(s$, " ") + 1))).sb, True), filter2$) > 0 Then
                pa$ = pa$ + ", " + ss$
            End If
            Else
                pa$ = pa$ + ", " + ss$
            End If
        End If
    Else
    pa$ = pa$ + ", " + ss$
    End If
    End If
Else
    If InStrRev(s$, ") ") > 0 Then
    pa$ = FixName(Left$(s$, InStrRev(s$, ")")))
    If Len(Filter$) > 0 Then
        If Not pa$ Like Filter$ Then
        pa$ = vbNullString
        ElseIf Len(filter2$) > 0 Then
            If Not InStr(myUcase(sbf(val(Mid$(s$, InStrRev(s$, ")") + 1))).sb, True), filter2$) > 0 Then
                pa$ = vbNullString
            End If
        End If
    End If
    ElseIf InStrRev(s$, " ") > 0 Then
    pa$ = FixName(Left$(s$, InStrRev(s$, " ") - 1))
    If Len(Filter$) > 0 Then
        If Not pa$ Like Filter$ Then
            pa$ = vbNullString
        ElseIf Len(filter2$) > 0 Then
            If Not InStr(myUcase(sbf(val(Mid$(s$, InStrRev(s$, " ") + 1))).sb, True), filter2$) > 0 Then
                pa$ = vbNullString
            End If
        End If
    End If
    End If
End If
Loop
ss$ = vbNullString
If pa$ <> "" Or Len(Filter$) > 0 Or Len(filter2$) > 0 Then
If here$ <> vbNullString Then If showlocal Then GoTo ponly
If pa$ <> "" Then
If Lang Then

pa$ = "In Memory: " & pa$ & vbCrLf & "        Use REMOVE to remove the right most, EDIT module_name to edit"
Else
pa$ = "Στη Μνήμη: " & pa$ & vbCrLf & "        Με τη ΔΙΑΓΡΑΦΗ θα σβήσεις το τελευταίο, με ΣΥΓΓΡΑΦΗ ή Σ όνομα_τμήματος θα γράψεις"
End If
End If
If here$ = vbNullString Then If showlocal Then GoTo ponly
End If


Set mDir = New recDir
mDir.IncludedFolders = False
mDir.Nofiles = False
mDir.TopFolder = mcd
mDir.SortType = Abs(FastSymbol(rest$, "!"))
frm$ = ExtractNameOnly(mDir.Dir2$(mcd, "GSB|GSB1", False), True)

If frm$ <> "" Then
If pa$ <> "" Then pa$ = pa$ & vbCrLf
If Lang Then
pa$ = pa$ & "On Disk: "
Else
pa$ = pa$ & "Στον Δίσκο: "
End If
End If
ss$ = vbNullString
Do While frm$ <> ""
s$ = frm$
If ss$ <> "" Then ss$ = ss$ & ", " & s$ Else ss$ = s$
s$ = mDir.Dir2
If s$ <> "" Then frm$ = ExtractNameOnly(s$, True) Else frm$ = vbNullString

Loop
Set mDir = Nothing
pa$ = pa$ & ss$
If ss$ <> "" Then
If here$ = vbNullString Then
If Lang Then
pa$ = pa$ & vbCrLf + Replace$("        Use LOAD 'module_name' to load, EDIT 'module_name.gsb' to edit on disk", "'", Chr(34))
If IsSupervisor Then pa$ = pa$ & ", WIN DIR$ for folders tasks"
Else
pa$ = pa$ & vbCrLf + Replace$("        Με ΦΟΡΤΩΣΕ ονομα_τμηματος φορτώνεις στη μνήμη, με Σ ή ΣΥΓΓΡΑΦΗ 'ονομα_τμηματος.gsb' διορθώνεις στο δίσκο", "'", Chr(34))
If IsSupervisor Then pa$ = pa$ & ", με ΣΥΣΤΗΜΑ ΚΑΤ$ ανοίγεις τον κατάλογο με τα αρχεία για εργασίες"
End If
End If
End If
' PRINT ONLY
ponly:
RepPlain bstack, bstack.Owner, pa$

End Function
Function FixName(s$) As String
Dim a() As String
If SecureNames Then
a() = Split(s$, "].")
If UBound(a()) = 1 Then
a(0) = GetName(a(0))
FixName = Join(a(), ".")
Else
FixName = s$
End If
Else
FixName = s$
End If

End Function

Function MyIcon(basestack As basetask, rest$) As Boolean
Dim s$, aPic As StdPicture, p, anyObj As Object, ok As Boolean
On Error Resume Next
    Dim aa As New cDIBSection
    If IsStrExp(basestack, rest$, s$) Then
        If CFname$(s$) <> "" Then
        
       s$ = CFname$(s$)
        Set aPic = LoadPicture(GetDosPath(s$))
        If aPic Is Nothing Then Exit Function
             ok = True
             If FileLen(GetDosPath(s$)) > 4000 Then
                 aa.CreateFromPicture LoadMyPicture(GetDosPath(s$), True, &H3B3B3B)
                AskDIBicon$ = DIBtoSTR(aa)
                Else
                AskDIBicon$ = ""
                End If
        End If
    ElseIf IsExp(basestack, rest$, p) Then
    
    If TypeOf basestack.lastobj Is mHandler Then
            Set anyObj = GetObjFromHandler(basestack.lastobj, ok)
            If ok Then
            Dim mm As MemBlock
            If Typename(anyObj) = "MemBlock" Then
                
                Set mm = anyObj
                mm.SubType = 30
                aa.CreateFromPicture mm.GetStdPicture(, , &H3B3B3B)
                ' aa.CreateFromPicture mm.GetStdPicture1(, , &H3B3B3B)  ' smooth
                AskDIBicon$ = DIBtoSTR(aa)
                mm.SubType = 300
                Set aPic = mm.GetStdPicture()
                
                

            Else
                AskDIBicon$ = ""
                Set aPic = anyObj
            End If
                
                MyIcon = True
            Else
            MyEr "No icon find", "Δεν βρήκα εικόνα"
                Exit Function
            End If
        Else
            Set aPic = Form2.icon
            End If
    Else
            Set aPic = Form2.icon
    End If
    If Not UseMe Is Nothing Then
        PlaceIcon aPic
        Set Form3.icon = aPic
        Else
        Set Form3.icon = aPic
    End If
    Set Form1.icon = aPic

    Set basestack.lastobj = Nothing
    Set basestack.lastpointer = Nothing
    MyIcon = ok
End Function
Private Function GetObjFromHandler(vv As Object, ok As Boolean) As Object
Dim mh As mHandler, vIndex As Long
If TypeOf vv Is mHandler Then
    Set mh = vv
    If mh.indirect >= 0 Then
    vIndex = mh.indirect
    If var2used < mh.indirect Then
    MyEr "weak reference out of scope", "η αναφορά είναι εκτός σκοπού"
    ok = False
    Exit Function
    End If
    
    Set GetObjFromHandler = var(mh.indirect)
    ok = True
    Else
    Set GetObjFromHandler = mh.objref
    ok = True
    End If
End If
End Function

Function ProcTitle(basestack As basetask, rest$, Lang As Long) As Boolean
Dim p As Variant, s$
If Form1.Visible Then Form1.TrueVisible = True
If IsStrExp(basestack, rest$, s$) Then
    
If FastSymbol(rest$, ",") Then

    If IsExp(basestack, rest$, p, , True) Then
        
        If p = 0 Then
            If LenB(s$) = 0 And Not UseMe Is Nothing Then
                UseMe.Hide
               ' Form1.TrueVisible = False
                
'
            Else
                If Not UseMe Is Nothing Then UseMe.Show
                If Not ttl Then Form1.Visible = False
                
                PlaceCaption s$
            End If
            If Not ttl Or Not UseMe Is Nothing Then
                
                Form1.Visible = False
               ' Form1.TrueVisible = Form1.Visible
                If s$ <> "" Then
                    Form1.CaptionW = s$
                Else
                    Form1.CaptionW = "M2000"
                End If
            Else
               If Not Form3.WindowState = 1 Then
                        If Not Form3.Visible Then
                        Form1.TrueVisible = Form1.Visible
                        Form1.Visible = False
                        If LenB(s$) = 0 Then
                        'Unload Form3: ttl = False
                        End If
                        ProcTitle = True
                        Exit Function
                        End If
                        
                        Form3.move VirtualScreenWidth() + 2000, VirtualScreenHeight() + 2000
                        Form3.skiptimer = True
                        Form3.WindowState = 1
                        Form1.TrueVisible = Form1.TrueVisible Or Form1.Visible
                        Form1.Visible = False
                        ProcTitle = True
                        Exit Function
                        
                   End If
                   
        End If
                   
        Else
NormalState:
        
        
         
        If UseMe Is Nothing Then
         If LenB(s$) = 0 Or Not Form1.Visible Then ProcTitle = True: Exit Function
            If Not ttl Then
            Load Form3
            If Form3.WindowState = 1 Then Form3.skiptimer = False: Form3.WindowState = 0
            Form3.move VirtualScreenWidth() + 2000, VirtualScreenHeight() + 2000: ttl = True
            
            End If
            Form3.Timer1.Interval = 30
            Form3.Timer1.enabled = False
            Form3.CaptionW = s$
            Form1.CaptionW = vbNullString
                Form1.TrueVisible = True
                If Not Form3.WindowState = 0 Then
                    Form3.Visible = True
                    
                    Form3.WindowState = 0
    
                        End If
        Else
           
            PlaceCaption s$
'            UseMe.Show
        End If
          
             mywait basestack, 100
             
             
             End If
        ProcTitle = True
        Exit Function
    Else
        ProcTitle = False
    End If
Else
    If Not UseMe Is Nothing Then UseMe.Show
    GoTo NormalState
End If
Else
    If UseMe Is Nothing Then
            If ttl Then
                Unload Form3
                ttl = False
            End If
    Else
            UseMe.SetExtCaption ""
            'Form1.CaptionW = "M2000"
    End If
End If
ProcTitle = True
End Function
Public Function myRegister(tp$) As String
    strTemp = String(MAX_FILENAME_LEN, Chr$(0))
    GetTempPath MAX_FILENAME_LEN, StrPtr(strTemp)
    strTemp = mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1))
Dim i As Long
i = FreeFile
Open strTemp & "tmp." & tp$ For Output As i
Print #i, "test"
Close i
' found me
Dim rl$
rl$ = PCall(strTemp & "tmp." & tp$)
If rl$ <> "" Then
rl$ = GetStrUntil(Chr(34), rl$)
End If
Sleep 10
KillFile strTemp & "tmp." & tp$
myRegister = Trim$(rl$)
End Function
Public Function MyShell(ww$, Optional way As VbAppWinStyle = vbNormalFocus, Optional param As String = vbNullString) As Long
Dim frm$, exst As Boolean, pexist As Boolean, pp$, EXE$
' logic

On Error GoTo 11111
If Is64bit Then Wow64EnableWow64FsRedirection False
again:
If ExtractType(ww$) <> "" Then

frm$ = ExtractPath(ww$) + ExtractName(ww$)
param = RTrim$(Mid$(ww$, Len(frm$) + 1) + " " + param)
ww$ = frm$
ElseIf ExtractPath(ww$) = vbNullString Then
Dim i As Long, j As Long
i = InStr(ww$, Chr(34))
j = InStrRev(ww$, Chr(34))
If j > i Then
param = Mid$(ww$, i, j - i + 1)
ww$ = Left$(ww$, i - 1)
End If

End If
If ww$ = vbNullString Then
If param <> "" Then
MyShell = Shell(Trim$(param), way)
If Is64bit Then Wow64EnableWow64FsRedirection True
Exit Function
End If
End If
If ExtractPath(ww$) = mylcasefILE(ww$) Then
' it is a path
ww$ = "a.@@@ " & ww$
Else
frm$ = CFname(ww$)
If ExtractName(frm$) <> ExtractName(ww$) Then
On Error Resume Next

'MyShell = Shell(Trim$(ww$ & " " & param), way)
EXE$ = Trim$(ww$)
MyShell = ShellExecute(0, 0, StrPtr(EXE$), StrPtr(param), 0, way)
'If Is64bit Then Wow64EnableWow64FsRedirection True
If Err.Number > 0 Then

Err.Clear
ww$ = PathFromApp(ww$)
If ww$ <> "" Then
EXE$ = Trim$(ww$)
MyShell = ShellExecute(0, 0, StrPtr(EXE$), StrPtr(param), 0, way)
End If
End If
If Is64bit Then Wow64EnableWow64FsRedirection True
Exit Function
End If
If CFname(ww$) <> "" Then ww$ = frm$: exst = True

pp$ = ExtractPath(ww$)
End If
If pp$ <> "" Then
pexist = True
ww$ = Mid$(ww$, Len(pp$) + 1)
End If
ww$ = ww$ & " "
EXE$ = vbNullString
If InStr(ww$, ".") > InStr(ww$, " ") Then
EXE$ = Left$(ww$, InStr(ww$, "."))
ww$ = Mid$(ww$, Len(EXE$) + 1)
End If
ww$ = ww$ & " "
EXE$ = EXE$ & Trim$(GetStrUntil(" ", ww$))
' until now we have all things splitted
EXE$ = mylcasefILE(EXE$)
' until now we have all things splitted
Select Case ExtractType(EXE$)
Case ""
If pexist Then
' this is not normal
' ***************ERROR*************
Else
' so we put exe by default
EXE$ = EXE$ & ".exe"
frm$ = PathFromApp(Trim$(EXE$ & " " & ww$))
If frm$ <> "" Then
MyShell = Shell(frm$, way)
If Is64bit Then Wow64EnableWow64FsRedirection True
Exit Function
Else
MyShell = Shell(Trim$(EXE$ & " " & ww$ & " " & param), way)
If Is64bit Then Wow64EnableWow64FsRedirection True
Exit Function
End If
End If
Case "exe", "bat", "com" ' can be run immediatly
If pexist Then
EXE$ = pp$ + EXE$
If param <> "" Then
'MyShell = Shell(Trim$(PP$ & EXE$ & " " & ww$ + " " + param), way)
If Form1.Visible And way = vbNormalFocus Then
MyShell = ShellExecute(Form1.hWnd, 0, StrPtr(EXE$), StrPtr(param), StrPtr(pp$), way)
Else
MyShell = ShellExecute(0, 0, StrPtr(EXE$), StrPtr(param), StrPtr(pp$), way)
End If
Else
If Form1.Visible And way = vbNormalFocus Then
MyShell = ShellExecute(Form1.hWnd, 0, StrPtr(EXE$), 0, StrPtr(pp$), way)
Else
MyShell = ShellExecute(0, 0, StrPtr(EXE$), 0, StrPtr(pp$), way)
End If
'MyShell = Shell(Trim$(PP$ & EXE$ & " " & ww$), way)
End If
If Is64bit Then Wow64EnableWow64FsRedirection True
Exit Function
Else
frm$ = PathFromApp(Trim$(EXE$ & " " & ww$))
If frm$ <> "" Then
MyShell = Shell(frm$, vbNormalFocus)
If Is64bit Then Wow64EnableWow64FsRedirection True
Exit Function
Else
MyShell = Shell(Trim$(EXE$ & " " & ww$), way)
If Is64bit Then Wow64EnableWow64FsRedirection True
Exit Function
End If
End If
Case "@@@"
'MyShell = Shell(RTrim$("explorer " & ww$), way)
EXE$ = "explorer"
If Form1.Visible Then
MyShell = ShellExecute(Form1.hWnd, 0, StrPtr(EXE$), StrPtr(ww$), 0, way)
Else
MyShell = ShellExecute(0, 0, StrPtr(EXE$), StrPtr(ww$), 0, way)
End If
If Is64bit Then Wow64EnableWow64FsRedirection True
Case Else ' its a document
pp$ = Replace$(pp$, "file:", "")
frm$ = PCall(pp$ & EXE$)
If frm$ <> "" Then
If AscW(frm$) = 34 Then
frm$ = frm$ & "@"
frm$ = Replace$(frm$, Chr(34) & "@", " " + param & Chr(34))
frm$ = Replace$(frm$, "@", "")
MyShell = Shell(Trim$(frm$), way)
Else
frm$ = PCall(pp$ & EXE$, param)

ww$ = frm$ & " " & ww$
GoTo again
End If
If Is64bit Then Wow64EnableWow64FsRedirection True
Exit Function
Else
End If
If Is64bit Then Wow64EnableWow64FsRedirection True
End Select
If MyShell <> 0 Then AppActivate MyShell
11111:
MyShell = 0
' its a document
End Function
Function newStart(basestack As basetask, rest$) As Boolean
Dim Scr As Object, s$, pa As Long
If HaltLevel > 0 Then rest$ = vbNullString: Exit Function
If Not basestack.IamChild And Not basestack.IamAnEvent Then
        If Check2Save Then
        newStart = True
            Exit Function
        End If
End If
Check2SaveModules = False
MyEr "", ""
NOEXECUTION = False
MOUT = False
byPassCallback = False
EditTabWidth = 6
tParam.cbSize = LenB(tParam)
tParam.iTabLength = 6
ReportTabWidth = 8
HaltLevel = -HaltLevel
newStart = True
Set Scr = basestack.Owner
SetNormal Scr
Targets = False
ReDim q(0) As target
Scr.forecolor = mycolor(11)
basestack.myBold = False
basestack.myitalic = False
pa = 0

            Err.Clear
            On Error Resume Next
If IsStrExp(basestack, rest$, s$) Then
            If s$ <> "" And s$ <> "*" Then MyFont = s$ Else MyFont = Scr.Font.Name
      
                Scr.Font.charset = 0
                Scr.Font.Name = MyFont
               If Not myLcase(MyFont) = myLcase(Scr.Font.Name) Then
               Scr.Font.charset = 1
               Scr.Font.Name = MyFont
               End If
               Sleep 1

                Scr.Font.charset = basestack.myCharSet
                    Form1.TEXT1.Font.charset = basestack.myCharSet

    Form1.List1.Font.charset = basestack.myCharSet

                Scr.FontBold = False
                Scr.FontItalic = False
            If Err.Number > 0 Then
                Err.Clear
                Scr.Font.Name = FFONT
                Scr.Font.charset = basestack.myCharSet
            End If
        StoreFont Scr.Font.Name, Scr.FontSize, Scr.Font.charset
        
        SetText Scr, -2, True
            s$ = vbNullString
            If FastSymbol(rest$, ",") Then
             
            If IsStrExp(basestack, rest$, s$) Then
            'rest$ = s$ & "}" & rest$
            If s$ <> "" Then s$ = ": " & s$ & "}"
            ElseIf FastSymbol(rest$, "{") Then
            s$ = ": " & block(rest$)
            If Not FastSymbol(rest$, "}") Then Set Scr = Nothing: newStart = False: Exit Function
                End If
                End If
            original Basestack1, s$  ' set...
            
Else
    MyEr "", ""
    closeAll ' we closed all files
    If AVIRUN Then MediaPlayer1.stopMovie
    PlaySoundNew ""
    
    If basestack.toprinter Then
    getnextpage
    End If
    Set Scr = Form1.DIS
  
    Form1.myBreak basestack
    basestack.toprinter = False
    players(DisForm).mypen = mycolor(PenOne)
    players(DisForm).mypentrans = 255
    players(DisForm).Paper = mycolor(PaperOne)
    players(DisForm).ReportTab = ReportTabWidth
    Form1.Cls
    
    original Basestack1, ""
    MyNew Basestack1, "", 1
    MyClear Basestack1, ""

    basestack.soros.Flush

End If

End Function
Sub newHide(basestack As basetask)
Dim Scr As Object
Set Scr = basestack.Owner
If Scr.Name = "DIS" Or Scr.Name = "dSprite" Then
Scr.Visible = False
End If
MyDoEvents1 Scr
Set Scr = Nothing
End Sub



Sub newshow(bstack As basetask)
On Error Resume Next
Dim Scr As Object
Set Scr = bstack.Owner
If Scr.Name = "DIS" Or Scr.Name = "dSprite" Then
If Not Form1.Visible And Form1.TrueVisible Then
If UseMe Is Nothing Then
Form3.skiptimer = True
Form3.Visible = True: If Form3.WindowState = 0 Then Form3.move VirtualScreenWidth() + 2000, VirtualScreenHeight() + 2000
mywait Basestack1, 100
Form3.CaptionWsilent = ExtractNameOnly(cLine)
Else
PlaceCaption ExtractNameOnly(cLine)
If Err.Number > 0 Then
Err.Clear: Set UseMe = Nothing
Form3.skiptimer = True
Form3.Visible = True: If Form3.WindowState = 0 Then Form3.move VirtualScreenWidth() + 2000, VirtualScreenHeight() + 2000
mywait Basestack1, 100
Form3.CaptionWsilent = ExtractNameOnly(cLine)
End If
End If
End If
If Form1.Visible = False Then
conthere:
    If ttl Then
        If Form3.WindowState = 1 Then
           Form1.Visible = True
            Form3.skiptimer = True
            Form3.Visible = True
            Form3.skiptimer = True
            Form3.WindowState = 0
              
            Do While Not Form1.Visible Or NOEXECUTION
                mywait bstack, 1
                Sleep 10
            Loop
            If Form3.WindowState = 0 Then Form3.move VirtualScreenWidth() + 2000, VirtualScreenHeight() + 2000
            Form3.Timer1.Interval = 20
            Form3.Show
            Sleep 50
            Form3.CaptionW = vbNullString
            PlaceCaption ""
            If Form3.Visible Then Form3.Refresh
            MyDoEvents1 Form3, True
            mywait bstack, 50
        End If
    Else
        Form1.Show , Form5
        mywait bstack, 5
    End If
Else
    If ttl Then
    ' we have title
        If Form3.WindowState = 1 Then
        Form3.Visible = True: Form3.WindowState = 0: If Form3.WindowState = 0 Then Form3.move VirtualScreenWidth() + 2000, VirtualScreenHeight() + 2000
            Do While Not Form1.Visible Or NOEXECUTION
           mywait bstack, 5
                Loop
        End If
    Else
     Form1.Show , Form5
     mywait bstack, 5
    End If
End If
If Typename(Scr) = "PictureBox" Then
If Scr.Parent.Visible = False Then
Scr.Parent.Visible = True
mywait bstack, 5
End If
End If
    Scr.Visible = True
        Do While Not Scr.Visible Or NOEXECUTION
mywait bstack, 5
        Loop
If Scr.Visible Then
Scr.SetFocus

End If
End If
Set Scr = Nothing
End Sub

Sub getfirstpage()
Dim try1 As Long
If UBound(MyDM) = 1 Then
PrinterDim pw, ph, psw, psh, pwox, phoy
End If

'If pwox > phoy Then mydpi = phoy Else mydpi = pwox
''mydpi = pwox / 4
If pwox <= phoy Then
mydpi = pwox
Else
mydpi = phoy
End If
''prFactor = 1
prFactor = mydpi / 600#
mydpi = 600 '288


again:
On Error Resume Next
' DC FROM PRINTER
'oprinter.EndPrint
oprinter.ClearUp
If oprinter.create(Int(psw / pwox * mydpi + 0.5), Int(psh / phoy * mydpi + 0.5)) Then
Form1.PrinterDocument1.backcolor = QBColor(15)
oprinter.WhiteBits
oprinter.GetDpi mydpi, mydpi
Form1.PrinterDocument1.Cls
oprinter.needHDC
Set Form1.PrinterDocument1 = hDCToPicture(oprinter.HDC1, 0, 0, oprinter.Width, oprinter.Height)
oprinter.FreeHDC
If Err > 0 And try1 < 2 Then
try1 = try1 + 1
prFactor = prFactor * 2
mydpi = mydpi / 2
GoTo again
End If
szFactor = mydpi * dv15 / 1440#
On Error Resume Next
Form1.PrinterDocument1.Refresh
Form1.PrinterDocument1.Scale (0, 0)-(Form1.ScaleX(Int(psw / pwox * mydpi + 0.5), 3, 1), Form1.ScaleY(Int(psh / phoy * mydpi + 0.5), 3, 1))
pnum = 0
End If
End Sub
Sub getnextpage()

If oprinter.Height = 0 Then
getfirstpage

Else
pnum = pnum + 1
With players(-2)
.curpos = 0
.currow = 0
.lastprint = False
.XGRAPH = 0
.YGRAPH = 0
End With
With Form1.PrinterDocument1
.currentX = 0
.currentY = 0
End With
oprinter.CopyPicturePrinter Form1.PrinterDocument1
oprinter.GetDpi mydpi, mydpi
oprinter.ThumbnailPaintPrinter 1, 100 * prFactor, False, True, True, , , , , , Form3.CaptionW '& " " & Str$(pnum)
oprinter.ClearBits Form1.PrinterDocument1.backcolor
oprinter.needHDC
Set Form1.PrinterDocument1 = hDCToPicture(oprinter.HDC1, 0, 0, oprinter.Width, oprinter.Height)
oprinter.FreeHDC

Form1.PrinterDocument1.Refresh
Form1.PrinterDocument1.Scale (0, 0)-(Form1.ScaleX(Int(psw / pwox * mydpi + 0.5), 3, 1), Form1.ScaleY(Int(psh / phoy * mydpi + 0.5), 3, 1))


End If
End Sub
Sub getenddoc()
pnum = pnum + 1
If prFactor = 0 Then prFactor = 1
Form1.Refresh
oprinter.CopyPicturePrinter Form1.PrinterDocument1
Form1.PrinterDocument1.Picture = LoadPicture("")
oprinter.ThumbnailPaintPrinter 1, 100 * prFactor, False, True, True, , , , , , Form3.CaptionW '& " " & Str$(pnum)
oprinter.ClearUp
oprinter.EndPrint



'Set oprinter = New cDIBSection
End Sub
Function MyDrawings(bstack As basetask, rest$, Lang As Long) As Boolean
MyDrawings = True
Dim mDir As recDir, s$, ss$, frm$
Set mDir = New recDir
mDir.IncludedFolders = False
mDir.Nofiles = False
mDir.TopFolder = mcd
mDir.SortType = Abs(FastSymbol(rest$, "!"))
frm$ = ExtractName(mDir.Dir2$(mcd, "WMF|EMF", False), True)
s$ = vbNullString
ss$ = vbNullString
Do While frm$ <> ""
s$ = frm$
If ss$ <> "" Then ss$ = ss$ & ", " & s$ Else ss$ = s$
s$ = mDir.Dir2
If s$ <> "" Then frm$ = ExtractName(s$, True) Else frm$ = vbNullString
Loop

If Lang Then
ss$ = "Drawings: " & ss$
Else
ss$ = "Σχέδια: " & ss$
End If
Set mDir = Nothing
RepPlain bstack, bstack.Owner, ss$
End Function
Function MyMovies(bstack As basetask, rest$, Lang As Long) As Boolean
MyMovies = True
Dim mDir As recDir, s$, ss$, frm$
Set mDir = New recDir
mDir.IncludedFolders = False
mDir.Nofiles = False
mDir.TopFolder = mcd
mDir.SortType = Abs(FastSymbol(rest$, "!"))
frm$ = ExtractNameOnly(mDir.Dir2$(mcd, "AVI", False), True)
s$ = vbNullString
ss$ = vbNullString
Do While frm$ <> ""
s$ = frm$
If ss$ <> "" Then ss$ = ss$ & ", " & s$ Else ss$ = s$
s$ = mDir.Dir2
If s$ <> "" Then frm$ = ExtractNameOnly(s$, True) Else frm$ = vbNullString
Loop
Set mDir = Nothing
If Lang Then
ss$ = "Movies: " & ss$
Else
ss$ = "Ταινίες: " & ss$
End If
RepPlain bstack, bstack.Owner, ss$
End Function
Function MyBitmaps(bstack As basetask, rest$, Lang As Long) As Boolean
MyBitmaps = True
Dim mDir As recDir, s$, ss$, frm$
Set mDir = New recDir
mDir.IncludedFolders = False
mDir.Nofiles = False
mDir.TopFolder = mcd
mDir.SortType = Abs(FastSymbol(rest$, "!"))
frm$ = ExtractName(mDir.Dir2$(mcd, "BMP|JPG|GIF|DIB|ICO|CUR|PNG|TIF", False), True)
s$ = vbNullString
ss$ = vbNullString
Do While frm$ <> ""
s$ = frm$
If ss$ <> "" Then ss$ = ss$ & ", " & s$ Else ss$ = s$
s$ = mDir.Dir2
If s$ <> "" Then frm$ = ExtractName(s$, True) Else frm$ = vbNullString
Loop
If Lang Then
ss$ = "Bitmaps: " & ss$
Else
ss$ = "Εικόνες: " & ss$
End If
Set mDir = Nothing
RepPlain bstack, bstack.Owner, ss$
End Function
Function ProcFKey(bstack As basetask, rest$, Lang As Long) As Boolean
Dim i As Long, p As Variant, s$, prive
If IsLabelSymbolNew(rest$, "ΚΑΘΑΡΟ", "CLEAR", Lang) Then
    For i = 1 To 24: FK$(i) = vbNullString: Next i
ElseIf IsExp(bstack, rest$, p) Then

    i = ((CLng(p) + 23) Mod 24) + 1
    If lookOne(rest$, "{") Then
         If IsStrExp(bstack, rest$, s$) Then
            FK$(i) = s$
        Else
            MissPar
            Exit Function
        End If
    ElseIf FastSymbol(rest$, ",") Then
        If IsStrExp(bstack, rest$, s$) Then
            FK$(i) = s$
        Else
            MissPar
            Exit Function
        End If
    Else
        prive = GetCode(bstack.Owner)
        PlainBaSket bstack.Owner, players(prive), FK$(i)
        crNew bstack, players(prive)
    End If
Else
    s$ = vbNullString: prive = GetCode(bstack.Owner)
    For i = 1 To 24
        If FK$(i) <> "" Then
            s$ = s$ + placeme$("ΚΛΕΙΔΙ", "FKEY", Lang) + Right$(" " & Str$(i), 3) & " [" & FK$(i) & "]" ' FKEY
            If i > 12 Then s$ = s$ + " SHIFT + F" & CStr(i - 12) Else s$ = s$ + " F" & CStr(i)
            s$ = s$ + vbCrLf
        End If
    Next i
    RepPlain bstack, bstack.Owner, s$
End If
ProcFKey = True
End Function
Function placeme$(gre$, Eng$, code As Long)
If code = 1 Then placeme$ = Eng$ Else placeme$ = gre$
End Function
Function MyScan(basestack As basetask, rest$) As Boolean
Dim p As Variant, y As Double, s$
ClearJoyAll
PollJoypadk
If GetForegroundWindow <> Form1.hWnd Or Not Targets Then
If IsExp(basestack, rest$, p, , True) Then

End If
MyScan = True
MyDoEvents0 basestack.Owner
If Fkey > 0 Then
If FK$(Fkey) <> "" Then
    s$ = FK$(Fkey)
    MyScan = interpret(Basestack1, s$)
Fkey = 0
End If
End If


 Exit Function
End If
If basestack.Owner.Visible = False Then basestack.Owner.Visible = True
basestack.Owner.SetFocus
NoAction = False

nomore = True
If IsExp(basestack, rest$, p, , True) Then
y = Timer + p

Do ' TOO
MyDoEvents
Loop Until NoAction Or Timer > y Or NOEXECUTION
    'End If
Else
Do ' TOO
 MyDoEvents
If Fkey > 0 Then
If FK$(Fkey) <> "" Then
rest$ = FK$(Fkey) + rest$
Fkey = 0
Exit Do
End If
End If
Loop Until NoAction Or NOEXECUTION
End If
nomore = False ' TOO

End Function

Function MyHelp(basestack As basetask, rest$, Lang As Long) As Boolean
Dim s$, s1$, aa As Boolean
If Not basestack.IamChild Or Not mHelp Or Not basestack.IamAnEvent Then
mHelp = False
abt = False
lastAboutHTitle = vbNullString
Dim i As Long
i = 1
If MaybeIsSymbol3lot(rest$, "?", i) Then
Mid$(rest$, 1, i) = space(i)
    fHelp basestack, "PRINT", Abs(pagio$ = "GREEK") + 1
    GoTo fhExit
ElseIf MaybeIsSymbol3lot(rest$, "@~$#", i) Then
s1$ = Mid$(rest$, i, 1)
Mid$(rest$, 1, i) = space(i)
i = i + 1
If MaybeIsSymbol3lot(rest$, "(", i) Then
s$ = s1$ + s$


FastSymbol rest$, ")"
    fHelp basestack, s$, Abs(pagio$ = "GREEK") + 1
    GoTo fhExit
End If

End If
If Abs(IsLabel(basestack, rest$, s$)) > 0 Then
    vH_title$ = vbNullString
    aa = AscW(s$ + Mid$(" Σ", Abs(pagio$ = "GREEK") + 1)) < 128
    If Len(s1$) > 0 Then s$ = s1$ + s$
    fHelp basestack, s$, aa
ElseIf Not ISSTRINGA(rest$, s$) Then
    nhelp basestack, Lang <> 1
Else
    fHelp basestack, s$, Lang = 1
End If
End If
fhExit:
MyHelp = True
End Function
Function ProcCls(basestack As basetask, rest$) As Boolean
'If basestack.toprinter Then Exit Function
Dim Scr As Object, p As Variant
 ProcCls = True
Set Scr = basestack.Owner

With players(GetCode(Scr))
If Not IsExp(basestack, rest$, p, , True) Then
p = -.Paper
Else
.Paper = mycolor(p)
End If



If FastSymbol(rest$, ",") Then
    If IsExp(basestack, rest$, p, , True) Then
    If Not basestack.toprinter Then
    If p < 0 Then p = .My + p
        .mysplit = MyRound(p)
        If .mysplit >= .My Then .mysplit = 0: .pageframe = 0
        .pageframe = Int((.My - .mysplit) * 2 / 3)

    End If
        Else
        ProcCls = False
        Set Scr = Nothing
        Exit Function
    End If
End If
If basestack.toprinter Then
If oprinter.Height > 0 Then
oprinter.Cls CLng(mycolor(.Paper))
oprinter.PaintPicture Form1.PrinterDocument1.Hdc
.curpos = 0
.currow = .mysplit
.lastprint = False
Scr.currentX = 0
Scr.currentY = 0
End If
Else
ClearScrNew Scr, players(GetCode(Scr)), CLng(mycolor(.Paper))
End If
If Form4Loaded Then
If Form4.Visible Then
If Not mHelp And Not abt Then vHelp
End If
End If
End With
Set Scr = Nothing
End Function

Public Sub DelTemp()
Dim tmp$
On Error Resume Next
While tempList2delete <> ""
If Not ISSTRINGA(tempList2delete, tmp$) Then Exit Sub
KillFile tmp$
Wend

End Sub
Public Function GetTempFileName() As String

   Dim sTmp    As String
   Dim sTmp2   As String
   Dim EXENAME As String
   EXENAME = App.EXENAME

   sTmp2 = GetTempPathgg
   sTmp = space(Len(sTmp2) + 256)
   Call GetTempFileNameW(StrPtr(sTmp2), StrPtr(EXENAME), UNIQUE_NAME, StrPtr(sTmp))
   GetTempFileName = Left$(sTmp, InStr(sTmp, Chr$(0)) - 1)
    tempList2delete = Sput(GetTempFileName) + tempList2delete
End Function
Private Function GetTempPathgg() As String
  
   Dim sTmp       As String
   Dim i          As Long
    Dim EM$
    
   i = GetTempPath(0, StrPtr(EM$))
   sTmp = space(i)

   Call GetTempPath(i, StrPtr(sTmp))
   GetTempPathgg = AddBackslash(Left$(sTmp, i - 1))

End Function
Public Function AddBackslash(s As String) As String

   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s & "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If

End Function
Function ProcCreateEmf(bstack As basetask, rest$, Lang As Long) As Boolean
Dim w, H  ' these are twips - need to convert to .01 mm
Dim F As Boolean, p As Variant, Col As Long, it As Long, ss$, x As Double, par As Boolean, prive As Long
Dim nd&, once As Boolean
ProcCreateEmf = True
prive = GetCode(bstack.Owner)
' skip for now
If IsExp(bstack, rest$, w, , True) Then
    If FastSymbol(rest$, ",") Then
           If Not IsExp(bstack, rest$, H, , True) Then MissNumExpr
    Else
    
    End If
End If
If FastSymbol(rest$, "{") Then
            ss$ = block(rest$)
            TraceStore bstack, nd&, rest$, 0
            If FastSymbol(rest$, "}") Then
                Call executeblock(it, bstack, ss$, False, once, , True)
                If it = 2 Then
                    If ss$ = "" Then
                        If once Then rest$ = ": Break": If trace Then WaitShow = 2: TestShowSub = vbNullString
                    Else
                        rest$ = ": Goto " + ss$
                        If trace Then WaitShow = 2: TestShowSub = rest$
                    End If
                    
                    it = 1
                End If
                If it <> 1 Then ProcCreateEmf = False: rest$ = ss$ + rest$
            Else
                MissPar
                ProcCreateEmf = False
                Exit Function
            End If
            bstack.addlen = nd&
        Else
            MissPar
            ProcCreateEmf = False
            Exit Function
        End If
        If Not IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then Exit Function

End Function
Function ProcPath(bstack As basetask, rest$, Lang As Long) As Boolean
Dim F As Boolean, p As Variant, Col As Long, it As Long, ss$, x As Double, par As Boolean, prive As Long
Dim OldGDILines As Boolean, region As Boolean, oldpathcolor As Long, oldpathfillstyle As Integer, nd&, once As Boolean
ProcPath = True
prive = GetCode(bstack.Owner)
F = IsLabelSymbolNew(rest$, "ΠΑΝΩ", "OVER", Lang)
If FastSymbol(rest$, "!") Then par = True

If IsExp(bstack, rest$, p, , True) Then
        Col = CLng(p)  ' using a fill color
        If FastSymbol(rest$, ",") Then
           If Not IsExp(bstack, rest$, x, , True) Then MissNumExpr
           Else
           x = vbSolid
           End If
        If FastSymbol(rest$, "{") Then

            ss$ = block(rest$)
            
            TraceStore bstack, nd&, rest$, 0
            If FastSymbol(rest$, "}") Then
            If MaybeIsSymbol(rest$, ";") Then
                    If MyTrim(ss$) = vbNullString Then GoTo contthere Else GoTo contthere2
            End If
                players(prive).pathgdi = players(prive).pathgdi + 1
                oldpathfillstyle = players(prive).pathfillstyle
                oldpathcolor = players(prive).pathcolor
                
                players(prive).pathcolor = mycolor(Col)
                players(prive).pathfillstyle = Int(x) Mod 8
                
                BeginPath bstack.Owner.Hdc
              '  If (par Or F) And GDILines Then OldGDILines = True: players(prive).NoGDI = True
                If (par Or region) Then players(prive).NoGDI = True: If GDILines Then OldGDILines = True
                Call executeblock(it, bstack, ss$, False, once, , True)

                
                players(prive).pathgdi = players(prive).pathgdi - 1
                If players(prive).pathgdi > 0 Then
                    players(prive).pathcolor = oldpathcolor
                    players(prive).pathfillstyle = oldpathfillstyle
                Else
                    If (par Or region) Then players(prive).NoGDI = False: GDILines = OldGDILines
                End If
                        
                ' what for 2 and 3 values
                EndPath bstack.Owner.Hdc
        
                bstack.Owner.fillstyle = Int(x) Mod 8
                bstack.Owner.fillcolor = mycolor(Col)
                Col = p ' from  6.3 change
                If par Then bstack.Owner.DrawMode = 7
                If F Then  ' from 8 rev 83
                      If bstack.Owner.fillstyle = 1 Then
                           StrokeAndFillPath bstack.Owner.Hdc
                        Else
                            FillPath bstack.Owner.Hdc
                          End If
                Else
                     StrokeAndFillPath bstack.Owner.Hdc         ' stroke and fill path
                End If
                If par Then bstack.Owner.DrawMode = 13
                bstack.Owner.fillstyle = vbFSTransparent
                If players(prive).pathgdi = 0 Then
                    players(prive).pathcolor = oldpathcolor
                    players(prive).pathfillstyle = oldpathfillstyle
                End If
                If it = 2 Then
                    If ss$ = "" Then
                        If once Then rest$ = ": Break": If trace Then WaitShow = 2: TestShowSub = vbNullString
                    Else
                        rest$ = ": Goto " + ss$
                        If trace Then WaitShow = 2: TestShowSub = rest$
                    End If
                    
                    it = 1
                End If
                If it <> 1 Then ProcPath = False: rest$ = ss$ + rest$
            Else
                MissPar
                ProcPath = False
            End If
            bstack.addlen = nd&

        Else
            MissPar
            ProcPath = False
        End If
    Exit Function
    Else
        If FastSymbol(rest$, "{") Then
            ss$ = block(rest$)
            TraceStore bstack, nd&, rest$, 0
            If MyTrim(ss$) = vbNullString Then
            ProcPath = FastSymbol(rest$, "}")
contthere:
              If FastSymbol(rest$, ";") Then
 
                  SelectClipRgn bstack.Owner.Hdc, 0&
              End If
              bstack.addlen = nd&
              Exit Function
            End If
        
            If FastSymbol(rest$, "}") Then
contthere2:
                If FastSymbol(rest$, ";") Then region = True
                
                oldpathfillstyle = players(prive).pathfillstyle
                oldpathcolor = players(prive).pathcolor
                
                players(prive).pathcolor = mycolor(Col)
                players(prive).pathfillstyle = Int(x) Mod 8
                BeginPath bstack.Owner.Hdc
                'If (par Or region) And GDILines Then OldGDILines = True: players(prive).NoGDI = True
                If (par Or region) Then players(prive).NoGDI = True: If GDILines Then OldGDILines = True
                Call executeblock(it, bstack, ss$, False, once, , , True)
                
                EndPath bstack.Owner.Hdc
                players(prive).pathgdi = players(prive).pathgdi - 1
                If players(prive).pathgdi > 0 Then
                    players(prive).pathcolor = oldpathcolor
                    players(prive).pathfillstyle = oldpathfillstyle
                Else
                    If (par Or region) Then players(prive).NoGDI = False: GDILines = OldGDILines
                End If
                bstack.Owner.fillstyle = vbSolid
                If region Then            ' path { block of commands };
                    
                    If F Then
                        SelectClipPath bstack.Owner.Hdc, 2
                    Else
                        SelectClipPath bstack.Owner.Hdc, RGN_COPY  ' make a clip path
                    End If
                
                    
                Else
                    If par Then bstack.Owner.DrawMode = 7
                    StrokePath bstack.Owner.Hdc               ' stroke path
                    If par Then bstack.Owner.DrawMode = 13
                End If
                 bstack.Owner.fillstyle = vbFSTransparent
                If it = 2 Then
                    If ss$ = "" Then
                    If once Then rest$ = ": Break": If trace Then WaitShow = 2: TestShowSub = vbNullString
                    Else
                    rest$ = ": Goto " + ss$
                    If trace Then WaitShow = 2: TestShowSub = rest$
                    End If
                    
                    it = 1
                End If
                If it <> 1 Then ProcPath = False: rest$ = ss$ + rest$
            Else
                MissPar
                ProcPath = False
            End If
            bstack.addlen = nd&
    Else
        MissPar
        ProcPath = False
    End If

End If

End Function
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Long

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Sub AddDirSep(strPathName As String)
    If Right(Trim(strPathName), Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR And _
       Right(Trim(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub
Function GetWindowsDir() As String
    Dim strBuf As String
    Const gintMAX_SIZE& = 255                        'Maximum buffer size
    strBuf = space$(gintMAX_SIZE)

    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf

        GetWindowsDir = strBuf
    Else
        GetWindowsDir = vbNullString
    End If
End Function
Sub Portrait(bstack As basetask)
Dim dummy As Object, try1 As Long
If UBound(MyDM) = 1 Then
PrinterDim pw, ph, psw, psh, pwox, phoy
End If
If pwox <= phoy Then
mydpi = pwox
Else
mydpi = phoy
End If
prFactor = mydpi / 600#
mydpi = 600
szFactor = mydpi * dv15 / 1440#

If Int(psw / pwox * mydpi + 0.5) / Int(psh / phoy * mydpi + 0.5) > 1 Then
    If oprinter.PrinterOn Then
        ChangeNowOrientationPortrait
        oprinter.ResetPageDM
        SwapPrinterDim pw, ph, psw, psh, pwox, phoy
        GoTo contnow
    Else
        ChangeOrientation dummy, Printer.DeviceName, MyDM()
        SwapPrinterDim pw, ph, psw, psh, pwox, phoy
        Exit Sub
    End If
    
Else
    Form1.PrinterDocument1.Cls
    Form1.Refresh
    Exit Sub
End If
contnow:
Dim thisprinter As New cDIBSection

If thisprinter.create(Int(psw / pwox * mydpi + 0.5), Int(psh / phoy * mydpi + 0.5)) Then
    thisprinter.ClearBits Form1.PrinterDocument1.backcolor
    
    thisprinter.GetDpi mydpi, mydpi
    Form1.PrinterDocument1.Cls
    thisprinter.needHDC
    Set Form1.PrinterDocument1 = hDCToPicture(thisprinter.HDC1, 0, 0, thisprinter.Width, thisprinter.Height)
    thisprinter.FreeHDC
    If Err > 0 And try1 < 2 Then
        try1 = try1 + 1
        prFactor = prFactor * 2
        mydpi = mydpi / 2
        GoTo contnow
    End If
    szFactor = mydpi * dv15 / 1440#
    On Error Resume Next
    Form1.Refresh
    Form1.PrinterDocument1.Scale (0, 0)-(Form1.ScaleX(Int(psw / pwox * mydpi + 0.5), 3, 1), Form1.ScaleY(Int(psh / phoy * mydpi + 0.5), 3, 1))
    thisprinter.CopyPrinter oprinter.PrinterHdc
    Set oprinter = thisprinter
ElseIf try1 < 2 Then
        try1 = try1 + 1
        prFactor = prFactor * 2
        mydpi = mydpi / 2
        GoTo contnow
End If

If bstack.toprinter Then
    SetText Form1.PrinterDocument1
    Else
    PlaceBasket Form1.PrinterDocument1, players(-2)
    SetText Form1.PrinterDocument1
End If
End Sub
Sub Landscape(bstack As basetask)
Dim dummy As Object, try1 As Long
If UBound(MyDM) = 1 Then
PrinterDim pw, ph, psw, psh, pwox, phoy
End If
If pwox <= phoy Then
mydpi = pwox
Else
mydpi = phoy
End If
prFactor = mydpi / 600#
mydpi = 600

If Int(psw / pwox * mydpi + 0.5) / Int(psh / phoy * mydpi + 0.5) < 1 Then
    If oprinter.PrinterOn Then
        ChangeNowOrientationLandscape
        oprinter.ResetPageDM
        SwapPrinterDim pw, ph, psw, psh, pwox, phoy
        GoTo contnow
    Else
        ChangeOrientation dummy, Printer.DeviceName, MyDM()
        SwapPrinterDim pw, ph, psw, psh, pwox, phoy
        Exit Sub
    End If
Else
    Form1.PrinterDocument1.Cls
    Form1.Refresh
    Exit Sub
End If

contnow:
Dim thisprinter As New cDIBSection

If thisprinter.create(Int(psw / pwox * mydpi + 0.5), Int(psh / phoy * mydpi + 0.5)) Then
    thisprinter.ClearBits Form1.PrinterDocument1.backcolor
    thisprinter.GetDpi mydpi, mydpi
    thisprinter.needHDC
    On Error Resume Next
    Set Form1.PrinterDocument1 = hDCToPicture(thisprinter.HDC1, 0, 0, thisprinter.Width, thisprinter.Height)
    thisprinter.FreeHDC
    If Err > 0 And try1 < 2 Then
        thisprinter.ClearUp
        try1 = try1 + 1
        prFactor = prFactor * 2
        mydpi = mydpi / 2
        GoTo contnow
    End If
    szFactor = mydpi * dv15 / 1440#
    On Error Resume Next
    Form1.Refresh
    Form1.PrinterDocument1.Scale (0, 0)-(Form1.ScaleX(Int(psw / pwox * (mydpi) + 0.5), 3, 1), Form1.ScaleY(Int(psh / phoy * mydpi + 0.5), 3, 1))
    thisprinter.CopyPrinter oprinter.PrinterHdc
    Set oprinter = thisprinter
ElseIf try1 < 2 Then
        try1 = try1 + 1
        prFactor = prFactor * 2
        mydpi = mydpi / 2
        GoTo contnow
End If

If bstack.toprinter Then
    SetText Form1.PrinterDocument1
    Else
    PlaceBasket Form1.PrinterDocument1, players(-2)
    SetText Form1.PrinterDocument1
End If

End Sub
Sub nhelp(bstack As basetask, Optional GREEK As Boolean = False)
Dim di As Object
Set di = bstack.Owner
If GREEK Then
Dim bb$
bb$ = "   ΕΛΛΗΝΙΚΑ ή ΛΑΤΙΝΙΚΑ για επιλογή κωδικοσελίδας για το τύπο εμφάνισης βοήθειας " & vbCrLf
bb$ = bb$ & "   Με Esc τερματίζει η εκτέλεση τμημάτων  " & vbCrLf
bb$ = bb$ & "   ctrl + f1 ανοίγει την βοήθεια, γράφοντας και επιλέγοντας βρίσκει" & vbCrLf
bb$ = bb$ & "   ctrl + c Τερματίζει την εκτέλεση και καθαρίζει" & vbCrLf
bb$ = bb$ & "   ctrl + οποιοδήποτε πλήκτρο ανοίγει τη βηματική εκτέλεση" & vbCrLf
bb$ = bb$ & "   pause/break κάνει ψυχρή εκκίνηση / δες ΒΟΗΘΕΙΑ ΑΡΧΗ" & vbCrLf
bb$ = bb$ & "   Σ ονομαΤμηματος ανοίγει τον διορθωτή για να γράψουμε πρόγραμμα" & vbCrLf
bb$ = bb$ & "   Σ ονομαΣυνάρτησης( ανοίγει τον διορθωτή για να γράψουμε συνάρτηση" & vbCrLf
bb$ = bb$ & "   Σ ονομαΣυνάρτησης$( ανοίγει τον διορθωτή για να γράψουμε συνάρτηση$" & vbCrLf
bb$ = bb$ & "   Τμηματα  [μας δείχνει τα τμήματα στη μνήμη και το δίσκο]" & vbCrLf
bb$ = bb$ & "   ΒΟΗΘΕΙΑ κατι (μας δίνει βοήθεια σε ξεχωριστό παράθυρο)" & vbCrLf
bb$ = bb$ & "   ? ή ΤΥΠΩΣΕ τυπώνει" & vbCrLf
bb$ = bb$ & "   δώσε την εντολή ΡΥΘΜΙΣΕΙΣ η ctrl+U για να αλλάξει την εξ ορισμού γραμματοσειρά και τα χρώματα" & vbCrLf
bb$ = bb$ & "   δώσε την εντολή ΕΛΕΓΧΟΣ για να δεις στοιχεία του διερμηνευτή" & vbCrLf
bb$ = bb$ & "   δώσε την εντολή ΤΕΛΟΣ για να τερματίσεις τον διερμηνευτή" & vbCrLf
Else
bb$ = "   GREEK or LATIN for choose the codepage for errors display" & vbCrLf
bb$ = bb$ & "   with LATIN all error messages are in ENGLISH language  " & vbCrLf
bb$ = bb$ & "   Esc escape execution" & vbCrLf
bb$ = bb$ & "   ctrl + f1 open help form, you can write and click for find" & vbCrLf
bb$ = bb$ & "   ctrl + c terminate execution, clear all" & vbCrLf
bb$ = bb$ & "   ctrl + anykey open for test" & vbCrLf
bb$ = bb$ & "   pause/break for break cold reset / look HELP START" & vbCrLf
bb$ = bb$ & "   EDIT modulename     [open editor for writing program]" & vbCrLf
bb$ = bb$ & "   EDIT functionname( [open editor for writing function()]" & vbCrLf
bb$ = bb$ & "   EDIT functionname$( [open editor for writing function$()]" & vbCrLf
bb$ = bb$ & "   MODULES for a list of modules in memory and on dik" & vbCrLf
bb$ = bb$ & "   use HELP writesomething [to find some help, open the help form]" & vbCrLf
bb$ = bb$ & "   ? or PRINT for printing" & vbCrLf
bb$ = bb$ & "   type SETTINGS or ctrl+U to change the default font and colors" & vbCrLf
bb$ = bb$ & "   type MONITOR for info about current state of Interpreter" & vbCrLf
bb$ = bb$ & "   type END and press enter to close this program" & vbCrLf
End If
wwPlain2 bstack, players(GetCode(bstack.Owner)), bb$, di.Width, 1000, True
crNew bstack, players(GetCode(bstack.Owner))
End Sub



 

Function GetWindowsfontDir() As String
    Dim strBuf As String
    Const gintMAX_SIZE& = 255                        'Maximum buffer size
    strBuf = space$(gintMAX_SIZE)

    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf
        strBuf = strBuf & "FONT"
        AddDirSep strBuf
    
        GetWindowsfontDir = strBuf
    Else
        GetWindowsfontDir = vbNullString
    End If
End Function
Sub GREEK(bstack As basetask)
On Error Resume Next
Clid = 1032
UserCodePage = 1253
DefBooleanString = ";Αληθές;Ψευδές"
NoUseDec = False
NowDec$ = ","
NowThou$ = "."
With Form1
   bstack.myCharSet = 161
If bstack.tolayer > 0 Then
    .dSprite(bstack.tolayer).Font.charset = 161
    ElseIf bstack.toback Then
    .Font.charset = 161
    Else
    .DIS.Font.charset = bstack.myCharSet
    .TEXT1.Font.charset = bstack.myCharSet
    .List1.Font.charset = bstack.myCharSet
   ' .List2.Font.CharSet = bstack.myCharSet
    End If
End With
pagio$ = "GREEK"
Clid = 1032
DefBooleanString = ";Αληθές;Ψευδές"
DialogSetupLang 0
With players(GetCode(bstack.Owner))
.charset = 161
End With
End Sub


Sub LATIN(bstack As basetask)
On Error Resume Next
Clid = 1033
UserCodePage = 1252
UserCodePage = 1253
NoUseDec = False
NowDec$ = "."
NowThou$ = ","
With Form1
bstack.myCharSet = 0
If bstack.tolayer > 0 Then
    .dSprite(bstack.tolayer).Font.charset = 0
    ElseIf bstack.toback Then
    .Font.charset = 0
    Else
    
    .DIS.Font.charset = bstack.myCharSet
    .TEXT1.Font.charset = bstack.myCharSet
    .List1.Font.charset = bstack.myCharSet
   ' .List2.Font.CharSet = bstack.myCharSet
    End If
End With
pagio$ = "LATIN"
Clid = 1033
DefBooleanString = ";\T\r\u\e;\F\a\l\s\e"
DialogSetupLang 1
With players(GetCode(bstack.Owner))
.charset = 0
End With
End Sub

Private Function GetLCIDFromKeyboard() As Long
    Dim Buffer As String, ret&, R&
    Buffer = String$(514, 0)
      R = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      R = val("&H" & Right(Hex(R), 4))
        ret = GetLocaleInfoW(R, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    GetLCIDFromKeyboard = CLng(val("&h" + Left$(Buffer, ret - 1)))
End Function
Public Function GetLCIDFromKeyboardLanguage() As String
    Dim Buffer As String, ret&, R&
    Buffer = String$(514, 0)
      R = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      R = val("&H" & Right(Hex(R), 4))
      'LOCALE_SENGLANGUAGE&
      If UserCodePage = DefCodePage Then ''
      ret = GetLocaleInfoW(R, LOCALE_SENGLANGUAGE&, StrPtr(Buffer), Len(Buffer))
      Else
        ret = GetLocaleInfoW(R, LOCALE_SLANGUAGE&, StrPtr(Buffer), Len(Buffer))
        End If
     If shortlang Then If ret > 3 Then ret = 4
     On Error Resume Next
    GetLCIDFromKeyboardLanguage = Left$(Buffer, ret - 1)

End Function
Public Function GetlocaleString(ByVal this As Long) As String
On Error GoTo 1234
    Dim Buffer As String, ret&, R&
    Buffer = String$(514, 0)
      
        ret = GetLocaleInfoW(Clid, this, StrPtr(Buffer), Len(Buffer))
    GetlocaleString = Left$(Buffer, ret - 1)
    
1234:
    
End Function
Public Function GetlocaleString2(ByVal this As Long, ByVal McLid As Long) As String
On Error GoTo 1234
    Dim Buffer As String, ret&, R&
    Buffer = String$(514, 0)
      
        ret = GetLocaleInfoW(McLid, this, StrPtr(Buffer), Len(Buffer))
    GetlocaleString2 = Left$(Buffer, ret - 1)
    
1234:
    
End Function

Public Function QueryDecString() As String
QueryDecString = GetDeflocaleString(14)
End Function

Function IsLabelOnly(a$, R$) As Long
Dim n$
If Len(a$) < 129 Then
    IsLabelOnly = IsLabelOnlyInner(a$, R$)
Else
    n$ = Left$(a$, 128)
    IsLabelOnly = IsLabelOnlyInner(n$, R$)
    If Len(n$) = 0 Then
        IsLabelOnly = IsLabelOnlyInner(a$, R$)
    Else
        a$ = Mid$(a$, 129 - Len(n$))
    End If
End If
End Function

Function IsLabelOnlyInner(a$, R$) As Long  ' ok
Dim rr&, one As Boolean, c$, dot&
R$ = vbNullString
If a$ = vbNullString Then IsLabelOnlyInner = 0: Exit Function
a$ = NLtrim$(a$)
Do While Len(a$) > 0
    c$ = Left$(a$, 1) 'ANYCHAR HERE
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 64  '"@"
            If R$ = vbNullString Then
                a$ = Mid$(a$, 2)
            ElseIf Mid$(a$, 2, 1) <> "(" And R$ <> "" Then
                R$ = R$ & "."
                a$ = Mid$(a$, 2)
            Else
                 IsLabelOnlyInner = 0: Exit Function
            End If
        Case 46 '"."
            If one Then
                Exit Do
            ElseIf R$ <> "" Then
                R$ = R$ & c$
                a$ = Mid$(a$, 2)
            ElseIf Not Mid$(a$, 2, 1) Like "[0-9]" Then
                If R$ <> "" Then
                    R$ = R$ & c$
                    rr& = 1
                Else
                    dot& = dot& + 1
                End If
                a$ = Mid$(a$, 2)
            Else
                If R$ = vbNullString And dot& > 0 Then
                    R$ = String$(dot& + 1, ".")
                    a$ = Mid$(a$, 2)
                    IsLabelOnlyInner = 1
                Else
                    IsLabelOnlyInner = 0
                End If
                Exit Function
            End If
        Case 92, 94, 123 To 126, 160 '"\","^", "{" To "~"
            Exit Do
        Case 48 To 57, 95 '"0" To "9", "_"
            If one Then
                Exit Do
            ElseIf R$ <> "" Then
                R$ = R$ & c$
                a$ = Mid$(a$, 2)
                rr& = 1 'is an identifier or floating point variable
            Else
                If dot& > 0 Then a$ = "." + a$: dot& = 0
                Exit Do
            End If
        Case Is < 0, Is > 64 ' >=A and negative
            If one Then
                Exit Do
            Else
                R$ = R$ & c$
                a$ = Mid$(a$, 2)
                rr& = 1 'is an identifier or floating point variable
            End If
        Case 38 ' "&"
            If R$ = vbNullString Then rr& = 2:    a$ = Mid$(a$, 2)
            Exit Do
        Case 36 ' "$"
            If one Then Exit Do
            If R$ <> "" Then
                one = True
                rr& = 3 ' is string variable
                R$ = R$ & c$
                a$ = Mid$(a$, 2)
            Else
                Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If R$ <> "" Then
                one = True
                rr& = 4 ' is long variable
                R$ = R$ & c$
                a$ = Mid$(a$, 2)
            Else
                Exit Do
            End If
        Case 40 ' "("
            If R$ <> "" Then
                If Mid$(a$, 2, 2) = ")@" Then
                    R$ = R$ & "()."
                    a$ = Mid$(a$, 4)
                Else
                    Select Case rr&
                    Case 1
                        rr& = 5 ' float array or function
                    Case 3
                        rr& = 6 'string array or function
                    Case 4
                        rr& = 7 ' long array
                    Case Else
                        Exit Do
                    End Select
                    R$ = R$ & c$
                    a$ = Mid$(a$, 2)
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Case Else
            Exit Do
        End Select
    Else
        If one Then
            Exit Do
        Else
            R$ = R$ & c$
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
        End If
    End If
Loop
IsLabelOnlyInner = rr&
End Function
Function ProcFind(basestack As basetask, rest$) As Boolean
Dim i As Long, s$, pppp As mArray, x1 As Long, y1 As Long
Dim p As Variant, Col As Long, frm$, ss$
ProcFind = True
    y1 = Abs(IsLabel(basestack, rest$, s$))
     
        If y1 = 6 Then
                If neoGetArray(basestack, s$, pppp) Then
                 If Not NeoGetArrayItem(pppp, basestack, s$, i, rest$) Then Exit Function
                Else
                    MissingDoc
                    Exit Function
                End If
    End If
    If FastSymbol(rest$, ",") Then
    If Not IsStrExp(basestack, rest$, frm$) Then
        MissStringExpr
        Exit Function
    End If
        ss$ = GetNextLine(frm$)
        SetNextLine frm$
    If frm$ <> "" Then
        MyEr "Search string with line breaks", "Αλφαριθμητικό αναζήτησης με αλλαγές γραμμών"
        Exit Function
    End If
    Else
        MissPar
        Exit Function
    End If
    
     If FastSymbol(rest$, ",") Then
        If Not IsExp(basestack, rest$, p) Then
            MissNumExpr
            Exit Function
        End If
        x1 = CLng(p)
     Else
        x1 = 0
    End If
    
        If y1 = 3 Then
            If GetVar(basestack, s$, i) Then
                If Typename(var(i)) = doc Then
                
                     x1 = var(i).FindStr(ss$, x1, y1, Col)
                     If x1 > 0 Then
                        basestack.soros.PushVal CDbl(Col)  ' CHAR IN PARAGRAPH
                        basestack.soros.PushVal CDbl(y1)   'PARAGRAPH ORDER ..NUMBER START FROM 1
                     End If
                        basestack.soros.PushVal CDbl(x1)   ' POSITION IN ALL DOCUMENT
                    
                Else
                    MissingDoc
                    Exit Function
                End If
            Else
                   MissFuncParameterStringVar
                    Exit Function
            End If
        ElseIf y1 = 6 Then
                    If pppp.ItemType(i) = doc Then
          
                        x1 = pppp.item(i).FindStr(ss$, x1, y1, Col)
                            If x1 > 0 Then
                               basestack.soros.PushVal CDbl(Col)  ' CHAR IN PARAGRAPH
                               basestack.soros.PushVal CDbl(y1)   'PARAGRAPH ORDER ..NUMBER START FROM 1
                            End If
                        basestack.soros.PushVal CDbl(x1)   ' POSITION IN ALL DOCUMENT
                    
              
                        Else
                         MissingDoc
                         Exit Function
                        End If
                    
    Else
                    
                MissPar
                Exit Function
    End If

End Function
Function ProcSaveDoc(entrypoint As Long, basestack As basetask, rest$) As Boolean
Dim dum As Boolean, w$, i As Long, s$, pppp As mArray, ss$, p As Variant
Dim x1 As Long, y1 As Long, doc1 As Document
If entrypoint = 1 Then dum = True  ' means documend append to file
    y1 = Abs(IsLabel(basestack, rest$, s$))
         If y1 = 6 Then
                If neoGetArray(basestack, s$, pppp) Then
                 If Not NeoGetArrayItem(pppp, basestack, s$, i, rest$) Then Exit Function
                Else
                MissingDoc
                Exit Function
                End If
    End If
    If FastSymbol(rest$, ",") Then
    If Not IsStrExp(basestack, rest$, w$) Then
        If Not dum Then
            If Not IsExp(basestack, rest$, p) Then GoTo sdmess
pass2:
                 If FastPureLabel(rest$, w$, , True) = 1 Then

                    If Not check2(w$, "ΩΣ", "AS") Then SyntaxError: Exit Function
                    
                    If Not Abs(IsLabel(basestack, rest$, w$)) = 3 Then MissingStrVar: Exit Function
                     If y1 = 3 Then
                    If Not GetVar(basestack, s$, i) Then Nosuchvariable s$: Exit Function
                    
                    
                    
                    If Not Typename(var(i)) = doc Then MissingDoc: Exit Function
                        Set doc1 = var(i)
                    ElseIf y1 = 6 Then
                    If Not pppp.ItemType(i) = doc Then MissingDoc: Exit Function
                        Set doc1 = pppp.item(i)
                    
                    End If       '
                    If Not GetVar(basestack, w$, i) Then i = globalvar(w$, "")
                    
                    If dum Then
                    var(i) = doc1.textDocFormated()
                    Else
                    var(i) = doc1.textDocFormated(p)
                    End If
                    ProcSaveDoc = True
                    Exit Function
                Else
                MissPar
                Exit Function
                End If
            
            End If
    
sdmess:
        MissStringExpr
        Exit Function
    End If
    ss$ = GetNextLine(w$)
    SetNextLine w$
    If w$ <> "" Then
    MyEr "filename with line breaks", "όνομα αρχείου με αλλαγές γραμμών"
    Exit Function
    End If
    ' check valid name
    If ExtractNameOnly(ss$, True) = vbNullString Then BadFilename: Exit Function
    If ExtractPath(ss$) = vbNullString Then
    ss$ = mylcasefILE(mcd + ss$)
    End If
    If ExtractType(ss$) = vbNullString Then ss$ = ss$ + ".txt"
    Else
    
                    dum = False
                GoTo pass2
    
    End If
    
     If FastSymbol(rest$, ",") Then
        If Not IsExp(basestack, rest$, p) Then    'type...for saving
        MissNumExpr
        Exit Function
        End If
        x1 = CLng(p)
        If x1 > 500 Then
        
        x1 = 0
        End If
     Else
        x1 = -5 ' ' 2 = utf-8 standard save mode
    End If
    
        If y1 = 3 Then
            If GetVar(basestack, s$, i) Then
                If Typename(var(i)) = doc Then
                         If x1 = -5 Then
                    If var(i).ListLoadedType <> 0 Then
                    x1 = var(i).ListLoadedType
                    Else
                    x1 = 2
                    End If
                    End If
                    If CanKillFile(ss$) Then
                    If x1 = 0 And p <> 0 Then
                    var(i).lcid = CLng(p)
                    x1 = 3
                    End If
                     If Not var(i).SaveUnicodeOrAnsi(ss$, x1, dum) Then
                       MyEr "can't save " + ss$, "δεν μπορώ να σώσω " + ss$
                      End If
                      Else
                      FilePathNotForUser
                      Exit Function
                      End If
                Else
                    MissingDoc
                    Exit Function
                End If
            Else
                   MissFuncParameterStringVar
                    Exit Function
            End If
        ElseIf y1 = 6 Then
                    If pppp.ItemType(i) = doc Then
                    If x1 = -5 Then
                    If pppp.item(i).ListLoadedType <> 0 Then
                    x1 = pppp.item(i).ListLoadedType
                    Else
                    x1 = 2
                    End If
                    End If
                    If CanKillFile(ss$) Then
                    If x1 = 0 And p <> 0 Then
                    pppp.item(i).lcid = CLng(p)
                    x1 = 3
                    End If
                     If Not pppp.item(i).SaveUnicodeOrAnsi(ss$, x1, dum) Then
                       MyEr "can't save " + ss$, "δεν μπορώ να σώσω " + ss$
                       Exit Function
                      End If
                      Else
                      FilePathNotForUser
                      Exit Function
                      End If
                        Else
                         MissingDoc
                         Exit Function
                        End If
                    
    Else
                    
                MissPar
                Exit Function
    End If
ProcSaveDoc = True
End Function
Function ProcWin(basestack As basetask, rest$) As Boolean
Dim s$, w$, x1 As Long
If IsSupervisor Then

x1 = Abs(IsLabelFileName(basestack, rest$, s$, , w$))

If x1 = 1 Then
s$ = w$
Else
rest$ = s$ + rest$
x1 = IsStrExp(basestack, rest$, s$)
End If

If x1 Then
On Error Resume Next

If s$ = ExtractPath$(s$) Then
MyShell "explorer " & Chr(34) + s$ & Chr(34)
Else
If IsSymbol(rest$, ",") Then
x1 = Abs(IsLabelFileName(basestack, rest$, (w$), , w$))
If x1 = 1 Then
If ExtractType(w$) = vbNullString Then w$ = w$ + ".gsb"
If ExtractPath(w$) = vbNullString Then w$ = mcd + w$
Else
rest$ = w$ + rest$
x1 = IsStrExp(basestack, rest$, w$)
End If
If x1 Then

MyShell s$, 1 - 5 * (IsSymbol(rest$, ";")), w$
Else
MissStringExpr
Exit Function
End If
Else
MyShell s$, 1 - 5 * (IsSymbol(rest$, ";"))
End If
End If
'***********************************************
End If
Else
BadCommand
Exit Function
End If
ProcWin = True
End Function

Function ProcDos(basestack As basetask, rest$) As Boolean
Dim s$, w$, x1 As Long, p As Variant
If IsSupervisor Then
On Error Resume Next

x1 = Abs(IsLabelFileName(basestack, rest$, s$, , w$))

If x1 = 1 Then
s$ = w$
Else
rest$ = s$ + rest$
x1 = IsStrExp(basestack, rest$, s$)
End If
If FastSymbol(rest$, ",") Then
If Not IsExp(basestack, rest$, p) Then MissNumExpr: Exit Function
Else
p = 300
End If
        If x1 Then
        
                    If FastSymbol(rest$, ";") Then
                                doslast = Shell("CMD /C " & s$, vbMinimizedNoFocus)
                    Else
                                doslast = Shell("CMD /K " & s$, vbNormalFocus)
                    End If
        Else
                    doslast = Shell("CMD", vbNormalFocus)
        End If

           MyDoEvents
        Sleep CLng(Abs(p))

Else
BadCommand
Exit Function
End If
ProcDos = True
End Function

Function ProcSpeech(basestack As basetask, rest$) As Boolean
Dim s$, p As Variant, dum As Boolean
If IsStrExp(basestack, rest$, s$) Then

If FastSymbol(rest$, "#") Then s$ = "<spell>" & s$ & "</spell>"
dum = FastSymbol(rest$, "!")
If FastSymbol(rest$, ",") Then ' get voice number
If Not IsExp(basestack, rest$, p) Then MissNumExpr:  Exit Function
SPEeCH s$, dum, CLng(p)
Else
SPEeCH s$, dum
End If
End If
ProcSpeech = True
End Function
Function ProcField(bstack As basetask, rest$, Lang As Long) As Boolean
Dim prive As Long, pppp As mArray, s$, it As Long, x As Double, y As Double, p As Variant
Dim i As Long, x1 As Long, y1 As Long, what$, par As Boolean
If Left$(Typename(bstack.Owner), 3) = "Gui" Then oxiforforms: Exit Function
ProcField = True
 prive = GetCode(bstack.Owner)
  With players(prive)
        If IsLabelSymbolNew(rest$, "ΝΕΟ", "NEW", Lang) Then
            If IsExp(bstack, rest$, p) Then Result = p Else Result = 0
            Exit Function
        End If
        If IsLabelSymbolNew(rest$, "ΣΥΝΘΗΜΑ", "PASSWORD", Lang) Then i = True
    
        If Not IsExp(bstack, rest$, x) Then x = .curpos
        If Not FastSymbol(rest$, ",") Then Exit Function
        If Not IsExp(bstack, rest$, y) Then y = .currow
        If FastSymbol(rest$, ",") Then
        If Not IsExp(bstack, rest$, p) Then Exit Function
        p = MyRound(p)
        End If
        If Not IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then Exit Function
        Select Case Abs(IsLabel(bstack, rest$, what$))
        Case 3
            par = False
            If Not GetVar(bstack, what$, it) Then it = globalvar(what$, String$(CLng(p), 32))
            s$ = var(it)
        Case 6
                 par = True
                If neoGetArray(bstack, what$, pppp) Then If Not NeoGetArrayItem(pppp, bstack, what$, it, rest$) Then Exit Function
                s$ = pppp.item(it)
        Case Else
        Exit Function
        End Select
        If p = 0 And s$ = vbNullString Then Exit Function
        If p = 0 Then p = Len(s$)
        s$ = Left$(s$, p)
        s$ = s$ & String$(p - Len(s$), " ")
        x1 = 1
        
        s$ = gf(bstack, y, x, s$, x1, y1, CBool(i))
        
        
        If y1 <> 99 Then
        LCTbasket bstack.Owner, players(prive), y + 1, 0
        End If
        Result = y1
        If par Then
                If pppp.ItemType(it) = doc Then
            Set pppp.item(it) = New Document
            If s$ <> "" Then pppp.item(it).textDoc = s$
            Else
            pppp.item(it) = s$
            End If
        Else
        CheckVar var(it), s$
        End If
        Exit Function
End With
End Function

Function ProcList(basestack As basetask, rest$, Lang As Long) As Boolean
Dim p As Variant, again$, Clsid() As GUID, LNames() As String, clsNumber As Long, i As Long, probe$, what$
Dim usemodule As Boolean, where As Long, page$
If IsLabelSymbolNew(rest$, "ΧΡΗΣΤΩΝ", "USERS", Lang) Then
ProcUsers basestack
ProcList = True
Exit Function
End If

If FastSymbol(rest$, "!") Then
mylist basestack, -2, Lang   ' proportional
ElseIf IsLabelSymbolNew(rest$, "COM", "COM", Lang) Then
If IsLabelSymbolNew(rest$, "ΣΤΟ", "TO", Lang) Then
If here$ <> "" Then
    MyEr "Only in command line interpreter", "Μόνο στον μεταφραστή γραμμής"
    Exit Function
End If

If FastPureLabel(rest$, what$, , True) = 0 Then
    MyEr "Expect a module name", "Περίμενα ένα όνομα τμήματος"
    Exit Function
End If
If Not subHash.Find(what$, where) Then
If Lang = 1 Then
where = ModuleSubAsap(what$, "\\ End for Automatic list" + vbCrLf)
Else
where = ModuleSubAsap(what$, "\\ Τέλος Αυτόματης λίσταq" + vbCrLf)
End If
End If
usemodule = True

End If

'' list com
'' List com to abcModule
If ObjectCatalog.count <> 0 Then again$ = "!"
' use Choose.object + to make it again
If ProcChooseObj(basestack, again$, Lang) Then
' list1 is a glist control
If Form1.List1.ListIndex = -1 Then ProcList = True: Exit Function
    If Form1.List1.listcount > 0 Then
                ObjectCatalog.Index = Form1.List1.ListIndex
        If Not usemodule Then
                basestack.soros.PushStr ObjectCatalog.Value
                ProcList = MyReport(basestack, "letter$", Lang)
        End If
        On Error GoTo there
        If GetAllCoclasses(ObjectCatalog.Value, Clsid(), LNames(), clsNumber) Then
        If Not usemodule Then If clsNumber > 0 Then ProcList = MyReport(basestack, "{CoClasses - Objects}", Lang)
        For i = 0 To clsNumber - 1
        If usemodule Then
            If Lang = 1 Then
                page$ = page$ + "Declare "
            Else
                page$ = page$ + "Όρισε "
            End If
            page$ = page$ + LNames(i) + " " + Chr$(34) + GetGUIDstr(Clsid(i)) + Chr$(34) + vbCrLf
        Else
        probe$ = strProgID(Clsid(i))
        basestack.soros.PushStr GetGUIDstr(Clsid(i))
        If probe$ = vbNullString Then
        basestack.soros.PushStr LNames(i)
        Else
        basestack.soros.PushStr probe$
        End If
        
        ProcList = MyReport(basestack, "quote$(letter$)+{, }+quote$(letter$)", Lang)
        End If
        Next i
         If usemodule Then
         sbf(where).sb = page$ + sbf(where).sb
         End If
        End If
      ProcList = True
    End If
Else
ProcList = True
End If
Exit Function
ElseIf IsExp(basestack, rest$, p) Then
mylist basestack, CLng(p), Lang

Else
mylist basestack, , Lang

  


End If
MyDoEvents1 basestack.Owner
ProcList = True
Exit Function
there:
MyEr Err.Description, Err.Description
Err.Clear
End Function
Function ProcLoad(basestack As basetask, rest$, Lang As Long) As Boolean
Dim x1 As Long, s$, w$, ss$, par As Boolean, vvl As Variant, Key$, par1 As Boolean, NoRun As Boolean
par1 = Not IsLabelSymbolNew(rest$, "ΝΕΟ", "NEW", Lang)
If par1 Then par1 = Not IsLabelSymbolNew(rest$, "ΝΕΑ", "NEW", Lang)  ' PLURAL FOR GREEK
NoRun = IsLabelSymbolNew(rest$, "ΤΜΗΜΑΤΑ", "MODULES", Lang)
ProcLoad = True
With basestack
If Not .IamChild And Not .IamAnEvent And Not .IamThread And Not .IamLambda Then
If Check2Save Then
    rest$ = vbNullString
    Exit Function
End If

If sb2used > 0 And Not NoRun Then
If MsgBoxN(IIf(pagio$ <> "GREEK", "There are modules/functions loaded, load anyway", "Υπάρχουν φορτωμένα τμήματα/συναρτήσεις, να φορτώσω όπως και να έχει"), 1) <> 1 Then

rest$ = vbNullString
Exit Function
End If
End If
End If
End With
Do
x1 = Abs(IsLabelFileName(basestack, rest$, s$, , w$))
If x1 = 1 Then
s$ = w$
Else
rest$ = s$ + rest$
x1 = IsStrExp(basestack, rest$, s$)
End If
If x1 <> 0 Then
   
    If par1 Then
        If loadcatalog.ExistKey(mcd + ExtractNameOnly(s$, True) & ".gsb") Then
            w$ = mcd + ExtractNameOnly(s$, True) & ".gsb"
            ss$ = loadcatalog.Value
            If loadcatalog.ExistKey(w$ + Chr(1)) Then Switches loadcatalog.Value
            par1 = False
            GoTo JUMPHERE
        ElseIf loadcatalog.ExistKey(mcd + ExtractNameOnly(w$, True) & ".gsb") Then
            w$ = mcd + ExtractNameOnly(w$, True) & ".gsb"
            ss$ = loadcatalog.Value
            If loadcatalog.ExistKey(w$ + Chr(1)) Then Switches loadcatalog.Value
            par1 = False
            GoTo JUMPHERE
        End If
    End If
    par1 = True
    If ExtractType(s$) <> "gsb" Then
        Key$ = mcd + ExtractNameOnly(s$, True) & ".gsb"
        s$ = CFname(mcd + ExtractNameOnly(s$, True) & ".gsb")
        If s$ = vbNullString Then
        Key$ = mcd + ExtractNameOnly(w$, True) & ".gsb"
        s$ = CFname(mcd + ExtractNameOnly(w$, True) & ".gsb")
        End If
        If s$ = vbNullString Then
        Key$ = mcd + ExtractNameOnly(w$, True) & ".gsb1"
        s$ = CFname(mcd + ExtractNameOnly(w$, True) & ".gsb1")
        If RenameFile2(s$, mcd + ExtractNameOnly(s$, True) + ".gsb") Then
         s$ = CFname(mcd + ExtractNameOnly(w$, True) & ".gsb")
        End If
        End If
     Else
        ss$ = s$
        If ExtractPath$(s$) = vbNullString Then
            s$ = ExtractName(s$, True)
            ss$ = Trim$(Mid$(ss$, Len(s$) + 1))
            s$ = mcd + s$
        Else
            s$ = ExtractPath(s$) + ExtractName(s$, True)
            ss$ = Trim$(Mid$(ss$, Len(s$) + 1))
        
        End If
        If CFname(s$) = vbNullString Then
        s$ = w$
        ss$ = s$
        If ExtractPath$(s$) = vbNullString Then
            s$ = ExtractName(s$, True)
            ss$ = Trim$(Mid$(ss$, Len(s$) + 1))
            s$ = mcd + s$
        Else
            s$ = ExtractPath(s$) + ExtractName(s$, True)
            ss$ = Trim$(Mid$(ss$, Len(s$) + 1))
        
        End If
        If CFname(s$) = vbNullString Then
            nosuchfile
            ProcLoad = False
            Exit Function
        End If
        End If
        If par1 Then
        If loadcatalog.ExistKey(Key$ + Chr$(1)) Then
            loadcatalog.Value = ss$
        Else
            loadcatalog.AddKey Key$ + Chr$(1), ss$
        End If
        End If
        Switches ss$
    End If
    If ExtractNameOnly(s$, True) = vbNullString Or LenB(CFname(s$)) = 0 Then
        nosuchfile
        ProcLoad = False
        Exit Function
    End If
    Dim oldclid As Long
    oldclid = Clid
    Clid = 1032
    ss$ = ReadUnicodeOrANSI(s$, True)
    ss$ = Join(Split(ss$, vbCrLf), vbLf)
    ss$ = Replace(ss$, vbCr, vbLf)
    ss$ = Join(Split(ss$, vbLf), vbCrLf)
    Clid = oldclid
    If ss$ <> "" Then
   If par1 Then
        If loadcatalog.ExistKey(Key$) Then
            loadcatalog.Value = ss$
        Else
               loadcatalog.AddKey Key$, ss$
        End If
    End If
    End If
JUMPHERE:
    If ss$ <> "" Then
    If Err.Number = 0 And Not (basestack.IamChild Or basestack.IamAnEvent) Then
       If Not NoRun Then LASTPROG$ = s$
        loadcatalog.Remove Key$: loadcatalog.Remove Key$ + Chr(1)
    End If
    If FastSymbol(rest$, ",") Then
        If IsStrExp(basestack, rest$, w$) Then
                ss$ = mycoder.decryptline(ss$, w$, (Len(ss$) / 2) Mod 33)
                If Abs(IsLabel(basestack, ss$, w$)) Then
                        If Not (Left$(ss$, 3) = ":" & vbCrLf) Then ProcLoad = False: Exit Function
                        If Not NORUN1 Then lckfrm = sb2used + 1
                        GoTo skipme2
                End If
        End If
End If
par = False

Do While MaybeIsSymbol(ss$, "\'[*")
SetNextLine ss$
par = True
Loop
' no more exclude tab
'ss$ = Replace(ss$, Chr$(9), "      ")

If ss$ <> "" Then
If par Then GoTo skipme
If (AscW(ss$) > 127 And myUcase(Left$(ss$, 5)) <> "ΚΛΑΣΗ" And myUcase(Left$(ss$, 5)) <> "ΤΜΗΜΑ" And myUcase(Left$(ss$, 9)) <> "ΣΥΝΑΡΤΗΣΗ") Or (((AscW(ss$) And &H4000) = &H4000)) Then
    ss$ = mycoder.must(ss$)
    If NORUN1 Then
        Clipboard.Clear
        SetTextData CF_UNICODETEXT, ss$
        basestack.LoadOnly = True
    End If

    If IsLabelA1("", ss$, w$) Then
        If Not (Left$(ss$, 3) = ":" & vbCrLf) Then ProcLoad = False: Exit Function
        'lock that module
        If Not NORUN1 Then lckfrm = sb2used + 1
    Else
        MOUT = True
    End If
Else
skipme:
    While FastSymbol(ss$, vbCrLf, , 2)
    
      ''  SleepWait 20
    Wend
    If Abs(IsLabel(basestack, ss$, w$)) Then
        If Not (Left$(ss$, 3) = ":" & vbCrLf) Then
        ss$ = w$ & " " & ss$
        End If
    End If
End If
skipme2:
vvl = CStr(vvl) + vbCrLf + ss$ & vbCrLf

End If
End If
End If
Loop Until MOUT Or Not IsSymbol(rest$, "&&", 2)
If lckfrm > 0 Then Resettimestamp
basestack.NoRun = NoRun
ProcLoad = interpret(basestack, CStr(vvl), Len(here$) > 0)
basestack.NoRun = False


End Function

Function MyNew(basestack As basetask, rest$, Lang As Long) As Boolean
MyNew = True
If HaltLevel > 0 Then Exit Function
If Not basestack.IamChild And Not basestack.IamAnEvent Then
        If Check2Save Then
            Exit Function
        End If
End If
Check2SaveModules = False
Resettimestamp
If (basestack.Process Is Nothing) And (basestack.Parent Is Nothing) Then
Set basestack.StaticCollection = Nothing 'New FastCollection
basestack.IamAnEvent = False
abt = False
Set loadcatalog = New FastCollection
LASTPROG$ = vbNullString
Randomize Timer
Set comhash = New sbHash
allcommands comhash
Set numid = New idHash
Set funid = New idHash
Set strid = New idHash
Set strfunid = New idHash
NumberId numid, funid
StringId strid, strfunid
NoOptimum = False
If Lang = 0 Then
sHelp "Μ2000 [ΒΟΗΘΕΙΑ]", "Γράψε ΤΕΛΟΣ για να βγεις από το πρόγραμμα" & vbCrLf & "Δες τα ΟΛΑ (κάνε κλικ στο ΟΛΑ)" & vbCrLf & "George Karras 2018", (ScrInfo(Console).Width - 1) * 3 / 5, (ScrInfo(Console).Height - 1) * 1 / 7
Else
sHelp "Μ2000 [HELP]", "Write END for exit from this program" & vbCrLf & "See ALL commands  (click on ALL)" & vbCrLf & "George Karras 2018", (ScrInfo(Console).Width - 1) * 3 / 5, (ScrInfo(Console).Height - 1) * 1 / 7
End If
NERR = False
lckfrm = 0
subHash.ReduceHash 0, sbf()
sb2used = 0
ReDim sbf(50) As modfun
TaskMaster.Dispose

CloseAllConnections
CleanupLibHandles
ProcPen basestack, ", 255"
' This is the INPUT END
If Not NOEDIT Then
NOEDIT = True
Else
If QRY Then QRY = False
End If
' restore DB.Provider for User
JetPrefixUser = JetPrefixHelp
JetPostfixUser = JetPostfixHelp
' SET ARRAY BASE TO ZERO
ArrBase = 0
End If
If Form1.Visible Then
If Not UseMe Is Nothing Then
UseMe.SetExtCaption "M2000"
End If
End If
End Function
Sub ClearCatalog()
Set ObjectCatalog = New FastCollection
End Sub
Function ProcChooseObj(bstack As basetask, rest$, Lang As Long) As Boolean
Dim F As Long, i As Long, s$, p As Variant
    Dim iSectCount As Long, iSect As Long, sSections() As String
    Dim iVerCount As Long, iVer As Long, sVersions() As String
      Dim iExeSectCount As Long, iExeSect As Long, sExeSect() As String
If Form4Loaded Then
If Form4.Visible Then
Form4.Visible = False
    If Form1.TEXT1.Visible Then
        Form1.TEXT1.SetFocus
    Else
        Form1.SetFocus
    End If
End If
End If
 Form1.List1.Clear
    If lookOne(rest$, "!") Then
            ObjectCatalog.Done = True
            For i = 0 To ObjectCatalog.count - 1
                ObjectCatalog.Index = i
                Form1.List1.additemFast ObjectCatalog.KeyToString
                
            Next i
        ProcChooseObj = MyMenu(2, bstack, rest$, Lang)
    Else
      
       Set ObjectCatalog = New FastCollection
       Dim cr As New cRegistry, first$, k As Long
       Dim bFoundExeSect As Boolean, ss$, mmdir As New recDir
       '' some code here are copies from VbScriptEditor
       '' http://www.codeproject.com/Articles/19986/VbScript-Editor-With-Intellisense
       cr.ClassKey = HKEY_CLASSES_ROOT
       cr.ValueType = REG_SZ
       cr.SectionKey = "TypeLib"
        If cr.EnumerateSections(sSections(), iSectCount) Then
            For iSect = 1 To iSectCount
                cr.SectionKey = "TypeLib\" & sSections(iSect)
                If cr.EnumerateSections(sVersions(), iVerCount) Then
             
                    For iVer = 1 To iVerCount
                    
                            cr.SectionKey = "TypeLib\" & sSections(iSect) & "\" & sVersions(iVer)
                            first$ = cr.Value
                            cr.EnumerateSections sExeSect(), iExeSectCount
                           ''    ObjectCatalog.AddKey cR.Value + sVersions(iVer) + CStr(k), cR.Value + " (" + sVersions(iVer) + ")" + CLSIDToProgID(sSections(iSect))
                            If iExeSectCount > 0 Then
                                       bFoundExeSect = False
                                        For iExeSect = 1 To iExeSectCount
                                            If IsNumeric(sExeSect(iExeSect)) Then
                                            
                                            cr.SectionKey = cr.SectionKey & "\" & sExeSect(iExeSect) & "\win32"
                                                bFoundExeSect = True
                                            Exit For
                     
                                            End If
                                    Next iExeSect
                                                
                                       If bFoundExeSect Then
                                            ss$ = cr.Value
                                            
'
                                             ss$ = ExtractPath(ss$, , True) + ExtractName(ss$, True)
                                             If mmdir.ExistFile(ss$) Then
                                             
                                                ObjectCatalog.AddKey first$ + " (" + sVersions(iVer) + ")", ss$
 
                                               
                                          End If
                                        End If
               
                            End If
                            
                       Next iVer
                     End If
            Next iSect
            ObjectCatalog.Sort
            ObjectCatalog.Done = True
            For i = 0 To ObjectCatalog.count - 1
                ObjectCatalog.Index = i
                Form1.List1.additemFast ObjectCatalog.KeyToString
                
            Next i
            
            ProcChooseObj = MyMenu(2, bstack, rest$, Lang)
        
        Else
            OutOfLimit
            ProcChooseObj = False
        End If
    End If
End Function
Function ProcChooseColor(bstack As basetask, rest$, Lang As Long) As Boolean
Dim p As Variant, i As Long, it As Long, Scr As Object
olamazi
With players(GetCode(bstack.Owner))
If IsExp(bstack, rest$, p) Then
i = CLng(p)
Else
i = -.mypen
End If
it = i
If i > 16 Then it = -it
If i > 0 And i < 16 Then i = QBColor(i)

If TypeOf bstack.Owner Is GuiM2000 Then
Set Scr = bstack.Owner
Else
    If Form1.Visible Then
    Set Scr = Form1
    Else
    Set Scr = Nothing
    End If
End If
DialogSetupLang Lang
If OpenColor(bstack, Scr, i) Then
bstack.soros.PushVal CDbl(-i)
Else
bstack.soros.PushVal CDbl(-it)

End If
Set Scr = Nothing
End With
ProcChooseColor = True
End Function

Function ProcDesktop(bstack As basetask, rest$, Lang As Long) As Boolean
Dim work1 As Boolean, oldleft As Long, oldtop As Long
Dim photo As cDIBSection, s$, p As Variant, x As Double, aPic As StdPicture
olamazi
If IsLabelSymbolNew(rest$, "ΕΙΚΟΝΑ", "IMAGE", Lang) Then
If IsStrExp(bstack, rest$, s$) Then
' FILL WIDTH  IMAGE
 If Left$(s$, 4) = "cDIB" And Len(s$) > 12 Then
 Set photo = New cDIBSection
 If Not cDib(s$, photo) Then MissCdibStr:  Exit Function
  photo.GetDpi 96, 96
  If form5iamloaded Then
  Form5.RestoreSizePos
  Form5.Cls
  photo.ThumbnailPaint Form5
  Else
  photo.ThumbnailPaint Form1
  End If
 Else
 If ExtractType(s$) = vbNullString Then s$ = s$ & ".jpg"
                    If CFname(s$) = vbNullString Then
                        s$ = mcd & s$
                        If CFname(s$) = vbNullString Then
                        BadFilename
                        Exit Function
                        End If
                    Else
                        s$ = CFname(s$)
                    End If
        If Len(s$) < 254 Then
        ' look for image to load
            Set photo = New cDIBSection
            If CFname(s$) <> "" Then
             s$ = CFname(s$)
                                       Set aPic = LoadMyPicture(GetDosPath(s$))
                If Not aPic Is Nothing Then
                
                    photo.CreateFromPicture aPic
                                           
                    If photo.bitsPerPixel <> 24 Then
                        Conv24 photo
                        Else
                        CheckOrientation photo, s$
                        End If
                       photo.GetDpi 96, 96
                       If form5iamloaded Then
                       Form5.RestoreSizePos
                       Form5.Cls
                       photo.ThumbnailPaint Form5
                       Else
                       photo.ThumbnailPaint Form1
                       End If
                    End If
                    End If
        Else
        BadFilename
        End If
        End If
 Set photo = Nothing
End If
ElseIf IsLabelSymbolNew(rest$, "ΚΡΥΨΕ", "HIDE", Lang) Then
If Not form5iamloaded Then
'
End If
Form5.backcolor = &H0 ' ALWAYS BLACK
Form5.Cls
SetTrans Form5, CByte(255), mycolor(0), True
ElseIf IsLabelSymbolNew(rest$, "ΚΑΘΑΡΗ", "CLEAR", Lang) Then
If form5iamloaded Then
Form5.RestoreSizePos
Form5.backcolor = &H0 ' ALWAYS BLACK
Form5.Cls

Set Form5.Picture = LoadPicture("")
Form5.Cls
SetTrans Form5, CByte(255), mycolor(-2)

Else
Form1.Cls
End If
Else
If Not Form1.Visible Then
work1 = True
oldleft = Form1.Left
oldtop = Form1.Left
Form1.move -Form1.Width - dv15, -Form1.Height - dv15
Form1.Visible = True
End If
If IsExp(bstack, rest$, p) Then
    If FastSymbol(rest$, ",") Then
        If IsExp(bstack, rest$, x) Then
        
        Form5.Visible = True
        Form5.ZOrder 1
        SetTrans Form1, CByte(p And &HFF), mycolor(x), True
        
        End If
    Else

    Form5.Visible = True
    'Form5.ZOrder 1
    SetTrans Form1, CByte(p And &HFF)
    End If
    Else
    CdESK
    End If
End If
If work1 Then
    Form1.Visible = False
    Form1.move oldleft, oldtop
End If

ProcDesktop = True
End Function
Function ProcFont(bstack As basetask, rest$, Lang As Long) As Boolean
Dim prive As Long, x1 As Long, s$
If IsLabelSymbolNew(rest$, "ΦΟΡΤΩΣΕ", "LOAD", Lang) Then
Do
If IsStrExp(bstack, rest$, s$) Then
    s$ = CFname$(s$)
    If s$ <> "" Then
        ProcFont = LoadFont(s$)
    End If
Else
    MissStringExpr
    ProcFont = False
    Exit Function
End If
Loop Until Not FastSymbol(rest$, ",")
ElseIf IsLabelSymbolNew(rest$, "ΔΙΑΓΡΑΦΗ", "REMOVE", Lang) Then
Do
If IsStrExp(bstack, rest$, s$) Then
    s$ = CFname$(s$)
    If s$ <> "" Then
        ProcFont = RemoveFont(s$)
    End If
Else
    MissStringExpr
    ProcFont = False
    Exit Function
End If
Loop Until Not FastSymbol(rest$, ",")
Else

    prive = GetCode(bstack.Owner)
    If IsStrExp(bstack, rest$, s$) Then
        On Error Resume Next
        x1 = bstack.Owner.Font.charset
        bstack.Owner.Font.Name = s$
        If Not (x1 = bstack.Owner.Font.charset) Then
            bstack.Owner.Font.charset = x1
        End If
    
        If LCase(bstack.Owner.Font.Name) <> LCase(s$) Then
        
            bstack.Owner.Font.Name = MyFont
            bstack.Owner.Font.charset = bstack.myCharSet
        End If
    End If
        StoreFont bstack.Owner.Font.Name, players(prive).SZ, bstack.Owner.Font.charset
        players(prive).FontName = bstack.Owner.Font.Name
        SetText bstack.Owner
        GetXYb bstack.Owner, players(prive), players(prive).curpos, players(prive).currow
End If
 ProcFont = True
End Function

Function ProcSubDir(basestack As basetask, rest$, Lang As Long) As Boolean
Dim x1 As Long, ss$, w$

x1 = Abs(IsLabelFileName(basestack, rest$, ss$, , w$))

If x1 = 1 Then
    ss$ = w$
ElseIf x1 = 0 Or x1 = 3 Or x1 = 6 Then
    rest$ = ss$ + rest$
    x1 = IsStrExp(basestack, rest$, ss$)
End If
If x1 <> 0 Then
    ss$ = mcd + ss$
    AddDirSep ss$
    If PathMakeDirs(ss$) Then
        mcd = ss$
        ProcSubDir = True
    Else
        BadPath
    End If
Else
MissDir
End If

End Function
Function ProcOpenFile(basestack As basetask, rest$, Lang As Long) As Boolean
Dim pa$, ss$, frm$, s$, w$, Scr As Object, x1 As Long, p As Variant, par As Boolean, F As Boolean
Dim aaa() As String, dum As Boolean
If IsSelectorInUse Then
SelectorInUse
Exit Function
End If
olamazi
frm$ = mcd
DialogSetupLang Lang

IsStrExp basestack, rest$, s$
If FastSymbol(rest$, ",") Then If IsStrExp(basestack, rest$, pa$) Then frm$ = pa$
If frm$ <> "" Then If Not isdir(frm$) Then NoSuchFolder: Exit Function
If FastSymbol(rest$, ",") Then If IsStrExp(basestack, rest$, pa$) Then ss$ = pa$
par = False
If FastSymbol(rest$, ",") Then If Not IsStrExp(basestack, rest$, w$) Then Exit Function
If FastSymbol(rest$, ",") Then If IsExp(basestack, rest$, p) Then par = p <> 0
If FastSymbol(rest$, ",") Then F = IsExp(basestack, rest$, p) Else p = 0  '' if f is false what???
 dum = p <> 0
If TypeOf basestack.Owner Is GuiM2000 Then
Set Scr = basestack.Owner
Else
 If Form1.Visible Then
Set Scr = Form1
Else
Set Scr = Nothing
End If
End If
If InStr(w$, "|") > 0 Then
    If InStr(w$, "(*.") > 0 Then
        aaa() = Split(w$, "(*.")
        w$ = vbNullString
        If UBound(aaa()) > LBound(aaa()) Then
            w$ = "|"
            For x1 = LBound(aaa()) + 1 To UBound(aaa())
                w$ = w$ & UCase(Left$(aaa(x1), InStr(aaa(x1), ")") - 1) & "|")
            Next x1
        End If
    Else
        aaa() = Split(w$, "|")
        w$ = vbNullString
        If UBound(aaa()) > LBound(aaa()) Then
            w$ = "|"
            For x1 = LBound(aaa()) To UBound(aaa())
                w$ = w$ & UCase(aaa(x1)) & "|"
            Next x1
        End If
    End If
End If

    If OpenDialog(basestack, Scr, frm$, s$, ss$, w$, Not par, dum) Then
     If multifileselection Then
        If ReturnListOfFiles <> "" Then
                aaa() = Split(ReturnListOfFiles, "#")
                If UBound(aaa()) > LBound(aaa()) Then
            
                        For x1 = UBound(aaa()) To LBound(aaa()) + 1 Step -1
                            basestack.soros.PushStr aaa(x1)
                        Next x1
                        basestack.soros.PushVal UBound(aaa()) - LBound(aaa())
                        basestack.soros.PushStr aaa(x1)
                 End If
            Else
            
     If isdir(ReturnFile) Then
     basestack.soros.PushStr ""
    Else
    basestack.soros.PushStr ReturnFile
    End If
        End If
    Else
    If isdir(ReturnFile) Then
     basestack.soros.PushStr ""
    Else
    basestack.soros.PushStr ReturnFile
    End If
    End If
    Else
    basestack.soros.PushStr ""
    End If

Set Scr = Nothing
ProcOpenFile = True
End Function

Function ProcTone(bstack As basetask, rest$) As Boolean
Dim p As Variant, sX As Double
If IsExp(bstack, rest$, p) Then
    If Not FastSymbol(rest$, ",") Then
       Beeper 1000, p
    ElseIf IsExp(bstack, rest$, sX) Then
    Beeper sX, p
    Else
    MyEr "wrong parameter", "λάθος παράμετρος"
    Exit Function
    End If
Else
Beeper 1000, 100
End If
ProcTone = True
End Function
Function ProcGradient(bstack As basetask, rest$) As Boolean
Dim x As Long, y As Long, p As Variant, trans As Long, useold As Boolean

ProcGradient = True
With players(GetCode(bstack.Owner))
trans = .mypentrans
useold = .NoGDI
If Not IsExp(bstack, rest$, p) Then x = rgb(255, 255, 255) Else x = mycolor(p)
If Not FastSymbol(rest$, ",") Then
y = 0
Else
If Not IsExp(bstack, rest$, p) Then y = 0 Else y = mycolor(p)
End If
If Not FastSymbol(rest$, ",") Then
If useold Then
TwoColorsGradient bstack.Owner, GRADIENT_FILL_RECT_V, GDIP_ARGB1(trans, x), GDIP_ARGB1(trans, y)
Else
GdiPlusGradient bstack.Owner.Hdc, 0, 0, bstack.Owner.Scalewidth / dv15 + 1, bstack.Owner.Scaleheight / dv15 + 1, GDIP_ARGB1(trans, x), GDIP_ARGB1(trans, y), 1
End If



Else
If Not IsExp(bstack, rest$, p) Then
ProcGradient = IfierVal: Exit Function
Else
If useold Then
TwoColorsGradient bstack.Owner, -CLng(p <> 0), GDIP_ARGB1(trans, x), GDIP_ARGB1(trans, y)
Else
If p = 0 Then
GdiPlusGradient bstack.Owner.Hdc, 0, 0, bstack.Owner.Scalewidth / dv15 + 1, bstack.Owner.Scaleheight / dv15 + 1, GDIP_ARGB1(trans, y), GDIP_ARGB1(trans, x), 0&
Else
GdiPlusGradient bstack.Owner.Hdc, 0, 0, bstack.Owner.Scalewidth / dv15 + 1, bstack.Owner.Scaleheight / dv15 + 1, GDIP_ARGB1(trans, x), GDIP_ARGB1(trans, y), 1&
End If
End If




End If
End If
End With
End Function
Public Function GetSpecialfolder(CSIDL As Long) As String
    Dim R As Long
    Dim IDL As ITEMIDLIST, NoError As Long, Path$
    'Get the special folder
    R = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If R = NoError Then
        'Create a buffer
        Path$ = space$(512)
        'Get the path from the IDList
        R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        'Remove the unnecessary chr$(0)'s
        GetSpecialfolder = mylcasefILE(Left$(Path, InStr(Path, Chr$(0)) - 1))
        Exit Function
    End If
    GetSpecialfolder = vbNullString
End Function
Sub ProcUsers(bstack As basetask)
Dim aDir As New recDir, ss$, a$, b$, n As Integer
aDir.IncludedFolders = True
aDir.Nofiles = True
aDir.TopFolder = GetSpecialfolder(CLng(26)) & "\M2000_USER\"
aDir.LevelStop = 1
aDir.SortType = 1

b$ = GetSpecialfolder(CLng(26)) & "\M2000_USER\"
n = Len(b$) + 2
a$ = Tcase(Mid$(mylcasefILE$(aDir.Dir2$(b$, "", False)), n))
b$ = vbNullString
While a$ <> ""
If InStr(a$, " ") > 0 Then a$ = "[" + Replace(a$, " ", ChrW(160)) + "]"
b$ = a$

a$ = Tcase(Mid$(aDir.Dir2, n))
If a$ <> "" Then ss$ = ss$ + b$ + ", "
Wend
If b$ <> "" Then ss$ = ss$ + b$
Dim Scr As Object, prive As Long
Set Scr = Form1.DIS 'bstack.Owner
prive = GetCode(Scr)
wwPlain2 bstack, players(prive), ss$, Scr.Width, 1000, True, , 3
End Sub

Function MyFrame(bstack As basetask, rest$) As Boolean
Dim prive As Long, x1 As Long, y1 As Long, Col As Long, p As Variant
Dim x As Double, y As Double, ss$
MyFrame = True
prive = GetCode(bstack.Owner)
With players(prive)
x1 = 1
y1 = 1
Col = .mypen
If FastSymbol(rest$, "@") Then
If FastSymbol(rest$, "(") Then
    If IsExp(bstack, rest$, p) Then x1 = Abs(p + .curpos) Mod (.mx + 1)
    If Not FastSymbol(rest$, ")") Then MissSymbol ")": Exit Function
Else
    If IsExp(bstack, rest$, p) Then x1 = Abs(p) Mod (.mx + 1)
End If
If FastSymbol(rest$, ",") Then
    If FastSymbol(rest$, "(") Then
        If IsExp(bstack, rest$, p) Then y1 = Abs(p + .currow - 1) Mod (.My + 1)
        If Not FastSymbol(rest$, ")") Then MissSymbol ")": Exit Function
    
    Else
        If IsExp(bstack, rest$, p) Then y1 = Abs(p) Mod (.My + 1)
    End If
    '
    
End If
y = 5
If FastSymbol(rest$, ",") Then If Not IsExp(bstack, rest$, y) Then y = 5
If FastSymbol(rest$, ",") Then
If IsExp(bstack, rest$, x) Then
If FastSymbol(rest$, ",") Then
If IsExp(bstack, rest$, p) Then
MyRect bstack.Owner, players(prive), (x1), (y1), (y), (x), (p)
Else
 MyFrame = False: MissNumExpr: Exit Function
End If
Else
MyRect bstack.Owner, players(prive), (x1), (y1), (y), (x)
End If
ElseIf IsStrExp(bstack, rest$, ss$) Then
MyRect bstack.Owner, players(prive), (x1), (y1), (y), ss$
Else
MyRect bstack.Owner, players(prive), (x1), (y1), 5, "?"
End If
Else
MyRect bstack.Owner, players(prive), (x1), (y1), 6, 0
End If
Else
If IsExp(bstack, rest$, p) Then x1 = Abs(p) Mod .mx
If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then y1 = Abs(p) Mod .My

x1 = x1 + .curpos - 1
y1 = y1 + .currow - 1
If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then BoxColorNew bstack.Owner, players(prive), x1, y1, (p)


If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then Col = p Else MyFrame = False: MissNumExpr: Exit Function


BoxBigNew bstack.Owner, players(prive), x1, y1, Col
End If
End With
MyDoEvents1 bstack.Owner


End Function
Function MyMark(bstack As basetask, rest$) As Boolean
Dim prive As Long, p As Variant, par As Boolean, x1 As Long, y1 As Long, Col As Long
MyMark = True
prive = GetCode(bstack.Owner)
With players(prive)
x1 = 1
y1 = 1
Col = .mypen
If IsExp(bstack, rest$, p) Then x1 = Abs(p) Mod .mx
If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then y1 = Abs(p) Mod .My
x1 = x1 + .curpos - 1
y1 = y1 + .currow - 1
If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then Col = p

If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then par = Not (p = 0)

CircleBig bstack.Owner, players(prive), x1, y1, Col, par
End With
MyDoEvents1 bstack.Owner


End Function
Function MyLineInput(bstack As basetask, rest$, Lang As Long) As Boolean
Dim F As Long, p As Variant, what$, it As Long, s$, i As Long, prive As Long, frm$
Dim pppp As mArray
If IsLabelSymbolNew(rest$, "ΕΙΣΑΓΩΓΗΣ", "INPUT", Lang) Then
If FastSymbol(rest$, "#") Then

    If Not IsExp(bstack, rest$, p) Then Exit Function
    If Not FastSymbol(rest$, ",") Then Exit Function
    F = CLng(MyMod(p, 512))
    Select Case Abs(IsLabel(bstack, rest$, what$))
    Case 3
    MyLineInput = True
    If uni(F) Then
    If Not getUniStringlINE(F, s$) Then MyLineInput = False: MyEr "Can't input, not UTF16LE", "Δεν μπορώ να εισάγω, όχι UTF16LE": Exit Function
    Else
    getAnsiStringlINE F, s$
    End If
    If GetVar(bstack, what$, i) Then
    CheckVar var(i), s$
    Else
    globalvar what$, s$
    End If
    Case 6
    If neoGetArray(bstack, what$, pppp) Then

    If Not NeoGetArrayItem(pppp, bstack, what$, it, rest$) Then Exit Function
    Else
    Exit Function
    End If
    MyLineInput = True
    If uni(F) Then
    If Not getUniStringlINE(F, s$) Then MyLineInput = False: MyEr "Can't input, not UTF16LE", "Δεν μπορώ να εισάγω, όχι UTF16LE": Exit Function
    Else
    getAnsiStringlINE F, s$
    End If
    If pppp.ItemType(it) = doc Then
    Set pppp.item(it) = New Document
    If s$ <> "" Then pppp.item(it).textDoc = s$
    Else
    pppp.item(it) = s$
    End If
    End Select
Else
If Not releasemouse Then If Not Form1.Visible Then newshow Basestack1
If bstack.toprinter = True Then oxiforPrinter:   Exit Function
If Left$(Typename(bstack.Owner), 3) = "Gui" Then oxiforforms: Exit Function
Select Case Abs(IsLabel(bstack, rest$, what$))
Case 3
           prive = GetCode(bstack.Owner)
                If players(prive).lastprint Then
                LCTbasket bstack.Owner, players(prive), players(prive).currow, players(prive).curpos
                players(prive).lastprint = False
                End If
QUERY bstack, frm$, s$, 1000, False

                If GetVar(bstack, what$, i) Then
                        CheckVar var(i), s$
                Else
                        globalvar what$, s$
                End If
                 MyLineInput = True
Case 6
If neoGetArray(bstack, what$, pppp) Then
                       If Not NeoGetArrayItem(pppp, bstack, what$, it, rest$) Then Exit Function
                Else
                 MyEr "No such array", "Δεν υπάρχει τέτοιος πίνακας"
                       Exit Function
                End If
                           prive = GetCode(bstack.Owner)
                If players(prive).lastprint Then
                LCTbasket bstack.Owner, players(prive), players(prive).currow, players(prive).curpos
                players(prive).lastprint = False
                End If
QUERY bstack, frm$, s$, 1000, False

 If pppp.ItemType(it) = doc Then
                Set pppp.item(it) = New Document
                        If s$ <> "" Then pppp.item(it).textDoc = s$
                Else
                        pppp.item(it) = s$
                End If
                 MyLineInput = True
End Select

End If


End If
End Function

Function MyLong(basestack As basetask, rest$, Lang As Long, Optional alocal As Boolean) As Boolean
Dim s$, what$, i As Long, p As Variant
MyLong = True

     Do While CheckTwoVal(Abs(IsLabel(basestack, rest$, what$)), 1, 4)
     If basestack.priveflag Then what$ = ChrW(&HFFBF) + what$
     If Not FastSymbol(rest$, "<") Then  ' get local var first
            If alocal Then
            i = globalvar(basestack.GroupName & what$, s$)     ' MAKE ONE  '

             GoTo makeitnow1
            ElseIf GetlocalVar(basestack.GroupName & what$, i) Then
            p = var(i)
            GoTo there01
            ElseIf GetVar(basestack, basestack.GroupName & what$, i) Then
             p = var(i)
            GoTo there01
            Else
            i = globalvar(basestack.GroupName & what$, s$)     ' MAKE ONE  '

             GoTo makeitnow1
            End If
            ElseIf GetVar(basestack, basestack.GroupName & what$, i) Then
            
there01:
                
                MakeitObjectLong var(i)
                On Error Resume Next
                Err.Clear
                CheckVarLong var(i), CLng(Int(p))
                If Err > 0 Then
                If Err.Number = 6 Then MyEr "overflowlong Long", "Υπερχείλιση  μακρύ"
                Err.Clear
                MyLong = False
                Exit Function
                End If
                GoTo there12
            Else
        
                i = globalvar(basestack.GroupName & what$, s$) ' MAKE ONE
                If i <> 0 Then
makeitnow1:
                    MakeitObjectLong var(i)
there12:
                    If FastSymbol(rest$, "=") Then
                        If IsExp(basestack, rest$, p, , True) Then
                          On Error Resume Next
                            Err.Clear
                            CheckVarLong var(i), CLng(Int(p))
                            If Err > 0 Then
                            If Err.Number = 6 Then MyEr "overflowlong Long", "Υπερχείλιση  μακρύ"
                            Err.Clear
                            MyLong = False
                            Exit Function
                            End If
                        Else
                            MissNumExpr
                            MyLong = False
                        End If
                    Else
                    ' DO NOTHING
                    End If
                End If
            End If
     
     If Not FastSymbol(rest$, ",") Then Exit Do
     Loop
End Function

Function ProcSoundRec(basestack As basetask, rest$, Lang As Long) As Boolean
' not tested yet...
Dim s$, p As Variant, ss$, x As Double, y As Double
Dim LRec As RecordMci
    
    
    If IsLabelSymbolNew(rest$, "ΝΕΑ", "NEW", Lang) Then
        Set sRec = New RecordMci
        Set LRec = sRec
        If Not LRec.HaveMic Then GoTo noMic
        LRec.Rec_Initialize
        If IsStrExp(basestack, rest$, s$) Then LRec.FileName = s$
        If FastSymbol(rest$, ",") Then
            If IsExp(basestack, rest$, p, , True) Then
                If IsLabelSymbolLatin(rest$, "STEREO") Then
                    LRec.Stereo
                Else
                    LRec.Mono
                End If
                If IsLabelSymbolLatin(rest$, "HIFI") Then
                    LRec.Bit16
                Else
                    LRec.Bit8
                End If
                LRec.QualityAny CDbl(p)
        End If
    Else
        If sRec Is Nothing Then
         Set sRec = New RecordMci
         If Not LRec.HaveMic Then GoTo noMic
        End If
        Set LRec = sRec
        LRec.RecFast
    End If
    ElseIf Not (sRec Is Nothing) Then
        Set LRec = sRec
    ss$ = vbNullString
    
    If IsLabelSymbolNewExp(rest$, "ΕΙΣΑΓΩΓΗ", "INSERT", Lang, ss$) Then
        LRec.Capture True
    ElseIf IsLabelSymbolNewExp(rest$, "ΑΛΛΑΓΗ", "OVERWRITE", Lang, ss$) Then
        LRec.ReCapture
    ElseIf IsLabelSymbolNewExp(rest$, "ΑΠΟΚΟΠΗ", "DELETE", Lang, ss$) Then
            If IsExp(basestack, rest$, x, , True) Then
            Else
                x = 0
            End If
            If IsLabelSymbolNew(rest$, "ΕΩΣ", "TO", Lang) Then
                If Not IsExp(basestack, rest$, y, , True) Then
                    y = LRec.getLengthInMS
                End If
            Else
                y = LRec.getLengthInMS
            End If
            LRec.CutRecordMs CDbl(x), CDbl(y)
    ElseIf IsLabelSymbolNewExp(rest$, "ΕΝΤΑΣΗ", "VOLUME", Lang, ss$) Then
                If IsExp(basestack, rest$, x, , True) Then
                    If x < 0 Then x = 0
                    If x > 100 Then x = 100
                    LRec.setVolume CLng(x)
                Else
                    LRec.setVolume 50&
                End If
    ElseIf IsLabelSymbolNewExp(rest$, "ΔΙΑΚΟΠΗ", "STOP", Lang, ss$) Then
        LRec.recStop
    ElseIf IsLabelSymbolNewExp(rest$, "ΔΟΚΙΜΗ", "TEST", Lang, ss$) Then
        LRec.recPlay
    ElseIf IsLabelSymbolNewExp(rest$, "ΘΕΣΗ", "POS", Lang, ss$) Then
        If LRec.isRecPlaying Then
            If IsExp(basestack, rest$, x, , True) Then
            LRec.recPlayFromMs x
            Else
            LRec.recPlay
            End If
        Else
        ' SEEK
            If IsExp(basestack, rest$, x, , True) Then
            LRec.oneMCI "seek capture to " & CStr(CLng(x))
            Else
            LRec.oneMCI "seek capture to 0"
            End If
        End If
    ElseIf IsLabelSymbolNewExp(rest$, "ΣΩΣΕ", "SAVE", Lang, ss$) Then
        If IsStrExp(basestack, rest$, s$) Then
            LRec.SaveAs s$
        Else
            LRec.Save
        End If
    ElseIf IsLabelSymbolNewExp(rest$, "ΚΛΕΙΣΕ", "END", Lang, ss$) Then
        Set sRec = Nothing
        Set LRec = Nothing
    End If
    Else
        
        MyEr "You don't have new recording", "Δεν έχεις ετοιμάσει νέα ηχογράφηση"
    End If
  ProcSoundRec = True
  Exit Function
noMic:
MissMic
  
End Function

Public Function ScanTarget(j() As target, ByVal x As Long, ByVal y As Long, ByVal myl As Long) As Long
Dim iu&, id&, i&, xx&, YY&

iu& = LBound(j())
id& = UBound(j())
ScanTarget = -1
For i& = iu& To id&
With j(i&)
If .Enable And .layer = myl Then
xx& = x \ .Xt
YY& = y \ .Yt
If .Lx <= xx& And .tx >= xx& And .ly <= YY& And .ty >= YY& Then
ScanTarget = i&
Exit For
End If
End If
End With
Next i&
End Function
Function ProcMedia(basestack As basetask, rest$, Lang As Long) As Boolean
Dim Scr As Object
Dim s$, ss$, x As Double, y As Double
Set Scr = basestack.Owner
On Error Resume Next
ProcMedia = True
If IsLabelSymbolNew(rest$, "ΦΟΡΤΩΣΕ", "LOAD", Lang) Then
            If AVIUP Then
                  AVI.GETLOST
                  MyDoEvents
            End If
            If IsStrExp(basestack, rest$, s$) Then
                If s$ <> "" Then
                    If ExtractType(s$) = vbNullString Then s$ = s$ & ".avi"
                    If CFname(s$) = vbNullString Then
                        s$ = mcd & s$: If CFname(s$) = vbNullString Then Exit Function
                    Else
                        s$ = CFname(s$)
                    End If
                Else
                    Set Scr = Nothing
                    ProcMedia = True  ' ??????????
                    Exit Function
                End If
                avifile = s$
                Load AVI
                If Not OsInfo.IsWindows8Point1OrGreater Then
                MediaPlayer1.playMovie
                MediaPlayer1.pauseMovie
                MediaPlayer1.setPositionTo 0
                End If
                
                Sleep 2
                MyDoEvents
                AVIRUN = False
                    If Form1.Visible Then Form1.SetFocus
                    'MediaPlayer1.setLeftVolume vol * 10
                    'MediaPlayer1.setRightVolume vol * 10
                    
                MediaPlayer1.sizeLocateMovie AVI.Left \ dv15, AVI.top \ dv15, AVI.Width \ dv15, AVI.Height \ dv15
                
            End If
            Set Scr = Nothing
            ProcMedia = True
            Exit Function
            
    ElseIf AVIUP Then
    ss$ = vbNullString
        If IsLabelSymbolNewExp(rest$, "ΔΕΙΞΕ", "SHOW", Lang, ss$) Then
            'If Not AVIRUN Then MediaPlayer1.playMovie: MediaPlayer1.pauseMovie
            
            
            If Scr.Name = "GuiM2000" Then
            If Scr.Visible Then
                
                AVI.Show , Scr
                'MediaPlayer1.sizeLocateMovie 0, 0, AVI.Width \ dv15, AVI.Height \ dv15 + 1
                MediaPlayer1.showMovie
      End If
                Set Scr = Nothing
                ProcMedia = True
                Exit Function
           
                
            Else
                If Form1.Visible Then
                AVI.Show , Form1
                Else
                AVI.Show , Form5
                End If
               'MediaPlayer1.sizeLocateMovie 0, 0, AVI.Width \ dv15, AVI.Height \ dv15 + 1
               MediaPlayer1.showMovie
                AVI.ZOrder 0
             AVI.SetFocus
                MyDoEvents
                Set Scr = Nothing
                ProcMedia = True
                Exit Function
       End If
        ElseIf IsLabelSymbolNewExp(rest$, "ΚΡΥΨΕ", "HIDE", Lang, ss$) Then
                AVI.Hide
                Set Scr = Nothing
                ProcMedia = True
                Exit Function
        ElseIf IsLabelSymbolNewExp(rest$, "ΚΡΑΤΗΣΕ", "PAUSE", Lang, ss$) Then
                If MediaPlayer1.isMoviePlaying Then MediaPlayer1.pauseMovie
                Set Scr = Nothing
                ProcMedia = True
                Exit Function
        ElseIf IsLabelSymbolNewExp(rest$, "ΠΑΙΞΕ", "PLAY", Lang, ss$) Then
        
                If Not AVIRUN Then
                AVI.Interval = MediaPlayer1.getLengthInMS - MediaPlayer1.getPositionInMS
                AVI.Avi2Up
                End If
        
                MyDoEvents
                Set Scr = Nothing
                ProcMedia = True
                Exit Function
        ElseIf IsLabelSymbolNewExp(rest$, "ΞΕΚΙΝΑ", "RESTART", Lang, ss$) Then
                    If Not MediaPlayer1.isMoviePlaying Then
                    
                         MediaPlayer1.playMovie
                         
                    Else
                         MediaPlayer1.resumeMovie
                    End If
                    MyDoEvents
                    AVIRUN = False
                    Set Scr = Nothing
                    ProcMedia = True
                    Exit Function
        ElseIf IsLabelSymbolNewExp(rest$, "ΣΤΟ", "TO", Lang, ss$) Then
                    If IsExp(basestack, rest$, x) Then
                        If MediaPlayer1.getLengthInMS > 0 Then MediaPlayer1.setPositionTo x
                        
                    End If
                    Set Scr = Nothing
                    ProcMedia = True
                    Exit Function

        End If
    ElseIf IsLabelSymbolNewExp(rest$, "ΚΡΥΨΕ", "HIDE", Lang, ss$) Then
    
    End If
    ss$ = vbNullString
' do nothing until here
If IsExp(basestack, rest$, x) Then
   
            If FastSymbol(rest$, ",") Then
    
             UseAviSize = False
    AviSizeX = 0
    AviSizeY = 0
    aviX = 0
    aviY = 0
    UseAviSize = False
    UseAviXY = True: aviX = CLng(x): aviY = 0
            If IsExp(basestack, rest$, y) Then aviY = CLng(y) Else ProcMedia = False: UseAviXY = False: aviX = 0
            Else ' SPECIAL
            If MediaPlayer1.getLengthInMS > 0 Then
                If x < 0 Then
                MediaPlayer1.pauseMovie
                AVIRUN = MediaPlayer1.isMoviePlaying
                If Scr.Name <> "Printer" Then
                If Scr.Visible Then Scr.SetFocus
                End If
                ElseIf x = 0 Then
                          
                MediaPlayer1.playMovie
                MyDoEvents
                Else
                MediaPlayer1.setPositionTo x
                End If
                ProcMedia = True
        Else
        ProcMedia = False
        End If
        Set Scr = Nothing
        Exit Function
            End If
            If aviX = 0 Then UseAviXY = False
            If FastSymbol(rest$, ",") Then
                    If IsExp(basestack, rest$, x) Then AviSizeX = CLng(x) Else rest$ = "," & rest$
                If FastSymbol(rest$, ",") Then
            If IsExp(basestack, rest$, x) Then AviSizeY = CLng(x) Else rest$ = "," & rest$
                End If
                UseAviSize = (Abs(AviSizeY) + Abs(AviSizeX)) <> 0 Or (aviX = 0 And aviY = 0)
                
            End If
            If Not FastSymbol(rest$, ",") Then
                   If AVIUP Then
                   If UseAviXY And UseAviSize Then
                   AVI.move aviX, aviY, AviSizeX, AviSizeY
                   MediaPlayer1.sizeLocateMovie 0, 0, AviSizeX \ dv15, AviSizeY \ dv15 + 1
                   ElseIf UseAviXY Then
                   AVI.move aviX, aviY
                    MediaPlayer1.sizeLocateMovie 0, 0, AVI.Width \ dv15, AVI.Height \ dv15 + 1
                    ElseIf UseAviSize Then
                     AVI.move AVI.Left, AVI.top, AviSizeX, AviSizeY
                   MediaPlayer1.sizeLocateMovie 0, 0, AviSizeX \ dv15, AviSizeY \ dv15 + 1
                    End If
                    If AVI.Visible Then AVI.Refresh
           Else
                   If AVIRUN Or AVIUP Then
                AVI.GETLOST
            End If
           
            End If
            Set Scr = Nothing
            Exit Function
            
            End If
      
ElseIf FastSymbol(rest$, ";") Then
'MediaPlayer1.closeMovie
    UseAviXY = False
    UseAviSize = False
    AviSizeX = 0
    AviSizeY = 0
    aviX = 0
    aviY = 0
    AVI.GETLOST
Else
 
'MediaPlayer1.closeMovie
If AVIRUN Or AVIUP Then
                AVI.GETLOST
              
            End If
  
End If

Do
ProcTask2 basestack

 If Not MediaPlayer1.isMoviePlaying Then AVIRUN = False
Loop Until Not AVIRUN Or NOEXECUTION

Do While IsStrExp(basestack, rest$, s$)
If s$ <> "" Then
If ExtractType(s$) = vbNullString Then s$ = s$ & ".avi"
    If CFname(s$) = vbNullString Then
        s$ = mcd & s$: If CFname(s$) = vbNullString Then Set Scr = Nothing: Exit Function

    Else
        s$ = CFname(s$)
    End If
    Else
    AVI.GETLOST
    Exit Do
End If
avifile = s$
Load AVI
AVI.Avi2Up



 AVI.Show
Sleep 5

If AVIRUN Then
If AVI.Height > 0 Then If Form1.Visible Then Form1.SetFocus
MediaPlayer1.setLeftVolume vol * 10
MediaPlayer1.setRightVolume vol * 10

End If
If FastSymbol(rest$, ",") Then
If AVIRUN Then
Do
 AVIRUN = MediaPlayer1.isMoviePlaying
 ProcTask2 basestack
' sleep 5

Loop Until AVIRUN = False Or NOEXECUTION
End If
Else
If FastSymbol(rest$, ";") Then
If AVIRUN Then
Do
 AVIRUN = MediaPlayer1.isMoviePlaying
ProcTask2 basestack
 ' sleep 5

Loop Until AVIRUN = False Or NOEXECUTION
End If
End If
Exit Do
End If
Loop
Set Scr = Nothing
Exit Function


End Function

Function Num2Str(p, FTXT As String) As String
        Dim s$
        If Not NoUseDec Then
                If OverideDec Then
                    s$ = Replace$(Format$(p, FTXT), GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                    s$ = Replace$(s$, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                    s$ = Replace$(s$, Chr(2), NowDec$)
                    Num2Str = Replace$(s$, Chr(3), NowThou$)
                ElseIf InStr(s$, NowDec$) > 0 And InStr(FTXT, ".") > 0 Then
                    Num2Str = Format$(p, FTXT)
                ElseIf InStr(s$, NowDec$) > 0 Then
                    s$ = Replace$(Format$(p, FTXT), NowDec$, Chr(2))
                    s$ = Replace$(s$, NowThou$, Chr(3))
                    s$ = Replace$(s$, Chr(2), ".")
                    Num2Str = Replace$(s$, Chr(3), ",")
                End If
        Else
            Num2Str = Format$(p, FTXT)
        End If
End Function


Function ProcImage(bstack As basetask, rest$, Lang As Long) As Boolean
Dim photo As cDIBSection, pppp As mArray, s$, x1 As Long, y1 As Long, w$, it As Long, p As Variant, part As Boolean, border As Long, titl$
Dim aPic As StdPicture, s1$, usecolorback As Boolean, ihavepic As Boolean, mem As MemBlock, cback As Long, usehandler As mHandler
ProcImage = True
part = IsLabelSymbolNew(rest$, "ΠΛΑΙΣΙΟ", "FRAME", Lang)
If IsStrExp(bstack, rest$, s$) Then
    ihavepic = Left$(s$, 4) = "cDIB" And Len(s$) > 12
    GoTo cont1
ElseIf IsExp(bstack, rest$, p) Then
    If Not bstack.lastobj Is Nothing Then
        If TypeOf bstack.lastobj Is mHandler Then
        Set usehandler = bstack.lastobj
            If usehandler.t1 = 2 Then
                Set mem = usehandler.objref
                Set usehandler = Nothing
                Set bstack.lastobj = Nothing
                If FastSymbol(rest$, "(") Then
                    If Not IsExp(bstack, rest$, p) Then MissNumExpr: Exit Function
               
                    usecolorback = True
                    cback = mycolor(p)
                    If Not FastSymbol(rest$, ")", True) Then Exit Function
                End If
                If mem Is Nothing Then
                GoTo errNoImage
                ElseIf usecolorback Then
                    Set aPic = mem.GetStdPicture(, , cback)
                Else
                    Set aPic = mem.GetStdPicture()
                End If
                Set bstack.lastobj = Nothing
                GoTo cont1
            End If
        End If
    End If
    Set usehandler = Nothing
    Set bstack.lastobj = Nothing
errNoImage:
    MyEr "No Image found", "Δεν βρήκα εικόνα"
    Exit Function
cont1:
    x1 = 0
    y1 = 0
    If aPic Is Nothing Then If Not ihavepic Then If ExtractType(s$) = vbNullString Then s$ = s$ + ".bmp"
    If part Then
        If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then x1 = p
        If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then y1 = p Else ProcImage = False: MissNumExpr: Exit Function
        If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then border = p Else ProcImage = False: MissNumExpr: Exit Function
        If FastSymbol(rest$, ",") Then If Not IsStrExp(bstack, rest$, titl$) Then ProcImage = False: MissStringExpr: Exit Function
        If Not aPic Is Nothing Then
            Set photo = New cDIBSection
            photo.CreateFromPicture aPic
            photo.GetDpi 96, 96
            If photo.Width = 0 Then
                Set photo = Nothing
                MissCdib
                ProcImage = False: Exit Function
            End If
            If photo.bitsPerPixel <> 24 Then Conv24 photo
            ThumbImageDib bstack.Owner, x1, y1, photo, border, dv15, titl$
            Set photo = Nothing
        ElseIf Not ihavepic Then
            Set photo = New cDIBSection
            If CFname(s$) <> "" Then
                s$ = CFname(s$)
                Set aPic = LoadMyPicture(GetDosPath(s$))
                If aPic Is Nothing Then Exit Function
                photo.CreateFromPicture aPic
                photo.GetDpi 96, 96
                If photo.Width = 0 Then
                    Set photo = Nothing
                    MissCdib
                    ProcImage = False: Exit Function
                End If
                CheckOrientation photo, s$
                If photo.bitsPerPixel <> 24 Then Conv24 photo
                ThumbImageDib bstack.Owner, x1, y1, photo, border, dv15, titl$
                Set photo = Nothing
            End If
        Else
           ThumbImage bstack.Owner, x1, y1, s$, border, dv15, titl$
        End If
    ElseIf IsLabelSymbolNew(rest$, "ΣΤΟ", "TO", Lang) Then
        If CFname(s$) <> "" Or ihavepic Or Not aPic Is Nothing Then
            Select Case Abs(IsLabel(bstack, rest$, w$))
            Case 3
                If GetVar(bstack, w$, it) Then
                    If FastSymbol(rest$, "(") Then
                        If IsExp(bstack, rest$, p) Then usecolorback = True: cback = mycolor(p)
                        If Not FastSymbol(rest$, ")", True) Then ProcImage = False: Exit Function
                    End If
                    Set photo = New cDIBSection
                    If Not (ihavepic Or Not aPic Is Nothing) Then
                        s$ = CFname(s$)
                        Set aPic = LoadMyPicture(s$, usecolorback, cback)
                    End If
                    If Not aPic Is Nothing Then
                        If aPic.Type = vbPicTypeIcon Then photo.backcolor = bstack.Owner.backcolor
                        If aPic.Type = 4 Then
                            If Not mem Is Nothing Then
                            If Not mem.SubType = 2 Then GoTo 1000
                            Set aPic = mem.GetStdPicture1(-1, -1, cback, True, True, True)
                            photo.CreateFromPicture aPic
                            GoTo 1010
                            Else
1000
                            With players(GetCode(bstack.Owner))
                                photo.emfSizeFactor = 1
                            End With
                            End If
                        End If
                        photo.ClearUp
                        photo.CreateFromPicture aPic, cback
1010
                        photo.GetDpi 96, 96
                        If photo.Width = 0 Then
                            Set photo = Nothing
                            MissCdib
                            ProcImage = False: Exit Function
                        End If
                        If Len(s$) > 0 Then CheckOrientation photo, s$
                        If photo.bitsPerPixel <> 24 Then Conv24 photo
                    ElseIf Not cDib(s$, photo) Then
                        Set photo = Nothing
                        ProcImage = False
                        MyEr "No Image Found", "Δεν βρήκε εικόνα"
                        Exit Function
                    End If
                    x1 = photo.Width
                    y1 = photo.Height
                    If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then x1 = x1 * p / 100#: y1 = y1 * p / 100#
                    If FastSymbol(rest$, ",") Then
                        If IsExp(bstack, rest$, p) Then
                            y1 = photo.Height * p / 100#
                        Else
                            Set photo = Nothing
                            ProcImage = False: MissNumExpr: Exit Function
                        End If
                    End If
                    If photo.Width > 0 Then
                    If Not (Abs(y1) = photo.Width And Abs(x1) = photo.Height) Then
                        If photo.BitmapType = 4 Then
                            Set aPic = photo.Picture()
                            p = Sqr(y1 * x1 / (photo.Width * photo.Height))
                            photo.ClearUp
                            photo.emfSizeFactor = p
                            photo.CreateFromPicture aPic, cback
                            photo.GetDpi 96, 96
                        End If
                        Set photo = photo.Resample(Abs(y1), Abs(x1))
                        
                        End If
                        var(it) = DIBtoSTR(photo)
                    Else
                        var(it) = vbNullString
                    End If
                    Set photo = Nothing
                Else
                    ProcImage = False
                    If w$ <> "" Then
                        Nosuchvariable w$
                    Else
                        MissingStrVar
                    End If
                End If
                Exit Function
            Case 6
    ' ΑΠΟ ΠΙΝΑΚΑ
                Dim W5 As Long
                If neoGetArray(bstack, w$, pppp) Then
                    If Not NeoGetArrayItem(pppp, bstack, w$, W5, rest$) Then ProcImage = False: MissNumExpr: Exit Function
                    If MyIsObject(pppp.item(W5)) Then
                        MyEr "can't copy image to " + pppp.ItemType(W5), "δεν μπορώ να αντιγράψω εικόνα σε " + pppp.ItemType(W5)
                        ProcImage = False
                        Exit Function
                    End If
                    If FastSymbol(rest$, "(") Then
                        If IsExp(bstack, rest$, p) Then usecolorback = True
                        If Not FastSymbol(rest$, ")", True) Then ProcImage = False: Exit Function
                    End If
                    Set photo = New cDIBSection
                    If Not (ihavepic Or Not aPic Is Nothing) Then
                        s$ = CFname(s$)
                        Set aPic = LoadMyPicture(s$, usecolorback, mycolor(p))
                    End If
                    If Not aPic Is Nothing Then
                        If aPic.Type = vbPicTypeIcon Then photo.backcolor = bstack.Owner.backcolor
                        photo.CreateFromPicture aPic
                        photo.GetDpi 96, 96
                        If photo.Width = 0 Then
                            Set photo = Nothing
                            MissCdib
                            ProcImage = False: Exit Function
                        End If
                        CheckOrientation photo, s$
                        If photo.bitsPerPixel <> 24 Then Conv24 photo
                    ElseIf Not cDib(s$, photo) Then
                            Set photo = Nothing
                            MissCdibStr
                            Exit Function
                    End If
                    x1 = photo.Width
                    y1 = photo.Height
                    If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then x1 = x1 * p / 100#: y1 = y1 * p / 100#
                    If FastSymbol(rest$, ",") Then
                        If IsExp(bstack, rest$, p) Then
                            y1 = photo.Height * p / 100#
                        Else
                            Set photo = Nothing
                         ProcImage = False: MissNumExpr: Exit Function
                        End If
                    End If
                    If photo.Width > 0 Then
                        Set photo = photo.Resample(y1, x1)
                        If MyIsObject(pppp.item(W5)) Then
                            MyEr "can't copy image to " + pppp.ItemType(W5), "δεν μπορώ να αντιγράψω εικόνα σε " + pppp.ItemType(W5)
                            ProcImage = False
                        Else
                            pppp.item(W5) = DIBtoSTR(photo)
                        End If
                    End If
                    Set photo = Nothing
                    Exit Function
                Else
                     ProcImage = False: MissingArray w$: Exit Function
                End If
            End Select
        Else
            MyEr "missing file, or image in string or in buffer", "λείπει αρχείο ή εικόνα σε αλφαριθμητικό ή σε διάρθρωση μνήμης"
            ProcImage = False
            Exit Function
        End If
    ElseIf IsLabelSymbolNew(rest$, "ΕΞΑΓΩΓΗ", "EXPORT", Lang) Then
        If IsStrExp(bstack, rest$, w$) Then
            If Not CanKillFile(w$) Then FilePathNotForUser:  Exit Function
            Set photo = New cDIBSection
            If Not (ihavepic Or Not aPic Is Nothing) Then
                 MyEr "No found Image in String or Buffer", "Δεν βρήκα εικόνα σε αλφαριθμητικό ή σε διάρθρωση μνήμης"
                 ProcImage = False
                 Exit Function
            ElseIf aPic Is Nothing Then
                If cDib(s$, photo) Then
cont4:
                    If FastSymbol(rest$, ",") Then
                        If IsExp(bstack, rest$, p) Then
                            x1 = (Abs(p) - 1) Mod 100 + 1
                            s$ = vbNullString
                            If FastSymbol(rest$, ",") Then If Not IsStrExp(bstack, rest$, s$) Then MissStringExpr: ProcImage = False: Exit Function
                            SaveJPG photo, ExtractPath(w$) + ExtractNameOnly(w$, True) & ".jpg", x1, s$
                        Else
                            Set photo = Nothing
                            ProcImage = False: MissNumExpr: Exit Function
                        End If
                    Else
                        photo.SaveDib ExtractPath(w$) + ExtractNameOnly(w$, True) & ".bmp"
                    End If
                Else
                 MyEr "No Proper Image in String", "Δεν βρήκα κατάλληλη εικόνα σε αλφαριθμητικό"
                 ProcImage = False
                 Exit Function
                End If
            Else
                photo.backcolor = &HFFFFFF
                photo.CreateFromPicture aPic
                photo.GetDpi 96, 96
                If photo.Width = 0 Then
                    Set photo = Nothing
                    MissCdib
                    ProcImage = False: Exit Function
                End If
                If photo.bitsPerPixel <> 24 Then Conv24 photo
                GoTo cont4
            End If
            Set photo = Nothing
            Exit Function
        Else
             ProcImage = False: MyEr "Missing Filename", "Δεν βρήκα όνομα αρχείου": Exit Function
        End If
ElseIf mem Is Nothing Then
    If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then x1 = p
    If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then y1 = p Else ProcImage = False: MissNumExpr: Exit Function
    SImage bstack.Owner, x1, y1, s$
Else
    Dim d1 As Object, maybeerror As Boolean
    Set d1 = bstack.Owner
    Dim mm As MetaDc, rr As Single, xSpot As Long, ySpot As Long
    x1 = -1
    y1 = -1
    If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then x1 = Abs(p)
    
    If FastSymbol(rest$, ",") Then
            If IsExp(bstack, rest$, p) Then
                y1 = Abs(p)
            Else
              maybeerror = True
            End If
    End If
    If TypeOf d1 Is MetaDc Then
        If mem.SubType = 2 Then
            If aPic Is Nothing Then Set aPic = mem.GetStdPicture
            Set mm = d1
            With players(GetCode(d1))
                If FastSymbol(rest$, ",") Then
                If IsExp(bstack, rest$, p) Then rr = p 'MyMod(p, 360)
                If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then xSpot = CSng(p)
                If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then ySpot = CSng(p) Else ProcImage = False: MissNumExpr: Exit Function
                If mem.IsWmf Then
                mm.PlayWmfInside aPic, .XGRAPH / dv15, .YGRAPH / dv15, x1, y1, rr, xSpot / dv15, ySpot / dv15
                Else
                mm.PlayEmfInside aPic, .XGRAPH / dv15, .YGRAPH / dv15, x1, y1, rr, xSpot / dv15, ySpot / dv15
                End If
                Else
                    mm.PaintPicture aPic, .XGRAPH, .YGRAPH, x1, y1
                End If
                
                Exit Function
            End With
        End If
        GoTo renderbitmap
    Else
        'If mem.SubType = 2 Then
        If FastSymbol(rest$, ",") Then
renderthere:
            If IsExp(bstack, rest$, p) Then rr = p 'MyMod(p, 360)
            If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then xSpot = CSng(p)
            If FastSymbol(rest$, ",") Then If IsExp(bstack, rest$, p) Then ySpot = CSng(p) Else ProcImage = False: MissNumExpr: Exit Function
            
             mem.DrawEmfToHdc bstack, xSpot, ySpot, rr, x1, y1
             
            
            Exit Function
        Else
            If maybeerror Then MissParam rest$: Exit Function
        End If
renderbitmap:
If MaybeIsSymbol(rest$, ",") Then GoTo renderthere
        If x1 > -1 Then x1 = d1.ScaleX(x1, 1, 3)
        If y1 > -1 Then y1 = d1.ScaleY(y1, 1, 3)
        With players(GetCode(d1))
            'If FastSymbol(rest$, ",") Then MyEr "Not for bitmap", "όχι για εικόνα σημείων": ProcImage = False: Exit Function
            
            mem.DrawImageToHdc d1, .XGRAPH \ dv15, .YGRAPH \ dv15, x1, y1
        End With
    End If
End If
End If
End Function
 
Function ProcPlayer(bstack As basetask, rest$, Lang As Long)
Dim par As Boolean, sX As Double, sY As Double, x As Double, y As Double, x1 As Integer, p As Variant, it As Long
Dim Col As Long, Scr As Object, what$, i As Long, s$, pppp As mArray, frm$, sxy As Double, orig As Long, zX As Variant, zY As Variant
If IsExp(bstack, rest$, p) Then
    If p = 0 Then   ' ZERO CLEAR ALL HARDWARE SPRITES
        ClrSprites
        ProcPlayer = True
        Exit Function
    End If
    If p < 1 Or p > 32 Then SyntaxError: ProcPlayer = False: Exit Function
    orig = CLng(p)
    it = FindSpriteByTag(orig)
    If FastSymbol(rest$, ",") Then
        If Not IsExp(bstack, rest$, x) Then  ' get new left or leave it empty
            If it = 0 Then
                x = 0
            Else
                x = Form1.dSprite(it).Left + players(it).x
            End If
            If FastSymbol(rest$, ",") Then
                If Not IsExp(bstack, rest$, y) Then MissNumExpr: ProcPlayer = False: Exit Function
            Else
                MissNumExpr
                ProcPlayer = False: Exit Function
            End If
        Else
            If FastSymbol(rest$, ",") Then   ' so ,, is "stay X where you are
                If Not IsExp(bstack, rest$, y) Then MissNumExpr: ProcPlayer = False: Exit Function
            Else
                If it = 0 Then
                    y = 0
                Else
                    y = Form1.dSprite(it).top + players(it).y
                End If
            End If
        End If
        If IsLabelSymbolNew(rest$, "ΜΕ", "USE", Lang) Then ' no need for coma
            Select Case Abs(IsLabel(bstack, rest$, what$))
            Case 3
                If GetVar(bstack, what$, i) Then s$ = var(i)
            Case 6
                If neoGetArray(bstack, what$, pppp) Then
                    
                    If Not NeoGetArrayItem(pppp, bstack, what$, it, rest$) Then
                        MissNumExpr
                        ProcPlayer = True: Exit Function
                    End If
                Else
                    MissNumExpr
                    ProcPlayer = False: Exit Function
                End If
                s$ = pppp.item(it)  ' get the sprite image
            Case Else
                MissNumExpr
                ProcPlayer = False: Exit Function
            End Select
            Col = 0
            sX = 0

            If FastSymbol(rest$, ",") Then  ' get image manipulators..
                    If IsExp(bstack, rest$, sY) Then
                     Col = mycolor(sY)
                    ' If col > 0 Then col = QBColor(col Mod 16) Else col = -col
                     ElseIf IsStrExp(bstack, rest$, frm$) Then
                     '' maybe is a mask
                     
                     Col = 0
                     Else
                     ProcPlayer = False: MissNumExpr: Exit Function
                    End If
                     
                        If FastSymbol(rest$, ",") Then
                            If IsExp(bstack, rest$, sX) Then
                          
                               Else
                            MissNumExpr
                            ProcPlayer = False: Exit Function
                            End If
                        Else
                
                        End If
                    
              End If
                    
                    If FastSymbol(rest$, ",") Then
                                If IsExp(bstack, rest$, sxy) Then
                                    If FastSymbol(rest$, ",") Then GoTo HOTSPOT
                                ElseIf FastSymbol(rest$, ",") Then
HOTSPOT:
                                IsExp bstack, rest$, zX
                                If FastSymbol(rest$, ",") Then
                                    If Not IsExp(bstack, rest$, zY) Then GoTo mis
                                End If
                                Else
mis:
                                    MissNumExpr
                                    ProcPlayer = False: Exit Function
                                End If
                    
                    
                    End If
                    
                   If IsLabelSymbolNew(rest$, "ΜΕΓΕΘΟΣ", "SIZE", Lang) Then
              If Not IsExp(bstack, rest$, sY) Then ProcPlayer = False: MissNumExpr: Exit Function
              Else
              sY = 1
              End If
                    
              ' so col, sx and sy are image manipulators
            it = GetNewSpriteObj(orig, s$, Col, CLng(sX), CSng(sY), CSng(sxy), frm$)
 On Error Resume Next
                players(it).HotSpotX = -zX * sY
                    players(it).HotSpotY = -zY * sY
 
            PosSprite orig, x - players(it).x + players(it).HotSpotX, y - players(it).y + players(it).HotSpotY
        Else ' without USE
      
    On Error Resume Next
         PosSprite orig, x - players(it).x + players(it).HotSpotX, y - players(it).y + players(it).HotSpotY
        
        End If
        Else ' without x, y
            If IsLabelSymbolNew(rest$, "ΜΕ", "USE", Lang) Then
        Select Case Abs(IsLabel(bstack, rest$, what$))
        Case 3
            If GetVar(bstack, what$, i) Then s$ = var(i)
        Case 6
             If neoGetArray(bstack, what$, pppp) Then
   
                If Not NeoGetArrayItem(pppp, bstack, what$, it, rest$) Then
                     ProcPlayer = False: MissNumExpr: Exit Function
                End If
            Else
                 ProcPlayer = False: MissNumExpr: Exit Function
            End If
            s$ = pppp.item(it)
        Case Else
             ProcPlayer = False: MissNumExpr: Exit Function
        End Select
        Col = 0 'rgb(255, 255, 255)
        sX = 0
    If FastSymbol(rest$, ",") Then
    
        If IsExp(bstack, rest$, sY) Then
            'col = CLng(sY)
            Col = mycolor(sY)
            'If col > 0 Then col = QBColor(col) Else col = -col
        ElseIf IsStrExp(bstack, rest$, frm$) Then
            '' maybe is a mask
            Col = 0
        Else
            ProcPlayer = False: MissNumExpr: Exit Function
        End If
        If FastSymbol(rest$, ",") Then
            If IsExp(bstack, rest$, sX) Then
            Else
                MissNumExpr
                ProcPlayer = False: Exit Function
            End If
        Else
        End If
    End If
            If FastSymbol(rest$, ",") Then
                If IsExp(bstack, rest$, sxy) Then
                Else
                    MissNumExpr
                    ProcPlayer = False: Exit Function
                End If
            Else
    
            End If
            If IsLabelSymbolNew(rest$, "ΜΕΓΕΘΟΣ", "SIZE", Lang) Then          ' SIZE WITHOUT COMMA
                If Not IsExp(bstack, rest$, sY) Then ProcPlayer = False: MissNumExpr: Exit Function
            Else
                sY = 1
            End If
    
            it = GetNewSpriteObj(orig, s$, Col, CLng(sX), CSng(sY), CSng(sxy), frm$)
            ' no USE no X, Y or X,Y USE ..
            ' only command
        ElseIf IsLabelSymbolNew(rest$, "ΔΕΙΞΕ", "SHOW", Lang) Then     ' SHOW
            SrpiteHideShow orig, (True)
        ElseIf IsLabelSymbolNew(rest$, "ΚΡΥΨΕ", "HIDE", Lang) Then        ' HIDE
            SrpiteHideShow orig, (False)
        ElseIf IsLabelSymbolNew(rest$, "ΠΑΝΩ", "OVER", Lang) Then      ' ΠΑΝΩ
            If IsExp(bstack, rest$, x) Then
                If orig <> x Then SpriteControlOver orig, CLng(x)
                
            Else
                ProcPlayer = False
            End If
        ElseIf IsLabelSymbolNew(rest$, "ΥΠΟ", "UNDER", Lang) Then
            If IsExp(bstack, rest$, x) Then
                If orig <> x Then SpriteControlUnder orig, CLng(x)
                
            Else
                ProcPlayer = False
            End If
        ElseIf IsLabelSymbolNew(rest$, "ΑΛΛΑΞΕ", "SWAP", Lang) Then       ' SWAP
            If IsExp(bstack, rest$, x) Then
                If orig <> x Then SpriteControl orig, CLng(x)
                
            Else
                ProcPlayer = False
            End If
        End If
    End If
End If
ProcPlayer = True
Exit Function
End Function


Sub ClearState()
Basestack1.IamAnEvent = False
abt = False
Set comhash = New sbHash
allcommands comhash
Set numid = New idHash
Set funid = New idHash
Set strid = New idHash
Set strfunid = New idHash
NumberId numid, funid
StringId strid, strfunid
NoOptimum = False
NERR = False
TaskMaster.Dispose
CloseAllConnections
CleanupLibHandles
If Not NOEDIT Then
NOEDIT = True
Else
If QRY Then QRY = False
End If
' restore DB.Provider for User
JetPrefixUser = JetPrefixHelp
JetPostfixUser = JetPostfixHelp
' SET ARRAY BASE TO ZERO
ArrBase = 0
End Sub

Function ProcPrinter(basestack As basetask, rest$) As Boolean
Dim xp As Printer, i As Long, p As Variant, x1 As Long, y1 As Long, x As Double, y As Double
Dim s$, ss$, F As Long, pa$, sX As Double, it As Long, ya As Long, AddTwipsTopL As Long, nd&
Dim Scr As Object
ProcPrinter = True
Set Scr = basestack.Owner
If basestack.toprinter Then Exit Function
 If ThereIsAPrinter = False Then Exit Function
If FastSymbol(rest$, "!") Then
olamazi
If ThereIsAPrinter Then
For Each xp In Printers
If xp.DeviceName = pname Then
Set Printer = xp
Exit For
End If
Next xp

If ShowProperties(Form1, Printer.DeviceName, MyDM()) Then
MyDoEvents
PrinterDim pw, ph, psw, psh, pwox, phoy
End If
End If
Exit Function
End If
If FastSymbol(rest$, "+") Then
Form1.List1.Clear
For Each xp In Printers
Form1.List1.additemFast xp.DeviceName & " (" & xp.port & ")"
Next xp
For i = 0 To Form1.List1.listcount - 1
If pname & " (" & port & ")" = Form1.List1.list(i) Then Form1.List1.ListIndex = i
Next i
Exit Function
ElseIf FastSymbol(rest$, "?") Then
Form1.List1.Clear
For Each xp In Printers
Form1.List1.additemFast xp.DeviceName & " (" & xp.port & ")"
Next xp
For i = 0 To Form1.List1.listcount - 1
If pname & " (" & port & ")" = Form1.List1.list(i) Then Form1.List1.ListIndex = i
Next i
Execute basestack, "menu !", True
If CDbl(Form1.List1.ListIndex + 1) > 0 Then
i = InStr(Form1.List1.ListValue, " (")
pname = Left$(Form1.List1.ListValue, i - 1)
port = Mid$(Form1.List1.ListValue, i + 2, InStr(i + 2, Form1.List1.ListValue, ")") - i - 2)
End If
For Each xp In Printers
If xp.DeviceName = pname And xp.port = port Then Set Printer = xp
Next xp
ReDim MyDM(1 To 1) As Byte
Exit Function
End If
getfirstpage
If Not IsExp(basestack, rest$, p) Then p = players(-2).SZ / szFactor

If FastSymbol(rest$, ",") Then
If IsStrExp(basestack, rest$, s$) Then
For Each xp In Printers
If xp.DeviceName & " (" & xp.port & ")" = s$ Then
pname = xp.DeviceName
port = xp.port
Set Printer = xp
ReDim MyDM(1 To 1) As Byte
If FastSymbol(rest$, ",") Then
If IsStrExp(basestack, rest$, ss$) Then
If ss$ <> "" Then
LoadArray MyDM(), ss$
End If
End If
Exit For
End If
End If
Next xp

Exit Function
End If
End If
If lookOne(rest$, "{") Then


 If ThereIsAPrinter = False Then Exit Function
If pname = vbNullString Then Exit Function
For Each xp In Printers
If xp.DeviceName = pname And xp.port = port Then Set Printer = xp
Next xp
getfirstpage

If players(-2).Xt = 0 Then

players(-2) = players(0)  'COPY dis
With players(-2)
players(-2).curpos = 0
players(-2).currow = 0
If p = 0 Then p = .SZ
szFactor = mydpi * dv15 / 1440#
.SZ = CSng(p * szFactor)
End With
PlaceBasket Form1.PrinterDocument1, players(-2)
SetText Form1.PrinterDocument1

Else
SetTextSZ Form1.PrinterDocument1, CSng(p * mydpi * dv15 / 1440)

LCTbasket Form1.PrinterDocument1, players(-2), 0, 0
'realfactor = CSng(players(-2).SZ) / prFactor / p
End If

With Printer   ' for no specific reason..I have to think it again
.currentX = 0
.currentY = 0
End With
basestack.toprinter = True
nd& = basestack.addlen
it = Execute(basestack, rest$, False, True, , True)
basestack.addlen = nd&
            If it = 2 Then
                        If rest$ = "" Then
                            rest$ = ": Break": If trace Then WaitShow = 2: TestShowSub = vbNullString
                        Else
                        rest$ = ": Goto " + rest$
                         If trace Then WaitShow = 2: TestShowSub = rest$
                        End If
                        it = 1
                        End If

If Not basestack.toprinter Then
pnum = 0
oprinter.ClearUp
Form1.PrinterDocument1.Picture = LoadPicture("")
Else
getenddoc
End If
basestack.toprinter = False
Set Scr = basestack.Owner
If it = 0 Then
ProcPrinter = False
End If

Else
PlainBaSket Scr, players(GetCode(Scr)), pname & " (" & port & ")"
crNew basestack, players(GetCode(Scr))
End If
Exit Function
End Function

Private Property Get xmlMonoNew() As XmlMono
    Dim m As New XmlMonoInternal, z As New XmlMono
    z.createTree m
    Set xmlMonoNew = z
End Property
Function IsCollide(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim r2 As Variant
If IsExp(bstack, a$, R, , True) Then
R = Fix(R)
If FastSymbol(a$, ",") Then
    If Not IsExp(bstack, a$, r2) Then: MissParam a$: IsCollide = False: Exit Function
    r2 = Fix(r2)
    If FastSymbol(a$, ",") Then
    R = SG * CollideArea(CLng(R), CLng(r2), bstack, a$)
    Else
    R = SG * CollidePlayers(CLng(R), CLng(r2))
    End If
Else
R = SG * CollidePlayers(CLng(R), CLng(100))
End If
   
    IsCollide = FastSymbol(a$, ")", True)
End If
End Function

Function IsParagr(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim s$, pp As Variant, dn As Long, w1 As Long, w2 As Long, pppp As mArray
    w1 = Abs(IsLabel(bstack, a$, s$))
                  
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) = doc Then
                        If Not FastSymbol(a$, ",") Then: MissParam a$: Exit Function
        
                       If IsExp(bstack, a$, pp, , True) Then
                                dn = CLng(Fix(pp))
                              R = SG * var(w1).ParagraphFromOrder(dn)           ''
                                 
                            
                                 Else
                                       MissNumExpr
                                        
                                        IsParagr = False
                                        Exit Function
                                 End If
         
               Else
                    MissingDoc
                                        
                                        IsParagr = False
                                        Exit Function
                End If
                
                IsParagr = FastSymbol(a$, ")", True)
            Else
                    
                    MissFuncParameterStringVarMacro a$
                    
            End If
        ElseIf w1 = 6 Then
                If neoGetArray(bstack, s$, pppp) Then
                 If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                                If Not FastSymbol(a$, ",") Then: MissParam a$: Exit Function
                            If IsExp(bstack, s$, pp, , True) Then
                                dn = CLng(Fix(pp))
                                 R = SG * pppp.item(w2).ParagraphFromOrder(dn)
                                 Else
                                        MissNumExpr
                                        
                                        IsParagr = False
                                        Exit Function
                                 End If
                  Else
                    MissingDoc
                                        
                                        IsParagr = False
                                        Exit Function
                End If
                    
                IsParagr = FastSymbol(a$, ")", True)
    Else
                    
                MissFuncParameterStringVarMacro a$
    End If
End Function

Public Function MyTitle$(basestack As basetask)
Static PREVT$

' On Error GoTo t1
If exWnd = 0 Then
PREVT$ = vbNullString
MyTitle$ = vbNullString
Exit Function
End If
If PREVT$ <> nnn$ Then
PREVT$ = nnn$

Form1.view1_StatusTextChange11 basestack, Trim$(nnn$)

End If
MyTitle$ = Trim$(PREVT$)

Exit Function
t1:
MyTitle$ = "???"
End Function
Public Function Originalusername()

Dim ss$
                 ss$ = UCase(userfiles)
                    DropLeft "\M2000_USER\", ss$
                    
If ss$ = vbNullString Then
Originalusername = UserName
Else
ss$ = Right$(userfiles, Len(ss$))
Originalusername = GetStrUntil("\", ss$)
End If
End Function
Public Function UserName()
Dim a$, b$, c$
a$ = GetSpecialfolder(0)
While a$ <> ""
c$ = b$
b$ = GetStrUntil("\", a$)
Wend
UserName = c$
End Function
Function IsDimension(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim s$, pppp As mArray, w1 As Long, p As Variant, anything As Object, pp As Variant, usehandler As mHandler
Dim k As Long
k = Abs(IsLabel(bstack, a$, s$))
Set bstack.lastobj = Nothing
If k > 4 Then
If neoGetArray(bstack, s$, pppp) Then
If pppp.Arr Then
If FastSymbol(a$, ")") Then
    If k = 6 Then
        IsStrExp bstack, s$ + ")", s$
    Else
        IsNumber bstack, s$ + ")", p
    End If
Else
bstack.tmpstr = s$ + Left$(a$, 1)
  BackPort a$
    If k = 6 Then
        IsStrExp bstack, a$, s$
    Else
        IsNumber bstack, a$, p
    End If
End If
getback:
If Not bstack.lastobj Is Nothing Then
If TypeOf bstack.lastobj Is mHandler Then
Set anything = bstack.lastobj
If CheckIsmArray(anything) Then Set bstack.lastobj = anything
Set anything = Nothing
End If
If Not (TypeOf bstack.lastobj Is mArray) Then NeedAnArray a$: Exit Function
Set pppp = bstack.lastobj
Set bstack.lastobj = Nothing
Else

IsDimension = False
End If
Else
   If NeoGetArrayItem(pppp, bstack, s$, w1, a$) Then
    If TypeOf pppp.GroupRef Is mHandler Then
        Set usehandler = pppp.GroupRef
        Set pppp = usehandler.objref.ValueObj
        Set usehandler = Nothing
    End If
  End If
End If
Else
      NeedAnArray a$: Exit Function
End If

 If pppp.Arr Then  '

    If FastSymbol(a$, ",") Then
        If IsExp(bstack, a$, p, , True) Then
            If p < 1 Then
                If Not FastSymbol(a$, ",") Then
                    pppp.GetDnum (0), pp, R
                    R = SG * -R
                ElseIf IsExp(bstack, a$, R, , True) Then
                    R = CLng(R) - 1
                    pppp.SerialItem pp, (R), 5
                    If R >= 0 And R < pp Then
                        pppp.GetDnum (R), pp, R
                        R = SG * -R
                    Else
                        CantReadDimension a$, s$
                        Exit Function
                   End If
                Else
                    missNumber
                    Exit Function
                End If
            Else
                'pppp.SerialItem PP, CLng(Fix(p) - 1), 6
                pppp.SerialItem pp, (R), 5
                p = p - 1
                If p >= 0 And p < pp Then
                    If FastSymbol(a$, ",") Then
                        If IsExp(bstack, a$, R, , True) Then
                            If R = 0 Then
                                pppp.GetDnum CLng(Fix(p)), pp, R
                                R = SG * -R
                            Else
                                pppp.GetDnum CLng(Fix(p)), pp, R
                                R = pp - R - 1
                            End If
                        End If
                    Else
                        pppp.GetDnum CLng(Fix(p)), pp, R
                        R = SG * pp
                    End If
                Else
                    CantReadDimension a$, s$
                    Exit Function
                End If
            
            End If
                IsDimension = FastSymbol(a$, ")", True)
        Else
            CantReadDimension a$, s$
        End If
        
      Else ' dimensions
      p = 0
      pppp.SerialItem pp, CLng(Fix(p)), 5
         R = SG * pp
              
              IsDimension = FastSymbol(a$, ")", True)
      End If
      Exit Function
      Else
       
       End If
ElseIf GetVar(bstack, s$, w1) Then
If IsExp(bstack, (s$), p) Then
    Set anything = bstack.lastobj
    If CheckIsmArray(anything) Then
        Set bstack.lastobj = anything
        Set anything = Nothing
    Else
    Set anything = Nothing

    End If
GoTo getback

End If
 Else
 
 End If
If IsStrExp(bstack, a$, s$) Then
    s$ = s$ & "("
    If neoGetArray(bstack, s$, pppp) Then
    
      If Not pppp.Arr Then
    
  If NeoGetArrayItem(pppp, bstack, s$, w1, a$, , False) Then
      If TypeOf pppp.GroupRef Is mHandler Then
        Set usehandler = pppp.GroupRef
        Set pppp = usehandler.objref.ValueObj
        Set usehandler = Nothing
    End If
  End If
  
   
  End If
  If Not pppp.Arr Then NeedAnArray a$: Exit Function
        If FastSymbol(a$, ",") Then
          If IsExp(bstack, a$, p, , True) Then
            If p < 1 Then
                R = SG * -pppp.myarrbase
            Else
                p = Fix(p)
                pppp.SerialItem pp, CLng(p - 1), 6
                R = SG * pp
            End If
            
            IsDimension = FastSymbol(a$, ")", True)
          Else
            CantReadDimension a$, s$
            End If
        Else ' dimensions
            p = 0
            pppp.SerialItem pp, CLng(p), 5
            R = SG * pp
            
            IsDimension = FastSymbol(a$, ")", True)
        End If
        Else
        CantFindArray a$, s$
    End If
Else
        CantFindArray a$, s$
End If
End Function

Function ProcLoadDoc(entrypoint As Long, basestack As basetask, rest$) As Boolean
Dim dum As Boolean, pppp As mArray, s$, i As Long, x1 As Long, y1 As Long, frm$, ss$
ProcLoadDoc = True
If entrypoint = 1 Then dum = True
      y1 = Abs(IsLabel(basestack, rest$, s$))

        If y1 = 6 Then
                If neoGetArray(basestack, s$, pppp) Then
                    If Not NeoGetArrayItem(pppp, basestack, s$, i, rest$) Then Exit Function
                Else
                    MissingDoc
                    Exit Function
                End If
    End If
    If FastSymbol(rest$, ",") Then
    If Not IsStrExp(basestack, rest$, frm$) Then
    MissStringExpr
    
    Exit Function
    End If
    ss$ = GetNextLine(frm$)
    SetNextLine frm$
    If frm$ <> "" Then
    MyEr "filename with line breaks", "όνομα αρχείου με αλλαγές γραμμών"
    
    End If
    ' check valid name
    If ExtractNameOnly(ss$, True) = vbNullString Then BadFilename:  Exit Function
    If ExtractPath(ss$) = vbNullString Then
    ss$ = mylcasefILE(mcd + ss$)
    End If
    If ExtractType(ss$) = vbNullString Then ss$ = ss$ + ".txt"
    Else
    
    MissPar
    Exit Function
    End If
    
   
        If y1 = 3 Then
            If GetVar(basestack, s$, i) Then
                If Typename(var(i)) = doc Then

                x1 = 2
                On Error Resume Next
                If FastSymbol(rest$, ",") Then
                Dim p As Variant
                If IsExp(basestack, rest$, p, , True) Then
                var(i).lcid = CLng(p)
                End If
                End If
            If basestack.Owner.Name = "GuiM2000" Then
                Set basestack.Owner.mDoc = var(i)
                var(i).ReadUnicodeOrANSI ss$, dum, x1
                Set basestack.Owner.mDoc = Nothing
            Else
                var(i).ReadUnicodeOrANSI ss$, dum, x1
                End If
                If Err.Number > 0 Then Err.Clear: Exit Function
                 var(i).ListLoadedType = x1
                 Exit Function
                Else
                    MissingDoc
                    
                End If
            Else
                   MissFuncParameterStringVar
                    
            End If
        ElseIf y1 = 6 Then
                    If pppp.ItemType(i) = doc Then
                                    x1 = 2
                pppp.item(i).ReadUnicodeOrANSI ss$, dum, x1
                 pppp.item(i).ListLoadedType = x1
                    
                        Else
                         MissingDoc
                         
                        End If
    Else
                MissPar
    End If
End Function


Function ProcRecursionLimit(basestack As basetask, rest$, Lang As Long) As Boolean
Dim p As Variant, prive As Long
ProcRecursionLimit = True
If IsExp(basestack, rest$, p, , True) Then
deep = Abs(MyRound(p))
If IsSymbol(rest$, ",") Then
If IsExp(basestack, rest$, p, , True) Then funcdeep = Abs(MyRound(p)) ' obsolate
End If
ElseIf IsSymbol(rest$, ",") Then
If IsExp(basestack, rest$, p, , True) Then funcdeep = Abs(MyRound(p)) ' obsolate
Else
prive = GetCode(basestack.Owner)
    If deep = 0 Then
            If Lang = 1 Then
            PlainBaSket basestack.Owner, players(prive), "NO RECURSION LIMIT FOR SUBRUTINES"
            Else
            PlainBaSket basestack.Owner, players(prive), "ΧΩΡΙΣ ΟΡΙΟ ΑΝΑΔΡΟΜΗΣ ΣΤΙΣ ΡΟΥΤΙΝΕΣ"
            End If
            
    Else
    If Lang = 1 Then
        PlainBaSket basestack.Owner, players(prive), "RECURSION LIMIT FOR SUBRUTINES " + CStr(deep)
         crNew basestack, players(prive)
         If m_bInIDE Then
         PlainBaSket basestack.Owner, players(prive), "RECURSION LIMIT FOR FUNCTIONS " + CStr(stacksize \ 2948 - 1)
         Else
        PlainBaSket basestack.Owner, players(prive), "RECURSION LIMIT FOR FUNCTIONS " + CStr(stacksize \ 9832 - 1)
        End If
    Else
        PlainBaSket basestack.Owner, players(prive), "ΟΡΙΟ ΑΝΑΔΡΟΜΗΣ ΣΤΙΣ ΡΟΥΤΙΝΕΣ " + CStr(deep)
         crNew basestack, players(prive)
          If m_bInIDE Then
          PlainBaSket basestack.Owner, players(prive), "ΟΡΙΟ ΑΝΑΔΡΟΜΗΣ ΣΤΙΣ ΣΥΝΑΡΤΗΣΕΙΣ " + CStr(stacksize \ 2948 - 1)
          Else
        PlainBaSket basestack.Owner, players(prive), "ΟΡΙΟ ΑΝΑΔΡΟΜΗΣ ΣΤΙΣ ΣΥΝΑΡΤΗΣΕΙΣ " + CStr(stacksize \ 9832 - 1)
        End If
    End If
    End If
      '  PlainBaSket basestack.Owner, players(prive), CStr(deep)
    crNew basestack, players(prive)
End If
If funcdeep < 128 Then funcdeep = 128
If funcdeep > 3260 Then funcdeep = 3260
End Function


Function ProcSalata(entrypoint As Long, basestack As basetask, rest$) As Boolean
Dim p As Variant, Scr As Object, prive As Long, s$
ProcSalata = True
On entrypoint GoTo charset, CodePage, Locale
Exit Function
charset:
If IsExp(basestack, rest$, p, , True) Then
On Error Resume Next
chr11:
    Set Scr = basestack.Owner
    prive = GetCode(Scr)
    Scr.Font.charset = CInt(p)
    Form1.TEXT1.Font.charset = Scr.Font.charset
    Form1.List1.Font.charset = Scr.Font.charset
      StoreFont Scr.Font.Name, players(prive).SZ, Scr.Font.charset
      players(prive).charset = Scr.Font.charset
          Set Scr = Nothing
End If
Exit Function
CodePage:
        If IsExp(basestack, rest$, p, , True) Then
        ' usercodepage for use compare.
        ' also change to form.
        On Error Resume Next
        If Not IsValidCodePage(CLng(p)) Then
            NoValidCodePage
            ProcSalata = True
        Exit Function
        End If
CHR222:
        UserCodePage = CLng(p)
        p = GetCharSet(CLng(p))
        
        GoTo chr11
        End If
Exit Function
Locale:
    On Error Resume Next
    If IsExp(basestack, rest$, p, , True) Then
        If CLng(p) <> 0 Then
        If GetCodePage(CLng(p)) = 0 Then
        ProcSalata = False
                        NoValidLocale
                        Exit Function
                    End If
    Clid = CLng(p)
    
            Else
            Clid = OsInfo.LangNonUnicodeCode
                    End If
    
    If Clid = 1032 Then
    DefBooleanString = ";Αληθές;Ψευδές"
    Else
    DefBooleanString = ";\T\r\u\e;\F\a\l\s\e"
    End If
    OverideDec = True
    NoUseDec = False
    NowDec$ = GetlocaleString(LOCALE_SDECIMAL)
    NowThou$ = GetlocaleString(LOCALE_STHOUSAND)
    p = GetCodePage(CLng(p))
    GoTo CHR222
    ElseIf IsStrExp(basestack, rest$, s$) Then
    ' format$ for true/false values
    If LenB(s$) = 0 Then DefBooleanString = ";\T\r\u\e;\F\a\l\s\e" Else DefBooleanString = s$
    End If
End Function
Function ProcScreenRes(basestack As basetask, rest$) As Boolean
Dim x As Double, y As Double
If IsExp(basestack, rest$, x) Then
    If FastSymbol(rest$, ",") Then
        If IsExp(basestack, rest$, y) Then
            ChangeScreenRes CLng(x), CLng(y)
        Else
            
            MissNumExpr
        End If
    Else
        
        SyntaxError
    End If
ElseIf FastSymbol(rest$, "!") Then
    ScreenRestore
Else
    
    MissNumExpr
End If
ProcScreenRes = True
End Function
Function ProcSaveAs(bstack As basetask, rest$, Lang As Long) As Boolean
Dim Scr As Object, frm$, pa$, s$, ss$, w$
If IsSelectorInUse Then
SelectorInUse
Exit Function
End If
olamazi

frm$ = mcd
DialogSetupLang Lang

IsStrExp bstack, rest$, s$
If FastSymbol(rest$, ",") Then If IsStrExp(bstack, rest$, pa$) Then frm$ = pa$
If frm$ <> "" Then If Not isdir(frm$) Then NoSuchFolder: Exit Function
If FastSymbol(rest$, ",") Then If IsStrExp(bstack, rest$, pa$) Then ss$ = pa$
If FastSymbol(rest$, ",") Then If Not IsStrExp(bstack, rest$, w$) Then Exit Function
olamazi
If TypeOf bstack.Owner Is GuiM2000 Then
Set Scr = bstack.Owner
Else
If Form1.Visible Then
Set Scr = Form1
Else
Set Scr = Nothing
End If
End If
' change for file type
If InStr(w$, "|") > 0 Then w$ = vbNullString  ' NOT COMBATIBLE..CHANGE TO ALL FILES
If SaveAsDialog(bstack, Scr, s$, frm$, ss$, w$) Then
bstack.soros.PushStr ReturnFile
Else
bstack.soros.PushStr ""
End If
Set Scr = Nothing
ProcSaveAs = True
End Function

Function ProcJoypad(basestack As basetask, rest$) As Boolean
Dim p As Variant
        If IsExp(basestack, rest$, p, , True) Then
        If Not StartJoypadk(MyRound(p)) Then
        ' ERROR
        MyEr "Joypad " & CStr(p) & " not exist", "η λαβή " & CStr(p) & " δεν υπάρχει"
        
        Exit Function
        End If
        While FastSymbol(rest$, ",")
        
         If IsExp(basestack, rest$, p, , True) Then
        If Not StartJoypadk(MyRound(p)) Then
        MyEr "Joypad " & CStr(p) & " not exist", "η λαβή " & CStr(p) & " δεν υπάρχει"
        
        Exit Function

        End If
        Else
        MyEr "Joypad Number?", "Αριθμός Λαβής?"
        
        Exit Function

        End If
        Wend
        
        Else
        FlushJoyAll
        End If
ProcJoypad = True
End Function
Function MyDelay(basestack As basetask, rest$) As Boolean
Dim p As Variant
If IsExp(basestack, rest$, p, , True) Then
mywait basestack, p
Else
mywait basestack, 0
End If
MyDelay = True

End Function
Function ProcTune(bstack As basetask, rest$) As Boolean
'' break async
Dim p As Variant, s$
If IsExp(bstack, rest$, p) Then
    If Not FastSymbol(rest$, ",") Then
        beeperBEAT = CLng(p)
    ElseIf IsStrExp(bstack, rest$, s$) Then
        beeperBEAT = CLng(p)
        PlayTune (s$)
    Else
        MyEr "wrong parameter", "λάθος παράμετρος"
        Exit Function
    End If
ElseIf IsStrExp(bstack, rest$, s$) Then
' B C D E F G
PlayTune (s$)
End If
ProcTune = True
End Function
Function ProcName(bstack As basetask, rest$, Lang As Long) As Boolean
Dim s$, w$, x1 As Long, y1 As Long, ss$
ProcName = True

x1 = Abs(IsLabelFileName(bstack, rest$, s$, , w$))

If x1 = 1 Then
s$ = w$
 If Not IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then ProcName = False: Exit Function

 y1 = Abs(IsLabelFileName(bstack, rest$, ss$, , w$))

 If y1 = 0 Then
rest$ = w$ + rest$
y1 = IsStrExp(bstack, rest$, ss$)
ElseIf y1 = 1 Then
ss$ = w$
End If
If y1 <> 0 Then
If Not CanKillFile(CFname(s$)) Then FilePathNotForUser: ProcName = False: Exit Function
If Not RenameFile(s$, ss$) Then NoRename

Exit Function
End If
Else
rest$ = s$ + rest$
End If
If IsStrExp(bstack, rest$, s$) Then
 If Not IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then ProcName = False: Exit Function
If IsStrExp(bstack, rest$, ss$) Then
On Error Resume Next
If Not RenameFile(s$, ss$) Then NoRename
On Error GoTo 0
Else
ProcName = False
Exit Function
End If
Else
ProcName = False
Exit Function
End If
End Function
Sub GetitObject(var, Optional cc, Optional serverclass1)
Dim aa As Object, b As GUID, serverclass As String
If IsMissing(cc) Then
If Left$(serverclass1, 1) = "{" Then
    serverclass = serverclass1
    serverclass = strProgIDfromSrting(serverclass)
Else
    serverclass = serverclass1
End If
If serverclass = "" Then Exit Sub
On Error Resume Next
Set aa = GetObject(, serverclass)
If Err Then
MyEr Err.Description, Err.Description
End If
ElseIf IsMissing(serverclass1) Then
On Error Resume Next
cc = GetDosPath((cc))
If cc <> "" Then
Set aa = GetObject(cc)
Else
MissFile
End If
If Err Then
MyEr Err.Description, Err.Description
End If
Else
On Error Resume Next
cc = GetDosPath((cc))
Set aa = GetObject(cc, serverclass1)
If Err Then
MyEr Err.Description, Err.Description
End If
End If
Set var = aa
End Sub
Function createAnobject(bstack As basetask, b$)
Dim s$, s1$, k As Integer, ob
Set ob = Nothing
If IsStrExp(bstack, b$, s$) Then
k = 1
End If
If FastSymbol(b$, ",") Then
    If IsStrExp(bstack, b$, s1$) Then
    k = k + 10
    End If
End If
If FastSymbol(b$, ")") Then
    Select Case k
    Case 1
        GetitObject ob, s$
    Case 10
        GetitObject ob, , s1$
    Case 11
        GetitObject ob, s$, s1$
    End Select
    If Not ob Is Nothing Then
        Set bstack.lastobj = ob
        createAnobject = True
    End If
End If
End Function
Function GetThisModuleName(R$) As Boolean
GetThisModuleName = True
R$ = GetName(here$)
If Len(R$) = 0 Then Exit Function
    If AscW(R$) = 8191 Then
        If InStr(here$, R$) > 0 Then
           R$ = GetName(Left$(here$, Len(here$) - Len(R$)))
           R$ = Left$(R$, Len(R$) - 1)
        Else
            R$ = sbf(val(Mid$(here$, rinstr(here$, "[") + 1))).sbgroup
            If Len(R$) > 0 Then
                R$ = Mid$(R$, rinstr(R$, ".", 2) + 1)
                R$ = Left$(R$, Len(R$) - 1)
            End If
        End If
    End If
    If InStr(R$, "[") > 0 Then R$ = GetName(R$)
    If InStr(R$, ").") Then R$ = Mid$(R$, InStr(R$, ").") + 2)
End Function
Public Function IsPoint(bstack As basetask, a$, R As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray
Dim r2 As Variant, r3 As Variant, r4 As Variant
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                    If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    If FastSymbol(a$, ",") Then
                        If Not IsExp(bstack, a$, r2, , True) Then: MissParam a$: Exit Function
                        If FastSymbol(a$, ",") Then
                            If Not IsExp(bstack, a$, r3, , True) Then: MissParam a$: Exit Function
                            If FastSymbol(a$, ",") Then
                                If Not IsExp(bstack, a$, r4, , True) Then: MissParam a$: Exit Function
                                    R = SetDIBPixel(var(w1), r2, r3, mycolor(r4))
                                Else
                                    R = GetDIBPixel(var(w1), r2, r3)
                                End If
                                If SG < 0 Then R = -R
                                IsPoint = FastSymbol(a$, ")", True)
                            Else
                                MissParam a$: Exit Function
                            End If
                        Else
                            MissParam a$: Exit Function
                        End If
                    Else
                    noImage a$
                    Exit Function
        End If
            Else
                    
                    MissFuncParameterStringVarMacro a$
                    
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    If FastSymbol(a$, ",") Then
                        If Not IsExp(bstack, a$, r2, , True) Then: MissParam a$: pppp.SwapItem w2, sV: Exit Function
                        If FastSymbol(a$, ",") Then
                            If Not IsExp(bstack, a$, r3, , True) Then: MissParam a$: pppp.SwapItem w2, sV: Exit Function
                            If FastSymbol(a$, ",") Then
                                If Not IsExp(bstack, a$, r4, , True) Then: MissParam a$: pppp.SwapItem w2, sV: Exit Function
                                R = SetDIBPixel(sV, r2, r3, mycolor(r4))
                            Else
                                R = GetDIBPixel(sV, r2, r3)
                            End If
                            If SG < 0 Then R = -R
                            pppp.SwapItem w2, sV
                            IsPoint = FastSymbol(a$, ")", True)
                        Else
                            pppp.SwapItem w2, sV
                            MissParam a$: Exit Function
                        End If
                    Else
                        pppp.SwapItem w2, sV
                        MissParam a$: Exit Function
                    End If
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End Function

Function StaticNew(bstack As basetask, b$, w$, Lang As Long) As Boolean
Dim p As Variant, ii As Long, ss$, usehandler As mHandler, H As Variant

If bstack.StaticCollection Is Nothing Then

Set bstack.StaticCollection = New FastCollection
If Not bstack.IamThread Then
    bstack.SetBacket "%_" + bstack.StaticInUse
End If
End If
Do
    Select Case IsLabel(bstack, b$, w$)
    Case 1
        If GetlocalVar(w$, ii) Then
            MyEr "Variable exist as local", "Η μεταβλητή υπάρχει ως τοπική"
            StaticNew = False
            Exit Function
        End If
        If Not bstack.ExistVar(w$) Then
            If FastSymbol(b$, "=") Then
                    If Not IsExp(bstack, b$, p) Then SyntaxError: Exit Function
                    Dim anything As Object
                    Set anything = bstack.lastobj
                    If CheckIsmArray(anything) Then
                            Set usehandler = New mHandler
                            With usehandler
                            .t1 = 3
                            Set .objref = anything
                            End With
                            Set p = usehandler
                            Set H = usehandler
                      bstack.SetVarobJvalue w$, H
                      Set usehandler = Nothing
                      Set H = Nothing
                    ElseIf CheckLastHandler(anything) Then
                        Set usehandler = anything
                        If usehandler.t1 = 2 Then
                            bstack.SetVarobJvalue w$, anything
                        ElseIf usehandler.t1 = 1 Then
                            bstack.SetVarobJvalue w$, anything
                        ElseIf usehandler.t1 = 3 Then
                            bstack.SetVarobJ w$, anything
                        ElseIf usehandler.t1 = 4 Then
                            Set p = usehandler
                            bstack.SetVarobJvalue w$, p
                        Else
                            GoTo conthere
                        End If
                    ElseIf Not bstack.lastobj Is Nothing Then
                        If TypeOf bstack.lastobj Is Group Then
                            If bstack.lastobj.IamApointer Then
                                If bstack.lastobj.link.IamFloatGroup Then
                                    bstack.SetVarobJvalue w$, bstack.lastobj
                                Else
                                    GoTo aaaa1
                                End If
                            Else
aaaa1:
                            Set bstack.lastobj = Nothing
                            MyEr "only for pointers for float groups", "μόνο για δείκτες σε μη στατικά αντικείμενα"
                            Exit Function
                        End If
                    End If
                Else
                    bstack.SetVar w$, p
                End If
                Set bstack.lastobj = Nothing
            ElseIf IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
               If IsLabelSymbolNew(b$, "ΑΡΙΘΜΟΣ", "DECIMAL", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                    p = CDec(p)
                ElseIf IsLabelSymbolNew(b$, "ΔΙΠΛΟΣ", "DOUBLE", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                    p = CDbl(p)
                ElseIf IsLabelSymbolNew(b$, "ΑΠΛΟΣ", "SINGLE", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                    p = CSng(p)
                ElseIf IsLabelSymbolNew(b$, "ΛΟΓΙΚΟΣ", "BOOLEAN", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                    p = CBool(p)
                ElseIf IsLabelSymbolNew(b$, "ΜΑΚΡΥΣ", "LONG", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                    p = CLng(p)
                ElseIf IsLabelSymbolNew(b$, "ΑΚΕΡΑΙΟΣ", "INTEGER", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                    p = CInt(p)
                ElseIf IsLabelSymbolNew(b$, "ΛΟΓΙΣΤΙΚΟ", "CURRENCY", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                        If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                    p = CCur(p)
                ElseIf IsEnumAs(bstack, b$, p) Then
                    bstack.SetVarobJ w$, p
                    GoTo aaa1
                Else
                    MyEr "No type found", "δεν βρήκα τύπο"
                    Exit Function
                End If
                bstack.SetVar w$, p
            ElseIf FastSymbol(b$, "->", , 2) Then
                If GetPointer(bstack, b$) Then
                    If bstack.lastpointer.IamFloatGroup Then
                        bstack.SetVarobJvalue w$, bstack.lastpointer
                    Else
                        Set bstack.lastpointer = Nothing
                        GoTo aaaa1
                    End If
                Else
                    GoTo aaaa1
                End If
            Else
               bstack.SetVar w$, p
            End If
            Set bstack.lastobj = Nothing
        ElseIf FastSymbol(b$, "=") Then
            ii = 1
            ss$ = aheadstatus(b$, False, ii)
            b$ = Mid$(b$, ii)
        ElseIf Fast2VarNoTrim(b$, "ΩΣ", 2, "AS", 2, 3, ii) Then
            ii = 1
            ss$ = aheadstatus(b$, False, ii)
            b$ = Mid$(b$, ii)
        ElseIf FastSymbol(b$, "->", , 2) Then
            ii = 1
            ss$ = aheadstatus(b$, False, ii)
            b$ = Mid$(b$, ii)
        End If
        StaticNew = True
    Case 3
        If Not bstack.ExistVar(w$) Then
            If FastSymbol(b$, "=") Then If Not IsStrExp(bstack, b$, ss$) Then SyntaxError: Exit Function
            bstack.SetVar w$, ss$
        ElseIf FastSymbol(b$, "=") Then
            ii = 1
            ss$ = aheadstatus(b$, False, ii)
            b$ = Mid$(b$, ii)
        End If
        StaticNew = True
    Case 4
        If Not bstack.ExistVar(w$) Then
            If FastSymbol(b$, "=") Then
            If Not IsExp(bstack, b$, p) Then SyntaxError: Exit Function
            ElseIf IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
    
               If IsLabelSymbolNew(b$, "ΑΡΙΘΜΟΣ", "DECIMAL", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                    If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                    p = CDec(p)
            ElseIf IsLabelSymbolNew(b$, "ΔΙΠΛΟΣ", "DOUBLE", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                    If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                p = CDbl(p)
            ElseIf IsLabelSymbolNew(b$, "ΑΠΛΟΣ", "SINGLE", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                    If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                p = CSng(p)
            ElseIf IsLabelSymbolNew(b$, "ΛΟΓΙΚΟΣ", "BOOLEAN", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                    If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                p = CBool(p)
            ElseIf IsLabelSymbolNew(b$, "ΜΑΚΡΥΣ", "LONG", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                    If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                p = CLng(p)
            ElseIf IsLabelSymbolNew(b$, "ΑΚΕΡΑΙΟΣ", "INTEGER", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                    If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If
                p = CInt(p)
            ElseIf IsLabelSymbolNew(b$, "ΛΟΓΙΣΤΙΚΟ", "CURRENCY", Lang, , , , False) Then
                    If FastSymbol(b$, "=") Then
                    If Not IsNumberD2(b$, p) Then missNumber: Exit Function
                    Else
                        p = 0
                    End If

                p = CCur(p)
            Else
            MyEr "No type found", "δεν βρήκα τύπο"
            Exit Function
            End If
            
            If Typename(p) <> "boolean" Then p = Int(p)
            End If
            bstack.SetVar w$, MyRound(p)
        ElseIf FastSymbol(b$, "=") Then
         ii = 1
            ss$ = aheadstatus(b$, False, ii)
        b$ = Mid$(b$, ii)
        ElseIf Fast2VarNoTrim(b$, "ΩΣ", 2, "AS", 2, 3, ii) Then
         ii = 1
            ss$ = aheadstatus(b$, False, ii)
        b$ = Mid$(b$, ii)
    End If
        StaticNew = True
        
    Case Else
conthere:
        MyErMacro b$, "No static for that type " + w$, "Όχι στατική για αυτό το τύπο " + w$
        Exit Function
    End Select
aaa1:
 Loop Until Not FastSymbol(b$, ",")
End Function

Function MyDocument(basestack As basetask, rest$, Lang As Long, Optional alocal As Boolean) As Boolean
Dim ss$, s$, what$, x1 As Long, i As Long, it As Long, pppp As mArray
MyDocument = True
ss$ = vbNullString
Do
    x1 = Abs(IsLabel(basestack, rest$, what$))
        If basestack.priveflag Then what$ = ChrW(&HFFBF) + what$
    If x1 = 3 Or x1 = 6 Then
        If x1 = 3 Then
        
                
                
            If Not FastSymbol(rest$, "<") Then  ' get local var first
            If alocal Then
            i = globalvar(basestack.GroupName & what$, s$)  ' MAKE ONE  '
             GoTo makeitnow
            ElseIf GetlocalVar(basestack.GroupName & what$, i) Then
            GoTo there0
            ElseIf GetVar(basestack, basestack.GroupName & what$, i) Then
            GoTo there0
            Else
            i = globalvar(basestack.GroupName & what$, s$)  ' MAKE ONE  '
             GoTo makeitnow
            End If
            ElseIf GetVar(basestack, basestack.GroupName & what$, i) Then
            
there0:
                s$ = var(i)
                MakeitObject var(i)
                CheckVar var(i), s$
                GoTo there1
            Else
        
                i = globalvar(basestack.GroupName & what$, s$) ' MAKE ONE
                If i <> 0 Then
makeitnow:
                    MakeitObject var(i)
there1:
                    If FastSymbol(rest$, "=") Then
                        If IsStrExp(basestack, rest$, s$) Then
                            CheckVar var(i), s$
                        Else
                            MissStringExpr
                            MyDocument = False
                        End If
                    Else
                    ' DO NOTHING
                    End If
                End If
            End If
        Else
            ' ARRAYf
            If neoGetArray(basestack, what$, pppp, here$ <> "") Then   ' basestack.GroupName &
                If Not NeoGetArrayItem(pppp, basestack, what$, it, rest$) Then MyDocument = False: Exit Function
                x1 = 0
                If Not MyIsObject(pppp.item(it)) Then
                    s$ = pppp.item(it)
                    Set pppp.item(it) = New Document
                    If s$ <> "" Then pppp.item(it).textDoc = s$
                    If FastSymbol(rest$, "=") Then
                        If IsStrExp(basestack, rest$, s$) Then
                            CheckVar pppp.item(it), s$
                        Else
                            MissStringExpr
                        MyDocument = False
                        End If
                    End If
                Else
                If FastSymbol(rest$, "=") Then
                    If IsStrExp(basestack, rest$, s$) Then
                        CheckVar pppp.item(it), s$
                    Else
                        MissStringExpr
                        MyDocument = False
                    End If
                Else
                    Exit Do
                   End If
                End If
                MyDocument = True
            Else
            MyErMacro rest$, "array has no dimension", "ο πίνακας δεν έχει οριστεί"
             MyDocument = False
             Exit Function
            End If
          
            End If
    Else
    SyntaxError
    MyDocument = False
    
    End If
    Loop Until Not FastSymbol(rest$, ",")

End Function

Function GrabFrame() As String
Dim p As New cDIBSection

p.CreateFromPicture hDCToPicture(GetDC(0), AVI.Left / DXP, AVI.top / DYP, AVI.Width / DXP, AVI.Height / DYP - 1)
If p.Height > 0 Then

GrabFrame = DIBtoSTR(p)
End If
End Function
