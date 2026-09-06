VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "x86"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' 32 Bit X86 Assembler
'
' Arne Elster 2007 / 2008


' TODO:
'       * support for instructions like 2-IMUL
'       * more complex expressions for all arguments
'       * Unsigned values (only possible with hex)


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal cBytes As Long _
)

Private labels As FastCollection
Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES  As Long = 16&
Private Const IMAGE_SIZEOF_SHORT_NAME           As Long = 8&
Private Const IMAGE_NT_OPTIONAL_HDR32_MAGIC     As Long = &H10B&
Private Const IMAGE_DOS_HDR16_MAGIC             As Long = &H5A4D&
Private Const IMAGE_DOS_HDR32_MAGIC             As Long = &H4550&
Private Const IMAGE_FILE_MACHINE_I386           As Long = &H14C&
Private Const IMAGE_REL_BASED_HIGHLOW           As Byte = 3

Private Const DOS_CODE_RELOCATIONS As String = _
        "0E1FBA0E00B409CD21B8014CCD21546869732070726" & _
        "F6772616D2063616E6E6F742062652072756E20696E" & _
        "20444F53206D6F64652E0D0D0A2400000000000000"

Private Const MEM_SECTION_SIZE  As Long = 4096
Private Const FILE_SECTION_SIZE As Long = 512

Private Const CHAR_AT           As Long = 64
Private Const CHAR_SPACE        As Long = 32
Private Const CHAR_LINEFEED     As Long = 10
Private Const CHAR_CARRIAGE     As Long = 13

Private Const CHAR_QUOTE        As Long = 34
Private Const CHAR_STOP         As Long = 46
Private Const CHAR_SEMICOLON    As Long = 59
Private Const CHAR_COLON        As Long = 58
Private Const CHAR_PLUS         As Long = 43
Private Const CHAR_MINUS        As Long = 45
Private Const CHAR_ASTERISK     As Long = 42
Private Const CHAR_AMPERSAND    As Long = 38
Private Const CHAR_SEPARATOR    As Long = 44
Private Const CHAR_UNDERSCORE   As Long = 95
Private Const CHAR_VERT_BAR     As Long = 124
Private Const CHAR_SHARP        As Long = 35

Private Const CHAR_BRACKET_L    As Long = 91
Private Const CHAR_BRACKET_R    As Long = 93
Private Const CHAR_PARENTH_L    As Long = 40
Private Const CHAR_PARENTH_R    As Long = 41

Private Const CHAR_NUMBER_0     As Long = 48
Private Const CHAR_NUMBER_7     As Long = 55

Private Const CHAR_NUMBER_9     As Long = 57

Private Const CHAR_ALPHA_UA     As Long = 65
Private Const CHAR_ALPHA_UZ     As Long = 90
Private Const CHAR_ALPHA_LA     As Long = 97
Private Const CHAR_ALPHA_LZ     As Long = 122

Private Const REG_COUNT         As Long = 24
Private Const MAX_PARAMETERS    As Long = 3
Private Const MAX_OPCODE_LEN    As Long = 4

Public Enum PESubsystem
    Subsystem_GUI = 2
    Subsystem_CUI = 3
End Enum

Private Enum SectionCharacteristics
    IMAGE_SCN_TYPE_NO_PAD = &H8&
    IMAGE_SCN_CNT_CODE = &H20&
    IMAGE_SCN_CNT_INITIALIZED_DATA = &H40&
    IMAGE_SCN_CNT_UNINITIALIZED_DATA = &H80&
    IMAGE_SCN_LNK_OTHER = &H100&
    IMAGE_SCN_LNK_INFO = &H200&
    IMAGE_SCN_LNK_REMOVE = &H800&
    IMAGE_SCN_LNK_COMDAT = &H1000&
    IMAGE_SCN_NO_DEFER_SPEC_EXC = &H4000&
    IMAGE_SCN_GPREL = &H8000&
    IMAGE_SCN_MEM_PURGEABLE = &H20000
    IMAGE_SCN_MEM_LOCKED = &H40000
    IMAGE_SCN_MEM_PRELOAD = &H80000
    IMAGE_SCN_LNK_NRELOC_OVFL = &H1000000
    IMAGE_SCN_MEM_DISCARDABLE = &H2000000
    IMAGE_SCN_MEM_NOT_CACHED = &H4000000
    IMAGE_SCN_MEM_NOT_PAGED = &H8000000
    IMAGE_SCN_MEM_SHARED = &H10000000
    IMAGE_SCN_MEM_EXECUTE = &H20000000
    IMAGE_SCN_MEM_READ = &H40000000
    IMAGE_SCN_MEM_WRITE = &H80000000
End Enum

Private Enum IMAGE_FILE_CHARACTERISTICS
    IMAGE_FILE_RELOCS_STRIPPED = &H1&
    IMAGE_FILE_EXECUTABLE_IMAGE = &H2&
    IMAGE_FILE_LINE_NUMS_STRIPPED = &H4&
    IMAGE_FILE_LOCAL_SYMS_STRIPPED = &H8&
    IMAGE_FILE_AGGRESSIVE_WS_TRIM = &H10&
    IMAGE_FILE_LARGE_ADDRESS_AWARE = &H20&
    IMAGE_FILE_16BIT_MACHINE = &H40&
    IMAGE_FILE_BYTES_REVERSED_LO = &H80&
    IMAGE_FILE_32BIT_MACHINE = &H100&
    IMAGE_FILE_DEBUG_STRIPPED = &H200&
    IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = &H400&
    IMAGE_FILE_NET_RUN_FROM_SWAP = &H800&
    IMAGE_FILE_SYSTEM = &H1000&
    IMAGE_FILE_DLL = &H2000&
    IMAGE_FILE_UP_SYSTEM_ONLY = &H4000&
    IMAGE_FILE_BYTES_REVERSED_HI = &H8000&
End Enum

Private Enum OptHeaderTbls
    ETableExport = 0
    ETableImport
    ETableResource
    ETableException
    ETableCertificate
    ETableRelocation
    ETableDebug
    ETableArchitecture
    ETableGlobalPtr
    ETableThreadStorage
    ETableLoadConfig
    ETableBoundImport
    ETableIAT
    ETableDelayImportDescriptor
    ETableCOMPlusRuntime
    ETableReserved
End Enum

Private Type ModRM
    Mod                         As Long
    rm                          As Long
    reg                         As Long
    disp                        As Long
    DispSize                    As ParamSize
End Type

Private Type SIB
    sscale                      As Long
    index                       As Long
    Base                        As Long
End Type

Private Type ASMLabel
    Name                        As String
    Instruction                 As Long
    Offset                      As Long
End Type

Private Type ASMExtern
    LibName                     As String
    Functions()                 As String
    FunctionCount               As Long
End Type

Private Type Pointer
    Registers(REG_COUNT - 1)    As Long
    UsedRegisters               As Long
    Displacement                As Long
    DispSize                    As ParamSize
End Type

Private Type PointerInfo
    TokenIndex                  As Long
    RegisterCount               As Long
    RegisterMultiples           As Boolean
    HasDisplacement             As Boolean
    DispSize                    As ParamSize
    Ptr                         As Pointer
End Type

Private Type RawData
    Size                        As ParamSize
    Values()                    As Long
    ValueCount                  As Long
End Type

Private Type ASMArgument
    TType                       As ParamType
    Size                        As ParamSize
    Pointer                     As PointerInfo
    Register                    As ASMRegisters
    FPURegister                 As ASMFPURegisters
    MMRegister                  As ASMXMMRegisters
    SymbolIndex                 As Long
    Value                       As Long
End Type

Private Type ASMInstruction
    Mnemonic                    As String
    Segment                     As ASMSegmentRegs
    Args(MAX_PARAMETERS - 1)    As ASMArgument
    OpCodeIndex                 As Long
    ArgCount                    As Long
    Size                        As Long
    Offset                      As Long
    Flags                       As OpCodePrefixes
    Data                        As RawData
    Line                        As Long
    Section                     As String
End Type

Private Type Scanner
    Source()                    As Byte
    Length                      As Long
    Position                    As Long
    Line                        As Long
    LinePos                     As Long
    Section                     As String
    NextIsEOI                   As Boolean
    LastWasEOI                  As Boolean
    NextToken                   As ASMToken
    CurToken                    As ASMToken
End Type


Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics             As Long
    TimeDateStamp               As Long
    MajorVersion                As Integer
    MinorVersion                As Integer
    Name                        As Long
    Base                        As Long
    NumberOfFunctions           As Long
    NumberOfNames               As Long
    AddressOfFunctions          As Long
    AddressOfNames              As Long
    AddressOfNameOrdinals       As Long
End Type

Private Type IMAGE_IMPORT_DIRECTORY
    ImportLookupTable           As Long
    TimeDateStamp               As Long
    ForwardChain                As Long
    ModuleName                  As Long
    ImportAddressTable          As Long
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress              As Long
    Size                        As Long
End Type

Private Type IMAGE_SECTION_HEADER
    SectionName(IMAGE_SIZEOF_SHORT_NAME - 1) As Byte
    VirtSizePhysAddr            As Long
    VirtualAddress              As Long
    SizeOfRawData               As Long
    PointerToRawData            As Long
    PointerToRelocations        As Long
    PointerToLinenumbers        As Long
    NumberOfRelocations         As Integer
    NumberOfLinenumbers         As Integer
    Characteristics             As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic                       As Integer
    MajorLinkerVersion          As Byte
    MinorLinkerVersion          As Byte
    SizeOfCode                  As Long
    SizeOfInitializedData       As Long
    SizeOfUninitializedData     As Long
    AddressOfEntryPoint         As Long
    BaseOfCode                  As Long
    BaseOfData                  As Long
    ImageBase                   As Long
    SectionAlignment            As Long
    FileAlignment               As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion           As Integer
    MinorImageVersion           As Integer
    MajorSubsystemVersion       As Integer
    MinorSubsystemVersion       As Integer
    Win32VersionValue           As Long
    SizeOfImage                 As Long
    SizeOfHeaders               As Long
    CheckSum                    As Long
    Subsystem                   As Integer
    DllCharacteristics          As Integer
    SizeOfStackReserve          As Long
    SizeOfStackCommit           As Long
    SizeOfHeapReserve           As Long
    SizeOfHeapCommit            As Long
    LoaderFlags                 As Long
    NumberOfRvaAndSizes         As Long
    DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1)  As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_DOS_HEADER
    Magic                       As Integer
    BytesInLastPage             As Integer
    Pages                       As Integer
    Relocations                 As Integer
    ParagraphsInHeader          As Integer
    MinAlloc                    As Integer
    MaxAlloc                    As Integer
    InitialSS                   As Integer
    InitialSP                   As Integer
    CheckSum                    As Integer
    InitialIP                   As Integer
    InitialCS                   As Integer
    RelocationTableFileAddress  As Integer
    OverlayNumber               As Integer
    Reserved1(3)                As Integer
    OEMIdentifier               As Integer
    OEMInformation              As Integer
    Reserved2(9)                As Integer
    NewHeaderOffset             As Long
End Type

Private Type IMAGE_FILE_HEADER
    Machine                     As Integer
    NumberOfSections            As Integer
    TimeDateStamp               As Long
    PointerToSymbolTable        As Long
    NumberOfSymbols             As Long
    SizeOfOptionalHeader        As Integer
    Characteristics             As Integer
End Type

Private Type IMAGE_NT_HEADERS
    Signature                   As Long
    FileHeader                  As IMAGE_FILE_HEADER
    OptionalHeader              As IMAGE_OPTIONAL_HEADER
End Type

Private m_udtScanner            As Scanner

Private m_clsTokens()           As ASMToken
Private m_lngTokenCount         As Long
Private m_lngCurToken           As Long

Private m_udtLabels()           As ASMLabel
Private m_lngLabelCount         As Long

Private m_udtExtern()           As ASMExtern
Private m_lngExternCount        As Long

Private m_udtInstrs()           As ASMInstruction
Private m_lngInstrCount         As Long

Private m_strLastError          As String
Private m_strLastErrorSection   As String
Private m_lngLastErrLine        As Long

Private m_btOutput()            As Byte
Private m_lngOutSize            As Long
Private m_lngOutPos             As Long

Private m_udeSubsystem          As PESubsystem

Private m_blnWritePE            As Boolean
' new for DLL
Private Type ASMExport
    Name                        As String
    LabelName                   As String
End Type

Private Type ASMRelocation
    Offset                      As Long
    Type                        As Byte
End Type
Private m_blnWriteDLL           As Boolean
Private m_udtExports()          As ASMExport
Private m_strDLLName            As String
Private m_lngExportSize         As Long
Private m_lngExportCount        As Long
Private m_btExportData()        As Byte

Private m_udtRelocs()           As ASMRelocation
Private m_lngRelocCount         As Long
Private m_lngRelocSize          As Long
Private m_btRelocData()         As Byte


Private m_lngPECodeSize         As Long

Private m_lngBaseAddress        As Long
Private m_SizeOfStackReserve    As Long

Private m_SizeOfStackCommit     As Long
Private m_SizeOfHeapReserve     As Long
Private m_SizeOfHeapCommit      As Long


Private Sub Class_Initialize()
    InitInstructions
    m_udeSubsystem = Subsystem_CUI
    m_SizeOfStackReserve = &H100000
    m_SizeOfStackCommit = &H1000
    m_SizeOfHeapReserve = &H100000
    m_SizeOfHeapCommit = &H1000
End Sub
Public Property Get LabelOffset(Name As String) As Long
    If labels Is Nothing Then Exit Property
    If labels.ExistKey(Name) Then
    With m_udtLabels(labels.Value)
        LabelOffset = .Offset - m_lngBaseAddress
    End With
    End If
End Property
Public Property Get LabelPtr(Name As String) As Long
    If labels Is Nothing Then Exit Property
    If labels.ExistKey(Name) Then
    With m_udtLabels(labels.Value)
        LabelPtr = .Offset
    End With
    End If
End Property
Public Property Get Subsystem() As PESubsystem
    Subsystem = m_udeSubsystem
End Property


Public Property Let Subsystem(ByVal lngVal As PESubsystem)
    m_udeSubsystem = lngVal
End Property


Public Property Get PEHeader() As Boolean
    PEHeader = m_blnWritePE
End Property


Public Property Let PEHeader(ByVal blnValue As Boolean)
    m_blnWritePE = blnValue
End Property

Private Sub AddRelocation(ByVal lngRVA As Long)
    ReDim Preserve m_udtRelocs(m_lngRelocCount) As ASMRelocation
    m_udtRelocs(m_lngRelocCount).Offset = lngRVA ' +RoundToSectionSize(GetPEHeaderSize(sections))..
    m_udtRelocs(m_lngRelocCount).Type = IMAGE_REL_BASED_HIGHLOW
    m_lngRelocCount = m_lngRelocCount + 1
End Sub

Public Property Get BaseAddress() As Long
    BaseAddress = m_lngBaseAddress
End Property


Public Property Let BaseAddress(ByVal lngVal As Long)
    If lngVal < 0 Then
    MyEr "Image Base < 0 invalid", "όχι αρνητική τιμή για βάση"
    lngVal = &H1000000
    End If
    ' ? LOW-LEVEL FIX: Enforce Windows 64 KB (0x10000) Allocation Granularity Alignment.
    ' If the passed address is not a multiple of 64 KB, we automatically round it UP
    ' to the next valid 64 KB boundary to ensure the Windows Loader accepts the file.
    Const IMAGE_BASE_ALIGNMENT As Long = &H10000  ' 64 KB in Hex
    
    If (lngVal Mod IMAGE_BASE_ALIGNMENT) <> 0 Then
        m_lngBaseAddress = lngVal + (IMAGE_BASE_ALIGNMENT - (lngVal Mod IMAGE_BASE_ALIGNMENT))
    Else
        m_lngBaseAddress = lngVal
    End If
End Property


Public Property Get LastErrorMessage() As String
    LastErrorMessage = m_strLastError
End Property


Public Property Get LastErrorSection() As String
    LastErrorSection = m_strLastErrorSection
End Property


Public Property Get LastErrorLine() As Long
    LastErrorLine = m_lngLastErrLine
End Property
Public Function GetOutPtr() As Long
    On Error Resume Next
    GetOutPtr = VarPtr(m_btOutput(0))
End Function

Public Function GetOutput() As Byte()
    GetOutput = m_btOutput
End Function


Public Property Get OutputSize() As Long
    OutputSize = m_lngOutSize
End Property


' 1. tokenize the input string
' 2. collect all labels
' 3. find all instructions
' 4. find OpCodes for instructions
'    get their sizes and calculate label offsets
' 5. now that the label offsets are known,
'    finally parse pointers
' 6. write instructions to output
' 7. if in PE mode, write IAT (import address table)
Public Function Assemble( _
    strASM As String, _
    Optional ByVal OnlySize As Boolean = False _
) As Boolean
    ' remove tabs
    '  ChrW(160)
    strASM = Replace(strASM, ChrW(160), " ")
    strASM = Replace(strASM, Chr$(9), " ")
    strASM = Replace(strASM, ";" + vbCr, "; " + vbCr)
    If InStr(1, strASM, "EQU", vbTextCompare) > 0 Then
        strASM = PreProcessAsm(strASM)
    End If
    Set labels = New FastCollection
    labels.UcaseKeys = True
    ScannerInit strASM
    
    m_lngTokenCount = 0
    m_lngLabelCount = 0
    m_lngInstrCount = 0
    m_lngCurToken = 0
    m_lngOutSize = 0
    m_lngOutPos = 0
    m_lngExternCount = 0
    
    m_strLastError = ""
    m_lngLastErrLine = 0
    
    TokenizeInput
    If FindLabels() Then
        If ParseInstructions() Then
            Dim Sections As Integer
            If m_blnWriteDLL Then Sections = 3
            If GetInstructionSizes(Sections) Then
                If OnlySize Then
                    Assemble = True
                Else
                    If ParsePointers() Then
                        If m_blnWritePE Or m_blnWriteDLL Then
                            If Not WritePEHeader() Then Exit Function
                        End If
                        
                        If AssembleInstructions() Then
                            If m_blnWritePE Or m_blnWriteDLL Then
                                If m_lngExternCount > 0 Then
                                OutputJumpTo RoundToMinSize(OutputPosition)
                                
                                WritePEImports Sections
                                End If
                                If m_blnWriteDLL Then
                                    ' Write Dynamic Link Exports Section
                                    If m_lngExportSize > 0 Then
                                        OutputJumpTo RoundToMinSize(OutputPosition)
                                        OutputMem VarPtr(m_btExportData(0)), m_lngExportSize
                                    End If
                                   ' If False Then
                                    ' =============================================================
                                    '  FIX: Safely generate and write relocation data
                                    ' =============================================================
                                    'If m_lngRelocSize > 0 Then
                                        Dim lngTempSize As Long
                                        
                                        ' Build the relocation byte array. The array resizing (ReDim)
                                        ' happens safely inside BuildRelocationSection.
                                        If m_lngExternCount > 0 Then
                                        m_btRelocData = BuildRelocationSection(Sections + 1, MEM_SECTION_SIZE, lngTempSize)
                                        Else
                                        m_btRelocData = BuildRelocationSection(Sections, MEM_SECTION_SIZE, lngTempSize)
                                        End If
                                        ' Jump to the pre-allocated space for the .reloc section
                                        OutputJumpTo RoundToMinSize(OutputPosition)
                                        
                                        ' Use the safely returned array to write bytes to the file
                                        OutputMem VarPtr(m_btRelocData(0)), lngTempSize
                                    'End If
                                    ' =============================================================
                                    'End If
                                End If
                                ' =============================================================
                                ' CALCULATE & PATCH THE PE CHECKSUM (DYNAMIC HYBRID METHOD)
                                ' =============================================================
                                Dim lngChecksumPos As Long
                                Dim lngNewChecksum As Long
                                Dim mzHdrTemp As IMAGE_DOS_HEADER
                                Dim PE_Offset As Long
                                
                                ' Safely read the PE signature offset from the MZ Header structure (e_lfanew)
                                CopyMemory mzHdrTemp, m_btOutput(0), Len(mzHdrTemp)
                                PE_Offset = mzHdrTemp.NewHeaderOffset
                                
                                ' OptionalHeader.CheckSum is exactly 88 bytes after the PE Signature start
                                ' (4 bytes Signature + 20 bytes FileHeader + 64 bytes inside OptionalHeader)
                                lngChecksumPos = PE_Offset + 4 + 20 + 64
                                
                                ' Compute the valid checksum over the fully generated file array
                                lngNewChecksum = CalculatePECheckSum(lngChecksumPos)
                                
                                ' Inject the calculated checksum back into the binary data array
                                CopyMemory m_btOutput(lngChecksumPos), lngNewChecksum, 4
                                ' =============================================================
                            End If
                            Assemble = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function


Private Sub WritePEImports(ByVal Sections As Integer)
    Dim lngRVAIAT   As Long
    Dim ntHdr       As IMAGE_NT_HEADERS
    Dim scHdr       As IMAGE_SECTION_HEADER
    
    lngRVAIAT = RoundToSectionSize(Len(ntHdr) + Len(scHdr) * Sections)
    lngRVAIAT = lngRVAIAT + RoundToSectionSize(m_lngPECodeSize)
    WriteIAT lngRVAIAT
    WriteIIDs lngRVAIAT
    WriteIAT lngRVAIAT
    WriteImportedNames
End Sub


Private Sub WriteImportedNames()
    Dim i       As Long
    Dim j       As Long
    
    For i = 0 To m_lngExternCount - 1
        For j = 0 To m_udtExtern(i).FunctionCount - 1
            OutputInteger 0
            WriteStr0Ev m_udtExtern(i).Functions(j)
        Next
        WriteStr0Ev m_udtExtern(i).LibName
    Next
End Sub


Private Sub WriteStr0Ev(ByVal strN As String)
    Dim btN()   As Byte
    
    btN = StrConv(strN & ChrW$(0), vbFromUnicode)
    OutputMem VarPtr(btN(0)), UBound(btN) + 1
    
    If (UBound(btN) + 1) Mod 2 = 1 Then
        OutputByte 0
    End If
End Sub


Private Sub WriteIIDs(ByVal Base As Long)
    Dim i       As Long
    Dim j       As Long
    Dim im      As IMAGE_IMPORT_DIRECTORY
    Dim eim     As IMAGE_IMPORT_DIRECTORY
    
    For i = 0 To m_lngExternCount - 1
        With im
            .ModuleName = Base + GetRelOfLibname(i)
            .ImportAddressTable = Base + GetIATLibStart(i)
            .ImportLookupTable = Base + GetILTLibStart(i)
        End With
        
        OutputMem VarPtr(im), Len(im)
    Next
    
    OutputMem VarPtr(eim), Len(eim)
End Sub


Private Sub WriteIAT(ByVal Base As Long)
    Dim i   As Long
    Dim j   As Long
    
    For i = 0 To m_lngExternCount - 1
        For j = 0 To m_udtExtern(i).FunctionCount - 1
            OutputLong Base + GetRelOfFncname(i, j)
        Next
        OutputLong 0
    Next
End Sub


' address of a function name in the imports section
' relative to the section's start
Private Function GetRelOfFncname(ByVal libdx As Long, ByVal fncidx As Long) As Long
    Dim i   As Long
    Dim j   As Long
    Dim SZ  As Long
    Dim im  As IMAGE_IMPORT_DIRECTORY
    
    For i = 0 To m_lngExternCount - 1
        SZ = SZ + 4 * (m_udtExtern(i).FunctionCount + 1) * 2
        SZ = SZ + Len(im)
    Next
    If m_lngExternCount > 0 Then SZ = SZ + Len(im)
    
    For i = 0 To m_lngExternCount - 1
        For j = 0 To m_udtExtern(i).FunctionCount - 1
            If (i = libdx) And (j = fncidx) Then
                GetRelOfFncname = SZ
                Exit Function
            Else
                SZ = SZ + 2 + EvenSize(Len(m_udtExtern(i).Functions(j)) + 1)
            End If
        Next
        SZ = SZ + EvenSize(Len(m_udtExtern(i).LibName) + 1)
    Next
End Function


' address of a library name in the imports section
' relative to the section's start
Private Function GetRelOfLibname(ByVal index As Long) As Long
    Dim i   As Long
    Dim j   As Long
    Dim SZ  As Long
    Dim im  As IMAGE_IMPORT_DIRECTORY
    
    For i = 0 To m_lngExternCount - 1
        SZ = SZ + 4 * (m_udtExtern(i).FunctionCount + 1) * 2
        SZ = SZ + Len(im)
    Next
    If m_lngExternCount > 0 Then SZ = SZ + Len(im)
    
    For i = 0 To m_lngExternCount - 1
        For j = 0 To m_udtExtern(i).FunctionCount - 1
            SZ = SZ + 2 + EvenSize(Len(m_udtExtern(i).Functions(j)) + 1)
        Next
        
        If i <> index Then
            SZ = SZ + EvenSize(Len(m_udtExtern(i).LibName) + 1)
        Else
            Exit For
        End If
    Next
    
    GetRelOfLibname = SZ
End Function


Private Function WritePEHeader() As Boolean
    Dim mzHdr   As IMAGE_DOS_HEADER
    Dim ntHdr   As IMAGE_NT_HEADERS
    Dim scHdr   As IMAGE_SECTION_HEADER
    Dim sdHdr   As IMAGE_SECTION_HEADER
    Dim seHdr   As IMAGE_SECTION_HEADER
    Dim srHdr   As IMAGE_SECTION_HEADER
    Dim impTbl  As IMAGE_IMPORT_DIRECTORY
    Dim lngImp  As Long
    Dim btDOS() As Byte
    
    Dim Sections As Integer
    
    If m_blnWriteDLL Then
        If GetLabelIndex("DLLMAIN") = -1 Then
            SetError "Entrypoint ""dllMain"" not found.", 0, ""
            Exit Function
        End If
        Sections = 3
    ElseIf GetLabelIndex("MAIN") = -1 Then
        SetError "Entrypoint ""Main"" not found.", 0, ""
        Exit Function
    Else
        Sections = 2
    End If
    lngImp = GetNeededImportsSize()
    Dim skipImport As Boolean
    If lngImp = 0 Then lngImp = 1: skipImport = True: Sections = Sections - 1

        
    With mzHdr
        .Magic = IMAGE_DOS_HDR16_MAGIC
        .BytesInLastPage = 144
        .Pages = 3
        .ParagraphsInHeader = 4
        .MaxAlloc = &HFFFF
        .InitialSP = &HB8
        .RelocationTableFileAddress = &H40
        .NewHeaderOffset = Len(mzHdr) + 64
    End With
    Dim more As Long
    With ntHdr
        .Signature = IMAGE_DOS_HDR32_MAGIC
        
        With .FileHeader
            .Machine = IMAGE_FILE_MACHINE_I386
            .NumberOfSections = Sections
            .SizeOfOptionalHeader = Len(ntHdr.OptionalHeader)
            
            If m_blnWriteDLL Then
                .Characteristics = IMAGE_FILE_EXECUTABLE_IMAGE Or IMAGE_FILE_DEBUG_STRIPPED Or IMAGE_FILE_LINE_NUMS_STRIPPED Or IMAGE_FILE_LOCAL_SYMS_STRIPPED Or IMAGE_FILE_32BIT_MACHINE Or IMAGE_FILE_DLL
            Else
                .Characteristics = IMAGE_FILE_RELOCS_STRIPPED Or IMAGE_FILE_LINE_NUMS_STRIPPED Or _
                               IMAGE_FILE_LOCAL_SYMS_STRIPPED Or IMAGE_FILE_EXECUTABLE_IMAGE Or _
                               IMAGE_FILE_32BIT_MACHINE Or IMAGE_FILE_DEBUG_STRIPPED Or _
                               IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP Or IMAGE_FILE_NET_RUN_FROM_SWAP
            End If
        End With
        
        With .OptionalHeader
            .Magic = IMAGE_NT_OPTIONAL_HDR32_MAGIC
            .SizeOfCode = RoundToMinSize(m_lngPECodeSize)  '
            .SizeOfInitializedData = RoundToMinSize(lngImp)
            .BaseOfCode = RoundToSectionSize(Len(ntHdr) + Len(scHdr) * Sections)
            .BaseOfData = .BaseOfCode + RoundToSectionSize(.SizeOfCode)
            If m_blnWriteDLL Then
                .AddressOfEntryPoint = m_udtLabels(GetLabelIndex("DLLMAIN")).Offset - m_lngBaseAddress
            Else
                .AddressOfEntryPoint = m_udtLabels(GetLabelIndex("MAIN")).Offset - m_lngBaseAddress
            End If
            .ImageBase = m_lngBaseAddress
            .SectionAlignment = MEM_SECTION_SIZE
            .FileAlignment = FILE_SECTION_SIZE
            .MajorOperatingSystemVersion = 4
            .MajorSubsystemVersion = 4
            ' here add size of export round to scection size
            If skipImport Then
            .SizeOfImage = .BaseOfData
            Else
            .SizeOfImage = .BaseOfData + RoundToSectionSize(lngImp)
            End If
            If m_blnWriteDLL Then
                .SizeOfImage = .SizeOfImage + RoundToSectionSize(m_lngExportSize)
               ' If False Then
                    If m_lngRelocSize > 0 Then
                       ' Sections = Sections + 1
                        '.BaseOfCode = .BaseOfCode + Len(scHdr)
                      '  .SizeOfImage = .SizeOfImage + RoundToSectionSize(m_lngRelocSize)
                    End If
               ' End If
            End If
            .DllCharacteristics = 0
            .SizeOfHeaders = GetPEHeaderSize(Sections)
            .Subsystem = m_udeSubsystem
            .SizeOfStackReserve = m_SizeOfStackReserve
            .SizeOfStackCommit = m_SizeOfStackCommit
            .SizeOfHeapReserve = m_SizeOfHeapReserve
            .SizeOfHeapCommit = m_SizeOfHeapCommit
            .NumberOfRvaAndSizes = 16
            Dim edataVirtualAddress As Long
            Dim relocVirtualAddress As Long
            
            If Not skipImport Then
                With .DataDirectory(ETableImport)
                    .VirtualAddress = ntHdr.OptionalHeader.BaseOfCode + _
                                      RoundToSectionSize(ntHdr.OptionalHeader.SizeOfCode) + _
                                      GetNeededIATSize()
                                      
                    .Size = Len(impTbl)
                End With
                
                With .DataDirectory(ETableIAT)
                    .VirtualAddress = ntHdr.OptionalHeader.BaseOfCode + _
                                      RoundToSectionSize(ntHdr.OptionalHeader.SizeOfCode)
                                      
                    .Size = GetNeededIATSize()
                    
                End With
                edataVirtualAddress = .DataDirectory(ETableIAT).VirtualAddress + RoundToSectionSize(.DataDirectory(ETableImport).Size + .DataDirectory(ETableIAT).Size)
            Else
                edataVirtualAddress = .BaseOfData
                
            End If
            If m_blnWriteDLL Then
                With .DataDirectory(ETableExport)
                    .VirtualAddress = edataVirtualAddress
                    .Size = m_lngExportSize
                    relocVirtualAddress = RoundToSectionSize(.Size) + .VirtualAddress
                End With
               'If False Then
                    With .DataDirectory(ETableRelocation)
                        .VirtualAddress = relocVirtualAddress
                        .Size = m_lngRelocSize
                    more = m_lngRelocSize
                    End With
                    .SizeOfHeaders = GetPEHeaderSize(Sections + 1)
                    
                    Sections = Sections + 1
                
               ' End If
            End If
            
        End With
        If .FileHeader.NumberOfSections < Sections Then
       .OptionalHeader.SizeOfImage = .OptionalHeader.SizeOfImage + more
        
        .FileHeader.NumberOfSections = Sections
        
        End If
    End With
    
    With scHdr
        WriteSectionName scHdr, ".code"
        .VirtSizePhysAddr = m_lngPECodeSize
        .VirtualAddress = ntHdr.OptionalHeader.BaseOfCode
        .SizeOfRawData = RoundToMinSize(m_lngPECodeSize)
        .PointerToRawData = GetPEHeaderSize(Sections)
        .Characteristics = IMAGE_SCN_CNT_CODE Or _
                           IMAGE_SCN_CNT_INITIALIZED_DATA Or _
                           IMAGE_SCN_MEM_EXECUTE Or _
                           IMAGE_SCN_MEM_READ Or _
                           IMAGE_SCN_MEM_WRITE
    End With
     If Not skipImport Then
    With sdHdr
        WriteSectionName sdHdr, ".rdata"
        .VirtSizePhysAddr = lngImp
        .VirtualAddress = ntHdr.OptionalHeader.BaseOfData
        .SizeOfRawData = RoundToMinSize(lngImp)
        .PointerToRawData = scHdr.PointerToRawData + scHdr.SizeOfRawData
        .Characteristics = IMAGE_SCN_MEM_READ Or _
                           IMAGE_SCN_MEM_WRITE
    End With
    End If
    If m_blnWriteDLL Then
        With seHdr
            WriteSectionName seHdr, ".edata"
            .VirtSizePhysAddr = m_lngExportSize
            .VirtualAddress = edataVirtualAddress
            .SizeOfRawData = RoundToMinSize(m_lngExportSize)
            If Not skipImport Then
            .PointerToRawData = sdHdr.PointerToRawData + sdHdr.SizeOfRawData
            Else
            
            .PointerToRawData = scHdr.PointerToRawData + scHdr.SizeOfRawData
            End If
            .Characteristics = &H40000040 ' INITIALIZED_DATA | READ
        End With
        If True Then
            With srHdr
                WriteSectionName srHdr, ".reloc"
                
                ' ? FIX: VirtualSize must represent the exact, non-aligned data size in memory.
                ' Using the precise m_lngRelocSize ensures Windows validation passes during ASLR relocation.
                .VirtSizePhysAddr = m_lngRelocSize
                
                .VirtualAddress = relocVirtualAddress
                
                ' ? FIX: Ensure SizeOfRawData is perfectly aligned to the disk sector size (FILE_SECTION_SIZE = 512).
                ' If there are no relocations, set it to 0 so the loader knows it's empty.
                If m_lngRelocSize > 0 Then
                    .SizeOfRawData = RoundToMinSize(m_lngRelocSize)
                Else
                    .SizeOfRawData = 0
                End If
                
                .PointerToRawData = seHdr.PointerToRawData + seHdr.SizeOfRawData
                .Characteristics = &H42000040 ' IMAGE_SCN_CNT_INITIALIZED_DATA | IMAGE_SCN_MEM_DISCARDABLE | IMAGE_SCN_MEM_READ
            End With
        Else
           With srHdr
                WriteSectionName srHdr, ".reloc"
                .VirtSizePhysAddr = m_lngRelocSize
                .VirtualAddress = relocVirtualAddress
                .SizeOfRawData = RoundToMinSize(m_lngRelocSize)
                .PointerToRawData = seHdr.PointerToRawData + seHdr.SizeOfRawData
                .Characteristics = &H42000040 ' INITIALIZED_DATA | DISCARDABLE | READ
            End With
        End If
    End If
    btDOS = HexToByte(DOS_CODE_RELOCATIONS)
    
    OutputMem VarPtr(mzHdr), Len(mzHdr) '
    OutputMem VarPtr(btDOS(0)), UBound(btDOS) + 1
    OutputMem VarPtr(ntHdr), Len(ntHdr)
    OutputMem VarPtr(scHdr), Len(scHdr)
    If Not skipImport Then
        OutputMem VarPtr(sdHdr), Len(sdHdr)
    End If
    If m_blnWriteDLL Then
        OutputMem VarPtr(seHdr), Len(seHdr)
        'If False Then
            OutputMem VarPtr(srHdr), Len(srHdr)
        'End If
    End If
    OutputJumpTo RoundToMinSize(OutputPosition)
    
    WritePEHeader = True
End Function


Private Sub WriteSectionName(sc As IMAGE_SECTION_HEADER, ByVal strName As String)
    Dim i   As Long
    
    For i = 1 To Len(strName)
        sc.SectionName(i - 1) = Asc(Mid$(strName, i, 1))
    Next
End Sub


Private Function GetPEHeaderSize(ByVal Sections As Integer) As Long
    Dim mzHdr   As IMAGE_DOS_HEADER
    Dim ntHdr   As IMAGE_NT_HEADERS
    Dim scHdr   As IMAGE_SECTION_HEADER

    
    GetPEHeaderSize = RoundToMinSize(Len(mzHdr) + 64 + Len(ntHdr) + Len(scHdr) * Sections)
End Function


' every instruction has prefixes (opt), an opcode and arguments (opt).
' write them to the output
Private Function AssembleInstructions() As Boolean
    Dim i       As Long
    Dim lngSz   As Long
    
    For i = 0 To m_lngInstrCount - 1
        lngSz = m_lngOutPos
        
        If m_udtInstrs(i).Mnemonic = "ALIGN" Then
            Dim idxNOP As Long
            For idxNOP = 1 To m_udtInstrs(i).Size
                OutputByte &H90 ' Write 0x90 (NOP) opcode
            Next idxNOP
        ElseIf m_udtInstrs(i).Data.Size <> BitsUnknown Then
            If Not RawDataOut(m_udtInstrs(i).Data) Then Exit Function
        Else
            InstructionOutPrefixes m_udtInstrs(i)
            InstructionOutOpCode m_udtInstrs(i)
            If Not InstructionOutArgs(m_udtInstrs(i)) Then Exit Function
        End If
        
        lngSz = m_lngOutPos - lngSz
        If lngSz <> m_udtInstrs(i).Size Then Err.Raise 123, , "invalid size after output"
    Next
    
    AssembleInstructions = True
End Function


' db, dw, dd strings
Private Function RawDataOut(Data As RawData) As Boolean
    Dim i   As Long
    
    For i = 0 To Data.ValueCount - 1
        OutputBytes Data.Values(i), Data.Size
    Next
    
    RawDataOut = True
End Function


' write arguments to output (modR/M, SIB, Pointers, ...)
' SIB is only written when a pointer has more then 2 registers
' or multiples of a register, or when ESP is used.
Private Function InstructionOutArgs(udtInstr As ASMInstruction) As Boolean
    Dim i               As Long
    Dim j               As Long
    
    Dim udtModRM        As ModRM
    Dim blnModRM        As Boolean
    Dim blnSIBNeeded    As Boolean
    
    Dim lngImmVal()     As Long
    Dim udeImmSize()    As ParamSize
    Dim lngImmValCnt    As Long
    Dim blnImmVal       As Boolean
    
    Dim lngDisplacement As Long
    Dim udeDispSize     As ParamSize
    Dim blnDisplacement As Boolean
    
    Dim lngSIBPtrIdx    As Long
    
    Dim blnGotXMMReg    As Boolean
    Dim lngMMVal        As Long
    
    With Instructions(udtInstr.OpCodeIndex)
        blnModRM = .ModRM
        If .RegOpExt > -1 Then
            udtModRM.reg = .RegOpExt
            blnGotXMMReg = True
        End If
    End With
    
    For i = 0 To udtInstr.ArgCount - 1
        With Instructions(udtInstr.OpCodeIndex).Parameters(i)
            If Not .Forced Then
            
                If (.PType = ParamImm) Or (.PType = ParamRel) Then
                    ' an instruction can have multiple immediates
                    ReDim Preserve lngImmVal(lngImmValCnt) As Long
                    ReDim Preserve udeImmSize(lngImmValCnt) As ParamSize
                    
                    If udtInstr.Args(i).TType = ParamImm Then
                        lngImmVal(lngImmValCnt) = udtInstr.Args(i).Value
                    ElseIf udtInstr.Args(i).TType = ParamRel Then
                        lngImmVal(lngImmValCnt) = m_udtLabels(udtInstr.Args(i).SymbolIndex).Offset
                    End If
                    
                    If .PType = ParamRel Then
                        ' if a relative value is needed, make the immediate
                        ' relative to the end of the current instruction
                        lngImmVal(lngImmValCnt) = lngImmVal(lngImmValCnt) - (udtInstr.Offset + udtInstr.Size)
                    End If
                    
                    If (SizesForInt(lngImmVal(lngImmValCnt)) And .Size) = 0 Then
                        SetError "Relative value too big for instruction", udtInstr.Line, udtInstr.Section
                        Exit Function
                    End If
                    
                        If (udtInstr.Args(i).TType And ParamExt) Then
                            If m_blnWriteDLL Or m_blnWritePE Then
                                AddRelocation m_lngOutPos
                            End If
                        End If
                    
                    
                    udeImmSize(lngImmValCnt) = .Size
                    lngImmValCnt = lngImmValCnt + 1
                    blnImmVal = True
                    
                ElseIf .PType = ParamReg Then
                    ' register must be put to ModR/M, else it would be "forced"
                    udtModRM.reg = ModRMRegNum(udtInstr.Args(i).Register)
                
                ElseIf .PType = ParamMM Then
                    Select Case udtInstr.Args(i).MMRegister
                        Case MM0, XMM0: lngMMVal = 0
                        Case mm1, XMM1: lngMMVal = 1
                        Case mm2, XMM2: lngMMVal = 2
                        Case MM3, XMM3: lngMMVal = 3
                        Case MM4, XMM4: lngMMVal = 4
                        Case MM5, XMM5: lngMMVal = 5
                        Case MM6, XMM6: lngMMVal = 6
                        Case MM7, XMM7: lngMMVal = 7
                    End Select
                    
                    If blnGotXMMReg Then
                        udtModRM.Mod = 3
                        udtModRM.rm = lngMMVal
                    Else
                        udtModRM.reg = lngMMVal
                        blnGotXMMReg = True
                    End If
                
                ElseIf (.PType = (ParamMem Or ParamReg)) Or _
                       (.PType = (ParamMem Or ParamMM)) Or _
                       (.PType = ParamMem) Then
                       
                    If (.PType = ParamMem) And (Not blnModRM) Then
                        ' !#! cann there also be more than one per instruction? !#!
                        lngDisplacement = udtInstr.Args(i).Pointer.Ptr.Displacement
                        udeDispSize = Bits32
                        blnDisplacement = True
                    Else
                        If udtInstr.Args(i).TType = ParamMem Then
                            Select Case udtInstr.Args(i).Pointer.RegisterCount
                                Case 0:
                                    ' no reigster in the pointer, only displacement possible
                                    If udtInstr.Args(i).Pointer.HasDisplacement Then
                                        udtModRM.disp = udtInstr.Args(i).Pointer.Ptr.Displacement
                                        udtModRM.DispSize = Bits32
                                        udtModRM.Mod = 0
                                        udtModRM.rm = 5
                                    End If
                                    
                                Case 1:
                                    ' 1 Register in the pointer, can be encoded with ModR/M
                                    ' if its not ESP or a multiple (reg*2/3/4/5/8/9)
                                    If udtInstr.Args(i).Pointer.HasDisplacement Then
                                        udtModRM.disp = udtInstr.Args(i).Pointer.Ptr.Displacement
                                        udtModRM.DispSize = udtInstr.Args(i).Pointer.Ptr.DispSize
                                        Select Case udtModRM.DispSize
                                            Case Bits8:     udtModRM.Mod = 1
                                            Case Else:      udtModRM.Mod = 2
                                        End Select
                                    Else
                                        udtModRM.Mod = 0
                                    End If
                                    
                                    For j = 0 To REG_COUNT - 1
                                        If udtInstr.Args(i).Pointer.Ptr.Registers(j) Then
                                            If udtInstr.Args(i).Pointer.Ptr.Registers(j) > 1 Then
                                                blnSIBNeeded = True
                                                lngSIBPtrIdx = i
                                            Else
                                                Select Case IdxToReg(j)
                                                    Case RegEAX: udtModRM.rm = 0
                                                    Case RegECX: udtModRM.rm = 1
                                                    Case RegEDX: udtModRM.rm = 2
                                                    Case RegEBX: udtModRM.rm = 3
                                                    Case RegEBP: udtModRM.rm = 5
                                                    Case RegESI: udtModRM.rm = 6
                                                    Case RegEDI: udtModRM.rm = 7
                                                    Case RegESP: blnSIBNeeded = True
                                                                 lngSIBPtrIdx = i
                                                End Select
                                            End If
                                            
                                            Exit For
                                        End If
                                    Next
                                    
                                Case 2:
                                    ' 2 registers, SIB needed
                                    If udtInstr.Args(i).Pointer.HasDisplacement Then
                                        udtModRM.disp = udtInstr.Args(i).Pointer.Ptr.Displacement
                                        udtModRM.DispSize = udtInstr.Args(i).Pointer.Ptr.DispSize
                                        Select Case udtModRM.DispSize
                                            Case Bits8:     udtModRM.Mod = 1
                                            Case Else:      udtModRM.Mod = 2
                                        End Select
                                    Else
                                        udtModRM.Mod = 0
                                    End If
                                    
                                    blnSIBNeeded = True
                                    lngSIBPtrIdx = i
                                    
                            End Select
                            
                        ElseIf udtInstr.Args(i).TType = ParamReg Then
                            ' encode second register in ModR/M
                            udtModRM.Mod = 3
                            udtModRM.rm = ModRMRegNum(udtInstr.Args(i).Register)
                        
                        ElseIf udtInstr.Args(i).TType = (ParamMem Or ParamExt) Then
                            lngDisplacement = udtInstr.Args(i).Pointer.Ptr.Displacement
                            udeDispSize = Bits32
                            udtModRM.rm = 5
                            blnDisplacement = True
                            'If m_blnWriteDLL Or m_blnWritePE Then
                            '    AddRelocation m_lngOutPos
                            'End If
                            'blnDisplacement = False
                            
                        ElseIf udtInstr.Args(i).TType = ParamMM Then
                            Select Case udtInstr.Args(i).MMRegister
                                Case MM0, XMM0: udtModRM.rm = 0
                                Case mm1, XMM1: udtModRM.rm = 1
                                Case mm2, XMM2: udtModRM.rm = 2
                                Case MM3, XMM3: udtModRM.rm = 3
                                Case MM4, XMM4: udtModRM.rm = 4
                                Case MM5, XMM5: udtModRM.rm = 5
                                Case MM6, XMM6: udtModRM.rm = 6
                                Case MM7, XMM7: udtModRM.rm = 7
                            End Select
                            udtModRM.Mod = 3
                        
                        End If
                    End If
                    
                End If
                
            End If
        End With
    Next
    
    If blnSIBNeeded Then
        udtModRM.rm = 4
        If Not WriteSIB(udtInstr, udtModRM, udtInstr.Args(lngSIBPtrIdx).Pointer.Ptr) Then
            Exit Function
        End If
    Else
        If blnModRM Then WriteModRM udtModRM
    End If
    
'    If blnDisplacement Then
'        If m_blnWritePE Or m_blnWriteDLL Then
'            AddRelocation m_lngOutPos
'       End If
'        OutputBytes lngDisplacement, udeDispSize
'    End If
    
'    If blnDisplacement Then
'        If m_blnWritePE Or m_blnWriteDLL Then
'            ' ? FIX: Only register a relocation if the displacement belongs to an External API (ParamExt)
'            ' or a global symbol/label. Do NOT record frame/stack offsets like [ebp + 8] or [esp + 4].
'            Dim blnIsGlobalSymbol As Boolean
'            blnIsGlobalSymbol = False
'
'            ' Scan the arguments to see if this instruction actually references an external or global symbol
'            For i = 0 To udtInstr.ArgCount - 1
'                If (udtInstr.Args(i).TType And ParamExt) Or (udtInstr.Args(i).TType = ParamRel) Then
'                    blnIsGlobalSymbol = True
'                    Exit For
'                End If
'            Next i
'
'            ' Only add to the .reloc table if it's a true global memory address change
'            If blnIsGlobalSymbol Then
'                AddRelocation m_lngOutPos
'            End If
'        End If
        
        ' Write the actual displacement bytes (like 8 or 0xC) safely to the file array
'        OutputBytes lngDisplacement, udeDispSize
'    End If

    If blnDisplacement Then
        If m_blnWritePE Or m_blnWriteDLL Then
            Dim blnIsGlobalReloc As Boolean
            blnIsGlobalReloc = False
            
            ' Check if this displacement is an external API or an absolute global variable
            For i = 0 To udtInstr.ArgCount - 1
                If (udtInstr.Args(i).TType And ParamExt) Then
                    blnIsGlobalReloc = True
                    Exit For
                ElseIf (udtInstr.Args(i).TType = ParamMem) Then
                    ' If it has no registers (RegisterCount = 0), it's an absolute global data pointer!
                    If udtInstr.Args(i).Pointer.RegisterCount = 0 Then
                        blnIsGlobalReloc = True
                        Exit For
                    End If
                End If
            Next i
            
            ' Only register in .reloc table if it's a true global address
            If blnIsGlobalReloc Then
                Debug.Print m_lngOutPos
                AddRelocation m_lngOutPos
            End If
        End If
        
        OutputBytes lngDisplacement, udeDispSize
    End If


    If blnImmVal Then
        For i = 0 To lngImmValCnt - 1
            OutputBytes lngImmVal(i), udeImmSize(i)
        Next
    End If
    
    If Instructions(udtInstr.OpCodeIndex).Now3DByte > -1 Then
        OutputByte Instructions(udtInstr.OpCodeIndex).Now3DByte
    End If
    
    InstructionOutArgs = True
End Function


Private Function WriteSIB( _
    udtInstr As ASMInstruction, _
    rm As ModRM, _
    Ptr As Pointer _
) As Boolean

    Dim udtSIB          As SIB
    Dim udeReg(1)       As ASMRegisters
    Dim lngRegCnt(1)    As Long
    Dim lngScale        As Long
    Dim lngBase         As Long
    Dim i               As Long
    Dim j               As Long
    
    ' find the used registers
    For i = 0 To REG_COUNT - 1
        If Ptr.Registers(i) Then
            udeReg(j) = IdxToReg(i)
            lngRegCnt(j) = Ptr.Registers(i)
            j = j + 1
        End If
    Next
    
    ' determine the scale register (can have a multiple)
    If lngRegCnt(0) >= 1 And lngRegCnt(1) = 1 Then
        lngScale = 0
    ElseIf lngRegCnt(1) >= 1 And lngRegCnt(0) = 1 Then
        lngScale = 1
    End If
    ' the base register is the other one ;)
    lngBase = 1 - lngScale
    
    If (Ptr.UsedRegisters = 1) And (lngRegCnt(0) = 1) Then
        ' olny one register used which isn't a multiple
        udtSIB.sscale = 0
        udtSIB.index = 4
        
        Select Case udeReg(0)
            Case RegEAX:    udtSIB.Base = 0
            Case RegECX:    udtSIB.Base = 1
            Case RegEDX:    udtSIB.Base = 2
            Case RegEBX:    udtSIB.Base = 3
            Case RegESP:    udtSIB.Base = 4  ' <= all the others encodable in ModR/M
            Case RegEBP:    udtSIB.Base = 5
            Case RegESI:    udtSIB.Base = 6
            Case RegEDI:    udtSIB.Base = 7
        End Select
        
    Else
    
        If (Ptr.UsedRegisters = 1) And (lngRegCnt(0) > 1) Then
            ' one register which is a multiple
            If lngRegCnt(0) = 2 Or _
               lngRegCnt(0) = 3 Or _
               lngRegCnt(0) = 5 Or _
               lngRegCnt(0) = 9 Then
               
                If lngRegCnt(0) = 2 Then
                    udtSIB.sscale = 0
                Else
                    udtSIB.sscale = GetFirstSetBitIdx(lngRegCnt(0) - 1)
                End If
                
                Select Case udeReg(0)
                    Case RegEAX:    udtSIB.index = 0:   udtSIB.Base = 0
                    Case RegECX:    udtSIB.index = 1:   udtSIB.Base = 1
                    Case RegEDX:    udtSIB.index = 2:   udtSIB.Base = 2
                    Case RegEBX:    udtSIB.index = 3:   udtSIB.Base = 3
                    Case RegESI:    udtSIB.index = 6:   udtSIB.Base = 6
                    Case RegEDI:    udtSIB.index = 7:   udtSIB.Base = 7
                    Case RegEBP:    udtSIB.index = 5:   udtSIB.Base = 5
                                    rm.Mod = 1
                    Case Else:
                        SetError "invalid multiple of a register in SIB", udtInstr.Line, udtInstr.Section
                        Exit Function
                End Select
                
            ElseIf lngRegCnt(0) = 4 Or lngRegCnt(0) = 8 Then
                ' if Mod of ModR/M byte would be > 0 here
                ' EBP+sbyte/sdword would be encoded, too
                rm.Mod = 0
                udtSIB.Base = 5
                udtSIB.sscale = GetFirstSetBitIdx(lngRegCnt(0))
                
                Select Case udeReg(0)
                    Case RegEAX:    udtSIB.index = 0
                    Case RegECX:    udtSIB.index = 1
                    Case RegEDX:    udtSIB.index = 2
                    Case RegEBX:    udtSIB.index = 3
                    Case RegEBP:    udtSIB.index = 5
                    Case RegESI:    udtSIB.index = 6
                    Case RegEDI:    udtSIB.index = 7
                    Case Else:
                        SetError "invalid multiples of register in SIB", udtInstr.Line, udtInstr.Section
                        Exit Function
                End Select
            
            Else
                SetError "invalid multiples of register in SIB", udtInstr.Line, udtInstr.Section
                Exit Function
                
            End If
            
        ElseIf Ptr.UsedRegisters = 2 Then
            ' 2 register in pointer
            
            Select Case lngRegCnt(lngScale)
                Case 1, 2, 4, 8:
                Case Else:
                    SetError "Possible multiples of scale register: 1, 2, 4, 8", udtInstr.Line, udtInstr.Section
                    Exit Function
            End Select
            
            If lngRegCnt(lngBase) <> 1 Then
                SetError "Base register mustn't have a multiple", udtInstr.Line, udtInstr.Section
                Exit Function
            End If
            
            ' ESP can only be encoded in the base, so it can't be the scale.
            ' Same thing for EBP, but the other way.
            If (lngRegCnt(lngScale) = 1) And (lngRegCnt(lngBase)) = 1 Then
                If (udeReg(lngScale) = RegESP) Or (udeReg(lngBase) = RegEBP) Then
                    lngScale = lngBase
                    lngBase = 1 - lngScale
                End If
            End If
            
            udtSIB.sscale = GetFirstSetBitIdx(lngRegCnt(lngScale))
            
            Select Case udeReg(lngScale)
                Case RegEAX:    udtSIB.index = 0
                Case RegECX:    udtSIB.index = 1
                Case RegEDX:    udtSIB.index = 2
                Case RegEBX:    udtSIB.index = 3
                Case RegEBP:    udtSIB.index = 5
                Case RegESI:    udtSIB.index = 6
                Case RegEDI:    udtSIB.index = 7
                Case Else:
                    SetError "invalid scale register", udtInstr.Line, udtInstr.Section
                    Exit Function
            End Select
            
            Select Case udeReg(lngBase)
                Case RegEAX:    udtSIB.Base = 0
                Case RegECX:    udtSIB.Base = 1
                Case RegEDX:    udtSIB.Base = 2
                Case RegEBX:    udtSIB.Base = 3
                Case RegESP:    udtSIB.Base = 4
                Case RegEBP:    udtSIB.Base = 5
                Case RegESI:    udtSIB.Base = 6
                Case RegEDI:    udtSIB.Base = 7
                Case Else:
                    SetError "invalid base register", udtInstr.Line, udtInstr.Section
                    Exit Function
            End Select
        End If
    End If
    
    OutputByte (rm.Mod * &H40) Or (rm.reg * &H8) Or rm.rm
    OutputByte (udtSIB.sscale * &H40) Or (udtSIB.index * &H8) Or udtSIB.Base
    If rm.DispSize <> BitsUnknown Then OutputBytes rm.disp, rm.DispSize
    
    WriteSIB = True
End Function


Private Function ModRMRegNum(ByVal reg As ASMRegisters) As Long
    Select Case reg
        Case RegAL, RegAX, RegEAX:  ModRMRegNum = 0
        Case RegCL, RegCX, RegECX:  ModRMRegNum = 1
        Case RegDL, RegDX, RegEDX:  ModRMRegNum = 2
        Case RegBL, RegBX, RegEBX:  ModRMRegNum = 3
        Case RegAH, RegSP, RegESP:  ModRMRegNum = 4
        Case RegCH, RegBP, RegEBP:  ModRMRegNum = 5
        Case RegDH, RegSI, RegESI:  ModRMRegNum = 6
        Case RegBH, RegDI, RegEDI:  ModRMRegNum = 7
    End Select
End Function


Private Sub WriteModRM(rm As ModRM)
    OutputByte (rm.Mod * &H40) Or (rm.reg * &H8) Or rm.rm
    If rm.DispSize <> BitsUnknown Then OutputBytes rm.disp, rm.DispSize
End Sub


Private Sub InstructionOutOpCode(udtInstr As ASMInstruction)
    Dim i As Long
    
    With Instructions(udtInstr.OpCodeIndex)
        For i = 0 To .OpCodeLen - 1
            OutputByte .Opcode(i)
        Next
    End With
End Sub


Private Sub InstructionOutPrefixes(udtInstr As ASMInstruction)
    With Instructions(udtInstr.OpCodeIndex)
        If (.Prefixes And PrefixFlgOperandSizeOverride) Then _
            OutputByte PREFIX_OPERAND_SIZE_OVERRIDE
        If (.Prefixes And PrefixFlgAddressSizeOverride) Then _
            OutputByte PREFIX_ADDRESS_SIZE_OVERRIDE
        If (.Prefixes And PrefixFlgBranchNotTaken) Then _
            OutputByte PREFIX_BRANCH_NOT_TAKEN
        If (.Prefixes And PrefixFlgBranchTaken) Then _
            OutputByte PREFIX_BRANCH_TAKEN
    End With
    
    With udtInstr
        If (.Flags And PrefixFlgLock) Then _
            OutputByte PREFIX_LOCK
        If (.Flags And PrefixFlgRep) Then _
            OutputByte PREFIX_REP
        If (.Flags And PrefixFlgRepne) Then _
            OutputByte PREFIX_REPNE
    End With
    
    Select Case udtInstr.Segment
        Case SegCS: OutputByte PREFIX_SEGMENT_CS
        Case SegDS: OutputByte PREFIX_SEGMENT_DS
        Case SegES: OutputByte PREFIX_SEGMENT_ES
        Case SegFS: OutputByte PREFIX_SEGMENT_FS
        Case SegGS: OutputByte PREFIX_SEGMENT_GS
        Case SegSS: OutputByte PREFIX_SEGMENT_SS
    End Select
End Sub


Private Sub OutputMem(ByVal Ptr As Long, ByVal bytes As Long)
    If m_lngOutPos + bytes > m_lngOutSize Then
        Err.Raise 123456, , "not enough space in output array"
    End If
    
    CopyMemory m_btOutput(m_lngOutPos), ByVal Ptr, bytes
    m_lngOutPos = m_lngOutPos + bytes
End Sub


Private Sub OutputJumpTo(ByVal lngVal As Long)
    If (lngVal >= m_lngOutSize) Or (lngVal < 0) Then
        Err.Raise 123456, , "new position out of bounds"
    End If
    
    m_lngOutPos = lngVal
End Sub


Private Property Get OutputPosition() As Long
    OutputPosition = m_lngOutPos
End Property


Private Sub OutputBytes(ByVal Value As Long, ByVal Size As ParamSize)
    Select Case Size
        Case Bits8:     OutputByte Value
        Case Bits16:    OutputInteger Value
        Case Bits32:    OutputLong Value
        Case Else:      Err.Raise 123456, , "invalid size"
    End Select
End Sub


Private Sub OutputByte(ByVal Value As Long)
    If Value < 0 Then
        m_btOutput(m_lngOutPos) = CByte(Value + 256)    ' Signed Byte
    Else
        m_btOutput(m_lngOutPos) = CByte(Value)
    End If
    
    m_lngOutPos = m_lngOutPos + 1
End Sub


Private Sub OutputInteger(ByVal Value As Long)
    CopyMemory m_btOutput(m_lngOutPos), Value, 2
    m_lngOutPos = m_lngOutPos + 2
End Sub


Private Sub OutputLong(ByVal Value As Long)
    CopyMemory m_btOutput(m_lngOutPos), Value, 4
    m_lngOutPos = m_lngOutPos + 4
End Sub


Private Function ParsePointers() As Boolean
    Dim i   As Long
    Dim j   As Long
    
    For i = 0 To m_lngInstrCount - 1
        For j = 0 To m_udtInstrs(i).ArgCount - 1
            With m_udtInstrs(i).Args(j)
                If .TType = ParamMem Then
                    m_lngCurToken = .Pointer.TokenIndex
                    If Not ParsePointer(.Pointer.Ptr) Then Exit Function
                End If
            End With
        Next
    Next
    
    ParsePointers = True
End Function


Private Function ParsePointer(Ptr As Pointer) As Boolean
    Dim lngSgn      As Long
    Dim lngTms      As Long
    Dim lngVal      As Long
    Dim i           As Long
    Dim lngReg      As Long
    Dim blnReg      As Boolean
    
    If Not Match(TokenBracketLeft) Then
        SetError """["" expected", token.Line, token.Section
        Exit Function
    End If
    
    If Match(TokenOpAdd) Then
        lngSgn = 1
    ElseIf Match(TokenOpSub) Then
        lngSgn = -1
    Else
        lngSgn = 1
    End If
    
    Do
        lngTms = 1
        blnReg = False
        
        Do
            Select Case token.TType
                Case TokenRegister:
                    blnReg = True
                    lngReg = RegToIdx(RegStrToReg(token.content))
                    Match TokenRegister
                Case TokenValue:
                    lngTms = lngTms * token.Value
                    Match TokenValue
                Case tokensymbol:
                    lngTms = lngTms * m_udtLabels(GetLabelIndex(token.content)).Offset
                    Match tokensymbol
            End Select
        Loop While Match(TokenOpMul)
        
        If blnReg Then
            Ptr.Registers(lngReg) = Ptr.Registers(lngReg) + lngSgn * lngTms
        Else
            Ptr.Displacement = Ptr.Displacement + lngSgn * lngTms
        End If
        
        If Match(TokenOpAdd) Then
            lngSgn = 1
        ElseIf Match(TokenOpSub) Then
            lngSgn = -1
        Else
            Exit Do
        End If
    Loop
    
    If Not Match(TokenBracketRight) Then
        SetError """]"" expected", token.Line, token.Section
    Else
        ParsePointer = True
    End If
End Function


Private Function GetInstructionSizes(Optional ByVal Sections As Integer = 2) As Boolean
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim Size        As Long
    Dim lngImpSz    As Long
    Dim blnFoundI   As Boolean

    For i = 0 To m_lngInstrCount - 1                ' all instructions in the source
        If m_udtInstrs(i).Data.Size <> BitsUnknown Then
            With m_udtInstrs(i)
                .Size = .Data.ValueCount * .Data.Size \ 8
            End With
        Else
            blnFoundI = False
            
            For j = 0 To InstructionCount - 1       ' all known instructions
                If StrComp(m_udtInstrs(i).Mnemonic, Instructions(j).Mnemonic, vbTextCompare) = 0 Then
                    blnFoundI = True
                    
                    If CompareInstrs(m_udtInstrs(i), Instructions(j)) Then
                        ' OpCode length + all used prefixes, ModR/M Byte
                        Size = Instructions(j).OpCodeLen + _
                               BitCount(Instructions(j).Prefixes Or m_udtInstrs(i).Flags) + _
                               Abs(m_udtInstrs(i).Segment <> SegUnknown) + _
                               Abs(Instructions(j).ModRM) + _
                               Abs(Instructions(j).Now3DByte > -1)
                        
                        ' immediates, displacement, SIB byte
                        If Instructions(j).ParamCount > 0 Then
                            For k = 0 To Instructions(j).ParamCount - 1
                                With Instructions(j).Parameters(k)
                                    If Not .Forced Then
                                    
                                        Select Case .PType
                                        
                                            Case ParamImm, ParamRel:
                                                Size = Size + .Size \ 8 ' Imm
                                            
                                            Case ParamMem Or ParamReg, ParamMem:
                                                If Not Instructions(j).ModRM And .PType = ParamMem Then
                                                    Size = Size + 4     ' Imm
                                                Else
                                                    If m_udtInstrs(i).Args(k).Pointer.HasDisplacement Then
                                                        Size = Size + m_udtInstrs(i).Args(k).Pointer.DispSize \ 8
                                                    End If
                                                    
                                                    If m_udtInstrs(i).Args(k).Pointer.RegisterCount = 2 Or _
                                                       m_udtInstrs(i).Args(k).Pointer.RegisterMultiples Then
                                                        Size = Size + 1 ' SIB
                                                    End If
                                                End If
                                                
                                        End Select
                                        
                                    End If
                                End With
                            Next
                        End If
                        
                        m_udtInstrs(i).OpCodeIndex = j
                        m_udtInstrs(i).Size = Size
                        Exit For
                    End If
                End If
            Next
            
            If j = InstructionCount Then
                If blnFoundI Then
                    SetError "invalid arguments", m_udtInstrs(i).Line, m_udtInstrs(i).Section
                Else
                    SetError "unknown instruction: " & m_udtInstrs(i).Mnemonic, m_udtInstrs(i).Line, m_udtInstrs(i).Section
                End If
                
                Exit Function
            End If
        End If
    Next
    ' =============================================================
    '  NEW: Scan instructions to find the exact number of relocations
    ' =============================================================
            ' UPDATED FIX: Scan for External APIs OR Absolute Global Memory Displacements
            ' Ignore stack frame offsets like [ebp + 8] (which have RegisterCount > 0)
            Dim idxInstr As Long
            Dim idxArg As Long
            Dim lngEstimatedRelocs As Long
            
            lngEstimatedRelocs = 0
            For idxInstr = 0 To m_lngInstrCount - 1
                For idxArg = 0 To m_udtInstrs(idxInstr).ArgCount - 1
                    
                    ' Case A: External Windows API Function Call
                    If (m_udtInstrs(idxInstr).Args(idxArg).TType And ParamExt) Then
                        lngEstimatedRelocs = lngEstimatedRelocs + 1
                        
                    ' Case B: Absolute Global Memory Access (like [data1] where RegisterCount = 0)
                    ElseIf (m_udtInstrs(idxInstr).Args(idxArg).TType = ParamMem) Then
                        If m_udtInstrs(idxInstr).Args(idxArg).Pointer.RegisterCount = 0 And _
                           m_udtInstrs(idxInstr).Args(idxArg).Pointer.HasDisplacement Then
                            lngEstimatedRelocs = lngEstimatedRelocs + 1
                        End If
                    End If
                    
                Next idxArg
            Next idxInstr

    
'    Dim idxInstr As Long
'    Dim idxArg As Long
'    Dim lngEstimatedRelocs As Long
'
'    lngEstimatedRelocs = 0
'            lngEstimatedRelocs = 0
'            For idxInstr = 0 To m_lngInstrCount - 1
'                For idxArg = 0 To m_udtInstrs(idxInstr).ArgCount - 1
'                    ' A true relocation is required ONLY if the instruction targets
'                    ' an external API function (ParamExt) or a global symbol/label.
'                    If (m_udtInstrs(idxInstr).Args(idxArg).TType And ParamExt) Or _
'                       (m_udtInstrs(idxInstr).Args(idxArg).TType = ParamRel) Then
'                        lngEstimatedRelocs = lngEstimatedRelocs + 1
'                    End If
'                Next idxArg
'            Next idxInstr
'
    If lngEstimatedRelocs = 0 And m_blnWriteDLL Then
        Sections = Sections - 1
    End If
                
    FillInOffsets Sections
    If m_blnWritePE Or m_blnWriteDLL Then
        m_lngPECodeSize = m_lngOutSize
        
        lngImpSz = GetNeededImportsSize()
        If lngImpSz = 0 Then
            m_lngOutSize = RoundToMinSize(GetPEHeaderSize(Sections)) + _
                           RoundToMinSize(m_lngPECodeSize) + _
                           RoundToMinSize(1)
        Else
            m_lngOutSize = RoundToMinSize(GetPEHeaderSize(Sections)) + _
                           RoundToMinSize(m_lngPECodeSize) + _
                           RoundToMinSize(lngImpSz)
        End If
        
        
            If m_blnWriteDLL Then
                FillInIAT RoundToSectionSize(GetPEHeaderSize(Sections + 1)) + _
                      RoundToSectionSize(m_lngPECodeSize)
                      
                ' 1. Calculate Export RVA and build the Exports section
                Dim lngExportRVA As Long
                lngExportRVA = RoundToSectionSize(GetPEHeaderSize(Sections)) + _
                               RoundToSectionSize(m_lngPECodeSize)
                If lngImpSz > 0 Then lngExportRVA = lngExportRVA + _
                               RoundToSectionSize(lngImpSz)
                               
                m_btExportData = BuildExportSection(lngExportRVA, m_lngExportSize)
                m_lngOutSize = RoundToSectionSize(m_lngOutSize) + m_lngExportSize
                
                ' REMOVED BuildRelocationSection FROM HERE
                ' Because m_lngRelocCount is still 0 at this stage.
                ' It will be generated later in Assemble() after instructions are assembled.
                ' Calculate exact size for the .reloc section based on the count
                ' Each relocation item takes 2 bytes. If there are items, we add 8 bytes
                ' for the Base Relocation Table Header (Page RVA + Block Size).
                If lngEstimatedRelocs > 0 Then
                    Dim lngTotalBytes As Long
                    lngTotalBytes = 8 + (lngEstimatedRelocs * 2)
                    
                    ' Ensure 4-byte DWORD padding alignment inside the block
                    If lngTotalBytes Mod 4 <> 0 Then lngTotalBytes = lngTotalBytes + 2
                    
                    m_lngRelocSize = lngTotalBytes
                    m_lngOutSize = RoundToSectionSize(m_lngOutSize) + m_lngRelocSize

                Else
                    Sections = Sections - 1
                    ' If no relocations are found, allocate the 8-byte dummy block
                    m_lngRelocSize = 8
                End If
                
                ' Add the precise relocation size to the final output file size
                                ' =============================================================
                
            Else
                FillInIAT RoundToSectionSize(GetPEHeaderSize(Sections)) + _
                          RoundToSectionSize(m_lngPECodeSize)
            End If
    End If
    If m_lngOutSize = 0 Then Exit Function
    ReDim m_btOutput(m_lngOutSize - 1) As Byte

    GetInstructionSizes = True
End Function
Private Function BuildRelocationSection(ByVal Sections As Integer, ByVal BaseRVA As Long, ByRef OutSize As Long) As Byte()
    Dim lngPageRVA As Long, lngPageStartPos As Long
    Dim lngWord As Integer
    
    If m_lngRelocCount = 0 Then
        Dim btDummy(7) As Byte
        lngPageRVA = BaseRVA
        lngWord = 8
        CopyMemory btDummy(0), lngPageRVA, 4
        CopyMemory btDummy(4), lngWord, 4
        OutSize = 8
        BuildRelocationSection = btDummy
        Exit Function
    End If
    
    Dim btOut()     As Byte
    Dim lngPos      As Long
    ReDim btOut(m_lngRelocCount * 2 + 100) As Byte
    Dim i As Long, j As Long
    ', headersize As Long
    
    'headersize = RoundToSectionSize(GetPEHeaderSize(Sections))
    
    ' Perfect Virtual Alignment for the Code Section based on PE-BEAR values
    ' Code section starts at Raw 1024 (&H400) and loads at Virtual 4096 (&H1000)
    Dim lngRawCodeStart As Long
    Dim lngVirtualCodeStart As Long
    
    lngRawCodeStart = RoundToMinSize(GetPEHeaderSize(Sections))         ' This equals 1024 (&H400)
    lngVirtualCodeStart = RoundToSectionSize(lngRawCodeStart) ' This equals 4096 (&H1000)
   ' lngRawCodeStart = 4096  ' .code VirtualAddress = 0x1000 (4096 bytes)
    
    For i = 0 To m_lngRelocCount - 1
        ' Formula: (RawFileOffset - lngRawCodeStart) + lngVirtualCodeStart
        ' 1st MOV (data1): (1056 - 1024) + 4096 = RVA 4128 (&H1020) -> Byte-Perfect!
        ' 2nd MOV (data1): (1062 - 1024) + 4096 = RVA 4134 (&H1026) -> Byte-Perfect!
        ' Extern Call:     (1054 - 1024) + 4096 = RVA 4126 (&H101E) -> Byte-Perfect!
        m_udtRelocs(i).Offset = (m_udtRelocs(i).Offset - lngRawCodeStart) + lngVirtualCodeStart
    Next i
    
    Dim udtTempReloc As ASMRelocation
    For i = 0 To m_lngRelocCount - 2
        For j = i + 1 To m_lngRelocCount - 1
            If m_udtRelocs(i).Offset > m_udtRelocs(j).Offset Then
                udtTempReloc = m_udtRelocs(i)
                m_udtRelocs(i) = m_udtRelocs(j)
                m_udtRelocs(j) = udtTempReloc
            End If
        Next j
    Next i
    
    i = 0
    Do While i < m_lngRelocCount
        lngPageRVA = m_udtRelocs(i).Offset And &HFFFFF000
        lngPageStartPos = lngPos
        
        CopyMemory btOut(lngPos), lngPageRVA, 4: lngPos = lngPos + 4
        lngPos = lngPos + 4
        
        Do While i < m_lngRelocCount
            If (m_udtRelocs(i).Offset And &HFFFFF000) <> lngPageRVA Then Exit Do
            
            lngWord = (m_udtRelocs(i).Offset And &HFFF) Or (CInt(IMAGE_REL_BASED_HIGHLOW) * &H1000)
            CopyMemory btOut(lngPos), lngWord, 2: lngPos = lngPos + 2
            i = i + 1
        Loop
        
        If (lngPos - lngPageStartPos) Mod 4 <> 0 Then
            lngWord = 0 ' IMAGE_REL_BASED_ABSOLUTE (Padding)
            CopyMemory btOut(lngPos), lngWord, 2: lngPos = lngPos + 2
        End If
        
        Dim lngBlockSize As Long
        lngBlockSize = lngPos - lngPageStartPos
        CopyMemory btOut(lngPageStartPos + 4), lngBlockSize, 4
    Loop
    
    ReDim Preserve btOut(lngPos - 1) As Byte
    OutSize = lngPos
    BuildRelocationSection = btOut
End Function

Private Function BuildExportSection(ByVal BaseRVA As Long, ByRef OutSize As Long) As Byte()
    If m_lngExportCount = 0 Then Exit Function
    Dim i As Long, j As Long
    Dim udtTemp As ASMExport
    ' sort
    For i = 0 To m_lngExportCount - 2
        For j = i + 1 To m_lngExportCount - 1
            If StrComp(m_udtExports(i).Name, m_udtExports(j).Name, vbBinaryCompare) > 0 Then
            udtTemp = m_udtExports(i)
            m_udtExports(i) = m_udtExports(j)
            m_udtExports(j) = udtTemp
        End If
        Next
    Next
    
    Dim lngNamesSize As Long
    Dim strDLLFile As String
    strDLLFile = IIf(m_strDLLName = "", "compiled_lib.dll", m_strDLLName) & ChrW$(0)
    lngNamesSize = EvenSize(Len(strDLLFile))
    For i = 0 To m_lngExportCount - 1
        lngNamesSize = lngNamesSize + EvenSize(Len(m_udtExports(i).Name & ChrW$(0)))
    Next
    Dim lngTotalSize As Long
    lngTotalSize = 40 + (4 * m_lngExportCount) + (4 * m_lngExportCount) + (2 * m_lngExportCount) + lngNamesSize
    Dim btOut() As Byte
    ReDim btOut(lngTotalSize - 1) As Byte
    Dim ptrHeader As Long: ptrHeader = 0
    Dim ptrEAT As Long: ptrEAT = 40
    Dim ptrNamePtr As Long: ptrNamePtr = ptrEAT + (4 * m_lngExportCount)
    Dim ptrOrdinals As Long: ptrOrdinals = ptrNamePtr + (4 * m_lngExportCount)
    Dim ptrStrings As Long: ptrStrings = ptrOrdinals + (2 * m_lngExportCount)
    Dim btStr() As Byte
    Dim lngDLLNameRVA As Long
    btStr = StrConv(strDLLFile, vbFromUnicode)
    CopyMemory btOut(ptrStrings), btStr(0), UBound(btStr) + 1
    lngDLLNameRVA = BaseRVA + ptrStrings
    ptrStrings = ptrStrings + EvenSize(UBound(btStr) + 1)
    Dim lngNameRVAs() As Long
    ReDim lngNameRVAs(m_lngExportCount - 1) As Long
    For i = 0 To m_lngExportCount - 1
        btStr = StrConv(m_udtExports(i).Name & ChrW$(0), vbFromUnicode)
        CopyMemory btOut(ptrStrings), btStr(0), UBound(btStr) + 1
        lngNameRVAs(i) = BaseRVA + ptrStrings
        ptrStrings = ptrStrings + EvenSize(UBound(btStr) + 1)
    Next
    Dim lngFuncRVA As Long
    Dim idxLabel As Long
    Dim wordOrdinal As Integer
    For i = 0 To m_lngExportCount - 1
        idxLabel = GetLabelIndex(m_udtExports(i).LabelName)
        If idxLabel <> -1 Then
            ' Since Offset already contains (ImageBase + RVA),
            ' we simply subtract the BaseAddress to get the exact memory RVA!
            lngFuncRVA = m_udtLabels(idxLabel).Offset - m_lngBaseAddress
        Else
            lngFuncRVA = 0
        End If
        
        CopyMemory btOut(ptrEAT + (i * 4)), lngFuncRVA, 4
        CopyMemory btOut(ptrNamePtr + (i * 4)), lngNameRVAs(i), 4
        wordOrdinal = CInt(i)
        CopyMemory btOut(ptrOrdinals + (i * 2)), wordOrdinal, 2
    Next

    Dim expDir As IMAGE_EXPORT_DIRECTORY
    With expDir
        .Name = lngDLLNameRVA
        .Base = 1
        .NumberOfFunctions = m_lngExportCount
        .NumberOfNames = m_lngExportCount
        .AddressOfFunctions = BaseRVA + 40
        .AddressOfNames = BaseRVA + 40 + (4 * m_lngExportCount)
        .AddressOfNameOrdinals = BaseRVA + 40 + (4 * m_lngExportCount) + (4 * m_lngExportCount)
    End With
    CopyMemory btOut(0), expDir, 40
    OutSize = lngTotalSize
    BuildExportSection = btOut
End Function



' calculate label- and instructionoffsets
Private Sub FillInOffsets(ByVal Sections As Integer)
    Dim i           As Long
    Dim j           As Long
    Dim lngPEOffset As Long
    
    If m_blnWritePE Then
        lngPEOffset = RoundToSectionSize(GetPEHeaderSize(Sections))
    End If
    
    For i = 0 To m_lngInstrCount - 1
        ' ? FIX: Dynamically calculate ALIGN size based on the current live address
        If m_udtInstrs(i).Mnemonic = "ALIGN" Then
            Dim lngCurrentVirtualAddr As Long
            Dim lngAlignBoundary As Long
            Dim lngPaddingNeeded As Long
            
            ' Calculate the actual memory address where this ALIGN instruction starts
            lngCurrentVirtualAddr = m_lngOutSize + m_lngBaseAddress + lngPEOffset
            lngAlignBoundary = m_udtInstrs(i).Args(0).Value
            
            If lngAlignBoundary > 0 Then
                If (lngCurrentVirtualAddr Mod lngAlignBoundary) <> 0 Then
                    ' Calculate exactly how many NOP bytes are needed to reach mod boundary = 0
                    lngPaddingNeeded = lngAlignBoundary - (lngCurrentVirtualAddr Mod lngAlignBoundary)
                    m_udtInstrs(i).Size = lngPaddingNeeded
                Else
                    m_udtInstrs(i).Size = 0
                End If
            End If
        End If

        ' Assign the calculated offset to the instruction
        m_udtInstrs(i).Offset = m_lngOutSize + m_lngBaseAddress + lngPEOffset
        m_lngOutSize = m_lngOutSize + m_udtInstrs(i).Size

        If j < m_lngLabelCount Then
            If m_udtLabels(j).Instruction = i Then
                m_udtLabels(j).Offset = m_udtInstrs(i).Offset
                j = j + 1
            End If
        End If
    Next
    
    If j < m_lngLabelCount Then
        With m_udtInstrs(m_lngInstrCount - 1)
            For j = j To m_lngLabelCount - 1
                m_udtLabels(j).Offset = .Offset
            Next
        End With
    End If
End Sub


' calculate jump addresses for imported functions
Private Sub FillInIAT(ByVal reladdr As Long)
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim libidx      As Long
    Dim fncidx      As Long
    
    For i = 0 To m_lngInstrCount - 1
        For j = 0 To m_udtInstrs(i).ArgCount - 1
            If (m_udtInstrs(i).Args(j).TType And ParamExt) Then
                libidx = (m_udtInstrs(i).Args(j).SymbolIndex \ &H10000) And &HFFFF&
                fncidx = m_udtInstrs(i).Args(j).SymbolIndex And &HFFFF&
                m_udtInstrs(i).Args(j).Pointer.Ptr.Displacement = GetExternRelOfFnc(libidx, fncidx, reladdr)
            End If
        Next
    Next
End Sub


' compare parsed instruction with one of the instruction set
Private Function CompareInstrs( _
    Src As ASMInstruction, _
    Comp As Instruction _
) As Boolean

    Dim i   As Long
    
    If Src.ArgCount = Comp.ParamCount Then
        For i = 0 To Src.ArgCount - 1
            ' imm and rel should be treated equal
            With Comp.Parameters(i)
                If (.PType And Src.Args(i).TType) = 0 Then
                    If Not (.PType = ParamImm And Src.Args(i).TType = ParamRel) Then
                        If Not (.PType = ParamRel And Src.Args(i).TType = ParamImm) Then
                            Exit Function
                        End If
                    End If
                End If
            End With
            
            If Comp.Parameters(i).Forced Then
                Select Case Comp.Parameters(i).PType
                    Case ParamReg:
                        If Src.Args(i).Register <> Comp.Parameters(i).Register Then
                            Exit Function
                        End If
                    Case ParamSTX:
                        If Src.Args(i).FPURegister <> Comp.Parameters(i).FPURegister Then
                            Exit Function
                        End If
                    Case ParamImm:
                        If Src.Args(i).Value <> Comp.Parameters(i).Value Then
                            Exit Function
                        End If
                End Select
            Else
                If Comp.Parameters(i).PType = ParamMem Then
                    If Not Comp.ModRM Then
                        If Src.Args(i).Pointer.RegisterCount > 0 Then
                            ' instruction mustn't have registers in the pointer
                            ' because ModR/M isn't allowed for it
                            Exit Function
                        End If
                    End If
                ElseIf (Comp.Parameters(i).PType And ParamMM) Then
                    If IsDefinite(Comp.Parameters(i).MMRegister) Then
                        If Comp.Parameters(i).MMRegister <> Src.Args(i).MMRegister Then
                            Exit Function
                        End If
                    Else
                        If Src.Args(i).TType = ParamMem Then
                            If (Comp.Parameters(i).PType And ParamMem) = 0 Then
                                Exit Function
                            End If
                        Else
                            If (Comp.Parameters(i).MMRegister And Src.Args(i).MMRegister) = 0 Then
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            
            If (Comp.Parameters(i).Size And Src.Args(i).Size) = 0 Then
                If Comp.Parameters(i).Size <> BitsUnknown Then
                    Exit Function
                End If
            End If
        Next
        
        CompareInstrs = True
    End If
End Function


' collect labels
Private Function FindLabels() As Boolean
    Dim i           As Long
    Dim lngInstrCnt As Long
    
    For i = 1 To m_lngTokenCount - 2
        If m_clsTokens(i).TType = tokensymbol Then
            If m_clsTokens(i + 1).TType = TokenOpColon Then
                If GetLabelIndex(m_clsTokens(i).content) > -1 Then
                    SetError "ambigious names: " & m_clsTokens(i).content, m_clsTokens(i).Line, m_clsTokens(i).Section
                    Exit Function
                End If
                AddLabel m_clsTokens(i).content, lngInstrCnt
            End If
        ElseIf m_clsTokens(i).TType = TokenOperator Then
            lngInstrCnt = lngInstrCnt + 1
        ElseIf m_clsTokens(i).TType = TokenRawData Then
            lngInstrCnt = lngInstrCnt + 1
        End If
    Next
    
    FindLabels = True
End Function


' skips labels because they're already collected
Private Function ParseInstructions() As Boolean
    If Not Match(TokenBeginOfInput) Then
        SetError "Unknown error occured while starting parsing", 0, ""
    Else
        Do While token.TType <> TokenEndOfInput
            Select Case token.TType
            
                Case TokenExtern:
                    If m_blnWritePE Then
                        Match TokenExtern
                        If Not ParseExtern() Then Exit Do
                    Else
                        SetError "Externs only allowed in PE mode", token.Line, token.Section
                        Exit Function
                    End If
            
                Case tokensymbol:
                    Match tokensymbol
                    If Not Match(TokenOpColon) Then
                        SetError """:"" expected after label ID", token.Line, token.Section
                        Exit Do
                    End If
                    
                Case TokenOperator, TokenKeyword:
                    If Not ParseInstruction Then Exit Do
                    
                Case TokenRawData:
                    If Not ParseRawData Then Exit Do
                    
                Case Else:
                    If token.TType <> TokenEndOfInstruction Then
                        SetError "Unexpected symbol: " & token.content, token.Line, token.Section
                        Exit Do
                    End If
                    
            End Select
            
            If Not Match(TokenEndOfInstruction) Then
                SetError "Unexpected end", token.Line, token.Section
                Exit Do
            Else
                ParseInstructions = token.TType = TokenEndOfInput
            End If
        Loop
    End If
End Function


Private Function AddExtern(ByVal LIB As String, ByVal fnc As String) As Boolean
    Dim i   As Long
    Dim j   As Long
    
    AddExtern = True
    
    For i = 0 To m_lngExternCount - 1
        If StrComp(m_udtExtern(i).LibName, LIB, vbTextCompare) = 0 Then
            For j = 0 To m_udtExtern(i).FunctionCount - 1
                If StrComp(m_udtExtern(i).Functions(j), fnc, vbTextCompare) = 0 Then
                    Exit Function
                End If
            Next
            
            With m_udtExtern(i)
                ReDim Preserve .Functions(.FunctionCount) As String
                .Functions(.FunctionCount) = fnc
                .FunctionCount = .FunctionCount + 1
            End With
            Exit Function
        End If
    Next
    
    If i = m_lngExternCount Then
        ReDim Preserve m_udtExtern(m_lngExternCount) As ASMExtern
        
        With m_udtExtern(m_lngExternCount)
            .LibName = LIB
            
            ReDim .Functions(0) As String
            .Functions(0) = fnc
            .FunctionCount = 1
        End With
        
        m_lngExternCount = m_lngExternCount + 1
    End If
End Function


Private Function ParseExtern() As Boolean
    Dim strLib  As String
    Dim strFnc  As String
        
    If token.TType <> TokenString Then
        SetError "Expected: string identifier for library name", token.Line, token.Section
        Exit Function
    Else
        strLib = token.content
        Match TokenString
        
        If Not Match(TokenSeparator) Then
            SetError "Libraryname and functionname have to be seperated through a "",""", token.Line, token.Section
            Exit Function
        Else
            If token.TType <> tokensymbol Then
                SetError "Expected: Name of the exporte function", token.Line, token.Section
                Exit Function
            Else
                strFnc = token.content
                Match tokensymbol
                
                If GetLabelIndex(strFnc) > -1 Then
                    SetError "Name not unique", token.Line, token.Section
                Else
                    ParseExtern = AddExtern(strLib, strFnc)
                End If
            End If
        End If
    End If
End Function


Private Function ParseRawData() As Boolean
    Dim udtInstr    As ASMInstruction
    Dim i           As Long
    Dim lngTemp     As Long
    Dim strOp       As String
    
    strOp = UCase$(token.content)
    
    ' Configure size based on data type
    Select Case strOp
        Case "DB":  udtInstr.Data.Size = Bits8
        Case "DW":  udtInstr.Data.Size = Bits16
        Case "DD":  udtInstr.Data.Size = Bits32
        Case "ALIGN": udtInstr.Data.Size = Bits8 ' NOPs are 1 byte each
    End Select
    Match TokenRawData
    ' ? FIX: Just parse and store the alignment boundary value (e.g., 4)
    If strOp = "ALIGN" Then
        udtInstr.Mnemonic = "ALIGN"
        If token.TType = TokenValue Then
            ' Store the alignment boundary (like 4) in the first argument's value
            udtInstr.Args(0).Value = token.Value
            Match TokenValue
        Else
            SetError "Alignment boundary expected (e.g., ALIGN 4)", token.Line, token.Section
            Exit Function
        End If
        
        ' At parse time, the size of this pseudo-instruction is temporary 0
        udtInstr.Size = 0
        AddInstruction udtInstr
        ParseRawData = True
        Exit Function
    End If
    
    ' Standard DB, DW, DD handling (Original Elster Code)
    With udtInstr.Data
        Do
            Select Case token.TType
                Case TokenValue:
                    ReDim Preserve .Values(.ValueCount) As Long
                    .Values(.ValueCount) = token.Value
                    .ValueCount = .ValueCount + 1
                    Match TokenValue
                    
                Case TokenString:
                    lngTemp = .ValueCount
                    ReDim Preserve .Values(.ValueCount + Len(token.content) - 1) As Long
                    If udtInstr.Data.Size = Bits8 Then
                        For i = 1 To Len(token.content)
                            .Values(lngTemp + i - 1) = Asc(Mid$(token.content, i, 1))
                        Next
                    Else
                        For i = 1 To Len(token.content)
                            .Values(lngTemp + i - 1) = AscW(Mid$(token.content, i, 1))
                        Next
                    End If
                    .ValueCount = .ValueCount + Len(token.content)
                    Match TokenString
                    
                Case Else:
                    SetError "Unexpected: " & token.content, udtInstr.Line, udtInstr.Section
                    Exit Function
            End Select
        Loop While Match(TokenSeparator)
    End With
    
    AddInstruction udtInstr
    ParseRawData = True
End Function

Private Function ParseInstruction() As Boolean
    Dim udtInstr    As ASMInstruction
    
    ' instruction can have multiple keywords like "lock rep"
    Do While token.TType = TokenKeyword
        Select Case UCase$(token.content)
            Case "REP":     udtInstr.Flags = udtInstr.Flags Or PrefixFlgRep
            Case "REPE":    udtInstr.Flags = udtInstr.Flags Or PrefixFlgRepe
            Case "REPZ":    udtInstr.Flags = udtInstr.Flags Or PrefixFlgRepz
            Case "REPNE":   udtInstr.Flags = udtInstr.Flags Or PrefixFlgRepne
            Case "REPNZ":   udtInstr.Flags = udtInstr.Flags Or PrefixFlgRepnz
            Case "LOCK":    udtInstr.Flags = udtInstr.Flags Or PrefixFlgLock
            Case Else:      SetError "Unknown keyword: " & token.content, token.Line, token.Section
                            Exit Function
        End Select
        
        Match TokenKeyword
    Loop
    
    If token.TType <> TokenOperator Then
        SetError "Operator expected", token.Line, token.Section
    Else
        udtInstr.Mnemonic = token.content
        udtInstr.Line = token.Line
        udtInstr.Section = token.Section
        
        Match TokenOperator
        
        If ParseArguments(udtInstr) Then
            AddInstruction udtInstr
            ParseInstruction = True
        End If
    End If
End Function


' Reads up to 3 arguments per instruction.
' First keywords are read, then the rest like labels, numerical values,
' registers or pointers.
Private Function ParseArguments(udtInstr As ASMInstruction) As Boolean
    Dim lngArgs     As Long
    Dim lngPtrIdx   As Long
    Dim blnCont     As Boolean
    
    lngPtrIdx = -1
    
    Do
        If lngArgs = 3 Then
            SetError "Too many arguments", udtInstr.Line, udtInstr.Section
            Exit Function
        End If
        
        If token.TType = TokenKeyword Then
            If Not ParseKeywords(udtInstr, lngArgs) Then Exit Function
        End If
        
        blnCont = False
        Select Case token.TType
            Case TokenRegister:
                If Not ParseArgumentRegister(udtInstr, udtInstr.Args(lngArgs)) Then Exit Function
            Case TokenFPUReg:
                If Not ParseArgumentFPUReg(udtInstr, udtInstr.Args(lngArgs)) Then Exit Function
            Case TokenMMRegister:
                If Not ParseArgumentMMReg(udtInstr, udtInstr.Args(lngArgs)) Then Exit Function
            Case tokensymbol:
                If Not ParseArgumentSymbol(udtInstr, udtInstr.Args(lngArgs)) Then Exit Function
            Case TokenValue:
                If Not ParseArgumentValue(udtInstr, udtInstr.Args(lngArgs)) Then Exit Function
            Case TokenBracketLeft:
                If Not ParseArgumentPtr(udtInstr, udtInstr.Args(lngArgs)) Then Exit Function
                lngPtrIdx = lngArgs
            Case TokenSegmentReg:
                If Not ParseSegmentReg(udtInstr) Then Exit Function
                blnCont = True
            Case TokenEndOfInstruction:
                Exit Do
            Case Else:
                SetError "Unexpected: " & token.content, udtInstr.Line, udtInstr.Section
                Exit Function
        End Select
        
        If Not blnCont Then lngArgs = lngArgs + 1
    Loop While Match(TokenSeparator) Or blnCont
    
    If lngPtrIdx > -1 Then
        If udtInstr.Args(lngPtrIdx).Size = BitsUnknown Then
            If lngArgs = 2 Then
                If IsDefinite(udtInstr.Args(1 - lngPtrIdx).Size) Or _
                   (udtInstr.Args(1 - lngPtrIdx).TType And ParamMM) Then
                    udtInstr.Args(lngPtrIdx).Size = udtInstr.Args(1 - lngPtrIdx).Size
                Else
                    SetError "Invalid pointer size", udtInstr.Line, udtInstr.Section
                    Exit Function
                End If
            End If
        End If
    End If
    
    udtInstr.ArgCount = lngArgs
    ParseArguments = True
End Function


Private Function ParseArgumentPtr( _
    udtInstr As ASMInstruction, _
    arg As ASMArgument _
) As Boolean

    arg.TType = ParamMem
    arg.Pointer.TokenIndex = m_lngCurToken
    ParseArgumentPtr = ValidatePointer(udtInstr, arg.Pointer)
End Function


Private Function ParseArgumentValue( _
    udtInstr As ASMInstruction, _
    arg As ASMArgument _
) As Boolean

    arg.TType = ParamImm
    arg.Value = token.Value
    
    If arg.Size = BitsUnknown Then
        arg.Size = SizesForInt(token.Value)
    Else
        If (arg.Size And SizesForInt(token.Value)) = 0 Then
            SetError "Incompatible integer sizes", udtInstr.Line, udtInstr.Section
            Exit Function
        End If
    End If
    
    Match TokenValue
    ParseArgumentValue = True
End Function


Private Function ParseArgumentSymbol( _
    udtInstr As ASMInstruction, _
    arg As ASMArgument _
) As Boolean

    Dim lngLblIdx   As Long
    Dim blnExtern   As Boolean
    
    lngLblIdx = GetLabelIndex(token.content)
    If lngLblIdx = -1 Then
        lngLblIdx = GetExternIndex(token.content)
        If lngLblIdx = -1 Then
            SetError "Unknown label: " & token.content, udtInstr.Line, udtInstr.Section
            Exit Function
        Else
            blnExtern = True
        End If
    End If
    
    If blnExtern Then
        arg.TType = ParamMem Or ParamExt
        arg.SymbolIndex = lngLblIdx
        arg.Size = Bits32
        
        arg.Pointer.HasDisplacement = True
        arg.Pointer.Ptr.DispSize = Bits32
        arg.Pointer.DispSize = Bits32
    Else
        arg.TType = ParamRel
        arg.SymbolIndex = lngLblIdx
        
        If arg.Size = BitsUnknown Then
            arg.Size = Bits32 Or Bits16 Or Bits8
        End If
    End If
    
    Match tokensymbol
    ParseArgumentSymbol = True
End Function


Private Function ParseArgumentMMReg( _
    udtInstr As ASMInstruction, _
    arg As ASMArgument _
) As Boolean

    arg.TType = ParamMM
    arg.MMRegister = MMRegStrToNum(token.content)
    
    If arg.Size <> BitsUnknown Then
        If (arg.Size And (Bits32 Or Bits64 Or Bits80 Or Bits128)) = 0 Then
            SetError "Incompatible sizes", udtInstr.Line, udtInstr.Section
            Exit Function
        End If
    Else
        If (arg.MMRegister < XMM0) Then
            arg.Size = Bits64
        Else
            arg.Size = Bits128
        End If
    End If
    
    Match TokenMMRegister
    ParseArgumentMMReg = True
End Function


Private Function ParseArgumentFPUReg( _
    udtInstr As ASMInstruction, _
    arg As ASMArgument _
) As Boolean

    arg.TType = ParamSTX
    arg.FPURegister = FPURegStrToNum(token.content)
    
    If arg.Size <> BitsUnknown Then
        If (arg.Size And (Bits32 Or Bits64 Or Bits80)) = 0 Then
            SetError "Incompatible sizes", udtInstr.Line, udtInstr.Section
            Exit Function
        End If
    Else
        arg.Size = Bits32 Or Bits64 Or Bits80
    End If
    
    Match TokenFPUReg
    ParseArgumentFPUReg = True
End Function


Private Function ParseArgumentRegister( _
    udtInstr As ASMInstruction, _
    arg As ASMArgument _
) As Boolean

    arg.TType = ParamReg
    arg.Register = RegStrToReg(token.content)
    
    If arg.Size <> BitsUnknown Then
        If arg.Size <> RegisterSize(arg.Register) Then
            SetError "Incompatible sizes", udtInstr.Line, udtInstr.Section
            Exit Function
        End If
    Else
        arg.Size = RegisterSize(arg.Register)
    End If
    
    Match TokenRegister
    ParseArgumentRegister = True
End Function


' Only sizes IDs valid at this position. "byte dword" causes an error at "dword"
Private Function ParseKeywords( _
    udtInstr As ASMInstruction, _
    argidx As Long _
) As Boolean

    Select Case UCase$(token.content)
        Case "BYTE":                udtInstr.Args(argidx).Size = Bits8
        Case "WORD":                udtInstr.Args(argidx).Size = Bits16
        Case "DWORD", "FLOAT":      udtInstr.Args(argidx).Size = Bits32
        Case "QWORD", "DOUBLE":     udtInstr.Args(argidx).Size = Bits64
        Case "EXTENDED":            udtInstr.Args(argidx).Size = Bits80
        Case "DQWORD":              udtInstr.Args(argidx).Size = Bits128
        Case Else:                  SetError "At this position invalid: " & token.content, udtInstr.Line, udtInstr.Section
                                    Exit Function
    End Select
    
    Match TokenKeyword
    ParseKeywords = True
End Function


Private Function ValidatePointer( _
    udtInstr As ASMInstruction, _
    Ptr As PointerInfo _
) As Boolean

    If Not Match(TokenBracketLeft) Then
        SetError "Expected: [", udtInstr.Line, udtInstr.Section
    Else
        If ValidateExpr(udtInstr, Ptr) Then
            If Not Match(TokenBracketRight) Then
                SetError """["" Expected", udtInstr.Line, udtInstr.Section
            Else
                ValidatePointer = True
            End If
        End If
    End If
End Function


' validate expressions in brackets ( e.g. "[eax*2+ebp-10]" )
Private Function ValidateExpr( _
    udtInstr As ASMInstruction, _
    Ptr As PointerInfo _
) As Boolean

    Dim lngRegCount(REG_COUNT - 1)  As Long
    Dim lngUsedRegs(1)              As Long
    Dim lngDisplacement             As Long
    Dim lngSign                     As Long
    Dim lngReg                      As Long
    Dim lngTms                      As Long
    Dim i                           As Long
    Dim blnHasSymbols               As Boolean
    Dim blnHasSymbol                As Boolean
    Dim blnHasRegister              As Boolean
    
    If Match(TokenOpAdd) Then
        lngSign = 1
    ElseIf Match(TokenOpSub) Then
        lngSign = -1
    Else
        lngSign = 1
    End If
    
    Do
        lngTms = 1
        blnHasSymbol = False
        blnHasRegister = False
        
        Do
            Select Case token.TType
                Case TokenRegister:
                    If blnHasRegister Then
                        SetError "Power of register not possible", udtInstr.Line, udtInstr.Section
                        Exit Function
                    End If
                    
                    If blnHasSymbol Then
                        SetError "Register*Label not possible", udtInstr.Line, udtInstr.Section
                        Exit Function
                    End If
                    
                    If RegStrToReg(token.content) < RegEAX Then
                        SetError "Only 32 bit registers allowed in pointer", udtInstr.Line, udtInstr.Section
                        Exit Function
                    End If
                    
                    lngReg = RegToIdx(RegStrToReg(token.content))
                    blnHasRegister = True
                    Match TokenRegister
                
                Case TokenValue:
                    lngTms = lngTms * token.Value
                    Match TokenValue
                
                Case tokensymbol:
                    If blnHasRegister Then
                        SetError "Register*Label not possible", udtInstr.Line, udtInstr.Section
                        Exit Function
                    End If
                    
                    If GetLabelIndex(token.content) = -1 Then
                        SetError "Unknown symbol: " & token.content, udtInstr.Line, udtInstr.Section
                        Exit Function
                    End If
                                        
                    blnHasSymbol = True
                    blnHasSymbols = True
                    Match tokensymbol
                
                Case Else:
                    SetError "Invalid in pointer: " & token.content, udtInstr.Line, udtInstr.Section
                    Exit Function
                    
            End Select
        Loop While Match(TokenOpMul)
        
        If blnHasRegister Then
            lngRegCount(lngReg) = lngRegCount(lngReg) + lngSign * lngTms
        Else
            lngDisplacement = lngDisplacement + lngSign * lngTms
        End If
        
        If Match(TokenOpAdd) Then
            lngSign = 1
        ElseIf Match(TokenOpSub) Then
            lngSign = -1
        Else
            Exit Do
        End If
    Loop
    
    For i = 0 To REG_COUNT - 1
        If lngRegCount(i) > 0 Then
            If lngRegCount(i) > 1 Then
                Ptr.RegisterMultiples = True
            Else
                ' If the register is ESP the SIB byte needs to be written.
                ' InstructionOutArgs() only writes it if RegisterMultiples is true,
                ' do so.
                If IdxToReg(i) = RegESP Then Ptr.RegisterMultiples = True
            End If
            
            Ptr.RegisterCount = Ptr.RegisterCount + 1
            If Ptr.RegisterCount > 2 Then Exit For
            lngUsedRegs(Ptr.RegisterCount - 1) = lngRegCount(i)
            
        ElseIf lngRegCount(i) < 0 Then
            SetError "Negative multiple of register", udtInstr.Line, udtInstr.Section
            Exit Function
            
        End If
    Next
    
    If blnHasSymbols Then
        Ptr.HasDisplacement = True
        Ptr.DispSize = Bits32
    Else
        If lngDisplacement <> 0 Then
            Ptr.HasDisplacement = True
            Ptr.DispSize = GetFirstSetBit(SizesForInt(lngDisplacement))
            If Ptr.DispSize > Bits8 Then
                ' Bits16 not possible because of ModR/M
                Ptr.DispSize = Bits32
            End If
        Else
            If lngRegCount(RegToIdx(RegEBP)) > 0 Then
                If Ptr.RegisterCount = 1 Then
                    ' EBP can only be encoded with a displacement,
                    ' so simply add a zero'ed 8 Bit displacement
                    Ptr.HasDisplacement = True
                    Ptr.DispSize = Bits8
                End If
            Else
                Ptr.HasDisplacement = False
                Ptr.DispSize = BitsUnknown
            End If
        End If
    End If
    
    If Ptr.RegisterCount = 1 Then
        If (lngUsedRegs(0) = 4) Or (lngUsedRegs(0) = 8) Then
            ' Reg*4/8 must be encoded with SIB,
            ' but in this case there is not base register (ptr.RegisterCount = 1).
            ' Because of this a 32 Bit displacement is needed
            Ptr.HasDisplacement = True
            Ptr.DispSize = Bits32
        End If
    End If
    
    If Ptr.RegisterCount > 2 Then
        SetError "Not more than two registers per pointer allowed", udtInstr.Line, udtInstr.Section
    Else
        Ptr.Ptr.UsedRegisters = Ptr.RegisterCount
        Ptr.Ptr.DispSize = Ptr.DispSize
        ValidateExpr = True
    End If
End Function


Private Function ParseSegmentReg(udtInstr As ASMInstruction) As Boolean
    If udtInstr.Segment <> SegUnknown Then
        SetError "Only one segment override per instruction", udtInstr.Line, udtInstr.Section
    Else
        udtInstr.Segment = SegStrToSeg(token.content)
        Match TokenSegmentReg
     
        If Match(TokenOpColon) Then
            ParseSegmentReg = True
        Else
            SetError "Expected: Colon after segment", udtInstr.Line, udtInstr.Section
        End If
    End If
End Function


Private Sub AddInstruction(udtInstr As ASMInstruction)
    ReDim Preserve m_udtInstrs(m_lngInstrCount) As ASMInstruction
    m_udtInstrs(m_lngInstrCount) = udtInstr
    m_lngInstrCount = m_lngInstrCount + 1
End Sub


' index of an imported function consists of two indexes
' library index is in the upper 16 bits, function index in the lower 16 bits
' of the returned Long
Private Function GetExternIndex(Name As String) As Long
    Dim i   As Long
    Dim j   As Long
    
    For i = 0 To m_lngExternCount - 1
        For j = 0 To m_udtExtern(i).FunctionCount - 1
            If StrComp(m_udtExtern(i).Functions(j), Name, vbTextCompare) = 0 Then
                GetExternIndex = i * &H10000 Or j
                Exit Function
            End If
        Next
    Next
    
    GetExternIndex = -1
End Function


' IAT address of an imported function
' reladdr should be the start of the IAT relative to the Image Base
Private Function GetExternRelOfFnc(ByVal libidx As Long, ByVal fncidx As Long, ByVal reladdr As Long) As Long
    Dim i   As Long
    Dim j   As Long
    Dim SZ  As Long
    
    For i = 0 To libidx - 1
        SZ = SZ + 4 * (m_udtExtern(i).FunctionCount + 1)
    Next
    SZ = SZ + 4 * fncidx
    
    GetExternRelOfFnc = SZ + m_lngBaseAddress + reladdr
End Function


' start of imported functions of a specific library
' in the import lookup table relative to IAT
Private Function GetILTLibStart(ByVal index As Long) As Long
    Dim im  As IMAGE_IMPORT_DIRECTORY
    Dim SZ  As Long
    
    SZ = GetNeededIATSize() + GetIATLibStart(index)
    SZ = SZ + Len(im) * (m_lngExternCount + 1)
    
    GetILTLibStart = SZ
End Function


' start of imported functions of a specific library in the IAT
' (which should fill the first bytes in the data section)
Private Function GetIATLibStart(ByVal index As Long) As Long
    Dim i   As Long
    Dim SZ  As Long
    
    For i = 0 To index - 1
        SZ = SZ + 4 * (m_udtExtern(i).FunctionCount + 1)
    Next
    
    GetIATLibStart = SZ
End Function


' size of IAT
Private Function GetNeededIATSize() As Long
    Dim i   As Long
    Dim SZ  As Long
    
    For i = 0 To m_lngExternCount - 1
        SZ = SZ + 4 * (m_udtExtern(i).FunctionCount + 1)
    Next
    
    GetNeededIATSize = SZ
End Function


' size of imports (IAT+ILT+names)
Private Function GetNeededImportsSize() As Long
    Dim i   As Long
    Dim j   As Long
    Dim SZ  As Long
    Dim im  As IMAGE_IMPORT_DIRECTORY
    
    For i = 0 To m_lngExternCount - 1
        SZ = SZ + EvenSize(Len(m_udtExtern(i).LibName) + 1)
        SZ = SZ + Len(im)
        
        For j = 0 To m_udtExtern(i).FunctionCount - 1
            SZ = SZ + EvenSize(Len(m_udtExtern(i).Functions(j)) + 1) + 2
            SZ = SZ + 4 * 2
        Next
        
        SZ = SZ + 4 * 2
    Next
    
    If m_lngExternCount > 0 Then
        SZ = SZ + Len(im)
    End If
    
    GetNeededImportsSize = SZ
End Function


Private Function RoundToSectionSize(ByVal lngVal As Long) As Long
    RoundToSectionSize = RoundTo(lngVal, MEM_SECTION_SIZE)
End Function


Private Function RoundToMinSize(ByVal lngVal As Long) As Long
    RoundToMinSize = RoundTo(lngVal, FILE_SECTION_SIZE)
End Function


Private Function RoundTo(ByVal lngVal As Long, ByVal lngMult As Long) As Long
    If lngVal Mod lngMult Then
        RoundTo = lngVal + (lngMult - lngVal Mod lngMult)
    Else
        RoundTo = lngVal
    End If
End Function


Private Function EvenSize(ByVal SZ As Long) As Long
    EvenSize = SZ + (SZ Mod 2)
End Function
Public Property Get OutputStrBytes() As String
    Dim Size    As Long
    Size = m_lngOutSize
    If Size Mod 2 = 1 Then
        OutputStrBytes = StrConv(space$(Size), vbFromUnicode)
    Else
        OutputStrBytes = space$(Size \ 2)
    End If
    CopyMemory StrPtr(OutputStrBytes), m_btOutput(0), Size
End Property

' hex string to byte array
Private Function HexToByte(ByVal strHex As String) As Byte()
    Dim bt()    As Byte
    Dim i       As Long
    
    ReDim bt(Len(strHex) \ 2 - 1) As Byte
    
    For i = 1 To Len(strHex) Step 2
        bt((i - 1) \ 2) = CByte("&H" & Mid$(strHex, i, 2))
    Next
    
    HexToByte = bt
End Function


Private Function GetLabelIndex(Name As String) As Long
    Dim i   As Long
    If labels.ExistKey(Name) Then
        GetLabelIndex = labels.Value
        Exit Function
    End If
    'For i = 0 To m_lngLabelCount - 1
    '    If StrComp(m_udtLabels(i).Name, Name, vbTextCompare) = 0 Then
    '        GetLabelIndex = i
    '        Exit Function
    '    End If
    'Next
    
    GetLabelIndex = -1
End Function


Private Sub SetError(ByVal strMsg As String, ByVal lngLine As Long, ByVal Section As String)
    m_strLastError = strMsg
    m_lngLastErrLine = lngLine
    m_strLastErrorSection = Section
End Sub


' match function for going through tokens after TokenizeInput()
' as there's need of doing a second pass
Private Function Match(ByVal tp As TokenType) As Boolean
    If m_clsTokens(m_lngCurToken).TType = tp Then
        m_lngCurToken = m_lngCurToken + 1
        Match = True
    End If
End Function


' Token property for accessing the current token after TokenizeInput()
' as there's need of doing a second pass
Private Property Get token() As ASMToken
    Set token = m_clsTokens(m_lngCurToken)
End Property


' collects all tokens and combines signs and numbers
Private Sub TokenizeInput()
    Dim blnInPtr    As Boolean
    Dim clsMinus    As ASMToken
    
    Do
        If clsMinus Is Nothing Then
            AddToken ScannerToken
        End If
        
        If ScannerToken.TType = TokenEndOfInput Then
            Exit Do
        Else
            ScannerGetNextToken
            Select Case ScannerToken.TType
                Case TokenBracketLeft:
                    blnInPtr = True
                Case TokenBracketRight:
                    blnInPtr = False
                Case TokenOpSub:
                    If Not blnInPtr Then Set clsMinus = ScannerToken
                Case TokenValue:
                    If Not clsMinus Is Nothing Then
                        ScannerToken.Value = -ScannerToken.Value
                        Set clsMinus = Nothing
                    End If
                Case Else:
                    If Not clsMinus Is Nothing Then
                        AddToken clsMinus
                        Set clsMinus = Nothing
                    End If
            End Select
        End If
    Loop
End Sub


Private Sub AddLabel(strName As String, ByVal lngInstr As Long)
    ReDim Preserve m_udtLabels(m_lngLabelCount) As ASMLabel
    If Not labels.ExistKey(strName) Then
        labels.AddKey strName, m_lngLabelCount
    End If
    With m_udtLabels(m_lngLabelCount)
        .Name = strName
        .Instruction = lngInstr
        .Offset = 0
    End With
    
    m_lngLabelCount = m_lngLabelCount + 1
End Sub


Private Sub AddToken(ByVal clsTk As ASMToken)
    ReDim Preserve m_clsTokens(m_lngTokenCount) As ASMToken
    Set m_clsTokens(m_lngTokenCount) = clsTk
    m_lngTokenCount = m_lngTokenCount + 1
End Sub


'#########################################################
'#########################################################
' Scanner functions


Private Function ScannerMatch(udeType As TokenType) As Boolean
    If m_udtScanner.CurToken.TType = udeType Then
        ScannerGetNextToken
        ScannerMatch = True
    End If
End Function


Private Property Get ScannerToken() As ASMToken
    Set ScannerToken = m_udtScanner.CurToken
End Property


Private Sub ScannerInit(Source As String)
    With m_udtScanner
        .Source = StrConv(Source & vbCrLf, vbFromUnicode)
        .Length = UBound(.Source) + 1
    
        .Position = 0
        .Line = 1
        .LinePos = 0
        .NextIsEOI = False
        .LastWasEOI = True
        .Section = "main"
    
        Set .NextToken = New ASMToken
        With .NextToken
            .TType = TokenBeginOfInput
            .Line = 1
            .Position = 0
        End With
    
        ScannerGetNextToken
    End With
End Sub


Private Sub ScannerGetNextToken()
    Dim clsToken    As ASMToken
    Dim i As Long
    
    Set clsToken = New ASMToken
    Set m_udtScanner.CurToken = m_udtScanner.NextToken
    
    If m_udtScanner.NextIsEOI Then
        clsToken.TType = TokenEndOfInstruction
        m_udtScanner.NextIsEOI = False
        m_udtScanner.LastWasEOI = True
    Else
        With m_udtScanner
            Do While (.Position < .Length) And (clsToken.TType = TokenUnknown)
                Select Case .Source(.Position)
                
                    Case CHAR_SPACE:
                        ' nichts tun
                        
                    Case CHAR_CARRIAGE:
                        clsToken.TType = TokenEndOfInstruction
                    
                    Case CHAR_VERT_BAR:
                        .Line = .Line + 1
                        .LinePos = .LinePos + 1
                        clsToken.TType = TokenEndOfInstruction

                    Case CHAR_LINEFEED:
                        .Line = .Line + 1
                        .LinePos = .LinePos + 1
                        
                    Case CHAR_SHARP:
                        .Section = ""
                        .Line = 0
                        .LinePos = 0
                        .Position = .Position + 1
                        
                        Do
                            .Section = .Section & ChrW$(.Source(.Position))
                            .Position = .Position + 1
                        Loop While .Source(.Position) <> CHAR_CARRIAGE
                        
                        .Position = .Position - 1
                        
                    Case CHAR_SEMICOLON:
                        Do
                            .Position = .Position + 1
                        Loop While .Source(.Position + 1) <> CHAR_CARRIAGE And _
                                   .Source(.Position + 1) <> CHAR_VERT_BAR
                            
                    Case CHAR_SEPARATOR:
                        clsToken.TType = TokenSeparator
                        clsToken.content = ","
                        
                    Case CHAR_COLON:
                        clsToken.TType = TokenOpColon
                        clsToken.content = ":"
                        
                    Case CHAR_PLUS:
                        clsToken.TType = TokenOpAdd
                        clsToken.content = "+"
                    
                    Case CHAR_MINUS:
                        clsToken.TType = TokenOpSub
                        clsToken.content = "-"
                    
                    Case CHAR_ASTERISK:
                        clsToken.TType = TokenOpMul
                        clsToken.content = "*"
                    
                    Case CHAR_BRACKET_L:
                        clsToken.TType = TokenBracketLeft
                        clsToken.content = "["
                    
                    Case CHAR_BRACKET_R:
                        clsToken.TType = TokenBracketRight
                        clsToken.content = "]"
                    
                    Case CHAR_NUMBER_0 To CHAR_NUMBER_9:
                        Select Case ChrW$(.Source(.Position + 1))
                            Case "x", "X", "χ", "Χ"
                                clsToken.content = "**"
                                .Position = .Position + 2
                        End Select
                        If clsToken.content = "**" Then
                            
                            If Not IsHexChar(.Source(.Position)) Then
                            clsToken.TType = tokensymbol
                            GoTo cont1231
                            End If
                            clsToken.content = "&H"
                            Do
                                If .Source(.Position) <> 95 Then
                                
                                clsToken.content = clsToken.content & ChrW$(.Source(.Position))
                                End If
                                .Position = .Position + 1
                            Loop While IsHexChar(.Source(.Position)) Or .Source(.Position) = 95
                           
                        Else
                            Do
                                If .Source(.Position) <> 95 Then
                                    clsToken.content = clsToken.content & ChrW$(.Source(.Position))
                                End If
                                .Position = .Position + 1
                            Loop While IsDigit(.Source(.Position)) Or .Source(.Position) = 95
                        End If
                        If .Source(.Position) <> CHAR_AMPERSAND Then
                            .Position = .Position - 1
                        End If
                        clsToken.TType = TokenValue
cont1231:
                    Case CHAR_AMPERSAND:
                        Select Case ChrW$(.Source(.Position + 1))
                        
                            Case "O", "o"
                                clsToken.content = ""
                                .Position = .Position + 2
                                If IsOctChar(.Source(.Position)) Then
                                    Do
                                        If .Source(.Position) <> 95 Then
                                        
                                        clsToken.content = clsToken.content & ChrW$(.Source(.Position))
                                        End If
                                        .Position = .Position + 1
                                    Loop While IsOctChar(.Source(.Position)) Or .Source(.Position) = 95
                                    If .Source(.Position) <> CHAR_AMPERSAND Then
                                        .Position = .Position - 1
                                    End If
                    
                                    clsToken.content = "&H" + PACKLNG$(val("&O" + clsToken.content))
                                clsToken.TType = TokenValue
                                Else
                                    clsToken.TType = TokenInvalid
                                    clsToken.content = "&"
                                End If
                                
                            Case "H", "h":
                                clsToken.content = "&H"
                                .Position = .Position + 2
                                If IsHexChar(.Source(.Position)) Then
                                    Do
                                        If .Source(.Position) <> 95 Then
                                            clsToken.content = clsToken.content & ChrW$(.Source(.Position))
                                        End If
                                        .Position = .Position + 1
                                    Loop While IsHexChar(.Source(.Position)) Or .Source(.Position) = 95
                                    If .Source(.Position) <> CHAR_AMPERSAND Then
                                        .Position = .Position - 1
                                    End If
                                    clsToken.TType = TokenValue
                                Else
                                    clsToken.TType = TokenInvalid
                                    clsToken.content = "&"
                                End If
                            Case Else:
                                clsToken.TType = TokenInvalid
                                clsToken.content = "&"
                                
                        End Select
                        
                    Case CHAR_QUOTE:
                        .Position = .Position + 1
                        Do While .Source(.Position) <> CHAR_QUOTE And _
                                 .Source(.Position) <> CHAR_CARRIAGE And _
                                 .Source(.Position) <> CHAR_LINEFEED
                        
                            clsToken.content = clsToken.content & ChrW$(.Source(.Position))
                            .Position = .Position + 1
                        Loop
                        clsToken.TType = TokenString
                    Case CHAR_AT
                        Do
                            clsToken.content = clsToken.content & ChrW$(.Source(.Position))
                            .Position = .Position + 1
                        Loop While IsAlphaNumeric(.Source(.Position))
                        Dim p
                            If IsNumberNew(Basestack1, Mid$(clsToken.content, 2), p, True) Then
                                If MemInt(VarPtr(p)) = vbString Then
                                    SwapVariant clsToken.content, p
                                    clsToken.TType = TokenString
                                Else
                                    clsToken.content = Trim(str$(CLng(p)))
                                    clsToken.TType = TokenValue
                                End If
                            Else
                                clsToken.TType = TokenInvalid
                            End If
                            .Position = .Position - 1
                        Case CHAR_AT, CHAR_ALPHA_UA To CHAR_ALPHA_UZ, CHAR_ALPHA_LA To CHAR_ALPHA_LZ:
                        Do
                            clsToken.content = clsToken.content & ChrW$(.Source(.Position))
                            .Position = .Position + 1
                        Loop While IsAlphaNumeric(.Source(.Position))
                        
                        If IsRegister(clsToken.content) Then
                            clsToken.TType = TokenRegister
                        ElseIf IsSegmentReg(clsToken.content) Then
                            clsToken.TType = TokenSegmentReg
                        ElseIf IsFPUReg(clsToken.content) Then
                            clsToken.TType = TokenFPUReg
                        ElseIf IsKeyword(clsToken.content) Then
                            clsToken.TType = TokenKeyword
                        ElseIf IsRawDataOp(clsToken.content) Then
                            clsToken.TType = TokenRawData
                        ElseIf IsMMReg(clsToken.content) Then
                            clsToken.TType = TokenMMRegister
                        ElseIf StrComp(clsToken.content, "extern", vbTextCompare) = 0 Then
                            clsToken.TType = TokenExtern
                            
                        Else
                            If GetlocalVar(clsToken.content, i) Then
                                'clsToken.content = Trim(str$(CLng(var(i))))
                                If MemInt(VarPtr(var(i))) = vbString Then
                                    clsToken.content = var(i)
                                
                                    clsToken.TType = TokenString
                                Else
                                    clsToken.content = Trim(str$(CLng(var(i))))
                                
                                    clsToken.TType = TokenValue
                                End If
                            Else
                                clsToken.TType = tokensymbol
                            End If
                         End If
                        
                        .Position = .Position - 1
                        
                    Case Else:
                        clsToken.TType = TokenInvalid
                        clsToken.content = ChrW$(.Source(.Position))
                        
                End Select
                
                If .Source(.Position) = CHAR_CARRIAGE Then
                    .LastWasEOI = True
                Else
                    If clsToken.TType <> TokenKeyword And _
                       clsToken.TType <> tokensymbol And _
                       .Source(.Position) <> CHAR_SPACE Then
                        .LastWasEOI = False
                    End If
                End If
                
                .Position = .Position + 1
            Loop
        
            If (clsToken.TType = TokenUnknown) And (.Position = .Length) Then
                clsToken.TType = TokenEndOfInput
            End If
        End With
    End If
    
    ' ** If the just scanned token is ":" then the previous token is a segment register
    '    or a label. In the last case the ":" must be followed by a line break.
    '    So add an invisible break (m_blnNextIsEOI).
    ' ** If the just scanned token is a string and the previous token is a line break
    '    or the begin of the input stream, then the current one probably is an operator.
    '    But it still can be followed by a colon and become a label in the next call to
    '    ScannerGetNextToken.
    ' ** ASMToken.Content kann contain hexadecimal or octal values.
    '    Supported is only VB style encoding (&H/&HO...)
    With clsToken
        If .TType = TokenOpColon Then
            If m_udtScanner.CurToken.TType = TokenOperator Then
                m_udtScanner.CurToken.TType = tokensymbol
                m_udtScanner.NextIsEOI = True
                m_udtScanner.LastWasEOI = False
            End If
        ElseIf .TType = tokensymbol Then
            If m_udtScanner.CurToken.TType = TokenEndOfInstruction Or _
               m_udtScanner.CurToken.TType = TokenBeginOfInput Or _
               m_udtScanner.LastWasEOI Then
                .TType = TokenOperator
                m_udtScanner.LastWasEOI = False
            End If
        ElseIf .TType = TokenValue Then
            On Error Resume Next
            .Value = CLng(.content)
            If Err Then .TType = TokenInvalid
            On Error GoTo 0
        End If

        .Line = m_udtScanner.Line
        .Section = m_udtScanner.Section
        .Position = (m_udtScanner.Position - Len(.content)) - m_udtScanner.LinePos + 1
    End With

    Set m_udtScanner.NextToken = clsToken
End Sub


Private Function IsDigit(ByVal bt As Byte) As Boolean
    IsDigit = (bt >= CHAR_NUMBER_0) And (bt <= CHAR_NUMBER_9)
End Function
Private Function IsOctChar(ByVal bt As Byte) As Boolean
    IsOctChar = (bt >= CHAR_NUMBER_0) And (bt <= CHAR_NUMBER_7)
End Function

Private Function IsHexChar(ByVal bt As Byte) As Boolean
    IsHexChar = ((bt >= CHAR_NUMBER_0) And (bt <= CHAR_NUMBER_9)) Or _
                ((bt >= CHAR_ALPHA_UA) And (bt <= CHAR_ALPHA_UA + 5)) Or _
                ((bt >= CHAR_ALPHA_LA) And (bt <= CHAR_ALPHA_LA + 5))
End Function


Private Function IsAlphaNumeric(ByVal bt As Byte) As Boolean
    IsAlphaNumeric = ((bt >= CHAR_NUMBER_0) And (bt <= CHAR_NUMBER_9)) Or _
                     ((bt >= CHAR_ALPHA_UA) And (bt <= CHAR_ALPHA_UZ)) Or _
                     ((bt >= CHAR_ALPHA_LA) And (bt <= CHAR_ALPHA_LZ)) Or _
                     (bt = CHAR_UNDERSCORE)
End Function


Private Function CalculatePECheckSum(ByVal ChecksumOffset As Long) As Long
    Dim lngSum      As Currency
    Dim i           As Long
    Dim lngWord     As Long
    Dim lngLength   As Long
    
    lngLength = m_lngOutSize
    lngSum = 0
    
    ' Loop through the file data in 2-byte (Word) blocks
    For i = 0 To lngLength - 1 Step 2
        ' Skip the 4 bytes of the CheckSum field position during calculation
        If i = ChecksumOffset Or i = ChecksumOffset + 2 Then
            lngWord = 0
        Else
            ' Read 2 bytes as a 16-bit unsigned word
            If i + 1 < lngLength Then
                lngWord = m_btOutput(i) + (CLng(m_btOutput(i + 1)) * 256&)
            Else
                lngWord = m_btOutput(i) ' Handle odd file length boundary if any
            End If
        End If
        
        lngSum = lngSum + lngWord
        
        ' Handle the 16-bit standard Internet Checksum carry fold-back
        If lngSum > 65535 Then
            lngSum = (lngSum And 65535) + Int(lngSum / 65536)
        End If
    Next i
    
    ' Final fold-back adjustments for any remaining carryover
    lngSum = (lngSum And 65535) + Int(lngSum / 65536)
    lngSum = (lngSum And 65535) + Int(lngSum / 65536)
    
    ' The official PE Checksum is the 16-bit word sum plus the total file size
    CalculatePECheckSum = CLng(lngSum) + lngLength
End Function
Public Sub AddExport(ByVal strFunctionName As String, ByVal strLabelName As String)
    ReDim Preserve m_udtExports(m_lngExportCount) As ASMExport
    m_udtExports(m_lngExportCount).Name = strFunctionName
    m_udtExports(m_lngExportCount).LabelName = strLabelName
    m_lngExportCount = m_lngExportCount + 1
End Sub

Public Property Get SizeOfStackReserve() As Long
SizeOfStackReserve = m_SizeOfStackReserve
End Property

Public Property Let SizeOfStackReserve(ByVal vNewValue As Long)
If vNewValue < &H100000 Then vNewValue = &H100000
m_SizeOfStackReserve = vNewValue
End Property
Public Property Get SizeOfStackCommit() As Long
SizeOfStackCommit = m_SizeOfStackCommit
End Property

Public Property Let SizeOfStackCommit(ByVal vNewValue As Long)
If vNewValue > m_SizeOfStackReserve Then vNewValue = m_SizeOfStackReserve
If vNewValue < &H1000 Then vNewValue = &H1000
m_SizeOfStackCommit = vNewValue
End Property
Public Property Get SizeOfHeapReserve() As Long
SizeOfHeapReserve = m_SizeOfHeapReserve
End Property

Public Property Let SizeOfHeapReserve(ByVal vNewValue As Long)
If vNewValue < &H100000 Then vNewValue = &H100000
m_SizeOfHeapReserve = vNewValue
End Property
Public Property Get SizeOfHeapCommit() As Long
SizeOfHeapCommit = m_SizeOfHeapCommit
End Property

Public Property Let SizeOfHeapCommit(ByVal vNewValue As Long)
If vNewValue > m_SizeOfHeapReserve Then vNewValue = m_SizeOfHeapReserve
If vNewValue < &H1000 Then vNewValue = &H1000
m_SizeOfHeapCommit = vNewValue
End Property
Public Property Get PEDll() As Boolean
    PEDll = m_blnWriteDLL
End Property

Public Property Let PEDll(ByVal blnValue As Boolean)
    m_blnWriteDLL = blnValue
    If blnValue Then m_blnWritePE = True
End Property

Public Property Get DLLName() As String
    DLLName = m_strDLLName
End Property

Public Property Let DLLName(ByVal strValue As String)
    m_strDLLName = strValue
End Property
