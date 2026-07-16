Attribute VB_Name = "ASMOpCodes"
Option Explicit


' X86 Assembler Instruktionen
'
' Arne Elster 2007 / 2008


' IMPORTANT FOR INSTRUCTIONS LIST:
'    FOR INSTRUCTIONS WITH A REL PARAM THE
'    ONE WITH THE BIGGEST REL VALUE GOES FIRST!
'    (for example see jmp instruction)
'
' TODO:
'   * for example PINSRW, rem32 => ptr needs DWORD keyword to be assembled,
'                                  but there aren't any other possibilities


Private Const MAX_PARAMETERS                As Long = 3
Private Const MAX_OPCODE_LEN                As Long = 4

Public Const PREFIX_LOCK                    As Long = &HF0
Public Const PREFIX_REPNE                   As Long = &HF2
Public Const PREFIX_REPNZ                   As Long = &HF2
Public Const PREFIX_REPE                    As Long = &HF3
Public Const PREFIX_REPZ                    As Long = &HF3
Public Const PREFIX_REP                     As Long = &HF3
Public Const PREFIX_SEGMENT_CS              As Long = &H2E
Public Const PREFIX_SEGMENT_SS              As Long = &H36
Public Const PREFIX_SEGMENT_DS              As Long = &H3E
Public Const PREFIX_SEGMENT_ES              As Long = &H26
Public Const PREFIX_SEGMENT_FS              As Long = &H64
Public Const PREFIX_SEGMENT_GS              As Long = &H65
Public Const PREFIX_BRANCH_TAKEN            As Long = &H2E
Public Const PREFIX_BRANCH_NOT_TAKEN        As Long = &H3E
Public Const PREFIX_OPERAND_SIZE_OVERRIDE   As Long = &H66
Public Const PREFIX_ADDRESS_SIZE_OVERRIDE   As Long = &H67

Public Const CONDITIONS     As String = _
    "A  B  C  E  G  L  S  Z  O  P  " & _
    "AE BE GE LE NA NB NC NE NG NL " & _
    "NO NP NS NZ PE PO NAE NBE NGE " & _
    "NLE"

Public Const SEGMENT_REGS   As String = _
    "CS DS ES FS GS SS"

Public Const Registers      As String = _
    "AL  BL  CL  DL  AH  BH  CH  DH " & _
    "AX  BX  CX  DX  BP  SP  DI  SI " & _
    "EAX EBX ECX EDX EBP ESP EDI ESI"

Public Const FPU_REGS       As String = _
    "ST0 ST1 ST2 ST3 ST4 ST5 ST6 ST7"

Public Const MM_REGS        As String = _
    "MM0 MM1 MM2 MM3 MM4 MM5 MM6 MM7 " & _
    "XMM0 XMM1 XMM2 XMM3 XMM4 XMM5 XMM6 XMM7"

Public Const KEYWORDS       As String = _
    "BYTE WORD DWORD QWORD DQWORD FLOAT " & _
    "DOUBLE EXTENDED " & _
    "LOCK REPNE REPNZ REPE REP"

Public Const RAW_DATA       As String = _
    "DB DW DD"

Public Enum OpCodePrefixes
    PrefixNone = &H0&
    PrefixFlgLock = &H1&
    PrefixFlgRepne = &H2&
    PrefixFlgRepnz = &H2&
    PrefixFlgRep = &H4&
    PrefixFlgRepe = &H4&
    PrefixFlgRepz = &H4&
    PrefixFlgSegmentCS = &H8&
    PrefixFlgSegmentSS = &H10&
    PrefixFlgSegmentDS = &H20&
    PrefixFlgSegmentES = &H40&
    PrefixFlgSegmentGS = &H80&
    PrefixFlgSegmentFS = &H100&
    PrefixFlgBranchNotTaken = &H200&
    PrefixFlgBranchTaken = &H400&
    PrefixFlgOperandSizeOverride = &H800&
    PrefixFlgAddressSizeOverride = &H1000&
End Enum

Public Enum TokenType
    TokenUnknown
    TokenBeginOfInput
    TokenOperator
    TokenKeyword
    TokenValue
    TokenSymbol
    TokenFPUReg
    TokenSegmentReg
    TokenRegister
    TokenMMRegister
    TokenSeparator
    TokenString
    TokenBracketLeft
    TokenBracketRight
    TokenOpAdd
    TokenOpSub
    TokenOpMul
    TokenOpColon
    TokenRawData
    TokenExtern
    TokenEndOfInstruction
    TokenEndOfInput
    TokenInvalid
End Enum

Public Enum ParamType
    ParamTypeUnknown = &H0
    ParamReg = &H1
    ParamRel = &H2
    ParamMem = &H4
    ParamImm = &H8
    ParamSTX = &H10
    ParamMM = &H20
    ParamExt = &H40
End Enum

Public Enum ParamSize
    BitsUnknown = 0
    Bits8 = 8
    Bits16 = 16
    Bits32 = 32
    Bits64 = 64
    Bits80 = 80
    Bits128 = 128
End Enum

Private Enum ExtType
    ExtNon = 0
    ExtFlt
    ExtReg
    ExtCon
    Ext3DN
End Enum

Private Enum SizeMod
    SizeModOvrd
    SizeModNone
End Enum

Public Enum ASMRegisters
    RegUnknown = &H0&
    RegAL = &H1&
    RegBL = &H2&
    RegCL = &H4&
    RegDL = &H8&
    RegAH = &H10&
    RegBH = &H20&
    RegCH = &H40&
    RegDH = &H80&
    RegAX = &H100&
    RegBX = &H200&
    RegCX = &H400&
    RegDX = &H800&
    RegBP = &H1000&
    RegSP = &H2000&
    RegDI = &H4000&
    RegSI = &H8000&
    RegEAX = &H10000
    RegEBX = &H20000
    RegECX = &H40000
    RegEDX = &H80000
    RegEBP = &H100000
    RegESP = &H200000
    RegEDI = &H400000
    RegESI = &H800000
End Enum

Public Enum ASMSegmentRegs
    SegUnknown = &H0
    SegCS = &H1
    SegDS = &H2
    SegES = &H4
    SegFS = &H8
    SegGS = &H10
    SegSS = &H20
End Enum

Public Enum ASMFPURegisters
    FP_UNKNOWN = -1
    FP_ST0 = 0
    FP_ST1
    FP_ST2
    FP_ST3
    FP_ST4
    FP_ST5
    FP_ST6
    FP_ST7
End Enum

Public Enum ASMXMMRegisters
    MM_Unknown = -1
    MM0 = &H1&
    MM1 = &H2&
    MM2 = &H4&
    MM3 = &H8&
    MM4 = &H10&
    MM5 = &H20&
    MM6 = &H40&
    MM7 = &H80&
    XMM0 = &H100&
    XMM1 = &H200&
    XMM2 = &H400&
    XMM3 = &H800&
    XMM4 = &H1000&
    XMM5 = &H2000&
    XMM6 = &H4000&
    XMM7 = &H8000&
End Enum

Private Type OpCode
    Bytes(MAX_OPCODE_LEN - 1)       As Byte
    ByteCount                       As Long
    RegOpExt                        As Long
End Type

' If Forced = True the parameter has a predefined (FPU) register
' or a numerical value. In this case it mustn't be assembled!
Public Type InstructionParam
    PType                           As ParamType
    size                            As ParamSize
    Register                        As ASMRegisters
    FPURegister                     As ASMFPURegisters
    MMRegister                      As ASMXMMRegisters
    Value                           As Long
    Forced                          As Boolean
End Type

Public Type Instruction
    Mnemonic                        As String
    Prefixes                        As OpCodePrefixes
    OpCode(MAX_OPCODE_LEN - 1)      As Byte
    OpCodeLen                       As Long
    RegOpExt                        As Long
    ModRM                           As Boolean
    Parameters(MAX_PARAMETERS - 1)  As InstructionParam
    ParamCount                      As Long
    Now3DByte                       As Long
End Type

Private m_strConditions()           As String
Private m_strSegments()             As String
Private m_strRegisters()            As String
Private m_strFPURegs()              As String
Private m_strMMRegs()               As String
Private m_strKeywords()             As String
Private m_strRawData()              As String

Public Instructions()               As Instruction
Public InstructionCount             As Long

Private m_blnInit                   As Boolean


Private Sub AddInstr(inst As Instruction)
    ReDim Preserve Instructions(InstructionCount) As Instruction
    Instructions(InstructionCount) = inst
    InstructionCount = InstructionCount + 1
End Sub


Public Sub InitInstructions()
    If Not m_blnInit Then
        m_strConditions = Split(RemoveWSDoubles(CONDITIONS), " ")
        m_strSegments = Split(RemoveWSDoubles(SEGMENT_REGS), " ")
        m_strRegisters = Split(RemoveWSDoubles(Registers), " ")
        m_strFPURegs = Split(RemoveWSDoubles(FPU_REGS), " ")
        m_strKeywords = Split(RemoveWSDoubles(KEYWORDS), " ")
        m_strRawData = Split(RemoveWSDoubles(RAW_DATA), " ")
        m_strMMRegs = Split(RemoveWSDoubles(MM_REGS), " ")

        AddArithmetics
        AddFlowCtrl
        AddMemory
        AddBits
        AddFloatingPoint
        AddOthers
        AddMMXSSE
        Add3DNow
        
        m_blnInit = True
    End If
End Sub


Private Sub Add3DNow()
    ' a prefix for immediates which marks that they should stay
    ' invisible until writing the output would be better
    Instruction "FEMMS      ", "0F 0E   ", SizeModNone, ExtNon
    Instruction "PAVGUSB    ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HBF"
    Instruction "PF2ID      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&H1D"
    Instruction "PFACC      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HAE"
    Instruction "PFADD      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&H9E"
    Instruction "PFCMPEQ    ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HB0"
    Instruction "PFCMPGE    ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&H90"
    Instruction "PFCMPGT    ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HA0"
    Instruction "PFMAX      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HA4"
    Instruction "PFMIN      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&H94"
    Instruction "PFMUL      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HB4"
    Instruction "PFRCP      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&H96"
    Instruction "PFRCPIT1   ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HA6"
    Instruction "PFRCPIT2   ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HB6"
    Instruction "PFRSQIT1   ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HA7"
    Instruction "PFRSQRT    ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&H97"
    Instruction "PFSUB      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&H9A"
    Instruction "PFSUBR     ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HAA"
    Instruction "PI2FD      ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&H0D"
    Instruction "PMULHRW    ", "0F 0F   ", SizeModNone, Ext3DN, "mm", "mm/mem", "&HB7"
    Instruction "PREFETCH   ", "0F 0D /0", SizeModNone, ExtNon, "mem"
    Instruction "PREFETCHW  ", "0F 0D /1", SizeModNone, ExtNon, "mem"
End Sub


Private Sub AddMMXSSE()
    Instruction "ADDPD      ", "66 0F 58    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "ADDPS      ", "0F 58       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "ADDSD      ", "F2 0F 58    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "ADDSS      ", "F3 0F 58    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "ADDSUBPD   ", "66 0F D0    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "ADDSUBPS   ", "F2 0F D0    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "ANDPD      ", "66 0F 54    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "ANDPS      ", "0F 54       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "ANDNPD     ", "66 0F 55    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "ANDNPS     ", "0F 55       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "CMPPD      ", "66 0F C2    ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "CMPPS      ", "0F C2       ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "CMPSD      ", "F2 0F C2    ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "CMPSS      ", "F3 0F C2    ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"

    Instruction "COMISD     ", "66 0F 2F    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "COMISS     ", "0F 2F       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "CVTDQ2PD   ", "F3 0F E6    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTDQ2PS   ", "0F 5B       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTPD2DQ   ", "F2 0F E6    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTPD2PI   ", "66 0F 2D    ", SizeModNone, ExtNon, "mm", "xmm/mem"
    Instruction "CVTPD2PS   ", "66 0F 5A    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTPI2PD   ", "66 0F 2A    ", SizeModNone, ExtNon, "xmm", "mm/mem"
    Instruction "CVTPI2PS   ", "0F 2A       ", SizeModNone, ExtNon, "xmm", "mm/mem"
    Instruction "CVTPS2DQ   ", "66 0F 5B    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTPS2PD   ", "0F 5A       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTPS2PI   ", "0F 2D       ", SizeModNone, ExtNon, "mm", "xmm/mem"
    Instruction "CVTSD2SI   ", "F2 0F 2D    ", SizeModNone, ExtNon, "reg32", "xmm/mem"
    Instruction "CVTSD2SS   ", "F2 0F 5A    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTSI2SD   ", "F2 0F 2A    ", SizeModNone, ExtNon, "xmm", "rem32"
    Instruction "CVTSI2SS   ", "F3 0F 2A    ", SizeModNone, ExtNon, "xmm", "rem32"
    Instruction "CVTSS2SD   ", "F3 0F 5A    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTSS2SI   ", "F3 0F 2D    ", SizeModNone, ExtNon, "reg32", "xmm/mem"
    Instruction "CVTTPD2PI  ", "66 0F 2C    ", SizeModNone, ExtNon, "mm", "xmm/mem"
    Instruction "CVTTPD2DQ  ", "66 0F E6    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTTPS2DQ  ", "F3 0F 5B    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "CVTTPS2PI  ", "0F 2C       ", SizeModNone, ExtNon, "mm", "xmm/mem"
    Instruction "CVTTSD2SI  ", "F2 0F 2C    ", SizeModNone, ExtNon, "reg32", "xmm/mem"
    Instruction "CVTTSS2SI  ", "F3 0F 2C    ", SizeModNone, ExtNon, "reg32", "xmm/mem"
    
    Instruction "DIVPD      ", "66 0F 5E    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "DIVPS      ", "0F 5E       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "DIVSD      ", "F2 0F 5E    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "DIVSS      ", "F3 0F 5E    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "EMMS       ", "0F 77       ", SizeModNone, ExtNon
    
    Instruction "HADDPD     ", "66 0F 7C    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "HADDPS     ", "F2 0F 7C    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "HSUBPD     ", "66 0F 7D    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "HSUBPS     ", "F2 0F 7D    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "LDDQU      ", "F2 0F F0    ", SizeModNone, ExtNon, "xmm", "mem    "
    Instruction "LDMXCSR    ", "0F AE /2    ", SizeModNone, ExtNon, "mem"
    Instruction "MASKMOVDQU ", "66 0F F7    ", SizeModNone, ExtNon, "xmm", "xmm"
    Instruction "MASKMOVQ   ", "0F F7       ", SizeModNone, ExtNon, "mm", "mm"
    
    Instruction "MAXPD      ", "66 0F 5F    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MAXPS      ", "0F 5F       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MAXSD      ", "F2 0F 5F    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MAXSS      ", "F3 0F 5F    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MINPD      ", "66 0F 5D    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MINPS      ", "0F 5D       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MINSD      ", "F2 0F 5D    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MINSS      ", "F3 0F 5D    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "MOVAPD     ", "66 0F 28    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVAPS     ", "0F 28       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVD       ", "0F 6E       ", SizeModNone, ExtNon, "mm ", "rem32"
    Instruction "MOVD       ", "0F 7E       ", SizeModNone, ExtNon, "rem32", "mm "
    Instruction "MOVD       ", "66 0F 6E    ", SizeModNone, ExtNon, "xmm", "rem32"
    Instruction "MOVD       ", "66 0F 7E    ", SizeModNone, ExtNon, "rem32", "xmm"
    
    Instruction "MOVDDUP    ", "F2 0F 12    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVDQA     ", "66 0F 6F    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVDQA     ", "66 0F 7F    ", SizeModNone, ExtNon, "xmm/mem", "xmm"
    Instruction "MOVDQU     ", "F3 0F 6F    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVDQU     ", "F3 0F 7F    ", SizeModNone, ExtNon, "xmm/mem", "xmm"
    Instruction "MOVDQ2Q    ", "F2 0F D6    ", SizeModNone, ExtNon, "mm", "xmm"
    Instruction "MOVHLPS    ", "0F 12       ", SizeModNone, ExtNon, "xmm", "xmm"
    Instruction "MOVHPD     ", "66 0F 16    ", SizeModNone, ExtNon, "xmm", "mem"
    Instruction "MOVHPD     ", "66 0F 17    ", SizeModNone, ExtNon, "mem", "xmm"
    Instruction "MOVHPS     ", "0F 16       ", SizeModNone, ExtNon, "xmm", "mem"
    Instruction "MOVHPS     ", "0F 17       ", SizeModNone, ExtNon, "mem", "xmm"
    Instruction "MOVLHPS    ", "0F 16       ", SizeModNone, ExtNon, "xmm", "xmm"
    Instruction "MOVLPD     ", "66 0F 12    ", SizeModNone, ExtNon, "xmm", "mem"
    Instruction "MOVLPD     ", "66 0F 13    ", SizeModNone, ExtNon, "mem", "xmm"
    Instruction "MOVLPS     ", "0F 12       ", SizeModNone, ExtNon, "xmm", "mem"
    Instruction "MOVLPS     ", "0F 13       ", SizeModNone, ExtNon, "mem", "xmm"
    
    Instruction "MOVMSKPD   ", "66 0F 50    ", SizeModNone, ExtNon, "reg32", "xmm"
    Instruction "MOVMSKPS   ", "0F 50       ", SizeModNone, ExtNon, "reg32", "xmm"
    Instruction "MOVNTDQ    ", "66 0F E7    ", SizeModNone, ExtNon, "mem", "xmm"
    Instruction "MOVNTI     ", "0F C3       ", SizeModNone, ExtNon, "mem", "reg32"
    Instruction "MOVNTPD    ", "66 0F 2B    ", SizeModNone, ExtNon, "mem", "xmm"
    Instruction "MOVNTPS    ", "0F 2B       ", SizeModNone, ExtNon, "mem", "xmm"
    Instruction "MOVNTQ     ", "0F E7       ", SizeModNone, ExtNon, "mem", "mm"
    Instruction "MOVQ       ", "0F 6F       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "MOVQ       ", "0F 7F       ", SizeModNone, ExtNon, "mm/mem", "mm"
    Instruction "MOVQ       ", "F3 0F 7E    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVQ       ", "66 0F D6    ", SizeModNone, ExtNon, "xmm/mem", "xmm"
    Instruction "MOVQ2DQ    ", "F3 0F D6    ", SizeModNone, ExtNon, "xmm", "mm"
    Instruction "MOVSD      ", "F2 0F 10    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVSD      ", "F2 0F 11    ", SizeModNone, ExtNon, "xmm/mem", "xmm"
    Instruction "MOVSHDUP   ", "F3 0F 16    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVSLDUP   ", "F3 0F 12    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVSS      ", "F3 0F 10    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVSS      ", "F3 0F 11    ", SizeModNone, ExtNon, "xmm/mem", "xmm"
    Instruction "MOVUPD     ", "66 0F 10    ", SizeModNone, ExtNon, "xmm", "mm/mem"
    Instruction "MOVUPD     ", "66 0F 11    ", SizeModNone, ExtNon, "xmm/mem", "xmm"
    Instruction "MOVUPS     ", "0F 10       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MOVUPS     ", "0F 11       ", SizeModNone, ExtNon, "xmm/mem", "xmm"
    Instruction "MULPD      ", "66 0F 59    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MULPS      ", "0F 59       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MULSD      ", "F2 0F 59    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "MULSS      ", "F3 0F 59    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "ORPD       ", "66 0F 56    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "ORPS       ", "0F 56       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PABSB      ", "0F 38 1C    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PABSB      ", "66 0F 38 1C ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PABSW      ", "0F 38 1D    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PABSW      ", "66 0F 38 1D ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PABSD      ", "0F 38 1E    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PABSD      ", "66 0F 38 1E ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PACKSSWB   ", "0F 63       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PACKSSWB   ", "66 0F 63    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PACKSSDW   ", "0F 6B       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PACKSSDW   ", "66 0F 6B    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PACKUSWB   ", "0F 67       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PACKUSWB   ", "66 0F 67    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PADDB      ", "0F FC       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PADDB      ", "66 0F FC    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PADDW      ", "0F FD       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PADDW      ", "66 0F FD    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PADDD      ", "0F FE       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PADDD      ", "66 0F FE    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PADDQ      ", "0F D4       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PADDQ      ", "66 0F D4    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PADDSB     ", "0F EC       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PADDSB     ", "66 0F EC    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PADDSW     ", "0F ED       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PADDSW     ", "66 0F ED    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PADDUSB    ", "0F DC       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PADDUSB    ", "66 0F DC    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PADDUSW    ", "0F DD       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PADDUSW    ", "66 0F DD    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PALIGNR    ", "0F 3A 0F    ", SizeModNone, ExtNon, "mm", "mm/mem", "imm08"
    Instruction "PALIGNR    ", "66 0F 3A 0F ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "PAND       ", "0F DB       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PAND       ", "66 0F DB    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PANDN      ", "0F DF       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PANDN      ", "66 0F DF    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PAVGB      ", "0F E0       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PAVGB      ", "66 0F E0    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PAVGW      ", "0F E3       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PAVGW      ", "66 0F E3    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PCMPEQB    ", "0F 74       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PCMPEQB    ", "66 0F 74    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PCMPEQW    ", "0F 75       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PCMPEQW    ", "66 0F 75    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PCMPEQD    ", "0F 76       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PCMPEQD    ", "66 0F 76    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PCMPGTB    ", "0F 64       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PCMPGTB    ", "66 0F 64    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PCMPGTW    ", "0F 65       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PCMPGTW    ", "66 0F 65    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PCMPGTD    ", "0F 66       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PCMPGTD    ", "66 0F 66    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PEXTRW     ", "0F C5       ", SizeModNone, ExtNon, "reg32", "mm", "imm08"
    Instruction "PEXTRW     ", "66 0F C5    ", SizeModNone, ExtNon, "reg32", "xmm", "imm08"
    Instruction "PHADDW     ", "0F 38 01    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PHADDW     ", "66 0F 38 01 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PHADDD     ", "0F 38 02    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PHADDD     ", "66 0F 38 02 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PHADDSW    ", "0F 38 03    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PHADDSW    ", "66 0F 38 03 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "PHSUBW     ", "0F 38 05    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PHSUBW     ", "66 0F 38 05 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PHSUBD     ", "0F 38 06    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PHSUBD     ", "66 0F 38 06 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PHSUBSW    ", "0F 38 07    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PHSUBSW    ", "66 0F 38 07 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PINSRW     ", "0F C4       ", SizeModNone, ExtNon, "mm", "rem32", "imm08"
    Instruction "PINSRW     ", "66 0F C4    ", SizeModNone, ExtNon, "xmm", "rem32", "imm08"
    Instruction "PMADDUBSW  ", "0F 38 04    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMADDUBSW  ", "66 0F 38 04 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMADDWD    ", "0F F5       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMADDWD    ", "66 0F F5    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMAXSW     ", "0F EE       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMAXSW     ", "66 0F EE    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMAXUB     ", "0F DE       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMAXUB     ", "66 0F DE    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMINSW     ", "0F EA       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMINSW     ", "66 0F EA    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMINUB     ", "0F DA       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMINUB     ", "66 0F DA    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMOVMSKB   ", "0F D7       ", SizeModNone, ExtNon, "reg32", "mm"
    Instruction "PMOVMSKB   ", "66 0F D7    ", SizeModNone, ExtNon, "reg32", "xmm"
    Instruction "PMULHRSW   ", "0F 38 0B    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMULHRSW   ", "66 0F 38 0B ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMULHUW    ", "0F E4       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMULHUW    ", "66 0F E4    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMULHW     ", "0F E5       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMULHW     ", "66 0F E5    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMULLW     ", "0F D5       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMULLW     ", "66 0F D5    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PMULUDQ    ", "0F F4       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PMULUDQ    ", "66 OF F4    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "POR        ", "0F EB       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "POR        ", "66 0F EB    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSADBW     ", "0F F6       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSADBW     ", "66 0F F6    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSHUFB     ", "0F 38 00    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSHUFB     ", "66 0F 38 00 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSHUFD     ", "66 0F 70    ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "PSHUFHW    ", "F3 0F 70    ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "PSHUFLW    ", "F2 0F 70    ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "PSHUFW     ", "0F 70       ", SizeModNone, ExtNon, "mm", "mm/mem", "imm08"
    
    Instruction "PSIGNB     ", "0F 38 08    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSIGNB     ", "66 0F 38 08 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSIGNW     ", "0F 38 09    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSIGNW     ", "66 0F 38 09 ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSIGND     ", "0F 38 0A    ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSIGND     ", "66 0F 38 0A ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSLLDQ     ", "66 0F 73 /7 ", SizeModNone, ExtNon, "xmm", "imm08"
    
    Instruction "PSLLW      ", "0F F1       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSLLW      ", "66 0F F1    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSLLW      ", "0F 71 /6    ", SizeModNone, ExtNon, "mm", "imm08"
    Instruction "PSLLW      ", "66 0F 71 /6 ", SizeModNone, ExtNon, "xmm", "imm08"
    Instruction "PSLLD      ", "0F F2       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSLLD      ", "66 0F F2    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSLLD      ", "0F 72 /6    ", SizeModNone, ExtNon, "mm", "imm08"
    Instruction "PSLLD      ", "66 0F 72 /6 ", SizeModNone, ExtNon, "xmm", "imm08"
    Instruction "PSLLQ      ", "0F F3       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSLLQ      ", "66 0F F3    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSLLQ      ", "0F 73 /6    ", SizeModNone, ExtNon, "mm", "imm08"
    Instruction "PSLLQ      ", "66 0F 73 /6 ", SizeModNone, ExtNon, "xmm", "imm08"
    
    Instruction "PSRAW      ", "0F E1       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSRAW      ", "66 0F E1    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSRAW      ", "0F 71 /4    ", SizeModNone, ExtNon, "mm", "imm08"
    Instruction "PSRAW      ", "66 0F 71 /4 ", SizeModNone, ExtNon, "xmm", "imm08"
    Instruction "PSRAD      ", "0F E2       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSRAD      ", "66 0F E2    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSRAD      ", "0F 72 /4    ", SizeModNone, ExtNon, "mm", "imm08"
    Instruction "PSRAD      ", "66 0F 72 /4 ", SizeModNone, ExtNon, "xmm", "imm08"
    Instruction "PSRLDQ     ", "66 0F 73 /3 ", SizeModNone, ExtNon, "xmm", "imm08"
    
    Instruction "PSRLW      ", "0F D1       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSRLW      ", "66 0F D1    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSRLW      ", "0F 71 /2    ", SizeModNone, ExtNon, "mm", "imm08"
    Instruction "PSRLW      ", "66 0F 71 /2 ", SizeModNone, ExtNon, "xmm", "imm08"
    Instruction "PSRLD      ", "0F D2       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSRLD      ", "66 0F D2    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSRLD      ", "0F 72 /2    ", SizeModNone, ExtNon, "mm", "imm08"
    Instruction "PSRLD      ", "66 0F 72 /2 ", SizeModNone, ExtNon, "xmm", "imm08"
    Instruction "PSRLQ      ", "0F D3       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSRLQ      ", "66 0F D3    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSRLQ      ", "0F 73 /2    ", SizeModNone, ExtNon, "mm", "imm08"
    Instruction "PSRLQ      ", "66 0F 73 /2 ", SizeModNone, ExtNon, "xmm", "imm08"
    
    Instruction "PSUBB      ", "0F F8       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSUBB      ", "66 0F F8    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSUBW      ", "0F F9       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSUBW      ", "66 0F F9    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSUBD      ", "0F FA       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSUBD      ", "66 0F FA    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSUBQ      ", "0F FB       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSUBQ      ", "66 0F FB    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "PSUBSB     ", "0F E8       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSUBSB     ", "66 0F E8    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSUBSW     ", "0F E9       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSUBSW     ", "66 0F E9    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "PSUBUSB    ", "0F D8       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSUBUSB    ", "66 0F D8    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PSUBUSW    ", "0F D9       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PSUBUSW    ", "66 0F D9    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PUNPCKHBW  ", "0F 68       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PUNPCKHBW  ", "66 0F 68    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PUNPCKHWD  ", "0F 69       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PUNPCKHWD  ", "66 0F 69    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PUNPCKHDQ  ", "0F 6A       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PUNPCKHDQ  ", "66 0F 6A    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PUNPCKHQDQ ", "66 0F 6D    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "PUNPCKLBW  ", "0F 60       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PUNPCKLBW  ", "66 0F 60    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PUNPCKLWD  ", "0F 61       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PUNPCKLWD  ", "66 0F 61    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PUNPCKLDQ  ", "0F 62       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PUNPCKLDQ  ", "66 0F 62    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "PUNPCKLQDQ ", "66 0F 6C    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "PXOR       ", "0F EF       ", SizeModNone, ExtNon, "mm", "mm/mem"
    Instruction "PXOR       ", "66 0F EF    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    
    Instruction "RCPPS      ", "0F 53       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "RCPSS      ", "F3 0F 53    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "RSQRTPS    ", "0F 52       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "RSQRTSS    ", "F3 0F 52    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "SHUFPD     ", "66 0F C6    ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "SHUFPS     ", "0F C6       ", SizeModNone, ExtNon, "xmm", "xmm/mem", "imm08"
    Instruction "SQRTPS     ", "0F 51       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "SQRTSD     ", "F2 0F 51    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "SQRTSS     ", "F3 0F 51    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "SUBPD      ", "66 0F 5C    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "SUBPS      ", "0F 5C       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "SUBSD      ", "F2 0F 5C    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "SUBSS      ", "F3 0F 5C    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "UCOMISD    ", "66 0F 2E    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "UCOMISS    ", "0F 2E       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "UNPCKHPD   ", "66 0F 15    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "UNPCKHPS   ", "0F 15       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "UNPCKLPD   ", "66 0F 14    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "UNPCKLPS   ", "0F 14       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "XORPD      ", "66 0F 57    ", SizeModNone, ExtNon, "xmm", "xmm/mem"
    Instruction "XORPS      ", "0F 57       ", SizeModNone, ExtNon, "xmm", "xmm/mem"
End Sub


Private Sub AddOthers()
    Instruction "AAA    ", "37      ", SizeModNone, ExtNon
    Instruction "AAD    ", "D5 0A   ", SizeModNone, ExtNon
    Instruction "AAM    ", "D4 0A   ", SizeModNone, ExtNon
    Instruction "AAS    ", "3F      ", SizeModNone, ExtNon
    Instruction "DAA    ", "27      ", SizeModNone, ExtNon
    Instruction "DAS    ", "2F      ", SizeModNone, ExtNon

    Instruction "BOUND  ", "62      ", SizeModOvrd, ExtNon, "reg16", "mem16"
    Instruction "BOUND  ", "62      ", SizeModNone, ExtNon, "reg32", "mem32"

    Instruction "LAHF   ", "9F      ", SizeModNone, ExtNon
    Instruction "SAHF   ", "9E      ", SizeModNone, ExtNon
    Instruction "LAR    ", "0F 02   ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "LAR    ", "0F 02   ", SizeModNone, ExtNon, "reg32", "rem32"

    Instruction "LEA    ", "8D      ", SizeModOvrd, ExtNon, "reg16", "mem  "
    Instruction "LEA    ", "8D      ", SizeModNone, ExtNon, "reg32", "mem  "

    Instruction "BSWAP  ", "0F C8   ", SizeModNone, ExtReg, "#32"

    Instruction "SCASB  ", "AE      ", SizeModNone, ExtNon
    Instruction "SCASW  ", "AF      ", SizeModOvrd, ExtNon
    Instruction "SCASD  ", "AF      ", SizeModNone, ExtNon
    
    Instruction "SMSW   ", "0F 01 /4", SizeModOvrd, ExtNon, "rem16"
    Instruction "SMSW   ", "0F 01 /4", SizeModNone, ExtNon, "rem32"
    
    Instruction "STC    ", "F9      ", SizeModNone, ExtNon
    Instruction "STD    ", "FD      ", SizeModNone, ExtNon
    Instruction "STI    ", "FB      ", SizeModNone, ExtNon
    
    Instruction "CBW    ", "98      ", SizeModOvrd, ExtNon
    Instruction "CWWDE  ", "98      ", SizeModNone, ExtNon
    Instruction "CLC    ", "F8      ", SizeModNone, ExtNon
    Instruction "CLD    ", "FC      ", SizeModNone, ExtNon
    Instruction "CMC    ", "F5      ", SizeModNone, ExtNon
    Instruction "CWD    ", "99      ", SizeModOvrd, ExtNon
    Instruction "CDQ    ", "99      ", SizeModNone, ExtNon
    
    Instruction "CPUID  ", "0F A2   ", SizeModNone, ExtNon

    Instruction "RDMSR  ", "0F 32   ", SizeModNone, ExtNon
    Instruction "RDPMC  ", "0F 33   ", SizeModNone, ExtNon
    Instruction "RDTSC  ", "0F 31   ", SizeModNone, ExtNon
End Sub


Private Sub AddMemory()
    Instruction "LODSB  ", "AC      ", SizeModNone, ExtNon
    Instruction "LODSW  ", "AD      ", SizeModOvrd, ExtNon
    Instruction "LODSD  ", "AD      ", SizeModNone, ExtNon
    
    Instruction "STOSB  ", "AA      ", SizeModNone, ExtNon
    Instruction "STOSW  ", "AB      ", SizeModOvrd, ExtNon
    Instruction "STOSD  ", "AB      ", SizeModNone, ExtNon

    Instruction "PUSH   ", "50      ", SizeModOvrd, ExtReg, "#16  "
    Instruction "PUSH   ", "50      ", SizeModNone, ExtReg, "#32  "
    Instruction "PUSH   ", "FF /6   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "PUSH   ", "FF /6   ", SizeModNone, ExtNon, "rem32"
    Instruction "PUSH   ", "6A      ", SizeModNone, ExtNon, "imm08"
    Instruction "PUSH   ", "68      ", SizeModOvrd, ExtNon, "imm16"
    Instruction "PUSH   ", "68      ", SizeModNone, ExtNon, "imm32"
    ' not segment pushs
    Instruction "PUSHA  ", "60      ", SizeModOvrd, ExtNon
    Instruction "PUSHAD ", "60      ", SizeModNone, ExtNon
    Instruction "PUSHF  ", "9C      ", SizeModOvrd, ExtNon
    Instruction "PUSHFD ", "9C      ", SizeModNone, ExtNon
    
    Instruction "POP    ", "58      ", SizeModOvrd, ExtReg, "#16  "
    Instruction "POP    ", "58      ", SizeModNone, ExtReg, "#32  "
    Instruction "POP    ", "8F /0   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "POP    ", "8F /0   ", SizeModNone, ExtNon, "rem32"
    ' no segment pops
    Instruction "POPA   ", "61      ", SizeModOvrd, ExtNon
    Instruction "POPAD  ", "61      ", SizeModNone, ExtNon
    Instruction "POPF   ", "9D      ", SizeModOvrd, ExtNon
    Instruction "POPFD  ", "9D      ", SizeModNone, ExtNon

    Instruction "XLAT   ", "D7      ", SizeModNone, ExtNon
    Instruction "XLATB  ", "D7      ", SizeModNone, ExtNon
    
    Instruction "XCHG   ", "90      ", SizeModOvrd, ExtReg, "AX   ", "#16  "
    Instruction "XCHG   ", "90      ", SizeModOvrd, ExtReg, "#16  ", "AX   "
    Instruction "XCHG   ", "90      ", SizeModNone, ExtReg, "EAX  ", "#32  "
    Instruction "XCHG   ", "90      ", SizeModNone, ExtReg, "#32  ", "EAX  "
    Instruction "XCHG   ", "86      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "XCHG   ", "86      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "XCHG   ", "87      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "XCHG   ", "87      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "XCHG   ", "87      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "XCHG   ", "87      ", SizeModNone, ExtNon, "reg32", "rem32"

    Instruction "MOV    ", "A0      ", SizeModNone, ExtNon, "AL   ", "mem08"
    Instruction "MOV    ", "A1      ", SizeModOvrd, ExtNon, "AX   ", "mem16"
    Instruction "MOV    ", "A1      ", SizeModNone, ExtNon, "EAX  ", "mem32"
    Instruction "MOV    ", "A2      ", SizeModNone, ExtNon, "mem08", "AL   "
    Instruction "MOV    ", "A3      ", SizeModOvrd, ExtNon, "mem16", "AX   "
    Instruction "MOV    ", "A3      ", SizeModNone, ExtNon, "mem32", "EAX  "
    Instruction "MOV    ", "B0      ", SizeModNone, ExtReg, "#08  ", "imm08"
    Instruction "MOV    ", "B8      ", SizeModOvrd, ExtReg, "#16  ", "imm16"
    Instruction "MOV    ", "B8      ", SizeModNone, ExtReg, "#32  ", "imm32"
    Instruction "MOV    ", "88      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "MOV    ", "89      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "MOV    ", "89      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "MOV    ", "8A      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "MOV    ", "8B      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "MOV    ", "8B      ", SizeModNone, ExtNon, "reg32", "rem32"
    Instruction "MOV    ", "C6 /0   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "MOV    ", "C7 /0   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "MOV    ", "C7 /0   ", SizeModNone, ExtNon, "rem32", "imm32"
    
    Instruction "MOVZX  ", "0F B6   ", SizeModOvrd, ExtNon, "reg16", "rem08"
    Instruction "MOVZX  ", "0F B6   ", SizeModNone, ExtNon, "reg32", "rem08"
    Instruction "MOVZX  ", "0F B7   ", SizeModNone, ExtNon, "reg32", "rem16"
    
    Instruction "MOVSX  ", "0F BE   ", SizeModOvrd, ExtNon, "reg16", "rem08"
    Instruction "MOVSX  ", "0F BE   ", SizeModNone, ExtNon, "reg32", "rem08"
    Instruction "MOVSX  ", "0F BF   ", SizeModNone, ExtNon, "reg32", "rem16"
End Sub


Private Sub AddFlowCtrl()
    Instruction "CMP    ", "3C      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "CMP    ", "3D      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "CMP    ", "3D      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "CMP    ", "80 /7   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "CMP    ", "81 /7   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "CMP    ", "81 /7   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "CMP    ", "83 /7   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "CMP    ", "83 /7   ", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "CMP    ", "38      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "CMP    ", "39      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "CMP    ", "39      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "CMP    ", "3A      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "CMP    ", "3B      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "CMP    ", "3B      ", SizeModNone, ExtNon, "reg32", "rem32"

    Instruction "CMOV   ", "0F 40   ", SizeModOvrd, ExtCon, "reg16", "rem16"
    Instruction "CMOV   ", "0F 40   ", SizeModNone, ExtCon, "reg32", "rem32"
    
    Instruction "CMPSB  ", "A6      ", SizeModNone, ExtNon
    Instruction "CMPSW  ", "A7      ", SizeModOvrd, ExtNon
    Instruction "CMPSD  ", "A7      ", SizeModNone, ExtNon
    
    Instruction "CMPXCHG", "0F B0   ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "CMPXCHG", "0F B1   ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "CMPXCHG", "0F B1   ", SizeModNone, ExtNon, "rem32", "reg32"
    
    Instruction "CMPXCHG8B", "0F C7 /1", SizeModNone, ExtNon, "mem"
    
    Instruction "J      ", "0F 80   ", SizeModNone, ExtCon, "rel32"
    Instruction "J      ", "0F 80   ", SizeModOvrd, ExtCon, "rel16"
    Instruction "J      ", "70      ", SizeModNone, ExtCon, "rel08"
    Instruction "JCXZ   ", "E3      ", SizeModOvrd, ExtNon, "rel08"
    Instruction "JECXZ  ", "E3      ", SizeModNone, ExtNon, "rel08"
    
    Instruction "JMP    ", "E9      ", SizeModNone, ExtNon, "rel32"
    Instruction "JMP    ", "E9      ", SizeModOvrd, ExtNon, "rel16"
    Instruction "JMP    ", "EB      ", SizeModNone, ExtNon, "rel08"
    Instruction "JMP    ", "FF /4   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "JMP    ", "FF /4   ", SizeModNone, ExtNon, "rem32"
    ' no far jumps (not needed in virtual mode?)

    Instruction "HLT    ", "F4      ", SizeModNone, ExtNon
    Instruction "NOP    ", "90      ", SizeModNone, ExtNon

    Instruction "RET    ", "C3      ", SizeModNone, ExtNon
    Instruction "RET    ", "C2      ", SizeModNone, ExtNon, "imm16"
    ' no far returns

    Instruction "CALL   ", "E8      ", SizeModNone, ExtNon, "rel32"
    Instruction "CALL   ", "E8      ", SizeModOvrd, ExtNon, "rel16"
    Instruction "CALL   ", "FF /2   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "CALL   ", "FF /2   ", SizeModNone, ExtNon, "rem32"
    ' no far calls

    Instruction "INT    ", "CC      ", SizeModNone, ExtNon, "3    "
    Instruction "INT    ", "CD      ", SizeModNone, ExtNon, "imm08"

    Instruction "ENTER  ", "C8      ", SizeModNone, ExtNon, "imm16", "imm08"
    Instruction "LEAVE  ", "C9      ", SizeModNone, ExtNon
    
    Instruction "SET    ", "0F 90   ", SizeModNone, ExtCon, "rem08"

   Instruction "LOOP    ", "E2      ", SizeModNone, ExtNon, "rel08"
   Instruction "LOOPE   ", "E1      ", SizeModNone, ExtNon, "rel08"
   Instruction "LOOPNE  ", "E0      ", SizeModNone, ExtNon, "rel08"
End Sub


Private Sub AddBits()
    Instruction "BT     ", "0F A3   ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "BT     ", "0F A3   ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "BT     ", "0F BA /4", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "BT     ", "0F BA /4", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "BTC    ", "0F BB   ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "BTC    ", "0F BB   ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "BTC    ", "0F BB /7", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "BTC    ", "0F BB /7", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "BTR    ", "0F B3   ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "BTR    ", "0F B3   ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "BTR    ", "0F BA /6", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "BTR    ", "0F BA /6", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "BTS    ", "0F AB   ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "BTS    ", "0F AB   ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "BTS    ", "0F BA /5", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "BTS    ", "0F BA /5", SizeModNone, ExtNon, "rem32", "imm08"

    Instruction "NEG    ", "F6 /3   ", SizeModNone, ExtNon, "rem08"
    Instruction "NEG    ", "F7 /3   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "NEG    ", "F7 /3   ", SizeModNone, ExtNon, "rem32"
    Instruction "NOT    ", "F6 /2   ", SizeModNone, ExtNon, "rem08"
    Instruction "NOT    ", "F7 /2   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "NOT    ", "F7 /2   ", SizeModNone, ExtNon, "rem32"

    Instruction "SHLD   ", "0F A4   ", SizeModOvrd, ExtNon, "rem16", "reg16", "imm08"
    Instruction "SHLD   ", "0F A4   ", SizeModNone, ExtNon, "rem32", "reg32", "imm08"
    Instruction "SHLD   ", "0F A5   ", SizeModOvrd, ExtNon, "rem16", "reg16", "CL   "
    Instruction "SHLD   ", "0F A5   ", SizeModNone, ExtNon, "rem32", "reg32", "CL   "
    
    Instruction "SHRD   ", "0F AC   ", SizeModOvrd, ExtNon, "rem16", "reg16", "imm08"
    Instruction "SHRD   ", "0F AC   ", SizeModNone, ExtNon, "rem32", "reg32", "imm08"
    Instruction "SHRD   ", "0F AD   ", SizeModOvrd, ExtNon, "rem16", "reg16", "CL   "
    Instruction "SHRD   ", "0F AD   ", SizeModNone, ExtNon, "rem32", "reg32", "CL   "

    Instruction "SHL    ", "D0 /4   ", SizeModNone, ExtNon, "rem08", "1    "
    Instruction "SHL    ", "D2 /4   ", SizeModNone, ExtNon, "rem08", "CL   "
    Instruction "SHL    ", "C0 /4   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "SHL    ", "D1 /4   ", SizeModOvrd, ExtNon, "rem16", "1    "
    Instruction "SHL    ", "D1 /4   ", SizeModNone, ExtNon, "rem32", "1    "
    Instruction "SHL    ", "D3 /4   ", SizeModOvrd, ExtNon, "rem16", "CL   "
    Instruction "SHL    ", "D3 /4   ", SizeModNone, ExtNon, "rem32", "CL   "
    Instruction "SHL    ", "C1 /4   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "SHL    ", "C1 /4   ", SizeModNone, ExtNon, "rem32", "imm08"
    
    Instruction "SAL    ", "D0 /4   ", SizeModNone, ExtNon, "rem08", "1    "
    Instruction "SAL    ", "D2 /4   ", SizeModNone, ExtNon, "rem08", "CL   "
    Instruction "SAL    ", "C0 /4   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "SAL    ", "D1 /4   ", SizeModOvrd, ExtNon, "rem16", "1    "
    Instruction "SAL    ", "D1 /4   ", SizeModNone, ExtNon, "rem32", "1    "
    Instruction "SAL    ", "D3 /4   ", SizeModOvrd, ExtNon, "rem16", "CL   "
    Instruction "SAL    ", "D3 /4   ", SizeModNone, ExtNon, "rem32", "CL   "
    Instruction "SAL    ", "C1 /4   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "SAL    ", "C1 /4   ", SizeModNone, ExtNon, "rem32", "imm08"
    
    Instruction "SAR    ", "D0 /7   ", SizeModNone, ExtNon, "rem08", "1    "
    Instruction "SAR    ", "D2 /7   ", SizeModNone, ExtNon, "rem08", "CL   "
    Instruction "SAR    ", "C0 /7   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "SAR    ", "D1 /7   ", SizeModOvrd, ExtNon, "rem16", "1    "
    Instruction "SAR    ", "D1 /7   ", SizeModNone, ExtNon, "rem32", "1    "
    Instruction "SAR    ", "D3 /7   ", SizeModOvrd, ExtNon, "rem16", "CL   "
    Instruction "SAR    ", "D3 /7   ", SizeModNone, ExtNon, "rem32", "CL   "
    Instruction "SAR    ", "C1 /7   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "SAR    ", "C1 /7   ", SizeModNone, ExtNon, "rem32", "imm08"
    
    Instruction "SHR    ", "D0 /5   ", SizeModNone, ExtNon, "rem08", "1    "
    Instruction "SHR    ", "D2 /5   ", SizeModNone, ExtNon, "rem08", "CL   "
    Instruction "SHR    ", "C0 /5   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "SHR    ", "D1 /5   ", SizeModOvrd, ExtNon, "rem16", "1    "
    Instruction "SHR    ", "D1 /5   ", SizeModNone, ExtNon, "rem32", "1    "
    Instruction "SHR    ", "D3 /5   ", SizeModOvrd, ExtNon, "rem16", "CL   "
    Instruction "SHR    ", "D3 /5   ", SizeModNone, ExtNon, "rem32", "CL   "
    Instruction "SHR    ", "C1 /5   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "SHR    ", "C1 /5   ", SizeModNone, ExtNon, "rem32", "imm08"
    
    Instruction "RCL    ", "D0 /2   ", SizeModNone, ExtNon, "rem08", "1    "
    Instruction "RCL    ", "D2 /2   ", SizeModNone, ExtNon, "rem08", "CL   "
    Instruction "RCL    ", "C0 /2   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "RCL    ", "D1 /2   ", SizeModOvrd, ExtNon, "rem16", "1    "
    Instruction "RCL    ", "D1 /2   ", SizeModNone, ExtNon, "rem32", "1    "
    Instruction "RCL    ", "D3 /2   ", SizeModOvrd, ExtNon, "rem16", "CL   "
    Instruction "RCL    ", "D3 /2   ", SizeModNone, ExtNon, "rem32", "CL   "
    Instruction "RCL    ", "C1 /2   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "RCL    ", "C1 /2   ", SizeModNone, ExtNon, "rem32", "imm08"
    
    Instruction "RCR    ", "D0 /3   ", SizeModNone, ExtNon, "rem08", "1    "
    Instruction "RCR    ", "D2 /3   ", SizeModNone, ExtNon, "rem08", "CL   "
    Instruction "RCR    ", "C0 /3   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "RCR    ", "D1 /3   ", SizeModOvrd, ExtNon, "rem16", "1    "
    Instruction "RCR    ", "D1 /3   ", SizeModNone, ExtNon, "rem32", "1    "
    Instruction "RCR    ", "D3 /3   ", SizeModOvrd, ExtNon, "rem16", "CL   "
    Instruction "RCR    ", "D3 /3   ", SizeModNone, ExtNon, "rem32", "CL   "
    Instruction "RCR    ", "C1 /3   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "RCR    ", "C1 /3   ", SizeModNone, ExtNon, "rem32", "imm08"
    
    Instruction "ROL    ", "D0 /0   ", SizeModNone, ExtNon, "rem08", "1    "
    Instruction "ROL    ", "D2 /0   ", SizeModNone, ExtNon, "rem08", "CL   "
    Instruction "ROL    ", "C0 /0   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "ROL    ", "D1 /0   ", SizeModOvrd, ExtNon, "rem16", "1    "
    Instruction "ROL    ", "D1 /0   ", SizeModNone, ExtNon, "rem32", "1    "
    Instruction "ROL    ", "D3 /0   ", SizeModOvrd, ExtNon, "rem16", "CL   "
    Instruction "ROL    ", "D3 /0   ", SizeModNone, ExtNon, "rem32", "CL   "
    Instruction "ROL    ", "C1 /0   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "ROL    ", "C1 /0   ", SizeModNone, ExtNon, "rem32", "imm08"
    
    Instruction "ROR    ", "D0 /1   ", SizeModNone, ExtNon, "rem08", "1    "
    Instruction "ROR    ", "D2 /1   ", SizeModNone, ExtNon, "rem08", "CL   "
    Instruction "ROR    ", "C0 /1   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "ROR    ", "D1 /1   ", SizeModOvrd, ExtNon, "rem16", "1    "
    Instruction "ROR    ", "D1 /1   ", SizeModNone, ExtNon, "rem32", "1    "
    Instruction "ROR    ", "D3 /1   ", SizeModOvrd, ExtNon, "rem16", "CL   "
    Instruction "ROR    ", "D3 /1   ", SizeModNone, ExtNon, "rem32", "CL   "
    Instruction "ROR    ", "C1 /1   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "ROR    ", "C1 /1   ", SizeModNone, ExtNon, "rem32", "imm08"

    Instruction "TEST   ", "A8      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "TEST   ", "A9      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "TEST   ", "A9      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "TEST   ", "F6 /0   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "TEST   ", "F7 /0   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "TEST   ", "F7 /0   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "TEST   ", "84      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "TEST   ", "85      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "TEST   ", "85      ", SizeModNone, ExtNon, "rem32", "reg32"
    
    Instruction "OR     ", "0C      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "OR     ", "0D      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "OR     ", "0D      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "OR     ", "80 /1   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "OR     ", "81 /1   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "OR     ", "81 /1   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "OR     ", "83 /1   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "OR     ", "83 /1   ", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "OR     ", "08      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "OR     ", "09      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "OR     ", "09      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "OR     ", "0A      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "OR     ", "0B      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "OR     ", "0B      ", SizeModNone, ExtNon, "reg32", "rem32"
    
    Instruction "AND    ", "24      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "AND    ", "25      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "AND    ", "25      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "AND    ", "80 /4   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "AND    ", "81 /4   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "AND    ", "81 /4   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "AND    ", "83 /4   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "AND    ", "83 /4   ", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "AND    ", "20      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "AND    ", "21      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "AND    ", "21      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "AND    ", "22      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "AND    ", "23      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "AND    ", "23      ", SizeModNone, ExtNon, "reg32", "rem32"
    
    Instruction "XOR    ", "34      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "XOR    ", "35      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "XOR    ", "35      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "XOR    ", "80 /6   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "XOR    ", "81 /6   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "XOR    ", "81 /6   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "XOR    ", "83 /6   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "XOR    ", "83 /6   ", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "XOR    ", "30      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "XOR    ", "31      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "XOR    ", "31      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "XOR    ", "32      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "XOR    ", "33      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "XOR    ", "33      ", SizeModNone, ExtNon, "reg32", "rem32"

    Instruction "BSF    ", "0F BC   ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "BSF    ", "0F BC   ", SizeModNone, ExtNon, "reg32", "rem32"
    Instruction "BSR    ", "0F BD   ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "BSR    ", "0F BD   ", SizeModNone, ExtNon, "reg32", "rem32"
End Sub


Private Sub AddFloatingPoint()
    Instruction "FADD   ", "D8 /0   ", SizeModNone, ExtNon, "rem32"
    Instruction "FADD   ", "DC /0   ", SizeModNone, ExtNon, "rem64"
    Instruction "FADD   ", "D8 C0   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FADD   ", "DC C0   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FADDP  ", "DE C0   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FADDP  ", "DE C1   ", SizeModNone, ExtNon
    Instruction "FIADD  ", "DA /0   ", SizeModNone, ExtNon, "rem32"
    Instruction "FIADD  ", "DE /0   ", SizeModNone, ExtNon, "rem16"
    
    Instruction "FABS   ", "D9 E1   ", SizeModNone, ExtNon
    Instruction "FBLD   ", "DF /4   ", SizeModNone, ExtNon, "mem  "
    Instruction "FBSTP  ", "DF /6   ", SizeModNone, ExtNon, "mem  "
    Instruction "FCHS   ", "D9 E0   ", SizeModNone, ExtNon
    Instruction "FCLEX  ", "9B DB E2", SizeModNone, ExtNon
    
    Instruction "FCMOVB ", "DA C0   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FCMOVE ", "DA C8   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FCMOVBE", "DA D0   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FCMOVU ", "DA D8   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FCMOVNB", "DB C0   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FCMOVNE", "DB C8   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FCMOVNBE", "DB D0  ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FCMOVNU", "DB D8   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    
    Instruction "FCOMI  ", "DB F0   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FCOMIP ", "DF F0   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FUCOMI ", "DB E8   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FUCOMIP", "DF E8   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    
    Instruction "FCOS   ", "D9 FF   ", SizeModNone, ExtNon
    Instruction "FSIN   ", "D9 FE   ", SizeModNone, ExtNon
    Instruction "FSINCOS", "D9 FB   ", SizeModNone, ExtNon
    Instruction "FDECSTP", "D9 F6   ", SizeModNone, ExtNon
    
    Instruction "FDIV   ", "D8 /6   ", SizeModNone, ExtNon, "mem32"
    Instruction "FDIV   ", "DC /6   ", SizeModNone, ExtNon, "mem64"
    Instruction "FDIV   ", "D8 F0   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FDIV   ", "DC F8   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FDIVP  ", "DE F8   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FDIVP  ", "DE F9   ", SizeModNone, ExtNon
    Instruction "FIDIV  ", "DA /6   ", SizeModNone, ExtNon, "mem32"
    Instruction "FIDIV  ", "DE /6   ", SizeModNone, ExtNon, "mem64"
    
    Instruction "FDIVR  ", "D8 /7   ", SizeModNone, ExtNon, "mem32"
    Instruction "FDIVR  ", "DC /7   ", SizeModNone, ExtNon, "mem64"
    Instruction "FDIVR  ", "D8 F8   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FDIVR  ", "DC F0   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FDIVRP ", "DE F0   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FDIVRP ", "DE F1   ", SizeModNone, ExtNon
    Instruction "FIDIVR ", "DA /7   ", SizeModNone, ExtNon, "mem32"
    Instruction "FIDIVR ", "DE /7   ", SizeModNone, ExtNon, "mem64"
    
    Instruction "FICOM  ", "DE /2   ", SizeModNone, ExtNon, "mem16"
    Instruction "FICOM  ", "DA /2   ", SizeModNone, ExtNon, "mem32"
    Instruction "FICOMP ", "DE /3   ", SizeModNone, ExtNon, "mem16"
    Instruction "FICOMP ", "DA /3   ", SizeModNone, ExtNon, "mem32"
    
    Instruction "FILD   ", "DF /0   ", SizeModNone, ExtNon, "mem16"
    Instruction "FILD   ", "DB /0   ", SizeModNone, ExtNon, "mem32"
    Instruction "FILD   ", "DF /5   ", SizeModNone, ExtNon, "mem64"
    
    Instruction "FINCSTP", "D9 F7   ", SizeModNone, ExtNon
    Instruction "FINIT  ", "9B DB E3", SizeModNone, ExtNon
    Instruction "FFREE  ", "DD C0   ", SizeModNone, ExtFlt, "st#  "

    Instruction "FIST   ", "DF /2   ", SizeModNone, ExtNon, "mem16"
    Instruction "FIST   ", "DB /2   ", SizeModNone, ExtNon, "mem32"
    Instruction "FISTP  ", "DF /3   ", SizeModNone, ExtNon, "mem16"
    Instruction "FISTP  ", "DB /3   ", SizeModNone, ExtNon, "mem32"
    Instruction "FISTP  ", "DF /7   ", SizeModNone, ExtNon, "mem64"
    
    Instruction "FISTTP ", "DF /1   ", SizeModNone, ExtNon, "mem16"
    Instruction "FISTTP ", "DB /1   ", SizeModNone, ExtNon, "mem32"
    Instruction "FISTTP ", "DD /1   ", SizeModNone, ExtNon, "mem64"
    
    Instruction "FLD    ", "D9 /0   ", SizeModNone, ExtNon, "mem32"
    Instruction "FLD    ", "DD /0   ", SizeModNone, ExtNon, "mem64"
    Instruction "FLD    ", "DB /5   ", SizeModNone, ExtNon, "mem80"
    Instruction "FLD    ", "D9 C0   ", SizeModNone, ExtFlt, "st#  "
    
    Instruction "FLD1   ", "D9 E8   ", SizeModNone, ExtNon
    Instruction "FLDL2T ", "D9 E9   ", SizeModNone, ExtNon
    Instruction "FLDL2E ", "D9 EA   ", SizeModNone, ExtNon
    Instruction "FLDPI  ", "D9 EB   ", SizeModNone, ExtNon
    Instruction "FLDLG2 ", "D9 EC   ", SizeModNone, ExtNon
    Instruction "FLDLN2 ", "D9 ED   ", SizeModNone, ExtNon
    Instruction "FLDZ   ", "D9 EE   ", SizeModNone, ExtNon
    
    Instruction "FLDCW  ", "D9 /5   ", SizeModNone, ExtNon, "mem  "
    Instruction "FLDENV ", "D9 /4   ", SizeModNone, ExtNon, "mem  "
    
    Instruction "FMUL   ", "D8 /1   ", SizeModNone, ExtNon, "mem32"
    Instruction "FMUL   ", "DC /1   ", SizeModNone, ExtNon, "mem64"
    Instruction "FMUL   ", "D8 C8   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FMUL   ", "DC C8   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FMUL   ", "DE C8   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FMULP  ", "DE C9   ", SizeModNone, ExtNon
    Instruction "FIMUL  ", "DA /1   ", SizeModNone, ExtNon, "mem32"
    Instruction "FIMUL  ", "DE /1   ", SizeModNone, ExtNon, "mem16"
    
    Instruction "FNOP   ", "D9 D0   ", SizeModNone, ExtNon
    Instruction "FPATAN ", "D9 F3   ", SizeModNone, ExtNon
    Instruction "FPREM  ", "D9 F8   ", SizeModNone, ExtNon
    Instruction "FPREM1 ", "D9 F5   ", SizeModNone, ExtNon
    Instruction "FPTAN  ", "D9 F2   ", SizeModNone, ExtNon
    Instruction "FRNDINT", "D9 FC   ", SizeModNone, ExtNon
    Instruction "FRSTOR ", "DD /4   ", SizeModNone, ExtNon, "mem  "
    Instruction "FSAVE  ", "9B DD /6", SizeModNone, ExtNon, "mem  "
    Instruction "FSCALE ", "D9 FD   ", SizeModNone, ExtNon
    Instruction "FSQRT  ", "D9 FA   ", SizeModNone, ExtNon
    
    Instruction "FST    ", "D9 /2   ", SizeModNone, ExtNon, "mem32"
    Instruction "FST    ", "DD /2   ", SizeModNone, ExtNon, "mem64"
    Instruction "FST    ", "DD D0   ", SizeModNone, ExtFlt, "st#  "
    Instruction "FSTP   ", "D9 /3   ", SizeModNone, ExtNon, "mem32"
    Instruction "FSTP   ", "DD /3   ", SizeModNone, ExtNon, "mem64"
    Instruction "FSTP   ", "DB /7   ", SizeModNone, ExtNon, "mem80"
    Instruction "FSTP   ", "DD D8   ", SizeModNone, ExtFlt, "st#  "
    
    Instruction "FSTCW  ", "9B D9 /7", SizeModNone, ExtNon, "mem  "
    Instruction "FSTENV ", "9B D9 /6", SizeModNone, ExtNon, "mem  "
    Instruction "FSTSW  ", "9B DD /7", SizeModNone, ExtNon, "mem  "
    Instruction "FSTSW  ", "9B DF E0", SizeModNone, ExtNon, "ax   "
    
    Instruction "FSUB   ", "D8 /4   ", SizeModNone, ExtNon, "mem32"
    Instruction "FSUB   ", "DC /4   ", SizeModNone, ExtNon, "mem64"
    Instruction "FSUB   ", "D8 E0   ", SizeModNone, ExtFlt, "st0  ", "st#  "
    Instruction "FSUB   ", "DC E8   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FSUBP  ", "DE E8   ", SizeModNone, ExtFlt, "st#  ", "st0  "
    Instruction "FSUBP  ", "DE E9   ", SizeModNone, ExtNon
    Instruction "FISUB  ", "DA /4   ", SizeModNone, ExtNon, "mem32"
    Instruction "FISUB  ", "DE /4   ", SizeModNone, ExtNon, "mem16"
    
    Instruction "FSUBR  ", "D8 /5   ", SizeModNone, ExtNon, "mem32"
    Instruction "FSUBR  ", "DC /5   ", SizeModNone, ExtNon, "mem64"
    Instruction "FSUBR  ", "D8 E8   ", SizeModNone, ExtNon, "st0  ", "st#  "
    Instruction "FSUBR  ", "DC E0   ", SizeModNone, ExtNon, "st#  ", "st0  "
    Instruction "FSUBR  ", "DE E0   ", SizeModNone, ExtNon, "st#  ", "st0  "
    Instruction "FSUBRP ", "DE E1   ", SizeModNone, ExtNon
    Instruction "FISUBR ", "DA /5   ", SizeModNone, ExtNon, "mem32"
    Instruction "FISUBR ", "DE /5   ", SizeModNone, ExtNon, "mem16"
    
    Instruction "FTST   ", "D9 E4   ", SizeModNone, ExtNon
    Instruction "FUCOM  ", "DD E0   ", SizeModNone, ExtFlt, "st#  "
    Instruction "FUCOM  ", "DD E1   ", SizeModNone, ExtNon
    Instruction "FUCOMP ", "DD E8   ", SizeModNone, ExtFlt, "st#  "
    Instruction "FUCOMP ", "DD E9   ", SizeModNone, ExtNon
    Instruction "FUCOMPP", "DA E9   ", SizeModNone, ExtNon
    Instruction "FXAM   ", "D9 E5   ", SizeModNone, ExtNon
    Instruction "FXCH   ", "D9 C8   ", SizeModNone, ExtFlt, "st#  "
    Instruction "FXCH   ", "D9 C9   ", SizeModNone, ExtNon
    Instruction "FXRSTOR", "0F AE /1", SizeModNone, ExtNon, "mem  "
    Instruction "FXSAVE ", "0F AE /0", SizeModNone, ExtNon, "mem  "
    Instruction "FXTRACT", "D9 F4   ", SizeModNone, ExtNon
    Instruction "FYL2X  ", "D9 F1   ", SizeModNone, ExtNon
    Instruction "FYL2XP1", "D9 F9   ", SizeModNone, ExtNon
    
    Instruction "FWAIT  ", "9B      ", SizeModNone, ExtNon
    Instruction "WAIT   ", "9B      ", SizeModNone, ExtNon
End Sub


Private Sub AddArithmetics()
    Instruction "INC    ", "40      ", SizeModOvrd, ExtReg, "#16  "
    Instruction "INC    ", "40      ", SizeModNone, ExtReg, "#32  "
    Instruction "INC    ", "FE /0   ", SizeModNone, ExtNon, "rem08"
    Instruction "INC    ", "FF /0   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "INC    ", "FF /0   ", SizeModNone, ExtNon, "rem32"
    
    Instruction "DEC    ", "48      ", SizeModOvrd, ExtReg, "#16  "
    Instruction "DEC    ", "48      ", SizeModNone, ExtReg, "#32  "
    Instruction "DEC    ", "FE /1   ", SizeModNone, ExtNon, "rem08"
    Instruction "DEC    ", "FF /1   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "DEC    ", "FF /1   ", SizeModNone, ExtNon, "rem32"

    Instruction "SBB    ", "1C      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "SBB    ", "1D      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "SBB    ", "1D      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "SBB    ", "80 /3   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "SBB    ", "81 /3   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "SBB    ", "81 /3   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "SBB    ", "83 /3   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "SBB    ", "83 /3   ", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "SBB    ", "18      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "SBB    ", "19      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "SBB    ", "19      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "SBB    ", "1A      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "SBB    ", "1B      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "SBB    ", "1B      ", SizeModNone, ExtNon, "reg32", "rem32"
    
    Instruction "IMUL   ", "F6 /5   ", SizeModNone, ExtNon, "rem08"
    Instruction "IMUL   ", "F7 /5   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "IMUL   ", "F7 /5   ", SizeModNone, ExtNon, "rem32"
    Instruction "IMUL   ", "0F AF   ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "IMUL   ", "0F AF   ", SizeModNone, ExtNon, "reg32", "rem32"
    Instruction "IMUL   ", "69      ", SizeModOvrd, ExtNon, "reg16", "rem16", "imm16"
    Instruction "IMUL   ", "69      ", SizeModNone, ExtNon, "reg32", "rem32", "imm32"
   'Instruction "IMUL   ", "69      ", SizeModOvrd, ExtNon, "reg16", "imm16"
   'Instruction "IMUL   ", "69      ", SizeModNone, ExtNon, "reg32", "imm32"
    Instruction "IMUL   ", "6B      ", SizeModOvrd, ExtNon, "reg16", "rem16", "imm08"
    Instruction "IMUL   ", "6B      ", SizeModNone, ExtNon, "reg32", "rem32", "imm08"
   'Instruction "IMUL   ", "6B      ", SizeModOvrd, ExtNon, "reg16", "imm08"
   'Instruction "IMUL   ", "6B      ", SizeModNone, ExtNon, "reg32", "imm08"
     
    Instruction "DIV    ", "F6 /6   ", SizeModNone, ExtNon, "rem08"
    Instruction "DIV    ", "F7 /6   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "DIV    ", "F7 /6   ", SizeModNone, ExtNon, "rem32"
     
    Instruction "MUL    ", "F6 /4   ", SizeModNone, ExtNon, "rem08"
    Instruction "MUL    ", "F7 /4   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "MUL    ", "F7 /4   ", SizeModNone, ExtNon, "rem32"
    
    Instruction "IDIV   ", "F6 /7   ", SizeModNone, ExtNon, "rem08"
    Instruction "IDIV   ", "F7 /7   ", SizeModOvrd, ExtNon, "rem16"
    Instruction "IDIV   ", "F7 /7   ", SizeModNone, ExtNon, "rem32"

    Instruction "XADD   ", "0F C0   ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "XADD   ", "0F C1   ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "XADD   ", "0F C1   ", SizeModNone, ExtNon, "rem32", "reg32"

    Instruction "ADC    ", "14      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "ADC    ", "15      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "ADC    ", "15      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "ADC    ", "80 /2   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "ADC    ", "81 /2   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "ADC    ", "81 /2   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "ADC    ", "83 /2   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "ADC    ", "83 /2   ", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "ADC    ", "10      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "ADC    ", "11      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "ADC    ", "11      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "ADC    ", "12      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "ADC    ", "13      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "ADC    ", "13      ", SizeModNone, ExtNon, "reg32", "rem32"
    
    Instruction "ADD    ", "04      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "ADD    ", "05      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "ADD    ", "05      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "ADD    ", "80 /0   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "ADD    ", "81 /0   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "ADD    ", "81 /0   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "ADD    ", "83 /0   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "ADD    ", "83 /0   ", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "ADD    ", "00      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "ADD    ", "01      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "ADD    ", "01      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "ADD    ", "02      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "ADD    ", "03      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "ADD    ", "03      ", SizeModNone, ExtNon, "reg32", "rem32"
    
    Instruction "SUB    ", "2C      ", SizeModNone, ExtNon, "AL   ", "imm08"
    Instruction "SUB    ", "2D      ", SizeModOvrd, ExtNon, "AX   ", "imm16"
    Instruction "SUB    ", "2D      ", SizeModNone, ExtNon, "EAX  ", "imm32"
    Instruction "SUB    ", "80 /5   ", SizeModNone, ExtNon, "rem08", "imm08"
    Instruction "SUB    ", "81 /5   ", SizeModOvrd, ExtNon, "rem16", "imm16"
    Instruction "SUB    ", "81 /5   ", SizeModNone, ExtNon, "rem32", "imm32"
    Instruction "SUB    ", "83 /5   ", SizeModOvrd, ExtNon, "rem16", "imm08"
    Instruction "SUB    ", "83 /5   ", SizeModNone, ExtNon, "rem32", "imm08"
    Instruction "SUB    ", "28      ", SizeModNone, ExtNon, "rem08", "reg08"
    Instruction "SUB    ", "29      ", SizeModOvrd, ExtNon, "rem16", "reg16"
    Instruction "SUB    ", "29      ", SizeModNone, ExtNon, "rem32", "reg32"
    Instruction "SUB    ", "2A      ", SizeModNone, ExtNon, "reg08", "rem08"
    Instruction "SUB    ", "2B      ", SizeModOvrd, ExtNon, "reg16", "rem16"
    Instruction "SUB    ", "2B      ", SizeModNone, ExtNon, "reg32", "rem32"
End Sub


Private Sub Instruction( _
    Mnemonic As String, _
    OpCode As String, _
    ByVal SizeModifier As SizeMod, _
    ByVal Ext As ExtType, _
    ParamArray Params() As Variant _
)

    Dim strArgs()   As String
    Dim i           As Long
    
    If UBound(Params) > -1 Then
        ReDim strArgs(UBound(Params)) As String
        
        For i = 0 To UBound(Params)
            strArgs(i) = Trim$(Params(i))
        Next
    End If
    
    Select Case Ext
        Case ExtNon: InstrDefault Trim$(Mnemonic), Trim$(OpCode), SizeModifier, strArgs, UBound(Params) + 1
        Case ExtReg: InstrRegExt Trim$(Mnemonic), Trim$(OpCode), SizeModifier, strArgs, UBound(Params) + 1
        Case ExtFlt: InstrFloatExt Trim$(Mnemonic), Trim$(OpCode), SizeModifier, strArgs, UBound(Params) + 1
        Case ExtCon: InstrCondition Trim$(Mnemonic), Trim$(OpCode), SizeModifier, strArgs, UBound(Params) + 1
        Case Ext3DN: Instr3DNow Trim$(Mnemonic), Trim$(OpCode), SizeModNone, strArgs, UBound(Params) + 1
    End Select
End Sub


' OpCode depends on the used register
Private Sub InstrRegExt( _
    Mnemonic As String, _
    Op As String, _
    ByVal SizeModifier As SizeMod, _
    Params() As String, _
    ParamCnt As Long _
)

    Dim udtInstr    As Instruction
    Dim i           As Long
    Dim lngRegPa    As Long
    Dim udeSize     As ParamSize

    udtInstr.Mnemonic = Mnemonic
    udtInstr.Prefixes = IIf(SizeModifier = SizeModOvrd, PrefixFlgOperandSizeOverride, 0)
    udtInstr.ParamCount = ParamCnt

    With ParseOpCode(Op)
        For i = 0 To MAX_OPCODE_LEN - 1
            udtInstr.OpCode(i) = .Bytes(i)
        Next
        
        udtInstr.OpCodeLen = .ByteCount
        udtInstr.RegOpExt = .RegOpExt
        udtInstr.Now3DByte = -1
        udtInstr.ModRM = .RegOpExt > -1
    End With
    
    For i = 0 To ParamCnt - 1
        If Left$(Params(i), 1) = "#" Then
            lngRegPa = i
            Select Case Mid$(Params(i), 2, 2)
                Case "08":  udeSize = Bits8
                Case "16":  udeSize = Bits16
                Case "32":  udeSize = Bits32
                Case Else:  Err.Raise 12345, , "Invalid size!"
            End Select
        Else
            udtInstr.Parameters(i) = ParseParameter(Params(i))
            If Not udtInstr.Parameters(i).Forced Then
                If udtInstr.Parameters(i).PType = ParamReg Or _
                   udtInstr.Parameters(i).PType = (ParamReg Or ParamMem) Then
                    udtInstr.ModRM = True
                End If
            End If
        End If
    Next
    
    For i = 0 To 7
        With udtInstr
            .Parameters(lngRegPa) = ParseParameter(GetRegExtRegName(i, udeSize))
            AddInstr udtInstr
            
            .OpCode(.OpCodeLen - 1) = .OpCode(.OpCodeLen - 1) + 1
        End With
    Next
End Sub


' OpCode depends on the used FPU register
Private Sub InstrFloatExt( _
    Mnemonic As String, _
    Op As String, _
    ByVal SizeModifier As SizeMod, _
    Params() As String, _
    ByVal ParamCnt As Long _
)

    Dim udtInstr    As Instruction
    Dim i           As Long
    Dim lngFlPa     As Long
    Dim lastbyte    As Byte

    udtInstr.Mnemonic = Mnemonic
    udtInstr.Prefixes = IIf(SizeModifier = SizeModOvrd, PrefixFlgOperandSizeOverride, 0)
    udtInstr.ParamCount = ParamCnt

    With ParseOpCode(Op)
        For i = 0 To MAX_OPCODE_LEN - 1
            udtInstr.OpCode(i) = .Bytes(i)
        Next
        
        udtInstr.OpCodeLen = .ByteCount
        udtInstr.RegOpExt = .RegOpExt
        udtInstr.Now3DByte = -1
        udtInstr.ModRM = .RegOpExt > -1
    End With
    lastbyte = udtInstr.OpCode(udtInstr.OpCodeLen - 1)
    
    For i = 0 To ParamCnt - 1
        If UCase$(Params(i)) = "ST#" Then
            lngFlPa = i
        Else
            udtInstr.Parameters(i) = ParseParameter(Params(i))
            If Not udtInstr.Parameters(i).Forced Then
                If udtInstr.Parameters(i).PType = ParamReg Or _
                   udtInstr.Parameters(i).PType = (ParamReg Or ParamMem) Then
                    udtInstr.ModRM = True
                End If
            End If
        End If
    Next
    
    For i = 0 To 7
        udtInstr.OpCode(udtInstr.OpCodeLen - 1) = lastbyte + i
        udtInstr.Parameters(lngFlPa) = ParseParameter("ST" & i)
        AddInstr udtInstr
    Next
End Sub


' conditional instruction (xxxZ, xxxNZ, xxxNGE, ...)
Private Sub InstrCondition( _
    Mnemonic As String, _
    Op As String, _
    ByVal SizeModifier As SizeMod, _
    Params() As String, _
    ByVal ParamCnt As Long _
)

    Dim udtInstr    As Instruction
    Dim conds()     As String
    Dim i           As Long
    Dim lastbyte    As Byte
    
    conds = Split(RemoveWSDoubles(CONDITIONS), " ")

    udtInstr.Prefixes = IIf(SizeModifier = SizeModOvrd, PrefixFlgOperandSizeOverride, 0)
    udtInstr.ParamCount = ParamCnt
    
    With ParseOpCode(Op)
        For i = 0 To MAX_OPCODE_LEN - 1
            udtInstr.OpCode(i) = .Bytes(i)
        Next
        
        udtInstr.OpCodeLen = .ByteCount
        udtInstr.RegOpExt = .RegOpExt
        udtInstr.Now3DByte = -1
        udtInstr.ModRM = .RegOpExt > -1
    End With
    lastbyte = udtInstr.OpCode(udtInstr.OpCodeLen - 1)
    
    For i = 0 To ParamCnt - 1
        udtInstr.Parameters(i) = ParseParameter(Params(i))
        If Not udtInstr.Parameters(i).Forced Then
            If udtInstr.Parameters(i).PType = ParamReg Or _
               udtInstr.Parameters(i).PType = (ParamReg Or ParamMem) Then
                udtInstr.ModRM = True
            End If
        End If
    Next

    For i = 0 To UBound(conds)
        udtInstr.Mnemonic = Mnemonic & conds(i)
        udtInstr.OpCode(udtInstr.OpCodeLen - 1) = lastbyte + ConditionOffset(conds(i))
        AddInstr udtInstr
    Next
End Sub


' 3DNow instruction has an immediate value defining the operation
Private Sub Instr3DNow( _
    Mnemonic As String, _
    Op As String, _
    ByVal SizeModifier As SizeMod, _
    Params() As String, _
    ByVal ParamCnt As Long _
)

    Dim udtInstr    As Instruction
    Dim i           As Long

    udtInstr.Mnemonic = Mnemonic
    udtInstr.Prefixes = IIf(SizeModifier = SizeModOvrd, PrefixFlgOperandSizeOverride, 0)
    udtInstr.ParamCount = ParamCnt - 1

    With ParseOpCode(Op)
        For i = 0 To MAX_OPCODE_LEN - 1
            udtInstr.OpCode(i) = .Bytes(i)
        Next
        
        udtInstr.OpCodeLen = .ByteCount
        udtInstr.RegOpExt = .RegOpExt
        udtInstr.ModRM = .RegOpExt > -1
    End With
    
    For i = 0 To ParamCnt - 2
        udtInstr.Parameters(i) = ParseParameter(Params(i))
        If Not udtInstr.Parameters(i).Forced Then
            If udtInstr.Parameters(i).PType = ParamReg Or _
               udtInstr.Parameters(i).PType = (ParamReg Or ParamMem) Or _
               (udtInstr.Parameters(i).PType And ParamMM) Then
                udtInstr.ModRM = True
            End If
        End If
    Next
    
    udtInstr.Now3DByte = CLng(Params(i))
    
    AddInstr udtInstr
End Sub


Private Sub InstrDefault( _
    Mnemonic As String, _
    Op As String, _
    ByVal SizeModifier As SizeMod, _
    Params() As String, _
    ByVal ParamCnt As Long _
)

    Dim udtInstr    As Instruction
    Dim i           As Long

    udtInstr.Mnemonic = Mnemonic
    udtInstr.Prefixes = IIf(SizeModifier = SizeModOvrd, PrefixFlgOperandSizeOverride, 0)
    udtInstr.ParamCount = ParamCnt

    With ParseOpCode(Op)
        For i = 0 To MAX_OPCODE_LEN - 1
            udtInstr.OpCode(i) = .Bytes(i)
        Next
        
        udtInstr.OpCodeLen = .ByteCount
        udtInstr.RegOpExt = .RegOpExt
        udtInstr.Now3DByte = -1
        udtInstr.ModRM = .RegOpExt > -1
    End With
    
    For i = 0 To ParamCnt - 1
        udtInstr.Parameters(i) = ParseParameter(Params(i))
        If Not udtInstr.Parameters(i).Forced Then
            If udtInstr.Parameters(i).PType = ParamReg Or _
               udtInstr.Parameters(i).PType = (ParamReg Or ParamMem) Or _
               (udtInstr.Parameters(i).PType And ParamMM) Then
                udtInstr.ModRM = True
            End If
        End If
    Next
    
    AddInstr udtInstr
End Sub


Private Function ParseParameter(param As String) As InstructionParam
    With ParseParameter
        If IsRegister(param) Then
            .Register = RegStrToReg(param)
            .size = RegisterSize(.Register)
            .PType = ParamReg
            .Forced = True
        ElseIf IsNumeric(param) Then
            .Value = CLng(param)
            .size = GetFirstSetBit(SizesForInt(.Value))
            .PType = ParamImm
            .Forced = True
        ElseIf IsFPUReg(param) Then
            .FPURegister = FPURegStrToNum(param)
            .size = Bits32 Or Bits64 Or Bits80
            .PType = ParamSTX
            .Forced = True
        Else
            If LCase$(Left$(param, 2)) = "mm" Or LCase$(Left$(param, 3)) = "xmm" Then
                .MMRegister = MMRegStrToNum(param)
                .PType = ParamMM
                
                If InStr(param, "/mem") > 0 Then
                    .PType = .PType Or ParamMem
                Else
                    If (.MMRegister And MM0) Then
                        .size = Bits64
                    Else
                        .size = Bits128
                    End If
                End If
            Else
                Select Case LCase$(Left$(param, 3))
                    Case "imm": .PType = ParamImm
                    Case "reg": .PType = ParamReg
                    Case "rel": .PType = ParamRel
                    Case "mem": .PType = ParamMem
                    Case "rem": .PType = ParamMem Or ParamReg
                End Select
            End If
            
            Select Case Right$(param, 2)
                Case "08":  .size = Bits8
                Case "16":  .size = Bits16
                Case "32":  .size = Bits32
                Case "64":  .size = Bits64
                Case "80":  .size = Bits80
            End Select
            
            Select Case Right$(param, 3)
                Case "128": .size = Bits128
            End Select
        End If
    End With
End Function


Private Function ParseOpCode(Op As String) As OpCode
    Dim i       As Long
    Dim strH    As String
    
    With ParseOpCode
        .RegOpExt = -1
        
        For i = 1 To Len(Op)
            Select Case Mid$(Op, i, 1)
                Case "0" To "9":    strH = strH & Mid$(Op, i, 1)
                Case "A" To "F":    strH = strH & Mid$(Op, i, 1)
                Case "a" To "f":    strH = strH & Mid$(Op, i, 1)
                Case "/":           .RegOpExt = CLng(Mid$(Op, i + 1, 1))
                                    i = i + 1
            End Select
            
            If Len(strH) = 2 Then
                .Bytes(.ByteCount) = CLng("&H" & strH)
                .ByteCount = .ByteCount + 1
                strH = ""
            End If
        Next
    End With
End Function


Public Function SizesForInt(ByVal lngVal As Long) As ParamSize
    If (lngVal >= -128 And lngVal <= 255) Then
        SizesForInt = Bits8 Or Bits16 Or Bits32 Or Bits64 Or Bits80
    ElseIf (lngVal >= -32768 And lngVal <= 65535) Then
        SizesForInt = Bits16 Or Bits32 Or Bits64 Or Bits80
    Else
        SizesForInt = Bits32 Or Bits64 Or Bits80
    End If
End Function


Public Function IsRawDataOp(strOp As String) As Boolean
    Dim i As Long
    
    For i = 0 To UBound(m_strRawData)
        If StrComp(m_strRawData(i), strOp, vbTextCompare) = 0 Then
            IsRawDataOp = True
            Exit Function
        End If
    Next
End Function


Public Function IsMMReg(strReg As String) As Boolean
    Dim i As Long
    
    For i = 0 To UBound(m_strMMRegs)
        If StrComp(m_strMMRegs(i), strReg, vbTextCompare) = 0 Then
            IsMMReg = True
            Exit Function
        End If
    Next
End Function


Public Function IsFPUReg(strReg As String) As Boolean
    Dim i As Long
    
    For i = 0 To UBound(m_strFPURegs)
        If StrComp(m_strFPURegs(i), strReg, vbTextCompare) = 0 Then
            IsFPUReg = True
            Exit Function
        End If
    Next
End Function


Public Function IsKeyword(strKey As String) As Boolean
    Dim i As Long
    
    For i = 0 To UBound(m_strKeywords)
        If StrComp(m_strKeywords(i), strKey, vbTextCompare) = 0 Then
            IsKeyword = True
            Exit Function
        End If
    Next
End Function


Public Function IsSegmentReg(strSeg As String) As Boolean
    Dim i As Long
    
    For i = 0 To UBound(m_strSegments)
        If StrComp(m_strSegments(i), strSeg, vbTextCompare) = 0 Then
            IsSegmentReg = True
            Exit Function
        End If
    Next
End Function


Public Function IsRegister(strReg As String) As Boolean
    Dim i As Long
    
    For i = 0 To UBound(m_strRegisters)
        If StrComp(m_strRegisters(i), strReg, vbTextCompare) = 0 Then
            IsRegister = True
            Exit Function
        End If
    Next
End Function


Public Function FPURegStrToNum(strST As String) As ASMFPURegisters
    If UCase$(Left$(strST, 2)) = "ST" Then
        FPURegStrToNum = CLng(Mid$(strST, 3, 1))
    Else
        FPURegStrToNum = FP_UNKNOWN
    End If
End Function


Public Function MMRegStrToNum(strMM As String) As ASMXMMRegisters
    Dim udeBase As ASMXMMRegisters
    Dim strNum  As String
    Dim i       As Long
    
    If UCase$(Left$(strMM, 3)) = "XMM" Then
        udeBase = XMM0
        strNum = Mid$(strMM, 4, 1)
    ElseIf UCase$(Left$(strMM, 2)) = "MM" Then
        udeBase = MM0
        strNum = Mid$(strMM, 3, 1)
    Else
        Err.Raise 12345, , "Kein (X)MM Register"
    End If
    
    If IsNumeric(strNum) Then
        udeBase = udeBase * 2 ^ CInt(strNum)
    Else
        If udeBase = XMM0 Then
            udeBase = XMM0 Or XMM1 Or XMM2 Or XMM3 Or XMM4 Or XMM5 Or XMM6 Or XMM7
        Else
            udeBase = MM0 Or MM1 Or MM2 Or MM3 Or MM4 Or MM5 Or MM6 Or MM7
        End If
    End If
    
    MMRegStrToNum = udeBase
End Function


' offset to add to the opcode of a conditional instruction for a specific condition
Public Function ConditionOffset(cc As String) As Long
    Select Case UCase$(cc)
        Case "O":               ConditionOffset = 0
        Case "NO":              ConditionOffset = 1
        Case "B", "C", "NAE":   ConditionOffset = 2
        Case "AE", "NB", "NC":  ConditionOffset = 3
        Case "E", "Z":          ConditionOffset = 4
        Case "NE", "NZ":        ConditionOffset = 5
        Case "BE", "NA":        ConditionOffset = 6
        Case "A", "NBE":        ConditionOffset = 7
        Case "S":               ConditionOffset = 8
        Case "NS":              ConditionOffset = 9
        Case "P", "PE":         ConditionOffset = 10
        Case "NP", "PO":        ConditionOffset = 11
        Case "L", "NGE":        ConditionOffset = 12
        Case "GE", "NL":        ConditionOffset = 13
        Case "LE", "NG":        ConditionOffset = 14
        Case "G", "NLE":        ConditionOffset = 15
    End Select
End Function


Public Function GetRegExtRegName(ByVal Offset As Long, ByVal size As ParamSize) As String
    Select Case size
        Case Bits8, Bits16, Bits32:
        Case Else:  Err.Raise 123456, , "GetRegExtRegName: invalid size"
    End Select
    
    Select Case Offset
        Case 0:
            Select Case size
                Case Bits8:     GetRegExtRegName = "AL"
                Case Bits16:    GetRegExtRegName = "AX"
                Case Bits32:    GetRegExtRegName = "EAX"
            End Select
        Case 1:
            Select Case size
                Case Bits8:     GetRegExtRegName = "CL"
                Case Bits16:    GetRegExtRegName = "CX"
                Case Bits32:    GetRegExtRegName = "ECX"
            End Select
        Case 2:
            Select Case size
                Case Bits8:     GetRegExtRegName = "DL"
                Case Bits16:    GetRegExtRegName = "DX"
                Case Bits32:    GetRegExtRegName = "EDX"
            End Select
        Case 3:
            Select Case size
                Case Bits8:     GetRegExtRegName = "BL"
                Case Bits16:    GetRegExtRegName = "BX"
                Case Bits32:    GetRegExtRegName = "EBX"
            End Select
        Case 4:
            Select Case size
                Case Bits8:     GetRegExtRegName = "AH"
                Case Bits16:    GetRegExtRegName = "SP"
                Case Bits32:    GetRegExtRegName = "ESP"
            End Select
        Case 5:
            Select Case size
                Case Bits8:     GetRegExtRegName = "CH"
                Case Bits16:    GetRegExtRegName = "BP"
                Case Bits32:    GetRegExtRegName = "EBP"
            End Select
        Case 6:
            Select Case size
                Case Bits8:     GetRegExtRegName = "DH"
                Case Bits16:    GetRegExtRegName = "SI"
                Case Bits32:    GetRegExtRegName = "ESI"
            End Select
        Case 7:
            Select Case size
                Case Bits8:     GetRegExtRegName = "BH"
                Case Bits16:    GetRegExtRegName = "DI"
                Case Bits32:    GetRegExtRegName = "EDI"
            End Select
        Case Else:
            Err.Raise 12345, , "GetRegExtRegName: Invalid offset"
    End Select
End Function


Public Function RegisterSize(reg As ASMRegisters) As ParamSize
    Select Case reg
        Case RegAL, RegAH, RegBL, RegBH, RegCL, RegCH, RegDL, RegDH:
            RegisterSize = Bits8
        Case RegAX, RegBX, RegCX, RegDX, RegBP, RegSP, RegDI, RegSI:
            RegisterSize = Bits16
        Case RegEAX, RegEBX, RegECX, RegEDX, RegEBP, RegESP, RegEDI, RegESI:
            RegisterSize = Bits32
        Case Else:
            RegisterSize = BitsUnknown
    End Select
End Function


Public Function SegStrToSeg(strSeg As String) As ASMSegmentRegs
    Select Case LCase$(strSeg)
        Case "cs":  SegStrToSeg = SegCS
        Case "ds":  SegStrToSeg = SegDS
        Case "es":  SegStrToSeg = SegES
        Case "fs":  SegStrToSeg = SegFS
        Case "gs":  SegStrToSeg = SegGS
        Case "ss":  SegStrToSeg = SegSS
        Case Else:  SegStrToSeg = SegUnknown
    End Select
End Function


Public Function IdxToReg(ByVal idx As Long) As ASMRegisters
    Select Case idx
        Case 0:     IdxToReg = RegAL
        Case 1:     IdxToReg = RegBL
        Case 2:     IdxToReg = RegCL
        Case 3:     IdxToReg = RegDL
        Case 4:     IdxToReg = RegAH
        Case 5:     IdxToReg = RegBH
        Case 6:     IdxToReg = RegCH
        Case 7:     IdxToReg = RegDH
        Case 8:     IdxToReg = RegBP
        Case 9:     IdxToReg = RegSP
        Case 10:    IdxToReg = RegDI
        Case 11:    IdxToReg = RegSI
        Case 12:    IdxToReg = RegAX
        Case 13:    IdxToReg = RegBX
        Case 14:    IdxToReg = RegCX
        Case 15:    IdxToReg = RegDX
        Case 16:    IdxToReg = RegEBP
        Case 17:    IdxToReg = RegESP
        Case 18:    IdxToReg = RegEDI
        Case 19:    IdxToReg = RegESI
        Case 20:    IdxToReg = RegEAX
        Case 21:    IdxToReg = RegEBX
        Case 22:    IdxToReg = RegECX
        Case 23:    IdxToReg = RegEDX
        Case Else:  IdxToReg = RegUnknown
    End Select
End Function


Public Function RegToIdx(ByVal udeReg As ASMRegisters) As Long
    Select Case udeReg
        Case RegAL:     RegToIdx = 0
        Case RegBL:     RegToIdx = 1
        Case RegCL:     RegToIdx = 2
        Case RegDL:     RegToIdx = 3
        Case RegAH:     RegToIdx = 4
        Case RegBH:     RegToIdx = 5
        Case RegCH:     RegToIdx = 6
        Case RegDH:     RegToIdx = 7
        Case RegBP:     RegToIdx = 8
        Case RegSP:     RegToIdx = 9
        Case RegDI:     RegToIdx = 10
        Case RegSI:     RegToIdx = 11
        Case RegAX:     RegToIdx = 12
        Case RegBX:     RegToIdx = 13
        Case RegCX:     RegToIdx = 14
        Case RegDX:     RegToIdx = 15
        Case RegEBP:    RegToIdx = 16
        Case RegESP:    RegToIdx = 17
        Case RegEDI:    RegToIdx = 18
        Case RegESI:    RegToIdx = 19
        Case RegEAX:    RegToIdx = 20
        Case RegEBX:    RegToIdx = 21
        Case RegECX:    RegToIdx = 22
        Case RegEDX:    RegToIdx = 23
        Case Else:      RegToIdx = -1
    End Select
End Function


Public Function RegStrToReg(strReg As String) As ASMRegisters
    Select Case LCase$(strReg)
        Case "al":  RegStrToReg = RegAL
        Case "ah":  RegStrToReg = RegAH
        Case "bl":  RegStrToReg = RegBL
        Case "bh":  RegStrToReg = RegBH
        Case "cl":  RegStrToReg = RegCL
        Case "ch":  RegStrToReg = RegCH
        Case "dl":  RegStrToReg = RegDL
        Case "dh":  RegStrToReg = RegDH
        Case "sp":  RegStrToReg = RegSP
        Case "bp":  RegStrToReg = RegBP
        Case "si":  RegStrToReg = RegSI
        Case "di":  RegStrToReg = RegDI
        Case "ax":  RegStrToReg = RegAX
        Case "bx":  RegStrToReg = RegBX
        Case "cx":  RegStrToReg = RegCX
        Case "dx":  RegStrToReg = RegDX
        Case "eax": RegStrToReg = RegEAX
        Case "ebx": RegStrToReg = RegEBX
        Case "ecx": RegStrToReg = RegECX
        Case "edx": RegStrToReg = RegEDX
        Case "esp": RegStrToReg = RegESP
        Case "ebp": RegStrToReg = RegEBP
        Case "esi": RegStrToReg = RegESI
        Case "edi": RegStrToReg = RegEDI
        Case Else:  RegStrToReg = RegUnknown
    End Select
End Function


Private Function RemoveWSDoubles(strWS As String) As String
    Do While InStr(strWS, "  ")
        strWS = Replace(strWS, "  ", " ")
    Loop
    
    RemoveWSDoubles = strWS
End Function


Public Function BitCount(ByVal num As Long) As Long
    Dim b   As Long
    Dim i   As Long
    Dim c   As Long
    
    b = 1
    
    For i = 0 To 30
        If num And b Then c = c + 1
        If i < 30 Then b = b * 2
    Next
    If num And &H80000000 Then c = c + 1
    
    BitCount = c
End Function


Public Function IsDefinite(ByVal num As Long) As Boolean
    Dim b   As Long
    Dim i   As Long
    Dim c   As Long
    
    b = 1
    
    For i = 0 To 30
        If num And b Then
            c = c + 1
            If c > 1 Then Exit For
        End If
        
        If i < 30 Then b = b * 2
    Next
    If num And &H80000000 Then c = c + 1
    
    IsDefinite = c = 1
End Function


Public Function GetFirstSetBitIdx(ByVal num As Long) As Long
    Dim i   As Long
    Dim b   As Long
    
    b = 1
    
    For i = 0 To 30
        If num And b Then
            GetFirstSetBitIdx = i
            Exit Function
        Else
            If i < 30 Then b = b * 2
        End If
    Next
    If num And &H80000000 Then GetFirstSetBitIdx = 31
End Function


Public Function GetFirstSetBit(ByVal num As Long) As Long
    Dim i   As Long
    Dim b   As Long
    
    b = 1
    
    For i = 0 To 30
        If num And b Then
            GetFirstSetBit = b
            Exit Function
        Else
            b = b * 2
        End If
    Next
    GetFirstSetBit = num And &H80000000
End Function


Public Function GetTokenName(ByVal tk As TokenType) As String
    Select Case tk
        Case TokenBeginOfInput:     GetTokenName = "BeginOfInp"
        Case TokenBracketLeft:      GetTokenName = "BracketLeft"
        Case TokenBracketRight:     GetTokenName = "BracketRight"
        Case TokenEndOfInput:       GetTokenName = "EndOfInp"
        Case TokenEndOfInstruction: GetTokenName = "EndOfInstr"
        Case TokenInvalid:          GetTokenName = "Invalid "
        Case TokenKeyword:          GetTokenName = "Keyword"
        Case TokenOpAdd:            GetTokenName = "OpAdd"
        Case TokenOpColon:          GetTokenName = "Colon"
        Case TokenOperator:         GetTokenName = "Operator"
        Case TokenOpMul:            GetTokenName = "OpMul"
        Case TokenOpSub:            GetTokenName = "OpSub"
        Case TokenRegister:         GetTokenName = "Register"
        Case TokenSegmentReg:       GetTokenName = "SegmentReg"
        Case TokenFPUReg:           GetTokenName = "FPUReg"
        Case TokenSeparator:        GetTokenName = "Separator"
        Case TokenSymbol:           GetTokenName = "Symbol"
        Case TokenUnknown:          GetTokenName = "Unknown"
        Case TokenValue:            GetTokenName = "Value"
        Case TokenMMRegister:       GetTokenName = "MMReg"
        Case Else:                  GetTokenName = "??????"
    End Select
End Function
