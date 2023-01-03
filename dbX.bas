Attribute VB_Name = "databaseX"
'This is the new version for ADO.
Option Explicit
'---- CursorTypeEnum Values ----
'Const adOpenForwardOnly = 0
'Const adOpenKeyset = 1
'Const adOpenDynamic = 2
'Const adOpenStatic = 3

'---- LockTypeEnum Values ----
'Const adLockReadOnly = 1
'Const adLockPessimistic = 2
'Const adLockOptimistic = 3
'Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
'Const adUseServer = 2
'Const adUseClient = 3
'ActiveX Data Objects (ADO)
Const adAddNew = &H1000400
Const adAffectAllChapters = 4
Const adAffectCurrent = 1
Const adAffectGroup = 2
Const adApproxPosition = &H4000
Const adArray = &H2000
Const adAsyncConnect = &H10
Const adAsyncExecute = &H10
Const adAsyncFetch = &H20
Const adAsyncFetchNonBlocking = &H40
Const adBigInt = 20
Const adBinary = 128
Const adBookmark = &H2000
Const adBookmarkCurrent = 0
Const adBookmarkFirst = 1
Const adBookmarkLast = 2
Const adBoolean = 11
Const adBSTR = 8
Const adChapter = 136
Const adChar = 129
Const adClipString = 2
Const adCmdFile = &H100
Const adCmdStoredProc = &H4
Const adCmdTable = &H2
Const adCmdTableDirect = &H200
Const adCmdText = &H1
Const adCmdUnknown = &H8
Const adCollectionRecord = 1
Const adCompareEqual = 1
Const adCompareGreaterThan = 2
Const adCompareLessThan = 0
Const adCompareNotComparable = 4
Const adCompareNotEqual = 3
Const adCopyAllowEmulation = 4
Const adCopyNonRecursive = 2
Const adCopyOverWrite = 1
Const adCopyUnspecified = -1
Const adCR = 13
Const adCreateCollection = &H2000
Const adCreateNonCollection = &H0
Const adCreateOverwrite = &H4000000
Const adCreateStructDoc = &H80000000
Const adCriteriaAllCols = 1
Const adCriteriaKey = 0
Const adCriteriaTimeStamp = 3
Const adCriteriaUpdCols = 2
Const adCRLF = -1
Const adCurrency = 6
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adDecimal = 14
Const adDefaultStream = -1
Const adDelayFetchFields = &H8000
Const adDelayFetchStream = &H4000
Const adDelete = &H1000800
Const adDouble = 5
Const adEditAdd = &H2
Const adEditDelete = &H4
Const adEditInProgress = &H1
Const adEditNone = &H0
Const adEmpty = 0
Const adErrBoundToCommand = &HE7B
Const adErrCannotComplete = &HE94
Const adErrCantChangeConnection = &HEA4
Const adErrCantChangeProvider = &HC94
Const adErrCantConvertvalue = &HE8C
Const adErrCantCreate = &HE8D
Const adErrCatalogNotSet = &HEA3
Const adErrColumnNotOnThisRow = &HE8E
Const adErrDataConversion = &HD5D
Const adErrDataOverflow = &HE89
Const adErrDelResOutOfScope = &HE9A
Const adErrDenyNotSupported = &HEA6
Const adErrDenyTypeNotSupported = &HEA7
Const adErrFeatureNotAvailable = &HCB3
Const adErrFieldsUpdateFailed = &HEA5
Const adErrIllegalOperation = &HC93
Const adErrIntegrityViolation = &HE87
Const adErrInTransaction = &HCAE
Const adErrInvalidArgument = &HBB9
Const adErrInvalidConnection = &HE7D
Const adErrInvalidParamInfo = &HE7C
Const adErrInvalidTransaction = &HE82
Const adErrInvalidURL = &HE91
Const adErrItemNotFound = &HCC1
Const adErrNoCurrentRecord = &HBCD
Const adErrNotReentrant = &HE7E
Const adErrObjectClosed = &HE78
Const adErrObjectInCollection = &HD27
Const adErrObjectNotSet = &HD5C
Const adErrObjectOpen = &HE79
Const adErrOpeningFile = &HBBA
Const adErrOperationCancelled = &HE80
Const adError = 10
Const adErrOutOfSpace = &HE96
Const adErrPermissionDenied = &HE88
Const adErrPropConflicting = &HE9E
Const adErrPropInvalidColumn = &HE9B
Const adErrPropInvalidOption = &HE9C
Const adErrPropInvalidValue = &HE9D
Const adErrPropNotAllSettable = &HE9F
Const adErrPropNotSet = &HEA0
Const adErrPropNotSettable = &HEA1
Const adErrPropNotSupported = &HEA2
Const adErrProviderFailed = &HBB8
Const adErrProviderNotFound = &HE7A
Const adErrReadFile = &HBBB
Const adErrResourceExists = &HE93
Const adErrResourceLocked = &HE92
Const adErrResourceOutOfScope = &HE97
Const adErrSchemaViolation = &HE8A
Const adErrSignMismatch = &HE8B
Const adErrStillConnecting = &HE81
Const adErrStillExecuting = &HE7F
Const adErrTreePermissionDenied = &HE90
Const adErrUnavailable = &HE98
Const adErrUnsafeOperation = &HE84
Const adErrURLDoesNotExist = &HE8F
Const adErrURLIntegrViolSetColumns = &HE8F
Const adErrURLNamedRowDoesNotExist = &HE99
Const adErrVolumeNotFound = &HE95
Const adErrWriteFile = &HBBC
Const adExecuteNoRecords = &H80
Const adFailIfNotExists = -1
Const adFieldAlreadyExists = 26
Const adFieldBadStatus = 12
Const adFieldCannotComplete = 20
Const adFieldCannotDeleteSource = 23
Const adFieldCantConvertValue = 2
Const adFieldCantCreate = 7
Const adFieldDataOverflow = 6
Const adFieldDefault = 13
Const adFieldDoesNotExist = 16
Const adFieldIgnore = 15
Const adFieldIntegrityViolation = 10
Const adFieldInvalidURL = 17
Const adFieldIsNull = 3
Const adFieldOK = 0
Const adFieldOutOfSpace = 22
Const adFieldPendingChange = &H40000
Const adFieldPendingDelete = &H20000
Const adFieldPendingInsert = &H10000
Const adFieldPendingUnknown = &H80000
Const adFieldPendingUnknownDelete = &H100000
Const adFieldPermissionDenied = 9
Const adFieldReadOnly = 24
Const adFieldResourceExists = 19
Const adFieldResourceLocked = 18
Const adFieldResourceOutOfScope = 25
Const adFieldSchemaViolation = 11
Const adFieldSignMismatch = 5
Const adFieldTruncated = 4
Const adFieldUnavailable = 8
Const adFieldVolumeNotFound = 21
Const adFileTime = 64
Const adFilterAffectedRecords = 2
Const adFilterConflictingRecords = 5
Const adFilterFetchedRecords = 3
Const adFilterNone = 0
Const adFilterPendingRecords = 1
Const adFind = &H80000
Const adFldCacheDeferred = &H1000
Const adFldFixed = &H10
Const adFldIsChapter = &H2000
Const adFldIsCollection = &H40000
Const adFldIsDefaultStream = &H20000
Const adFldIsNullable = &H20
Const adFldIsRowURL = &H10000
Const adFldKeyColumn = &H8000
Const adFldLong = &H80
Const adFldMayBeNull = &H40
Const adFldMayDefer = &H2
Const adFldNegativeScale = &H4000
Const adFldRowID = &H100
Const adFldRowVersion = &H200
Const adFldUnknownUpdatable = &H8
Const adFldUpdatable = &H4
Const adGetRowsRest = -1
Const adGUID = 72
Const adHoldRecords = &H100
Const adIDispatch = 9
Const adIndex = &H800000
Const adInteger = 3
Const adIUnknown = 13
Const adLF = 10
Const adLockBatchOptimistic = 4
Const adLockOptimistic = 3
Const adLockPessimistic = 2
Const adLockReadOnly = 1
Const adLongVarBinary = 205
Const adLongVarChar = 201
Const adLongVarWChar = 203
Const adMarshalAll = 0
Const adMarshalModifiedOnly = 1
Const adModeRead = 1
Const adModeReadWrite = 3
Const adModeRecursive = &H400000
Const adModeShareDenyNone = &H10
Const adModeShareDenyRead = 4
Const adModeShareDenyWrite = 8
Const adModeShareExclusive = &HC
Const adModeUnknown = 0
Const adModeWrite = 2
Const adMoveAllowEmulation = 4
Const adMoveDontUpdateLinks = 2
Const adMoveOverWrite = 1
Const adMovePrevious = &H200
Const adMoveUnspecified = -1
Const adNotify = &H40000
Const adNumeric = 131
Const adOpenAsync = &H1000
Const adOpenDynamic = 2
Const adOpenForwardOnly = 0
Const adOpenIfExists = &H2000000
Const adOpenKeyset = 1
Const adOpenRecordUnspecified = -1
Const adOpenSource = &H800000
Const adOpenStatic = 3
Const adOpenStreamAsync = 1
Const adOpenStreamFromRecord = 4
Const adOpenStreamUnspecified = -1
Const adParamInput = &H1
Const adParamInputOutput = &H3
Const adParamLong = &H80
Const adParamNullable = &H40
Const adParamOutput = &H2
Const adParamReturnValue = &H4
Const adParamSigned = &H10
Const adParamUnknown = &H0
Const adPersistADTG = 0
Const adPersistXML = 1
Const adPosBOF = -2
Const adPosEOF = -3
Const adPosUnknown = -1
Const adPriorityAboveNormal = 4
Const adPriorityBelowNormal = 2
Const adPriorityHighest = 5
Const adPriorityLowest = 1
Const adPriorityNormal = 3
Const adPromptAlways = 1
Const adPromptComplete = 2
Const adPromptCompleteRequired = 3
Const adPromptNever = 4
Const adPropNotSupported = &H0
Const adPropOptional = &H2
Const adPropRead = &H200
Const adPropRequired = &H1
Const adPropVariant = 138
Const adPropWrite = &H400
Const adReadAll = -1
Const adReadLine = -2
Const adRecalcAlways = 1
Const adRecalcUpFront = 0
Const adRecCanceled = &H100
Const adRecCantRelease = &H400
Const adRecConcurrencyViolation = &H800
Const adRecDBDeleted = &H40000
Const adRecDeleted = &H4
Const adRecIntegrityViolation = &H1000
Const adRecInvalid = &H10
Const adRecMaxChangesExceeded = &H2000
Const adRecModified = &H2
Const adRecMultipleChanges = &H40
Const adRecNew = &H1
Const adRecObjectOpen = &H4000
Const adRecOK = &H0
Const adRecordURL = -2
Const adRecOutOfMemory = &H8000
Const adRecPendingChanges = &H80
Const adRecPermissionDenied = &H10000
Const adRecSchemaViolation = &H20000
Const adRecUnmodified = &H8
Const adResync = &H20000
Const adResyncAllValues = 2
Const adResyncUnderlyingValues = 1
Const adRsnAddNew = 1
Const adRsnClose = 9
Const adRsnDelete = 2
Const adRsnFirstChange = 11
Const adRsnMove = 10
Const adRsnMoveFirst = 12
Const adRsnMoveLast = 15
Const adRsnMoveNext = 13
Const adRsnMovePrevious = 14
Const adRsnRequery = 7
Const adRsnResynch = 8
Const adRsnUndoAddNew = 5
Const adRsnUndoDelete = 6
Const adRsnUndoUpdate = 4
Const adRsnUpdate = 3
Const adSaveCreateNotExist = 1
Const adSaveCreateOverWrite = 2
Const adSchemaAsserts = 0
Const adSchemaCatalogs = 1
Const adSchemaCharacterSets = 2
Const adSchemaCheckConstraints = 5
Const adSchemaCollations = 3
Const adSchemaColumnPrivileges = 13
Const adSchemaColumns = 4
Const adSchemaColumnsDomainUsage = 11
Const adSchemaConstraintColumnUsage = 6
Const adSchemaConstraintTableUsage = 7
Const adSchemaCubes = 32
Const adSchemaDBInfoKeywords = 30
Const adSchemaDBInfoLiterals = 31
Const adSchemaDimensions = 33
Const adSchemaForeignKeys = 27
Const adSchemaHierarchies = 34
Const adSchemaIndexes = 12
Const adSchemaKeyColumnUsage = 8
Const adSchemaLevels = 35
Const adSchemaMeasures = 36
Const adSchemaMembers = 38
Const adSchemaPrimaryKeys = 28
Const adSchemaProcedureColumns = 29
Const adSchemaProcedureParameters = 26
Const adSchemaProcedures = 16
Const adSchemaProperties = 37
Const adSchemaProviderSpecific = -1
Const adSchemaProviderTypes = 22
Const adSchemaReferentialConstraints = 9
Const adSchemaSchemata = 17
Const adSchemaSQLLanguages = 18
Const adSchemaStatistics = 19
Const adSchemaTableConstraints = 10
Const adSchemaTablePrivileges = 14
Const adSchemaTables = 20
Const adSchemaTranslations = 21
Const adSchemaTrustees = 39
Const adSchemaUsagePrivileges = 15
Const adSchemaViewColumnUsage = 24
Const adSchemaViews = 23
Const adSchemaViewTableUsage = 25
Const adSearchBackward = -1
Const adSearchForward = 1
Const adSeek = &H400000
Const adSeekAfter = &H8
Const adSeekAfterEQ = &H4
Const adSeekBefore = &H20
Const adSeekBeforeEQ = &H10
Const adSeekFirstEQ = &H1
Const adSeekLastEQ = &H2
Const adSimpleRecord = 0
Const adSingle = 4
Const adSmallInt = 2
Const adStateClosed = &H0
Const adStateConnecting = &H2
Const adStateExecuting = &H4
Const adStateFetching = &H8
Const adStateOpen = &H1
Const adStatusCancel = &H4
Const adStatusCantDeny = &H3
Const adStatusErrorsOccurred = &H2
Const adStatusOK = &H1
Const adStatusUnwantedEvent = &H5
Const adStructDoc = 2
Const adTinyInt = 16
Const adTypeBinary = 1
Const adTypeText = 2
Const adUnsignedBigInt = 21
Const adUnsignedInt = 19
Const adUnsignedSmallInt = 18
Const adUnsignedTinyInt = 17
Const adUpdate = &H1008000
Const adUpdateBatch = &H10000
Const adUseClient = 3
Const adUserDefined = 132
Const adUseServer = 2
Const adVarBinary = 204
Const adVarChar = 200
Const adVariant = 12
Const adVarNumeric = 139
Const adVarWChar = 202
Const adWChar = 130
Const adWriteChar = 0
Const adWriteLine = 1
Const adwrnSecurityDialog = &HE85
Const adwrnSecurityDialogHeader = &HE86
Const adXactAbortRetaining = &H40000
Const adXactBrowse = &H100
Const adXactChaos = &H10
Const adXactCommitRetaining = &H20000
Const adXactCursorStability = &H1000
Const adXactIsolated = &H100000
Const adXactReadCommitted = &H1000
Const adXactReadUncommitted = &H100
Const adXactRepeatableRead = &H10000
Const adXactSerializable = &H100000
Const adXactUnspecified = &HFFFFFFFF

'ADC / ADO Constants
Const adcExecAsync = 2
Const adcExecSync = 1
Const adcFetchAsync = 3
Const adcFetchBackground = 2
Const adcFetchUpFront = 1
Const adcReadyStateComplete = 4
Const adcReadyStateInteractive = 3
Const adcReadyStateLoaded = 2
Public ArrBase As Long
Dim AABB As Long
Dim conCollection As FastCollection
Dim Init As Boolean
'  to be changed User and UserPassword
Public JetPrefixUser As String
Public JetPostfixUser As String
Public JetPrefix As String
Public JetPostfix As String
'old Microsoft.Jet.OLEDB.4.0
' Microsoft.ACE.OLEDB.12.0
Public Const JetPrefixOld = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public Const JetPostfixOld = ";Jet OLEDB:Database Password=100101;"
Public Const JetPrefixHelp = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
Public Const JetPostfixHelp = ";Jet OLEDB:Database Password=100101;"
Public DBUser As String ' '= VbNullString ' "admin"  ' or ""
Public DBUserPassword   As String ''= VbNullString
Public extDBUser As String ' '= VbNullString ' "admin"  ' or ""
Public extDBUserPassword   As String ''= VbNullString
Public DBtype As String ' can be mdb or something else
Public Const DBtypeHelp = ".mdb" 'allways help has an mdb as type"
Public Const DBSecurityOFF = ";Persist Security Info=False"

Private Declare Function MoveFileW Lib "kernel32.dll" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Public Sub KillFile(sFilenName As String)
DeleteFileW StrPtr(sFilenName)
End Sub

Public Function MoveFile(pOldPath As String, pNewPath As String)

    MoveFileW StrPtr(pOldPath), StrPtr(pNewPath)
    
End Function
Public Function isdir(f$) As Boolean
On Error Resume Next
Dim mm As New recDir
Dim lookfirst As Boolean
Dim Pad$
If f$ = vbNullString Then Exit Function
If f$ = "." Then f$ = mcd
If InStr(f$, "\..") > 0 Or f$ = ".." Or Left$(f$, 3) = "..\" Then
If Right$(f$, 1) <> "\" Then
Pad$ = ExtractPath(f$ + "\", True, True)
Else
Pad$ = ExtractPath(f$, True, True)
End If
If Pad$ = vbNullString Then
If Right$(f$, 1) <> "\" Then
Pad$ = ExtractPath(mcd + f$ + "\", True)
Else
Pad$ = ExtractPath(mcd + f$, True)
End If
End If
lookfirst = mm.isdir(Pad$)
If lookfirst Then f$ = Pad$
Else
f$ = mylcasefILE(f$)
lookfirst = mm.isdir(f$)
If Not lookfirst Then

Pad$ = mcd + f$

lookfirst = mm.isdir(Pad$)
If lookfirst Then f$ = Pad$

End If
End If
isdir = lookfirst
End Function
Public Sub fHelp(bstack As basetask, D$, Optional Eng As Boolean = False)
Static a As Long, aa As Long, where As Long, there As Long, no_par As Long, dum As Long
Dim SQL$, b$, p$, c$, gp$, R As Double, bb As Long, i As Long
Dim CD As String, doriginal$, monitor As Long, rr$
D$ = Replace(D$, " ", ChrW(160))
On Error GoTo E5
'ON ERROR GoTo 0
If Not Form4.Visible Then
monitor = FindFormSScreen(Form1)
Else
monitor = FindFormSScreen(Form4)
End If
If HelpLastWidth > ScrInfo(monitor).Width Then HelpLastWidth = -1
doriginal$ = D$
D$ = Replace(D$, "!", "")
If D$ <> "" Then If Right$(D$, 1) = "(" Then D$ = D$ + ")"
If D$ = vbNullString Or D$ = "F12" Then
D$ = vbNullString
If Right$(D$, 1) = "(" Then D$ = D$ + ")"
p$ = subHash.Show

While ISSTRINGA(p$, c$)
'IsLabelA "", c$, b$
b$ = GetName(GetStrUntil(" ", c$))

If Right$(b$, 1) = "(" Then b$ = b$ + ")"
If gp$ <> "" Then gp$ = b$ + ", " + gp$ Else gp$ = b$
Wend
If vH_title$ <> "" Then b$ = "<| " + vH_title$ + vbCrLf + vbCrLf Else b$ = vbNullString
If Eng Then
        sHelp "User Modules/Functions [F12]", b$ + gp$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
Else
        sHelp "Τμήματα/Συναρτήσεις Χρήστη [F12]", b$ + gp$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
End If
vHelp Not Form4.Visible
Exit Sub
ElseIf GetSub(D$, i) Then
GoTo conthere
ElseIf GetlocalSubExtra(D$, i) Or D$ = here$ Then
conthere:
If D$ = here$ Then i = bstack.OriginalCode
If vH_title$ <> "" Then
b$ = "<| " + vH_title$ + vbCrLf + vbCrLf
Else
If Eng Then
b$ = "<| User Modules/Functions [F12]" + vbCrLf + vbCrLf
Else
b$ = "<| Τμήματα/Συναρτήσεις Χρήστη [F12]" + vbCrLf + vbCrLf
End If
End If
If Right$(D$, 1) = ")" Then

If Eng Then c$ = "[Function]" Else c$ = "[Συνάρτηση]"
Else
If Eng Then c$ = "[Module]" Else c$ = "[Function]"
End If

Dim ss$
    ss$ = GetNextLine((SBcode(i)))
    If Left$(ss$, 10) = "'11001EDIT" Then
    
    ss$ = Mid$(SBcode(i), Len(ss$) + 3)
    Else
     ss$ = SBcode(i)
     End If
        sHelp D$, c$ + "  " + b$ + ss$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
    
        vHelp Not Form4.Visible
Exit Sub
End If

CD = App.Path
AddDirSep CD
If UseMDBHELP Then

JetPrefix = JetPrefixHelp
JetPostfix = JetPostfixHelp
DBUser = vbNullString
DBUserPassword = vbNullString

Dim sec$

p$ = Chr(34)
c$ = ","
D$ = doriginal$
If Right$(D$, 2) = "()" Then D$ = Left$(D$, Len(D$) - 1)
If Left$(D$, 1) = "#" Then
If AscW(Mid$(D$, 2, 1) + " ") < 128 Then
SQL$ = "SELECT * FROM [COMMANDS] WHERE ENGLISH >= '" + UCase(D$) + "'"
Else
SQL$ = "SELECT * FROM [COMMANDS] WHERE DESCRIPTION >= '" + myUcase(D$, True) + "'"
End If
Else
If AscW(D$ + " ") < 128 Then
SQL$ = "SELECT * FROM [COMMANDS] WHERE ENGLISH >= '" + UCase(D$) + "'"
Else
SQL$ = "SELECT * FROM [COMMANDS] WHERE DESCRIPTION >= '" + myUcase(D$, True) + "'"
End If
End If
b$ = mylcasefILE(CD + "help2000")
getrow bstack, p$ + b$ + p$ + c$ + p$ + SQL$ + p$ + ",1," + p$ + p$ + c$ + p$ + p$, False, , , True
SQL$ = p$ + b$ + p$ + c$ + p$ + "GROUP" + p$
If bstack.IsNumber(R) Then
If bstack.IsString(gp$) Then
If bstack.IsString(b$) Then
If bstack.IsString(p$) Then
If bstack.IsNumber(R) Then
getrow bstack, SQL$ + "," + CStr(1) + "," + Chr(34) + "GROUPNUM" + Chr(34) + "," + Str$(R), False, , , True
If bstack.IsNumber(R) Then
If bstack.IsNumber(R) Then
If bstack.IsString(c$) Then
' nothing
checkit:

        If Right$(gp$, 1) = "(" Then gp$ = gp$ + ")": p$ = p$ + ")"
        
        If Eng Then
            sec$ = "Identifier: " + p$ + ", Gr: " + gp$ + vbCrLf
            gp$ = p$
        Else
            sec$ = "Αναγνωριστικό: " + gp$ + ", En: " + p$ + vbCrLf
        End If
        If vH_title$ <> "" Then
            If vH_title$ = gp$ And Form4.Visible = True Then GoTo E5
        End If
        bb = InStr(b$, "__<ENG>__")
        If bb > 0 Then
            If Eng Then
            c$ = "List [" + NLtrim$(Mid$(c$, InStr(c$, ",") + 1)) + "]"
                b$ = Mid$(b$, bb + 11)
            Else
            c$ = "Λίστα [" + Mid$(c$, 1, InStr(c$, ",") - 1) + "]"
                b$ = Left$(b$, bb - 1)
            End If
            Else
             c$ = "Λίστα [" + Mid$(c$, 1, InStr(c$, ",") - 1) + "], List [" + NLtrim$(Mid$(c$, InStr(c$, ",") + 1)) + "]"
        End If
        If vH_title$ <> "" Then b$ = "<| " + vH_title$ + vbCrLf + vbCrLf + b$ Else b$ = vbCrLf + b$
        
        sHelp gp$, sec$ + c$ + "  " + b$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
    
        vHelp Not Form4.Visible
      End If
    
    End If
End If

End If
End If
End If
End If
End If
Else
If HelpFile.DocLines = 0 Then
    rr$ = mylcasefILE(CD + "help2000utf8.dat")
    

HelpFile.ReadUnicodeOrANSI rr$
If HelpFile.DocLines = 0 Then Exit Sub
a = val(HelpFile.TextParagraph(1))
aa = val(HelpFile.TextParagraph(2 + a))
where = HelpFile.FindStr(HelpFile.TextParagraph(2 + a + 1 + 2 * aa), 1, no_par, dum)
End If

If where = 0 Then Exit Sub
If HelpFile.FindStr("\" + myUcase(doriginal$, True) + "!", where, there, dum) Then
GoTo th111
ElseIf HelpFile.FindStr("\" + myUcase(doriginal$, True), where, there, dum) Then
th111:
D$ = HelpFile.TextParagraph(there)
Eng = InStr(D$, "- ") <> 0
dum = InStr(D$, "!")
If Eng Then
c$ = HelpFile.TextParagraph(1 + val(Mid$(D$, dum + 1)))
p$ = Mid$(D$, 2, dum - 2)
b$ = HelpFile.TextParagraph(2 + a + aa + val(Mid$(D$, InStr(D$, "- ") + 2)))
gp$ = Mid$(b$, 4, InStr(b$, "\") - 4)
b$ = EscapeStrToString(Mid$(b$, InStr(b$, "\") + 4))
Else
c$ = HelpFile.TextParagraph(1 + val(Mid$(D$, dum + 1)))
'c$ = Mid$(c$, 1, InStr(c$, "," + Chr$(160)) - 1)
gp$ = Mid$(D$, 2, dum - 2)
b$ = HelpFile.TextParagraph(2 + a + 1 + there - no_par)
p$ = Mid$(b$, 4, InStr(b$, "\") - 4)
b$ = EscapeStrToString(Mid$(b$, InStr(b$, "\") + 4))


End If


        If Right$(gp$, 1) = "(" Then gp$ = gp$ + ")": p$ = p$ + ")"
        
        If Eng Then
            sec$ = "Identifier: " + p$ + ", Gr: " + gp$ + vbCrLf
            gp$ = p$
        Else
            sec$ = "Αναγνωριστικό: " + gp$ + ", En: " + p$ + vbCrLf
        End If
        If vH_title$ <> "" Then
            If vH_title$ = gp$ And Form4.Visible = True Then GoTo E5
        End If
       
        If Eng Then
        c$ = "List [" + NLtrim$(Mid$(c$, InStr(c$, ",") + 1)) + "]"
            
        Else
        c$ = "Λίστα [" + Mid$(c$, 1, InStr(c$, ",") - 1) + "]"
            
        End If
          
        If vH_title$ <> "" Then b$ = "<| " + vH_title$ + vbCrLf + vbCrLf + b$ Else b$ = vbCrLf + b$
        
        sHelp gp$, sec$ + c$ + "  " + b$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
    
        vHelp Not Form4.Visible



End If
End If

E5:
JetPrefix = JetPrefixUser
JetPostfix = JetPostfixUser
DBUser = extDBUser
DBUserPassword = extDBUserPassword
Err.Clear
End Sub
Public Function inames(i As Long, Lang As Long) As String
If (i And &H3) <> 1 Then
Select Case Lang
Case 1

inames = "DESCENDING"
Case Else
inames = "ΦΘΙΝΟΥΣΑ"
End Select
Else
Select Case Lang
Case 1
inames = "ASCENDING"
Case Else
inames = "ΑΥΞΟΥΣΑ"
End Select

End If

End Function
Public Function fnames(i As Long, Lang As Long) As String
Select Case i
Case 1  '.........11
    Select Case Lang
    Case 1
    fnames = "BOOLEAN"
    Case Else
     fnames = "ΛΟΓΙΚΟΣ"
    End Select
    Exit Function
Case 2  ' ..........16
    Select Case Lang
    Case 1
    fnames = "BYTE"
    Case Else
     fnames = "ΨΗΦΙΟ"
    End Select
   Exit Function

Case 3  '............2
        Select Case Lang
    Case 1
    fnames = "INTEGER"
    Case Else
     fnames = "ΑΚΕΡΑΙΟΣ"
    End Select
   Exit Function
Case 4 '..........3
        Select Case Lang
    Case 1
    fnames = "LONG"
    Case Else
     fnames = "ΜΑΚΡΥΣ"
    End Select
   Exit Function
 
Case 5 ' .............6
        Select Case Lang
    Case 1
    fnames = "CURRENCY"
    Case Else
     fnames = "ΛΟΓΙΣΤΙΚΟΣ"
    End Select
   Exit Function

Case 6 ' ............4
    Select Case Lang
    Case 1
    fnames = "SINGLE"
    Case Else
     fnames = "ΑΠΛΟΣ"
    End Select
   Exit Function

Case 7 '...........5
    Select Case Lang
    Case 1
    fnames = "DOUBLE"
    Case Else
     fnames = "ΔΙΠΛΟΣ"
    End Select
   Exit Function
Case 8 ' ..........7
    Select Case Lang
    Case 1
    fnames = "DATEFIELD"
    Case Else
     fnames = "ΗΜΕΡΟΜΗΝΙΑ"
    End Select
   Exit Function
Case 9 '.....................128
    Select Case Lang
    Case 1
    fnames = "BINARY"
    Case Else
     fnames = "ΔΥΑΔΙΚΟ"
    End Select
   Exit Function
Case 10 '..........................................202
    Select Case Lang
    Case 1
    fnames = "TEXT"
    Case Else
     fnames = "ΚΕΙΜΕΝΟ"
    End Select
   Exit Function
Case 11 '...........205
    fnames = "OLE"
    Exit Function
Case 12 '...........................202
    Select Case Lang
    Case 1
    fnames = "MEMO"
    Case Else
     fnames = "ΥΠΟΜΝΗΜΑ"
    End Select
Case Else
fnames = Trim$(Str$(i))
End Select
End Function

Public Function NewBase(bstackstr As basetask, R$, Lang As Long) As Boolean
Dim base As String, othersettings As String, vv
If FastSymbol(R$, "1") Then
ArrBase = 1
NewBase = True
Exit Function
ElseIf FastSymbol(R$, "0") Then
ArrBase = 0
NewBase = True
Exit Function
End If
If IsLabelSymbolNew(R$, "ΤΟΠΙΚΗ", "LOCAL", Lang) Then
'
If Not IsStrExp(bstackstr, R$, base, False) Then
MissStringExpr
Exit Function
Else
If getone(base, vv) Then
Set vv = New Mk2Base
changeone base, vv
Else
Set vv = New Mk2Base
PushOne base, vv
End If
NewBase = True
End If
Else
If Not IsStrExp(bstackstr, R$, base, False) Then
MissStringExpr
Exit Function
End If
If FastSymbol(R$, ",") Then
If Not IsStrExp(bstackstr, R$, othersettings, False) Then
MissStringExpr
Exit Function ' make it to give error
End If
End If
 On Error Resume Next
 If Left$(base, 1) = "(" Or JetPostfix = ";" Then Exit Function ' we can't create in ODBC
If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
If ExtractType(base) = vbNullString Then base = base + ".mdb"

If CFname((base)) <> "" Then
 If Not CanKillFile(base) Then FilePathNotForUser: Exit Function
' check to see if is our
RemoveOneConn base
If CheckMine(base) Then
KillFile base
Err.Clear
Else
MyEr "Can 't delete the Base", "Δεν μπορώ να διαγράψω τη βάση"

Exit Function
End If
End If

CreateObject("ADOX.Catalog").create (JetPrefix + base + JetPostfix + othersettings)   'create a new, empty *.mdb-File
NewBase = True
End If
End Function
Public Function TABLENAMES_mk2(base As String, bstackstr As basetask, R$, Lang As Long) As Boolean
Dim vv, tables As FastCollection, stac1 As New mStiva, pppp As mArray, i As Long
Dim mb As Mk2Base, Tablename As String, fieldlist As FastCollection, ma As mArray, j As Long
Dim indexField As FastCollection, fields As Long
If getone(base, vv) Then
Set mb = vv

Set tables = mb.tables()
If FastSymbol(R$, ",") Then
    If Not IsStrExp(bstackstr, R$, Tablename, False) Then
    MissStringExpr
    Exit Function
    End If
    If Not tables.ExistKey(myUcase(Tablename, True)) Then
    MyEr "table not exist", "ο πίνακας δεν υπάρχει"
    Exit Function
    End If
    If tables.sValue <> 0 Then
        MyEr "table not exist", "ο πίνακας δεν υπάρχει"
        Exit Function
    End If
    Set pppp = tables.ValueObj
    
    Set fieldlist = pppp.item(1)
    For i = 0 To fieldlist.Count - 1
        fieldlist.index = i
        fieldlist.Done = True
        If fieldlist.sValue = -1 Then
            fields = fields + 1
            Set ma = fieldlist.ValueObj
            stac1.DataStr ma.item(0)
            stac1.DataStr fnames(ma.item(1), Lang)
            stac1.DataVal ma.item(2)
            
            fieldlist.Done = False
        
        End If
    Next i
    Set indexField = pppp.item(3)
    
    If indexField.Count > 0 Then
    stac1.DataVal indexField.Count
    For j = 0 To pppp.item(3).Count - 1
    indexField.index = j
    indexField.Done = True
    
   
    stac1.DataStr (Split(indexField.Value, " ")(0))
    stac1.DataStr inames(-CLng(1 <> (val(Split(indexField.Value, " ")(1) And 1))), Lang)
    
    indexField.Done = False
    Next j
    End If
stac1.PushVal indexField.Count
stac1.PushVal fields
Else
For i = 0 To tables.Count - 1
tables.index = i
tables.Done = True
If tables.sValue = 0 Then
Set pppp = tables.ValueObj
    stac1.DataStr pppp.item(20)
    
    stac1.DataVal pppp.item(3).Count
End If
tables.Done = False
Next i
stac1.PushVal stac1.Count \ 2
End If

End If
bstackstr.soros.MergeTop stac1
TABLENAMES_mk2 = True

End Function
Public Function TABLENAMES(base As String, bstackstr As basetask, R$, Lang As Long) As Boolean
Dim Tablename As String, scope As Long, cnt As Long, srl As Long, stac1 As New mStiva, vv
Dim myBase  ' variant
scope = 1
If Len(base) > 0 Then
If getone(base, vv) Then
If TypeOf vv Is Mk2Base Then
TABLENAMES = TABLENAMES_mk2(base, bstackstr, R$, Lang)
Exit Function
End If
End If
End If
If FastSymbol(R$, ",") Then
If IsStrExp(bstackstr, R$, Tablename, False) Then
scope = 2

End If
End If


    Dim vindx As Boolean

    On Error Resume Next
            If Left$(base, 1) = "(" Or JetPostfix = ";" Then
        'skip this
        Else
            If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
            If ExtractType(base) = vbNullString Then base = base + ".mdb"
            If Not CanKillFile(base) Then FilePathNotForUser: Exit Function
        End If
    If True Then
        On Error Resume Next
        If Not getone(base, myBase) Then
            Set myBase = CreateObject("ADODB.Connection")
            If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                srl = DriveSerial(Left$(base, 3))
                If srl = 0 And Not GetDosPath(base) = vbNullString Then
                    If Lang = 0 Then
                        If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " + ExtractName(base, True)) = vbCancel Then Exit Function
                    Else
                        If Not ask("Put CD/Disk with file " + ExtractName(base, True)) = vbCancel Then Exit Function
                    End If
                End If
                If myBase = vbNullString Then
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.open JetPrefix + JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Function
                        End If
                    Else
                        myBase.open JetPrefix + GetDosPath(base) + ";Mode=Share Deny Write" + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF      'open the Connection
                    End If
                End If
                If Err.Number > 0 Then
                    Do While srl <> DriveSerial(Left$(base, 3))
                        If Lang = 0 Then
                            If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " + CStr(srl) + " στον οδηγό " + Left$(base, 1)) = vbCancel Then Exit Do
                        Else
                            If ask("Put CD/Disk with serial number " + CStr(srl) + " in drive " + Left$(base, 1)) = vbCancel Then Exit Do
                        End If
                    Loop
                    If srl = DriveSerial(Left$(base, 3)) Then
                        Err.Clear
                        If myBase = vbNullString Then myBase.open JetPrefix + GetDosPath(base) + ";Mode=Share Deny Write" + JetPostfix + "User Id=" + DBUser + ";Password=" + DBSecurityOFF       'open the Connection
                    End If
                End If
            Else
                If myBase = vbNullString Then
                ' check if we have ODBC
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.open JetPrefix + JetPostfix
                        If Err.Number Then
                            MyEr Err.Description, Err.Description
                            Exit Function
                        End If
                    Else
                        Err.Clear
                        myBase.open JetPrefix + GetDosPath(base) + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.Clear
                           myBase.open JetPrefixOld + GetDosPath(base) + JetPostfixOld + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                End If
        End If
        If Err.Number > 0 Then GoTo g102
        PushOne base, myBase
    End If
  Dim cat, TBL, rs
     Dim i As Long, j As Long, k As Long, KB As Boolean
  
           Set rs = CreateObject("ADODB.Recordset")
        Set TBL = CreateObject("ADOX.TABLE")
           Set cat = CreateObject("ADOX.Catalog")
           Set cat.ActiveConnection = myBase
           If cat.ActiveConnection.errors.Count > 0 Then
           MyEr "Can't connect to Base", "Δεν μπορώ να συνδεθώ με τη βάση"
           Exit Function
           End If
        If cat.tables.Count > 0 Then
        For Each TBL In cat.tables
        
        If TBL.Type = "TABLE" Then
        vindx = False
        KB = False
        If scope <> 2 Then
        
        cnt = cnt + 1
                            stac1.DataStr TBL.Name
                       If TBL.indexes.Count > 0 Then
                                         For j = 0 To TBL.indexes.Count - 1
                                                   With TBL.indexes(j)
                                                   If (.unique = False) And (.indexnulls = 0) Then
                                                        KB = True
                                                  Exit For
             '
                                                       End If
                                                   End With
                                                Next j
                                              If KB Then
                    
                                                     stac1.DataVal CDbl(1)
                                                     
                                                Else
                                                    stac1.DataVal CDbl(0)
                                                End If
                                               
                                           
                                            Else
                                            stac1.DataVal CDbl(0)
                                        End If
         ElseIf Tablename = TBL.Name Then
         cnt = 1
                     rs.open "Select * From [" + TBL.Name + "] ;", myBase, 3, 4 'adOpenStatic, adLockBatchOptimistic
                                         stac1.Flush
                                        stac1.DataVal CDbl(rs.fields.Count)
                                        If TBL.indexes.Count > 0 Then
                                         For j = 0 To TBL.indexes.Count - 1
                                                   With TBL.indexes(j)
                                                   If (.unique = False) And (.indexnulls = 0) Then
                                                   vindx = True
                                                   Exit For
                                                       End If
                                                   End With
                                                Next j
                                                If vindx Then
                                                
                                                     stac1.DataVal CDbl(1)
                                                Else
                                                    stac1.DataVal CDbl(0)
                                                End If
                                            Else
                                            stac1.DataVal CDbl(0)
                                        End If
                     For i = 0 To rs.fields.Count - 1
                     With rs.fields(i)
                             stac1.DataStr .Name
                             If .Type = 203 And .DEFINEDSIZE >= 536870910# Then
                             
                                         If Lang = 1 Then
                                        stac1.DataStr "MEMO"
                                        Else
                                        stac1.DataStr "ΥΠΟΜΝΗΜΑ"
                                        End If
                                        
                                        stac1.DataVal CDbl(0)
                            
                             ElseIf .Type = 205 Then
                                       
                                            stac1.DataStr "OLE"
                                       
                                       
                                            stac1.DataVal CDbl(0)
                                     ElseIf .Type = 202 And .DEFINEDSIZE <> 536870910# Then
                                            If Lang = 1 Then
                                            stac1.DataStr "TEXT"
                                            Else
                                            stac1.DataStr "ΚΕΙΜΕΝΟ"
                                            End If
                                            stac1.DataVal CDbl(.DEFINEDSIZE)
                                    
                             Else
                                        stac1.DataStr ftype(.Type, Lang)
                                        On Error GoTo 10000
                                        If .properties("ISAUTOINCREMENT") Then
                                        stac1.DataVal CDbl(-1)
                                        Else
10000
                                        stac1.DataVal CDbl(.DEFINEDSIZE)
                                        On Error Resume Next
                                        End If

                                        
                             End If
                     End With
                     Next i
                     rs.Close
                     If vindx Then
                    If TBL.indexes.Count > 0 Then
                             For j = 0 To TBL.indexes.Count - 1
                          With TBL.indexes(j)
                          If (.unique = False) And (.indexnulls = 0) Then
                          stac1.DataVal CDbl(.Columns.Count)
                          For k = 0 To .Columns.Count - 1
                            stac1.DataStr .Columns(k).Name
                             stac1.DataStr inames(.Columns(k).sortorder, Lang)
                          Next k
                             Exit For
                             
                             End If
                          End With
                       Next j
                    End If
                     End If
             End If
             End If
            
                                     
                         
               Next TBL
               Set TBL = Nothing
    End If
    If scope = 1 Then
    stac1.PushVal CDbl(cnt)
    Else
    If cnt = 0 Then
     MyEr "No such TABLE in DATABASE", "Δεν υπάρχει τέτοιο αρχείο στη βάση δεδομένων"
    End If
    End If
     bstackstr.soros.MergeTop stac1
     TABLENAMES = True
     Else
     RemoveOneConn myBase
     MyEr "No such DATABASE", "Δεν υπάρχει τέτοια βάση δεδομένων"
    End If
g102:
End Function

Public Function append_table(bstackstr As basetask, base As String, R$, ed As Boolean, Optional Lang As Long = -1) As Boolean
Dim table$, i&, par$, ok As Boolean, TT, t As Double, j&, vv, p_acc As mArray, acc As mArray
Dim gindex As Long
Dim mb As Mk2Base, tables As FastCollection, pppp As mArray, Temp As mArray, fieldlist As FastCollection
Dim indexField As FastCollection, allkey$, prevkey$

 ok = False
If getone(base, vv) Then
If Not TypeOf vv Is Mk2Base Then GoTo noMk2
Else
 GoTo noMk2
End If
If FastSymbol(R$, ",", True) Then
    If IsStrExp(bstackstr, R$, table$, False) Then
    
    If Lang <> -1 Then If IsLabelSymbolNew(R$, "ΣΤΟ", "TO", Lang) Then If IsExp(bstackstr, R$, t) Then gindex = CLng(t) Else SyntaxError

    
        If FastSymbol(R$, ",", True) Then
        ok = True
        End If
    Else
        MissStringExpr
    End If
End If
If Not ok Then Exit Function


Set mb = vv
Set tables = mb.tables
If tables.ExistKey(myUcase(table$, True)) Then
Set pppp = tables.ValueObj
If gindex > 0 Then
Set tables = pppp.item(2)
Set indexField = pppp.item(3)
Set fieldlist = pppp.item(4)
If indexField.Count > 0 Then
    If gindex > fieldlist.Count + 1 Then gindex = fieldlist.Count + 1
    fieldlist.index = gindex - 1
    fieldlist.Done = True
    tables.index = fieldlist.Value
    tables.Done = True
    fieldlist.Done = False
    Set p_acc = tables.ValueObj
    prevkey$ = fieldlist.KeyToString
    p_acc.CopyArray acc
    tables.Done = False
Else
    If gindex > tables.Count + 1 Then gindex = tables.Count + 1
    tables.index = gindex - 1
    tables.Done = True
    Set p_acc = tables.ValueObj
    p_acc.CopyArray acc
End If

Else
Set p_acc = pppp.item(18)
p_acc.CopyArray acc
End If
Set p_acc = pppp.item(17)
i& = 0
Set fieldlist = pppp.item(1)
Do
TT = 0
fieldlist.index = p_acc.item(i&)
fieldlist.Done = True
Set Temp = fieldlist.ValueObj

If FastSymbol(R$, ",") Then
If i& = 0 And Temp.item(4) Then
    If Not ed Then
        acc.item(p_acc.item(i&)) = Temp.item(5)
        Temp.item(5) = Temp.item(5) + 1
    End If
End If
ElseIf IsStrExp(bstackstr, R$, par$, False) Then
    If i& = 0 And Temp.item(4) Then
        If Not ed Then
            acc.item(p_acc.item(i&)) = Temp.item(4)
            Temp.item(4) = Temp.item(4) + 1
        End If
    ElseIf Temp.item(2) > 0 Then
        acc.item(p_acc.item(i&)) = RealLeft(par$, Temp.item(2))
    Else
        acc.item(p_acc.item(i&)) = par$
    End If
    If Not FastSymbol(R$, ",") Then Exit Do
ElseIf IsExp(bstackstr, R$, TT, False) Then
    If i& = 0 And Temp.item(4) Then
        If Not ed Then
            acc.item(p_acc.item(i&)) = Temp.item(4)
            Temp.item(4) = Temp.item(4) + 1
        End If
    Else
        acc.item(p_acc.item(i&)) = TT
    End If
    If Not FastSymbol(R$, ",") Then Exit Do
End If
i& = i& + 1
Loop Until i& = p_acc.Count
' now we append to table, without checking the keys (its a fault)
Set tables = pppp.item(2)
Set indexField = pppp.item(3)
Set fieldlist = pppp.item(4)
If indexField.Count > 0 Then
For j& = 0 To indexField.Count - 1
    indexField.index = j&
    indexField.Done = True
    If Len(allkey$) = 0 Then
    allkey$ = indexField.Normalize(acc.item(indexField.sValue))
    Else
    allkey$ = allkey$ + ChrW(1) + indexField.Normalize(acc.item(indexField.sValue))
    End If
Next j&
If Len(prevkey$) > 0 Then
    If prevkey$ <> allkey$ Then
        If fieldlist.ExistKey(allkey$) Then
            MyEr "index key not unique", "Το κλειδί δεν είναι μοναδικό"
            Exit Function
        End If
        fieldlist.index = gindex - 1
        fieldlist.Done = True
        tables.index = fieldlist.Value
        ' I have to add logic here for calculating the final bytes for file
        ' so if the previous is less than the current, we can save to same plase
        ' else we have to delete this and make a new entry
        Set tables.ValueObj = acc
        fieldlist.AddKey allkey$, tables.index
        fieldlist.Remove prevkey$
        If pppp.item(5) = 2 Then
            fieldlist.SortDes
        Else
            fieldlist.Sort
        End If
    
    Else
        fieldlist.index = gindex - 1
        fieldlist.Done = True
        tables.index = fieldlist.Value
        tables.Done = True
        fieldlist.Done = False
        Set tables.ValueObj = acc
        
    End If
    append_table = True
    Exit Function
Else
    If fieldlist.ExistKey(allkey$) Then
        MyEr "index key not unique", "Το κλειδί δεν είναι μοναδικό"
        Exit Function
    End If
    fieldlist.AddKey allkey$, tables.Count
    If pppp.item(5) = 2 Then
        fieldlist.SortDes
    Else
        fieldlist.Sort
    End If
End If
End If
If gindex > 0 Then
    tables.index = gindex - 1
    tables.Done = True
    Set tables.ValueObj = acc
Else
    tables.AddKey tables.Count + 1, acc
    pppp.item(21) = pppp.item(21) + 1
End If
tables.Done = False
append_table = True
End If
Exit Function

noMk2:
If FastSymbol(R$, ",", True) Then
If IsStrExp(bstackstr, R$, table$, False) Then
ok = True
End If
End If
If Not ok Then Exit Function
If Lang <> -1 Then If IsLabelSymbolNew(R$, "ΣΤΟ", "TO", Lang) Then If IsExp(bstackstr, R$, t, False) Then gindex = CLng(t) Else SyntaxError
Dim Id$
  If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
End If





If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base + ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Function
End If
          On Error Resume Next
          Dim myBase
          
               If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Function
                Else
                If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                    myBase.open JetPrefix + JetPostfix
                    If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Function
                    End If
                Else
                        Err.Clear
                        myBase.open JetPrefix + GetDosPath(base) + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.Clear
                           myBase.open JetPrefixOld + GetDosPath(base) + JetPostfixOld + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                End If
                PushOne base, myBase
            End If
           Err.Clear
         
         '  If Err.Number > 0 Then GoTo thh
           
           
         '  Set rec = myBase.OpenRecordset(table$, dbOpenDynaset)
          Dim rec, LL$
          
           Set rec = CreateObject("ADODB.Recordset")
            Err.Clear
           rec.open Id$, myBase, 3, 4 'adOpenStatic, adLockBatchOptimistic

 If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.open = LL$
 PushOne base, myBase
 Err.Clear
rec.open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description + " " + Id$, Err.Description + " " + Id$
Exit Function
End If
End If
   
   
If ed Then
If gindex > 0 Then
Err.Clear
    rec.MoveLast
    rec.MoveFirst
    rec.absoluteposition = gindex '  - 1
    If Err.Number <> 0 Then
    MyEr "Wrong index for table " + table$, "Λάθος δείκτης για αρχείο " + table$
    End If
ElseIf rec.EOF Then
    MyEr "Record not found", "Η Εγγραφή δεν βρέθηκε"
    append_table = False
    Exit Function
Else

    rec.MoveLast
End If
' rec.Edit  no need for undo
Else
rec.AddNew
End If
i& = 0
While FastSymbol(R$, ",")
If ed Then
    While FastSymbol(R$, ",")
    i& = i& + 1
    Wend
End If
If IsStrExp(bstackstr, R$, par$) Then
    rec.fields(i&) = par$
ElseIf IsExp(bstackstr, R$, t) Then
    If vartype(t) = vbString Then
    rec.fields(i&) = t
    Else
    rec.fields(i&) = CStr(t)   '??? convert to a standard format
    End If
End If

i& = i& + 1
Wend
Err.Clear
rec.UpdateBatch  ' update be an updatebatch
If Err.Number > 0 Then
MyEr "Can't append " + Err.Description, "Αδυναμία προσθήκης:" + Err.Description
append_table = False
Else
append_table = True
End If

End Function
Public Sub getrow(bstackstr As basetask, R$, Optional ERL As Boolean = True, Optional Search$ = " = ", Optional Lang As Long = 0, Optional IamHelpFile As Boolean = False)
Dim stat As Long
Dim base As String, table$, from As Long, first$, Second$, ok As Boolean, fr As Double, stac1$, p, i&
Dim mb As Mk2Base, tables As FastCollection, pppp As mArray, Temp As mArray, fieldlist As FastCollection, ii As Long, pppp1 As mArray
Dim vv, IndexList As FastCollection, many As Long, topi As Long, temp2 As mArray, ret As Boolean
Dim LastRead As Long
ok = False
If IsStrExp(bstackstr, R$, base, False) Then
    If getone(base, vv) Then
        If Not TypeOf vv Is Mk2Base Then GoTo noMk2
    Else
     GoTo noMk2
    End If
    If FastSymbol(R$, ",") Then
        If IsStrExp(bstackstr, R$, table$, False) Then
            If FastSymbol(R$, ",") Then
                If IsExp(bstackstr, R$, fr) Then
                    from = CLng(fr)
                    ok = True
                    If FastSymbol(R$, ",") Then
                        ok = False
                        If IsStrExp(bstackstr, R$, first$, False) Then
                            If FastSymbol(R$, ",") Then
                                If Search$ = vbNullString Then
                                If Not IsStrExp(bstackstr, R$, Search$, False) Then
                                Search$ = " = "
                                End If
                                If Not FastSymbol(R$, ",", True) Then Exit Sub
                                End If
                                If IsExp(bstackstr, R$, p) Then
                                    ok = True
                                ElseIf IsStrExp(bstackstr, R$, Second$) Then
                                    p = Second$
                                    ok = True
                                End If
                                If ok Then
                                
                                Search$ = Trim$(UCase(Search$))
                                
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If ok Then
        Set mb = vv
        Set tables = mb.tables
        If tables.ExistKey(myUcase(table$, True)) Then
            Set pppp = tables.ValueObj
            If pppp.item(21) = 0 Then
                bstackstr.soros.PushVal 0
                Exit Sub
            End If
            If from > pppp.item(21) Then from = 1
            Set tables = pppp.item(2)
            Set IndexList = pppp.item(4)
            many = pppp.item(21)
            If Len(first$) > 0 Then
                    Set fieldlist = pppp.item(1)
                    If fieldlist.ExistKey(myUcase(first$, True)) Then
                        If fieldlist.sValue <> -1 Then GoTo aError
                        i& = fieldlist.index
                        ii = 0
                        topi = tables.Count
                        many = 0
                        If IndexList.Count > 0 Then
                        topi = IndexList.Count
                        While ii < topi
                            IndexList.index = ii
                            IndexList.Done = True
                            tables.index = IndexList.Value
                            tables.Done = True
                            If tables.sValue = 0 Then
                                Set Temp = tables.ValueObj
                                If Temp.Count > i& Then
                                If Temp.CompareItem(i&, p, Search$, ret) Then
                                    If ret Then
                                        from = from - 1
                                        If temp2 Is Nothing Then Set temp2 = Temp: LastRead = ii
                                        If from = 0 Then Set pppp1 = Temp: LastRead = ii
                                        
                                        many = many + 1
                                    End If
                                Else
                                MyEr "compare can't processed", "η σύγκριση δεν μπορεί να γίνει"
                                Exit Sub
                                End If
                                End If
                            End If
                            ii = ii + 1
                        Wend
                        Else
                        While ii < topi
                            tables.index = ii
                            tables.Done = True
                            If tables.sValue = 0 Then
                                Set Temp = tables.ValueObj
                                If Temp.Count > i& Then
                                If Temp.CompareItem(i&, p, Search$, ret) Then
                                
                                    If ret Then
                                        from = from - 1
                                        If temp2 Is Nothing Then Set temp2 = Temp
                                        If from = 0 Then Set pppp1 = Temp
                                        
                                        many = many + 1
                                    End If
                                Else
                                MyEr "compare can't processed", "η σύγκριση δεν μπορεί να γίνει"
                                Exit Sub
                                End If
                                End If
                            End If
                            ii = ii + 1
                        Wend
                        End If
                        If Not pppp1 Is Nothing Then Set Temp = pppp1 Else Set Temp = temp2
                        If many = 0 Then
                        bstackstr.soros.PushVal 0
                        Exit Sub
                        
                        End If
                        GoTo cont001
                    Else
aError:
                        MyEr "no field " + first$ + " found", "Δεν βρήκα το πεδίο " + first$
                        Exit Sub
                    End If
                

            End If

            ii = 0
            If IndexList.Count > 0 Then
                topi = IndexList.Count
                
                While from > 0 And ii < topi
                    IndexList.index = ii
                    IndexList.Done = True
                    tables.index = IndexList.Value
                    tables.Done = True
                    If tables.sValue = 0 Then
                        from = from - 1
                        If from = 0 Then Set Temp = tables.ValueObj: LastRead = ii
                    End If
                    ii = ii + 1
                Wend
            
            
            
            Else
            topi = tables.Count
            
            While from > 0 And ii < topi
            
                tables.index = ii
                tables.Done = True
                If tables.sValue = 0 Then
                    from = from - 1
                    If from = 0 Then Set Temp = tables.ValueObj: LastRead = ii
                End If
                ii = ii + 1
            Wend
            End If
cont001:
            pppp.item(22) = -1
            If Not Temp Is Nothing Then
                pppp.item(22) = LastRead
                Set pppp1 = pppp.item(17)  ' field indexes
                For ii = pppp1.Count - 1 To 0 Step -1
                    from = pppp1.item(ii)
                    If Temp.IsStringItem(from) Then
                        bstackstr.soros.PushStrVariant Temp.item(from)
                    Else
                        bstackstr.soros.PushVal Temp.item(from)
                    End If
                Next ii
                bstackstr.soros.PushVal many
            End If
        End If
    End If
    Exit Sub

noMk2:
If FastSymbol(R$, ",") Then
If IsStrExp(bstackstr, R$, table$, False) Then
If FastSymbol(R$, ",") Then
If IsExp(bstackstr, R$, fr, , True) Then
from = CLng(fr)
If FastSymbol(R$, ",") Then
If IsStrExp(bstackstr, R$, first$, False) Then
If FastSymbol(R$, ",") Then
If Search$ = vbNullString Then
    If IsStrExp(bstackstr, R$, Search$, False) Then
    Search$ = " " + Search$ + " "
        If FastSymbol(R$, ",") Then
                If IsExp(bstackstr, R$, p, , True) Then
                If vartype(p) = vbString Then
                Second$ = p
                GoTo grw123
                Else
                Second$ = Search$ + Str$(p)
                End If
                ok = True
            ElseIf IsStrExp(bstackstr, R$, Second$) Then
grw123:
            If InStr(Second$, "'") > 0 Then
                Second$ = Search$ + Chr(34) + Second$ + Chr(34)
            Else
                Second$ = Search$ + "'" + Second$ + "'"
                End If
                ok = True
            End If
        End If
 
        End If
    Else
     If IsExp(bstackstr, R$, p) Then
            If vartype(p) = vbString Then
            Second$ = p
            GoTo qtw3345
            End If
            Second$ = Search$ + Str$(p)
            ok = True
            ElseIf IsStrExp(bstackstr, R$, Second$) Then
qtw3345:
                      If InStr(Second$, "'") > 0 Then
                Second$ = Search$ + Chr(34) + Second$ + Chr(34)
            Else
                Second$ = Search$ + "'" + Second$ + "'"
                End If
            ok = True
        End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
'Dim wrkDefault As Workspace,
Dim myBase  ' as variant


Dim rec   '  as variant  too  - As Recordset
Dim srl As Long
On Error Resume Next
' new addition to handle ODBC
' base=""
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this

Else
If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
If ExtractType(base) = vbNullString Then base = base + ".mdb"
If Not IamHelpFile Then If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If

g05:
Err.Clear
   On Error Resume Next
Dim Id$
   
      If first$ = vbNullString Then
If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
  End If
   Else
Id$ = "SELECT * FROM [" + table$ + "] WHERE [" + first$ + "] " + Second$
 End If

   If Not getone(base, myBase) Then
   
      Set myBase = CreateObject("ADODB.Connection")
   
      
    If DriveType(Left$(base, 3)) = "Cd-Rom" Then
        srl = DriveSerial(Left$(base, 3))
        If srl = 0 And Not GetDosPath(base) = vbNullString Then
                If Lang = 0 Then
                    If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " + ExtractName(base, True)) = vbCancel Then Exit Sub
                Else
                    If Not ask("Put CD/Disk with file " + ExtractName(base, True)) = vbCancel Then Exit Sub
                End If
         End If

 
 '  If mybase = VbNullString Then ' mybase.Mode = adShareDenyWrite
   If myBase = vbNullString Then myBase.open JetPrefix + GetDosPath(base) + ";Mode=Share Deny Write" + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection

            If Err.Number > 0 Then
            
            Do While srl <> DriveSerial(Left$(base, 3))
                If Lang = 0 Then
                If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " + CStr(srl) + " στον οδηγό " + Left$(base, 1)) = vbCancel Then Exit Do
                Else
                If ask("Put CD/Disk with serial number " + CStr(srl) + " in drive " + Left$(base, 1)) = vbCancel Then Exit Do
                End If
            Loop
            If srl = DriveSerial(Left$(base, 3)) Then
            Err.Clear
        If myBase = vbNullString Then myBase.open JetPrefix + GetDosPath(base) + ";Mode=Share Deny Write" + JetPostfix + "User Id=" + DBUser + ";Password=" + DBSecurityOFF      'open the Connection
        
            End If
        
        End If
    Else
'     myBase.Open JetPrefix + """" + GetDosPath(BASE) + """" + ";Jet OLEDB:Database Password=100101;User Id=" + DBUser  + ";Password=" + DBUserPassword + ";" +  DBSecurityOFF  'open the Connection
 If myBase = vbNullString Then
 If Left$(base, 1) = "(" Or JetPostfix = ";" Then
 myBase.open JetPrefix + JetPostfix
 If Err.Number Then
 MyEr Err.Description, Err.Description
 Exit Sub
 End If
 Else
        Err.Clear
        myBase.open JetPrefix + GetDosPath(base) + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
        If Err.Number = -2147467259 Then
           Err.Clear
           myBase.open JetPrefixOld + GetDosPath(base) + JetPostfixOld + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
           If Err.Number = 0 Then
               JetPrefix = JetPrefixOld
               JetPostfix = JetPostfixOld
           Else
               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
           End If
        ElseIf Err.Number <> 0 Then
        
            MyEr "Read row, open base: " + Err.Description, "Διάβασμα γραμμής, άνοιγμα βάσης: " + Err.Description
            Err.Clear
            Exit Sub
        End If
 End If
 End If


    End If

   If Err.Number > 0 Then GoTo g10
   
      PushOne base, myBase
      
      End If

Dim LL$
   Set rec = CreateObject("ADODB.Recordset")
 Err.Clear
 If myBase.Mode = 0 Then myBase.open
 Err.Clear
  rec.open Id$, myBase, 3, 4
If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.open = LL$
 PushOne base, myBase
 Err.Clear
rec.open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description + " " + Id$, Err.Description + " " + Id$
Exit Sub
End If
End If

   

   
  If rec.EOF Then
   ' stack$(BASESTACK) = " 0" + stack$(BASESTACK)
   bstackstr.soros.PushVal CDbl(0)
   rec.Close
  myBase.Close
    
    Exit Sub
  End If
  rec.MoveLast
  ii = rec.RecordCount

If ii <> 0 Then
If from >= 0 Then
  rec.MoveFirst
    If ii >= from Then
  rec.move from - 1
  End If
End If

    For i& = rec.fields.Count - 1 To 0 Step -1
    On Error Resume Next
    Err.Clear
    stat = rec.fields(i&).status
    If Err Then
        Err.Clear
         
    ElseIf stat > 1 Then
        bstackstr.soros.PushUndefine
        GoTo contNext
    End If
   Select Case rec.fields(i&).Type
Case 1, 2, 3, 4, 5, 6

    If myIsNull(rec.fields(i&)) Then
        bstackstr.soros.PushUndefine
    Else
    stat = rec.fields(i&).status
    If stat = 2 Then
    bstackstr.soros.PushUndefine
    Else
        bstackstr.soros.PushVal CDbl(rec.fields(i&))
        End If
    End If
Case 7
If myIsNull(rec.fields(i&)) Then
    
     bstackstr.soros.PushStr ""
 Else
  
   bstackstr.soros.PushStr CStr(CDate(rec.fields(i&)))
  End If


Case 130, 8, 203, 202
If myIsNull(rec.fields(i&)) Then
    
     bstackstr.soros.PushStr ""
 Else
  
   bstackstr.soros.PushStrVariant rec.fields(i&)
  End If
Case 11, 12 ' this is the binary field so we can save unicode there
   Case Else
'
   bstackstr.soros.PushStr "?"
 End Select
contNext:
   Next i&
   End If
   
   bstackstr.soros.PushVal CDbl(ii)


Exit Sub
g10:
If ERL Then
If Lang = 0 Then
If ask("Το ερώτημα SQL δεν μπορεί να ολοκληρωθεί" + vbCrLf + table$, True) = vbRetry Then GoTo g05
Else
If ask("SQL can't complete" + vbCrLf + table$) = vbRetry Then GoTo g05
End If
Err.Clear
MyErMacro R$, "Can't read a database table :" + table$, "Δεν μπορώ να διαβάσω πίνακα :" + table$
End If
On Error Resume Next


End Sub

Public Sub GetNames(bstackstr As basetask, R$, bv As Object, Lang)
Dim base As String, table$, from As Long, many As Long, ok As Boolean, fr As Double, stac1$, i&
ok = False
Dim vv
If IsStrExp(bstackstr, R$, base, False) Then
    If getone(base, vv) Then
        If Not TypeOf vv Is Mk2Base Then GoTo noMk2
    Else
     GoTo noMk2
    End If
    MyEr "not for m2k base yet", "όχι για βάσεις m2k ακόμα"
    Exit Sub

noMk2:
If FastSymbol(R$, ",") Then
If IsStrExp(bstackstr, R$, table$, False) Then
If FastSymbol(R$, ",") Then
If IsExp(bstackstr, R$, fr) Then
from = CLng(fr)
If FastSymbol(R$, ",") Then
If IsExp(bstackstr, R$, fr) Then
many = CLng(fr)

ok = True
End If
End If
End If
End If
End If
End If
End If
Dim ii As Long
Dim myBase ' variant
Dim rec
Dim srl As Long
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base + ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
Dim Id$
  If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
End If

     If Not getone(base, myBase) Then
   
      Set myBase = CreateObject("ADODB.Connection")
   
   
   If DriveType(Left$(base, 3)) = "Cd-Rom" Then
       srl = DriveSerial(Left$(base, 3))
    If srl = 0 And Not GetDosPath(base) = vbNullString Then
    
       If Lang = 0 Then
    If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " + ExtractName(base, True)) = vbCancel Then Exit Sub
    Else
      If Not ask("Put CD/Disk with file " + ExtractName(base, True)) = vbCancel Then Exit Sub
    End If
     End If

     myBase.open JetPrefix + GetDosPath(base) + ";Mode=Share Deny Write" + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF    'open the Connection

               If Err.Number > 0 Then
        
            Do While srl <> DriveSerial(Left$(base, 3))
            If Lang = 0 Then
            If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " + CStr(srl) + " στον οδηγό " + Left$(base, 1)) = vbCancel Then Exit Do
            Else
            If ask("Put CD/Disk with serial number " + CStr(srl) + " in drive " + Left$(base, 1)) = vbCancel Then Exit Do
            End If
            Loop
            If srl = DriveSerial(Left$(base, 3)) Then
            Err.Clear
   myBase.open JetPrefix + GetDosPath(base) + ";Mode=Share Deny Write" + JetPostfix + "User Id=" + DBUser + ";Password=" + DBSecurityOFF   'open the Connection
                
            End If
        
        End If
   Else
    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
 myBase.open JetPrefix + JetPostfix
 If Err.Number Then
 MyEr Err.Description, Err.Descnullription
 Exit Sub
 End If
 Else
        Err.Clear
        myBase.open JetPrefix + GetDosPath(base) + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
        If Err.Number = -2147467259 Then
           Err.Clear
           myBase.open JetPrefixOld + GetDosPath(base) + JetPostfixOld + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
           If Err.Number = 0 Then
               JetPrefix = JetPrefixOld
               JetPostfix = JetPostfixOld
           Else
               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
           End If
        End If
End If
End If
On Error GoTo g101
      PushOne base, myBase
      
      End If
 Dim LL$
   Set rec = CreateObject("ADODB.Recordset")
    Err.Clear
     rec.open Id$, myBase, 3, 4
      If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.open = LL$
 PushOne base, myBase
 Err.Clear
rec.open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description + " " + Id$, Err.Description + " " + Id$
Exit Sub
End If
End If


 ' DBEngine.Idle dbRefreshCache

  If rec.EOF Then
   ''''''''''''''''' stack$(BASESTACK) = " 0" + stack$(BASESTACK)
bstackstr.soros.PushVal CDbl(0)
  Exit Sub
 
'    wrkDefault.Close
  End If
  rec.MoveLast
  ii = rec.RecordCount

If ii <> 0 Then
If from >= 0 Then
  rec.MoveFirst
    If ii >= from Then
  rec.move from - 1
  End If
End If
If many + from - 1 > ii Then many = ii - from + 1
bstackstr.soros.PushVal CDbl(ii)
''''''''''''''''' stack$(BASESTACK) = " " + Trim$(Str$(II)) + stack$(BASESTACK)

    For i& = 1 To many
    bv.additemFast CStr(rec.fields(0))   ' USING gList
    
    If i& < many Then rec.MoveNext
    Next
  End If
rec.Close
'myBase.Close

Exit Sub
g101:
MyErMacro R$, "Can't read a table from database", "Δεν μπορώ να διαβάσω ένα πίνακα βάσης δεδομένων"

'myBase.Close
End Sub
Public Sub CommExecAndTimeOut(bstackstr As basetask, R$)
Dim base As String, com2execute As String, comTimeOut As Double
Dim ok As Boolean
comTimeOut = 30
If IsStrExp(bstackstr, R$, base, False) Then
    If FastSymbol(R$, ",") Then
        If IsStrExp(bstackstr, R$, com2execute, False) Then
        ok = True
            If FastSymbol(R$, ",") Then
                If Not IsExp(bstackstr, R$, comTimeOut) Then
                ok = False
                End If
            End If
        End If
    End If
End If
If Not ok Then Exit Sub
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base + ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If

Dim myBase, rs As Object
    
On Error Resume Next
If Not getone(base, myBase) Then
   
    Set myBase = CreateObject("ADODB.Connection")
      
    If DriveType(Left$(base, 3)) = "Cd-Rom" Then
    ' we can do NOTHING...
        MyEr "Can't execute command in a CD-ROM", "Δεν μπορώ εκτελέσω εντολή στη βάση δεδομένων σε CD-ROM"
        Exit Sub
    Else
        If Left$(base, 1) = "(" Or JetPostfix = ";" Then
            myBase.open JetPrefix + JetPostfix
            If Err.Number Then
            MyEr Err.Description, Err.Description
            Exit Sub
            End If
        Else
            Err.Clear
            myBase.open JetPrefix + GetDosPath(base) + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
            If Err.Number = -2147467259 Then
               Err.Clear
               myBase.open JetPrefixOld + GetDosPath(base) + JetPostfixOld + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
               If Err.Number = 0 Then
                   JetPrefix = JetPrefixOld
                   JetPostfix = JetPostfixOld
               Else
                   MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
               End If
            End If
        End If
    End If
    PushOne base, myBase
End If
Dim erdesc$
Err.Clear
If comTimeOut >= 10 Then myBase.CommandTimeout = CLng(comTimeOut)
If Err.Number > 0 Then Err.Clear: myBase.errors.Clear
com2execute = Replace(com2execute, Chr(9), " ")
com2execute = Replace(com2execute, vbCrLf, "")
com2execute = Replace(com2execute, ";", vbCrLf)
Dim commands() As String, i As Long, mm As mStiva, aa As Object
commands() = Split(com2execute + vbCrLf, vbCrLf)
Set mm = New mStiva
For i = LBound(commands()) To UBound(commands())

    If Len(MyTrim(commands(i))) > 0 Then
        ProcTask2 bstackstr  'to allow threads to run at background.
        Set rs = myBase.Execute(commands(i))
        If Typename(rs) = "Recordset" Then
            If rs.fields.Count > 0 Then
                Set aa = rs
                mm.DataObj aa
                Set aa = Nothing
                Set rs = Nothing
            End If
        End If
        If myBase.errors.Count <> 0 Then Exit For
    End If
Next i

If mm.Total > 0 Then bstackstr.soros.MergeTop mm
If myBase.errors.Count <> 0 Then
    For i = 0 To myBase.errors.Count - 1
        erdesc$ = erdesc$ + myBase.errors(i)
    Next i
        MyEr "Can't execute command:" + erdesc$, "Δεν μπορώ να εκτελέσω την εντολή:" + erdesc$
    myBase.errors.Clear
End If
End Sub





Public Function MyOrder(bstackstr As basetask, R$, Lang As Long) As Boolean
Dim base As String, Tablename As String, fs As String, i&, o As Double, ok As Boolean
Dim pppp As mArray
ok = False
Dim mb As Mk2Base, vv, tables As FastCollection, param As mStiva2
If Not IsStrExp(bstackstr, R$, base, False) Then
MissStringExpr
Exit Function
End If

If getone(base, vv) Then
If Not TypeOf vv Is Mk2Base Then GoTo noMk2

    If FastSymbol(R$, ",", True) Then
        If IsStrExp(bstackstr, R$, Tablename, False) Then
            ok = True
        Else
            MissStringExpr
        End If
    End If
If Not ok Then Exit Function

Set mb = vv
Set tables = mb.tables
If Not tables.ExistKey(myUcase(Tablename, True)) Then
MyEr "table not exist", "ο πίνακας δεν υπάρχει"
Exit Function
End If

Set param = New mStiva2
    If FastSymbol(R$, ",") Then
        Do
           If IsStrExp(bstackstr, R$, fs, False) Then
            If FastSymbol(R$, ",", True) Then
            If IsExp(bstackstr, R$, o) Then
            
            param.DataStr fs
            param.DataVal o
            Else
            MissNumExpr
            Exit Function
            End If
            End If
            Else
            MissStringExpr
            Exit Function
            End If
          Loop Until Not FastSymbol(R$, ",")
    End If
        param.PushStr Tablename
        MyOrder = mb.AddIndexes_(param)
Exit Function
End If
noMk2:

    If FastSymbol(R$, ",", True) Then
        If IsStrExp(bstackstr, R$, Tablename, False) Then
            ok = True
        Else
            MissStringExpr
        End If
    End If


If Not ok Then Exit Function
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base + ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Function
End If
    
    Dim myBase
    
    On Error Resume Next
       If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Function
                Else
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.open JetPrefix + JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Function
                        End If
                    Else
                        Err.Clear
                        myBase.open JetPrefix + GetDosPath(base) + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.Clear
                           myBase.open JetPrefixOld + GetDosPath(base) + JetPostfixOld + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                 
                End If
                PushOne base, myBase
            End If
           Err.Clear
           Dim LL$, mcat, pIndex, mtable
           Dim okntable As Boolean
          
            Err.Clear
            Set mcat = CreateObject("ADOX.Catalog")
            mcat.ActiveConnection = myBase

            

        If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.open = LL$
 PushOne base, myBase
 Err.Clear
            Set mcat = CreateObject("ADOX.Catalog")
            mcat.ActiveConnection = myBase
            

If Err.Number Then
MyEr Err.Description + " " + Tablename, Err.Description + " " + Tablename
Exit Function
End If
End If
Err.Clear
mcat.tables(Tablename).indexes("ndx").Remove
Err.Clear
mcat.tables(Tablename).indexes.Refresh

   If mcat.tables.Count > 0 Then
   okntable = True
        For Each mtable In mcat.tables
        If mtable.Type = "TABLE" Then
        If mtable.Name = Tablename Then
        okntable = False
        Exit For
        End If
        End If
        Next mtable
'        Set mtable = Nothing
        If okntable Then GoTo t111
Else
t111:
MyEr "No tables in Database " + ExtractNameOnly(base, True), "Δεν υπάρχουν αρχεία στη βάση δεδομένων " + ExtractNameOnly(base, True)
Exit Function
End If
' now we have mtable from mybase
If mtable Is Nothing Then
Else
 mtable.indexes("ndx").Remove  ' remove the old index/
 End If
 Err.Clear
 If mcat.ActiveConnection.errors.Count > 0 Then
 mcat.ActiveConnection.errors.Clear
 End If
 Err.Clear
   Set pIndex = CreateObject("ADOX.Index")
    pIndex.Name = "ndx"  ' standard
    pIndex.indexnulls = 0 ' standrard
  
        While FastSymbol(R$, ",")
        If IsStrExp(bstackstr, R$, fs, False) Then
        If FastSymbol(R$, ",") Then
        If IsExp(bstackstr, R$, o, False) Then
        
        pIndex.Columns.Append fs
        If o = 0 Then
        pIndex.Columns(fs).sortorder = CLng(1)
        Else
        pIndex.Columns(fs).sortorder = CLng(2)
        End If
        End If
        End If
                 
        End If
        Wend
        If pIndex.Columns.Count > 0 Then
        If mtable.indexes.Count = 1 Then
        mtable.indexes.Delete pIndex.Name
        End If
        mtable.indexes.Append pIndex
        
             If Err.Number Then
          '   mtable.Append pIndex
         MyEr Err.Description, Err.Description
         Exit Function
        End If

mcat.tables.Append mtable
Err.Clear
mcat.tables.Refresh
End If

MyOrder = True
End Function
Public Function NewTable(bstackstr As basetask, R$, Lang As Long) As Boolean
'BASE As String, tablename As String, ParamArray flds()
Dim base As String, Tablename As String, fs As String, i&, n As Double, l As Double, ok As Boolean
Dim vv, mb As Mk2Base, param As mStiva2, oldl As Double
ok = False

If Not IsStrExp(bstackstr, R$, base, False) Then
    MissStringExpr
    Exit Function
Else
    If getone(base, vv) Then
    ' what
    If Not TypeOf vv Is Mk2Base Then GoTo noMk2
    Else
        GoTo noMk2
    End If
    If Not FastSymbol(R$, ",", True) Then
    Exit Function
    End If
    If Not IsStrExp(bstackstr, R$, Tablename, False) Then
    MissStringExpr
    Exit Function
    Else
    Set mb = vv
    Set param = New mStiva2
    param.DataVal Tablename
    mb.AddTables_ param
    param.Flush
    
    If Not FastSymbol(R$, ",") Then
    NewTable = True
    Exit Function
    End If
    Do
    NewTable = False
    If Not IsStrExp(bstackstr, R$, fs, False) Then
    MissStringExpr
    Exit Function
    End If
    If Not FastSymbol(R$, ",", True) Then
    Exit Function
    End If
    If Not IsExp(bstackstr, R$, n) Then
    MissNumExpr
    Exit Function
    End If
    If Not FastSymbol(R$, ",", True) Then
    Exit Function
    End If
    If Not IsExp(bstackstr, R$, l) Then
    MissNumExpr
    Exit Function
    End If
    param.DataStr fs
    param.DataVal n
    param.DataVal l
    Loop Until Not FastSymbol(R$, ",")
    param.PushVal Tablename
    NewTable = mb.AddFields_(param)
    End If
    End If
Exit Function
noMk2:
If FastSymbol(R$, ",", True) Then
If IsStrExp(bstackstr, R$, Tablename, False) Then
ok = True
Else
MissStringExpr
End If
End If


If Not ok Then Exit Function
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base + ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Function
End If
    Dim okndx As Boolean, okntable As Boolean, one_ok As Boolean
    ' Dim wrkDefault As Workspace
    Dim myBase ' As Database
    Err.Clear
    On Error Resume Next
                   If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Function
                Else
                If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                    myBase.open JetPrefix + JetPostfix
                    If Err.Number Then
                    MyEr Err.Description, Err.Description
                    Exit Function
                    End If
                Else
                    Err.Clear
                    myBase.open JetPrefix + GetDosPath(base) + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                    If Err.Number = -2147467259 Then
                       Err.Clear
                       myBase.open JetPrefixOld + GetDosPath(base) + JetPostfixOld + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                       If Err.Number = 0 Then
                           JetPrefix = JetPrefixOld
                           JetPostfix = JetPostfixOld
                       Else
                           MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                       End If
                    End If
                End If
                End If
                PushOne base, myBase
            End If
           Err.Clear

    On Error Resume Next
   okntable = True
Dim cat, mtable, LL$
  Set cat = CreateObject("ADOX.Catalog")
           Set cat.ActiveConnection = myBase


If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.open = LL$
 PushOne base, myBase
 Err.Clear
 Set cat.ActiveConnection = myBase
If Err.Number Then
MyEr Err.Description + " " + mtable, Err.Description + " " + mtable
Exit Function
End If
End If

    Set mtable = CreateObject("ADOX.TABLE")
         Set mtable.parentcatalog = cat
' check if table exist

           If cat.tables.Count > 0 Then
        For Each mtable In cat.tables
          If mtable.Type = "TABLE" Then
        If mtable.Name = Tablename Then
        okntable = False
        Exit For
        End If
        End If
        Next mtable
       If okntable Then
       Set mtable = CreateObject("ADOX.TABLE")      ' get a fresh one
        mtable.Name = Tablename
        Set mtable.parentcatalog = cat
       End If
    
    
 With mtable.Columns

                Do While FastSymbol(R$, ",")
                
                        If IsStrExp(bstackstr, R$, fs, False) Then
                        one_ok = True
                                If FastSymbol(R$, ",") Then
                                        If IsExp(bstackstr, R$, n) Then
                                
                                            If FastSymbol(R$, ",") Then
                                                If IsExp(bstackstr, R$, l) Then
                                                If n = 1 Then n = 11: l = 0
                                                If n = 2 Then n = 16: l = 0
                                                If n = 3 Then n = 2: l = 0
                                                If n = 4 Then
                                                oldl = l
                                                n = 3: l = 0
                                                End If
                                                If n = 5 Then n = 6: l = 0
                                                If n = 6 Then n = 4: l = 0
                                                If n = 7 Then n = 5: l = 0
                                                If n = 8 Then n = 7: l = 0
                                                If n = 9 Then n = 128
                                                If n = 10 Then n = 202
                                                If n = 12 Then n = 203: l = 0
                                                If n = 16 Then n = 14: l = 0
                                                    If l <> 0 Then
                                                
                                                     .Append fs, n, l
                                           
                                                    Else
                                                     .Append fs, n
                                                     If n = 3 And oldl = -1 Then
                                                     mtable.Columns.item(0).properties("AutoIncrement") = True
                                                   
                                                     End If
                                                    End If
                                        
                                                End If
                                            End If
                                        End If
                        
                                End If
                
                        End If
                
                Loop
               
End With
        If okntable Then
        
        cat.tables.Append mtable
        If Err.Number Then
        If Err.Number = -2147217859 Then
        Err.Clear
        Else
         MyEr Err.Description, Err.Description
         Exit Function
        End If
        
        End If
        cat.tables.Refresh
        ElseIf Not one_ok Then
        cat.tables.Delete Tablename
        cat.tables.Refresh
        End If
        
' may the objects find the creator...

NewTable = okntable
End If

End Function


Sub BaseCompact(bstackstr As basetask, R$)

Dim base As String, conn, BASE2 As String, realtype$
If Not IsStrExp(bstackstr, R$, base, False) Then
MissParam R$
Else
If FastSymbol(R$, ",") Then
If Not IsStrExp(bstackstr, R$, BASE2, False) Then
MissParam R$
Exit Sub
End If
End If
'only for mdb
If Left$(base, 1) = "(" Or JetPostfix = ";" Then Exit Sub ' we can't compact in ODBC use control panel

''If JetPrefix <> JetPrefixHelp Then Exit Sub
  On Error Resume Next
  
If ExtractPath(base) = vbNullString Then
base = mylcasefILE(mcd + base)
Else
  If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
realtype$ = mylcasefILE(Trim$(ExtractType(base)))
If realtype$ <> "" Then
    base = ExtractPath(base, True) + ExtractNameOnly(base, True)
    If BASE2 = vbNullString Then BASE2 = strTemp + LTrim$(Str(Timer)) + "_0." + realtype$ Else BASE2 = ExtractPath(BASE2) + LTrim$(Str(Timer)) + "_0." + realtype$
    Set conn = CreateObject("JRO.JetEngine")
    base = base + "." + realtype$

   conn.CompactDatabase JetPrefix + base + JetPostfixUser, _
                                GetStrUntil(";", "" + JetPrefix) + _
                                GetStrUntil(":", "" + JetPostfix) + ":Engine Type=5;" + _
                                "Data Source=" + BASE2 + JetPostfixUser
                                

    
    If Err.Number = 0 Then
    If ExtractPath(base) <> ExtractPath(BASE2) Then
       KillFile base
       Sleep 50
        If Err.Number = 0 Then
            MoveFile BASE2, base
            Sleep 50

        Else
            If GetDosPath(BASE2) <> "" Then KillFile BASE2
        End If
    
    Else
        KillFile base
        MoveFile BASE2, base
            Sleep 50
    
    End If
       
    
    
    
    Else
      
      
 
      MyErMacro R$, "Can't compact databese " + ExtractName(base, True) + "." + " use a back up", "Πρόβλημα με την βάση " + ExtractName(base, True) + ".mdb χρησιμοποίησε ένα σωσμένο αρχείο"
      End If
      Err.Clear
    End If
End If
End Sub

Public Function DELfields(bstackstr As basetask, R$) As Boolean
Dim base$, table$, first$, Second$, ok As Boolean, p As Double, vv, usehandler As mHandler
ok = False
If IsExp(bstackstr, R$, p) Then
If bstackstr.lastobj Is Nothing Then
GoTo ee1
End If

If Not TypeOf bstackstr.lastobj Is mHandler Then
GoTo ee1
Else
Set usehandler = bstackstr.lastobj
If Not usehandler.t1 = 1 Then
ee1:
MyEr "Expected Inventory", "Περίμενα Κατάσταση"
Exit Function
End If
End If
Dim aa As FastCollection
Set aa = usehandler.objref
If aa.StructLen > 0 Then
MyEr "Structure members are ReadOnly", "Τα μέλη της δομής είναι μόνο για ανάγνωση"
Exit Function
End If
Set bstackstr.lastobj = Nothing
Set usehandler = Nothing
Do While FastSymbol(R$, ",")
ok = False
If IsExp(bstackstr, R$, p) Then
aa.Remove p
If Not aa.Done Then MyEr "Key not exist", "Δεν υπάρχει τέτοιο κλειδί": Exit Do
ok = True
ElseIf IsStrExp(bstackstr, R$, first$, False) Then
aa.Remove first$
If Not aa.Done Then MyEr "Key not exist", "Δεν υπάρχει τέτοιο κλειδί": Exit Do
ok = True
Else
    Exit Do
End If
Loop
DELfields = ok
Set aa = Nothing
Exit Function

ElseIf IsStrExp(bstackstr, R$, base$, False) Then

    If getone(base, vv) Then
        If Not TypeOf vv Is Mk2Base Then GoTo noMk2
    Else
     GoTo noMk2
    End If
    MyEr "not for m2k base yet", "όχι για βάσεις m2k ακόμα"
    Exit Function
noMk2:
If FastSymbol(R$, ",") Then
If IsStrExp(bstackstr, R$, table$, False) Then
If FastSymbol(R$, ",") Then
If IsStrExp(bstackstr, R$, first$, False) Then
If FastSymbol(R$, ",") Then
If IsStrExp(bstackstr, R$, Second$, False) Then
ok = True

           If InStr(Second$, "'") > 0 Then
                Second$ = Chr(34) + Second$ + Chr(34)
            Else
                Second$ = "'" + Second$ + "'"
                End If
ElseIf IsExp(bstackstr, R$, p) Then
ok = True
    If CheckInt64(p) Then
        Second$ = CStr(p)
    ElseIf vartype(p) = vbString Then
        Second$ = LTrim$(p)
    Else
        Second$ = LTrim$(Str(p))
    End If
Else
MissParam R$
End If
Else
MissParam R$

End If
Else
MissParam R$

End If
Else
MissParam R$

End If
Else
MissParam R$
End If
Else
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this we can 't killfile the base for odbc
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base + ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: DELfields = False: Exit Function
    If CheckMine(base) Then KillFile base: DELfields = True: Exit Function
    
End If

End If
Else
MissParam R$
End If
If Not ok Then DELfields = False: Exit Function
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base + ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: DELfields = False: Exit Function
End If

Dim myBase
   On Error Resume Next
                   If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Function
                Else
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.open JetPrefix + JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        DELfields = False: Exit Function
                        End If
                    Else
                        Err.Clear
                        myBase.open JetPrefix + GetDosPath(base) + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.Clear
                           myBase.open JetPrefixOld + GetDosPath(base) + JetPostfixOld + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                End If
                PushOne base, myBase
            End If
           Err.Clear

    On Error Resume Next
Dim rec
   
   
   
   If first$ = vbNullString Then
   MyEr "Nothing to delete", "Τίποτα για να σβήσω"
   DELfields = False
   Exit Function
   Else
   myBase.errors.Clear
   myBase.Execute "DELETE * FROM [" + table$ + "] WHERE " + first$ + " = " + Second$
   If myBase.errors.Count > 0 Then
   MyEr "Can't delete " + table$, "Δεν μπορώ να διαγράψω"
   Else
    DELfields = True
   End If
   
   End If
   Set rec = Nothing

End Function

Function CheckMine(DBFileName) As Boolean
' M2000 changed to ADO...

Dim Cnn1
 Set Cnn1 = CreateObject("ADODB.Connection")

 On Error Resume Next
 Cnn1.open JetPrefix + DBFileName + ";Jet OLEDB:Database Password=;User Id=" + DBUser + ";Password=" + DBUserPassword + ";"  ' +  DBSecurityOFF 'open the Connection
 If Err Then
 Err.Clear
 Cnn1.open JetPrefix + DBFileName + JetPostfix + "User Id=" + DBUser + ";Password=" + DBUserPassword + ";" + DBSecurityOFF    'open the Connection
 If Err Then
 Else
 CheckMine = True
 End If
 Cnn1.Close
 Else
 End If
End Function


Public Sub PushOne(conname As String, v As Variant)
On Error Resume Next
conCollection.AddKey conname, v
'Set v = conCollection(conname)
End Sub
Sub CloseAllConnections()
Dim v As Variant, bb As Boolean
On Error Resume Next
If Not Init Then Exit Sub
If conCollection.Count > 0 Then
Dim i As Long
Err.Clear
For i = conCollection.Count - 1 To 0 Step -1
On Error Resume Next
conCollection.index = i
If conCollection.IsObj Then
If TypeOf conCollection.ValueObj Is Mk2Base Then
' do nothing just throw
With conCollection.ValueObj
.Close_
End With
Else
With conCollection.ValueObj
bb = .ConnectionString <> ""
If Err.Number = 0 Then
If .Mode > 0 Then
If .state = 1 Then
   .Close
ElseIf .state = 2 Then
    .Close
ElseIf .state > 2 Then
Call .Cancel
.Close
End If
    
End If
End If
End With
End If
End If
conCollection.Remove conCollection.KeyToString
Err.Clear

Next i
Set conCollection = New FastCollection
End If
Err.Clear
End Sub
Public Sub RemoveOneConn(conname)
On Error Resume Next
Dim vv, mb As Mk2Base
If conCollection Is Nothing Then Exit Sub
If Not conCollection.ExistKey(conname) Then
    conname = mylcasefILE(conname)
    If ExtractPath(conname) = vbNullString Then conname = mylcasefILE(mcd + conname)
    If ExtractType(CStr(conname)) = vbNullString Then conname = mylcasefILE(conname + ".mdb")
    If conCollection.ExistKey(conname) Then
    
    GoTo conthere
    End If
    Exit Sub
Else
conthere:
    Set vv = conCollection.ValueObj
    If TypeOf vv Is Mk2Base Then
        Set mb = vv
        mb.Close_
    ElseIf vv.ConnectionString <> "" Then
    
    If Err.Number = 0 And vv.Mode <> 0 Then vv.Close
    Err.Clear
    End If
    conCollection.Remove conname
    Err.Clear
End If
End Sub
Private Function getone(conname As String, this As Variant) As Boolean
On Error Resume Next
InitMe
If conCollection.ExistKey(conname) Then
Set this = conCollection.ValueObj
getone = True
End If
End Function
Private Sub changeone(conname As String, this As Variant)
On Error Resume Next
InitMe
If conCollection.ExistKey(conname) Then
Set conCollection.ValueObj = this
End If
End Sub
Public Function getone2(conname As String, this As Variant) As Boolean
On Error Resume Next
InitMe

If conCollection.ExistKey(conname) Then
Set this = conCollection.ValueObj
getone2 = True
End If
End Function
Private Sub InitMe()
If Init Then Exit Sub
Set conCollection = New FastCollection
Init = True
End Sub
Function ftype(ByVal a As Long, Lang As Long) As String
Select Case Lang
Case 0
Select Case a
    Case 0
ftype = "ΑΔΕΙΟ"
    Case 2
ftype = "ΑΚΕΡΑΙΟΣ"
    Case 3
ftype = "ΜΑΚΡΥΣ"
    Case 4
ftype = "ΑΠΛΟΣ"
    Case 5
ftype = "ΔΙΠΛΟΣ"
    Case 6
ftype = "ΛΟΓΙΣΤΙΚΟΣ"
    Case 7
ftype = "ΗΜΕΡΟΜΗΝΙΑ"
    Case 8
ftype = "BSTR"
    Case 9
ftype = "IDISPATCH"
    Case 10
ftype = "ERROR"
    Case 11
ftype = "ΛΟΓΙΚΟΣ"
    Case 12
ftype = "VARIANT"
    Case 13
ftype = "IUNKNOWN"
    Case 14
ftype = "DECIMAL"
    Case 16
ftype = "ΨΗΦΙΟ"
    Case 17
ftype = "UNSIGNEDTINYINT"
    Case 18
ftype = "UNSIGNEDSMALLINT"
    Case 19
ftype = "UNSIGNEDINT"
    Case 20
ftype = "BIGINT"
    Case 21
ftype = "UNSIGNEDBIGINT"
    Case 64
ftype = "FILETIME"
    Case 72
ftype = "GUID"
    Case 128
ftype = "BINARY"
    Case 129
ftype = "CHAR"
    Case 130
ftype = "WCHAR"
    Case 131
ftype = "NUMERIC"
    Case 132
ftype = "USERDEFINED"
    Case 133
ftype = "DBDATE"
    Case 134
ftype = "DBTIME"
    Case 135
ftype = "ΗΜΕΡΟΜΗΝΙΑ" 'DBTIMESTAMP
    Case 136
ftype = "CHAPTER"
    Case 138
ftype = "PROPVARIANT"
    Case 139
ftype = "VARNUMERIC"
    Case 200
ftype = "VARCHAR"
    Case 201
ftype = "LONGVARCHAR"
    Case 202
ftype = "ΚΕΙΜΕΝΟ" '"VARWCHAR"
    Case 203
ftype = "LONGVARWCHAR"
    Case 204
ftype = "ΔΥΑΔΙΚΟ"  ' "VARBINARY"
    Case 205
ftype = "OLE" '"LONGVARBINARY"
    Case 8192
ftype = "ARRAY"
Case Else
ftype = "????"


End Select

Case Else  ' this is for 1
Select Case a
    Case 0
ftype = "EMPTY"
    Case 2
ftype = "INTEGER"
    Case 3
ftype = "LONG"
    Case 4
ftype = "SINGLE"
    Case 5
ftype = "DOUBLE"
    Case 6
ftype = "CURRENCY"
    Case 7
ftype = "DATE"
    Case 8
ftype = "BSTR"
    Case 9
ftype = "IDISPATCH"
    Case 10
ftype = "ERROR"
    Case 11
ftype = "BOOLEAN"
    Case 12
ftype = "VARIANT"
    Case 13
ftype = "IUNKNOWN"
    Case 14
ftype = "DECIMAL"
    Case 16
ftype = "BYTE"
    Case 17
ftype = "UNSIGNEDTINYINT"
    Case 18
ftype = "UNSIGNEDSMALLINT"
    Case 19
ftype = "UNSIGNEDINT"
    Case 20
ftype = "BIGINT"
    Case 21
ftype = "UNSIGNEDBIGINT"
    Case 64
ftype = "FILETIME"
    Case 72
ftype = "GUID"
    Case 128
ftype = "BINARY"
    Case 129
ftype = "CHAR"
    Case 130
ftype = "WCHAR"
    Case 131
ftype = "NUMERIC"
    Case 132
ftype = "USERDEFINED"
    Case 133
ftype = "DBDATE"
    Case 134
ftype = "DBTIME"
    Case 135
ftype = "DBTIMESTAMP"
    Case 136
ftype = "CHAPTER"
    Case 138
ftype = "PROPVARIANT"
    Case 139
ftype = "VARNUMERIC"
    Case 200
ftype = "VARCHAR"
    Case 201
ftype = "LONGVARCHAR"
    Case 202
ftype = "VARWCHAR"
    Case 203
ftype = "LONGVARWCHAR"
    Case 204
ftype = "VARBINARY"
    Case 205
ftype = "OLE"
    Case 8192
ftype = "ARRAY"


Case Else
ftype = "????"
End Select
End Select
End Function
Sub GeneralErrorReport(aBasBase As Variant)
Dim errorObject

 For Each errorObject In aBasBase.ActiveConnection.errors
 'Debug.Print "Description :"; errorObject.Description
 'Debug.Print "Number:"; Hex(errorObject.Number)
 Next
End Sub

Function Digits(a$, Label$, notrim As Boolean) As Boolean
Dim a1 As Long, LI As Long, A2 As Long
LI = Len(a$)

If LI > 0 Then
If notrim Then
a1 = 1
Else
a1 = MyTrimL(a$)
End If
A2 = a1
If a1 > LI Then a$ = vbNullString: Exit Function
'If LI > 5 + A2 Then LI = 4 + A2
If Mid$(a$, a1, 1) Like "[0-9]" Then
Do While a1 <= LI
a1 = a1 + 1
If Not Mid$(a$, a1, 1) Like "[0-9]" Then Exit Do

Loop
Label$ = Mid$(a$, A2, a1 - A2): a$ = Mid$(a$, a1)
Digits = True
End If

End If
End Function
Public Sub SQL()
Dim a$, R$, k As Long, r1$, waittablename As Boolean
Dim closepar As Long, getonename As Boolean
a$ = "SELECT * FROM [COMMANDS alfa] WHERE ENGLISH LIKE '" + "MODULE%" + "' AND GROUPNUM = 3"
a$ = "SELECT [ENGLISH] FROM COMMANDS WHERE GROUPNUM =" + Str$(100) + " ORDER BY [ENGLISH]"
a$ = "SELECT DISTINCT customer_name FROM depositor d WHERE NOT EXISTS ( SELECT * FROM borrower b WHERE b.customer_name = d.customer_name);"
Debug.Print a$
If IsLabelOnly(a$, R$) Then
Debug.Print "Command:"; R$
Debug.Print a$

If FastSymbol(a$, "*") Then Debug.Print "ALL"

Do
k = Len(a$)
While FastSymbol(a$, "[")
R$ = Left$(a$, InStr(a$, "]") - 1)
Debug.Print "Field:"; R$
a$ = Mid$(a$, Len(R$) + 2)
Wend
While FastSymbol(a$, "'")
R$ = Left$(a$, InStr(a$, "'") - 1)
Debug.Print "string", R$
a$ = Mid$(a$, Len(R$) + 2)
Wend
If IsLabelOnly(a$, R$) Then
R$ = UCase(R$)
Select Case R$
Case "EXISTS"
Debug.Print "FLAG??:";
Case "FROM"
getonename = False
Debug.Print "Command:";
waittablename = True
Case "GROUP"
getonename = False
If IsLabelOnly(a$, r1$) Then
r1$ = UCase(r1$)
If Not r1$ = "BY" Then Debug.Print "Missing BY", a$: Exit Sub
R$ = R$ + " " + r1$
Debug.Print "Command:";
waittablename = False
End If
Case "ORDER"
getonename = False
If IsLabelOnly(a$, r1$) Then
r1$ = UCase(r1$)
If Not r1$ = "BY" Then Debug.Print "Missing BY", a$: Exit Sub
R$ = R$ + " " + r1$
Debug.Print "Command:";
waittablename = False
End If
Case "WHERE", "HAVING"
getonename = False
Debug.Print "Command:";
waittablename = False
Case "DISTINCT"
Debug.Print "Without Doublicates: ";
Case "AS"
getonename = False
R$ = ""
Debug.Print "Name the result: ";
Case "MIN(", "MAX(", "COUNT("
getonename = False
Debug.Print "Aggregate Function: ";
closepar = 1
Case "AVG(", "SUM("   ' numeric inputs
getonename = False
Debug.Print "Aggregate Function (for numeric inputs): ";
closepar = 1
Case "ASC", "DESC"
Debug.Print "Order Attribute: ";
Case "AND", "NOT"
getonename = False
Debug.Print "Logic Operator: ";
Case "LIKE"
getonename = False
Debug.Print "Operator: ";
Case Else
If getonename Then
Debug.Print "Alias:";
Else
If waittablename Then
Debug.Print "Table:";
Else
Debug.Print "Field:";
End If
getonename = True
End If
End Select
Debug.Print R$
End If
a$ = LTrim(a$)
R$ = ""
Do
If Len(a$) = 0 Then Exit Do
If InStr("<=>-+*/", Left$(a$, 1)) > 0 Then
R$ = R$ + Left$(a$, 1): a$ = Mid$(a$, 2)
Else
Exit Do
End If
Loop
If Len(R$) > 0 Then Debug.Print "operator: "; R$: getonename = False
If Digits(a$, R$, False) Then
    If Left$(a$, 1) = "." Then
        R$ = R$ + "."
        If Digits(a$, r1$, True) Then
            R$ = R$ + r1$
            If InStr("eE", Left$(a$, 1)) > 0 Then
                If InStr("-+", Mid$(a$, 2, 1)) > 0 Then
                    r1$ = "E" + Mid$(a$, 2, 1)
                    If Digits(Mid$(a$, 3), (r1$), True) Then
                        R$ = R$ + r1$
                        a$ = Mid$(a$, 3)
                        Digits a$, r1$, True
                        R$ = R$ + r1$
                    End If
                Else
                    r1$ = "E"
                    If Digits(Mid$(a$, 2), (r1$), True) Then
                        R$ = R$ + r1$
                        a$ = Mid$(a$, 2)
                        Digits a$, r1$, True
                        R$ = R$ + r1$
                    End If
                End If
            End If
        End If
    End If
    Debug.Print "Number:", R$: getonename = True
End If
While FastSymbol(a$, vbCrLf)
Wend
If FastSymbol(a$, "(") Then Debug.Print "Open Parenthesis": closepar = closepar + 1: getonename = False
If closepar > 0 Then If FastSymbol(a$, ")") Then Debug.Print "Close Parenthesis": closepar = closepar - 1
If FastSymbol(a$, ",") Then Debug.Print "Another Item (,)": getonename = False
If FastSymbol(a$, ";") Then Exit Do
Loop Until k = Len(a$)
End If
End Sub
Function ProcDBprovider(bstack As basetask, rest$, Lang As Long) As Boolean
Dim pa$, s$, ss$
ProcDBprovider = True
If IsStrExp(bstack, rest$, pa$) Then
If pa$ = vbNullString Then
JetPrefixUser = JetPrefixHelp
JetPostfixUser = JetPostfixHelp
Else
' DB.PROVIDER "Microsoft.ACE.OLEDB.12.0","Jet OLEDB","100101"
' DB.PROVIDER "Microsoft.Jet.OLEDB.4.0", "Jet OLEDB", "100101"
' DB.PROVIDER "dns=testme;Uid=admin;Pwd=12alfa45", "ODBC", "100101"
' use (name) for database name

 JetPrefixUser = "Provider=" + pa$ + ";Data Source="  ' normal
    If FastSymbol(rest$, ",") Then
       If IsStrExp(bstack, rest$, s$) Then
          If s$ = vbNullString Then
             ProcDBprovider = False
          ElseIf UCase(s$) = "ODBC" Or UCase(s$) = "PATH" Then
                If FastSymbol(rest$, ",") Then
                 If IsStrExp(bstack, rest$, ss$) Then
                 JetPrefixUser = pa$ + ";Password=" + ss$
                 Else
                 JetPrefixUser = pa$ + ";Password="
                 End If
                Else
                JetPrefixUser = pa$
                End If
                JetPostfixUser = ";"
          Else
          
             If FastSymbol(rest$, ",") Then
                If IsStrExp(bstack, rest$, ss$) Then
                   If ss$ = vbNullString Then
                       JetPostfixUser = ";" + s$ + ":Database Password=100101;"
                   Else
                       JetPostfixUser = ";" + s$ + ":Database Password=" + ss$ + ";"
                   
                   End If
                    
                Else
                    ProcDBprovider = False
                End If
             Else
                JetPostfixUser = ";" + s$ + ":Database Password=100101;"
             End If
          End If
        Else
         ProcDBprovider = False
       End If
    Else
       JetPostfixUser = JetPostfixHelp

    End If
 End If
 Else
 JetPrefixUser = JetPrefixHelp
 
End If
JetPostfix = JetPostfixUser
JetPrefix = JetPrefixUser
End Function
Function ProcDBUSER(bstack As basetask, rest$, Lang As Long) As Boolean
Dim s$
ProcDBUSER = -True
    If IsStrExp(bstack, rest$, s$) Then
        If s$ = vbNullString Then
            extDBUser = vbNullString
            extDBUserPassword = vbNullString
        Else
            extDBUser = s$
        End If
        If FastSymbol(rest$, ",") Then
            If Not IsStrExp(bstack, rest$, extDBUserPassword) Then
                extDBUserPassword = vbNullString
            End If
            DBUser = extDBUser
            DBUserPassword = extDBUserPassword
        End If
    End If

End Function
