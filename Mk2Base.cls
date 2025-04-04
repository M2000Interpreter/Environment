VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mk2Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This is the Mk2 new format for database,
' to replace the access database for help file, as primary target
Option Explicit
' this is the file id, where we place as first 16 bytes
Const b4 = "mk2-0001" ' these the first 16 bytes, as utf16le, type and version
' this is the verification code, so we have to check this, written as 16bytes
Const magic = "{336C0F04-D0E4-7DF9-3F2B-E592A9936DE3}"
' then then GUID for the base written from 33 byte in file.
' If this change, then we can't read or write to base, until we get the schema of the base.
' This can be change if we open the base for exclusive work
' so every time we open the base for shared use,
' we look the GUID if it is the same as the one we have here
' if is the same, means that we have the proper schema
' When we use in memory base, we can do anything, but the GUID may change frequently
' So when we place the in memory base to a file, we set the last Guid.
' this is the procedure when we define the base.
' The actuall save of base happen when we close the base, with the option to save to file.

' Basic operations are: Open Base from file
' (this return to Open Base In Memory, and set the flag SaveAtClose)
' Open Base In Memory. This is for temporary use, as a record set
' We can populate records in a Base in memory, opening a Base from File (as source), and then
' we setup the query string from source, and we make a list of records as the output of query execution
' this list of records marshal the source to get a field row (the record), then using the

Private Type GUID
    d001 As Long
    d002 As Long
    d003 As Long
    d004 As Long
End Type
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Dst As Any, Src As Any, ByVal BLen As Long)
Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
                         ByVal lpszCLSID As Long, _
                         ByRef Clsid As GUID) As Long
Private Declare Function StringFromCLSID Lib "ole32.dll" ( _
                         ByRef Clsid As GUID, _
                         lpszProgID As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
                            (ByVal hMem As Long)
Private Declare Sub MakeGuid Lib "advapi32.dll" Alias "SystemFunction036" (dest As GUID, ByVal lSize As Long)
Private mGuid As GUID
Dim BaseTable As New FastCollection
Dim UseofIndex As Boolean
Dim baseName As String, mInMemory As Boolean
Private Function GetBStrFromPtr(lpSrc As Long) As String
Dim slen As Long
  If lpSrc = 0 Then Exit Function
  slen = lstrlenW(lpSrc)
  If slen Then GetBStrFromPtr = space$(slen) Else Exit Function
   RtlMoveMemory ByVal StrPtr(GetBStrFromPtr), ByVal lpSrc, slen * 2
End Function
Private Function GetGUIDstr(G As GUID) As String
Dim ret As Long, here As Long
ret = StringFromCLSID(G, here)
If ret Then Exit Function
GetGUIDstr = GetBStrFromPtr(here)
CoTaskMemFree here
End Function
Private Function SetGUID(StrGuid) As GUID
Dim tGUID As GUID
CLSIDFromString StrPtr(StrGuid), tGUID
SetGUID = tGUID
End Function

'Open(aName$), open the database, creating the index file
Friend Sub OpenBase_(aName$)
' We can open a new, so the indexes can be none
End Sub
Friend Function AddTables_(ParamList As mStiva2) As Boolean
Dim i As Integer, Tablename, pppp As tuple, fieldlist As FastCollection
Dim M As GUID, originalname As String
Dim goodlist As New mStiva2
Do While ParamList.count > 0
    originalname = ParamList.PopVal
    If originalname = vbNullString Then Exit Function
    Tablename = myUcase(originalname, True)
    If BaseTable.ExistKey(Tablename) Then
        ' not accepted, the name exist
        Exit Function
    Else
        goodlist.DataVal Tablename
    End If
Loop

Do While goodlist.count > 0
    Set pppp = New tuple
    pppp.PushDim (30), (0)  ' 30 items, of anytype, we have many spare items
    pppp.PushEnd
    MakeGuid M, LenB(M)
    ' 0 the GUID
    pppp.item(0) = GetGUIDstr(M)
    Set fieldlist = New FastCollection
    fieldlist.AllowAnyKey
    Set pppp.item(1) = fieldlist ' the fields collection
    Set pppp.item(2) = New FastCollection  ' records
 
    Set pppp.item(3) = New FastCollection   'fields for indexes

    Set pppp.item(4) = New FastCollection  ' the indexes keys are here
  
    pppp.item(5) = 0
    ' for recordsets
    ' each recordset has an array of items
    pppp.item(6) = 0  ' when is 1 we have a Query.
    ' a Query is a table from other table(s).
    ' in pppp.item(1) has the actual table fields
    ' in pppp.item(2) has the actual table records
    ' (3) and (4) and (17) not used
    ' (7) has the fields indexes for query
    ' (8) has the
    
    
    ' pppp.item(17) is the fields indexes (maybe we delete some fields)
    ' pppp.item(18) is the accumulator
    pppp.item(20) = originalname
    ' pppp.item(21) the records number
    ' pppp.item(22) the last record which we read
    ' pppp.item(23) the last record which we change/append
    
    BaseTable.AddKey goodlist.PopVal, pppp

Loop
AddTables_ = True
End Function
Friend Function RemoveTables_(ParamList As mStiva2) As Boolean
Dim goodlist As New mStiva2, Tablename
Do While ParamList.count > 0
    Tablename = ParamList.PopVal
    If Tablename = vbNullString Then Exit Function
    If BaseTable.ExistKey(Tablename) Then
        goodlist.DataVal Tablename
    Else
        ' not accepted, the name not exist
        Exit Function
    End If
Loop
Do While goodlist.count > 0
    BaseTable.Remove goodlist.PopVal
Loop
RemoveTables_ = True
End Function
Friend Function AddFields_(ParamList As mStiva2) As Boolean
Dim Tablename, fieldname, pppp As tuple, fieldlist As FastCollection, fl As String, ma As tuple
Dim originalname As String, i As Long, acc As Long, pppp1 As tuple, pppp2 As tuple, j As Long
Dim fieldtype As Long
If ParamList.count > 0 Then Tablename = ParamList.PopVal
If Tablename = vbNullString Then Exit Function
Tablename = myUcase(Tablename, True)
If BaseTable.Find(Tablename) Then
    Set pppp = BaseTable.ValueObj

    Set fieldlist = pppp.item(1)
    Do While ParamList.count <> 0

    originalname = ParamList.PopVal
    fl = myUcase(originalname, True)
    If fieldlist.ExistKey(fl) Then
     Set ma = fieldlist.ValueObj
     ' pppp.item(21) is the record number
     If ma.item(4) And pppp.item(21) > 0 Then
     ' error
     MyEr "you can change an autoincrement field", "��� ������� �� �������� ��� ����� �� ��������� ���������"
     Exit Function
     End If
    fieldlist.sValue = 0  ' delete old field
    End If
    Set ma = New tuple
    ma.PushDim (10), (0)  ' 10 items
    ma.PushEnd
    ma.item(0) = originalname
    fieldtype = ParamList.PopVal
    ma.item(1) = (fieldtype And 511)
    ma.item(3) = (fieldtype And 1024) = 1024  ' can take null
    ma.item(6) = (fieldtype And 2048) = 2048  ' foreign_key
    If ma.item(6) Then
    ' here we have to get parameters, "base!table" or something
    End If
    If (fieldtype And 512) = 512 Then
        ' get default value
       
        
        ma.item(7) = ParamList.PopVal
    End If
    ma.item(2) = ParamList.PopVal
    ' autoincrement
    If ma.item(2) = -1 And ma.item(1) < 10 Then ma.item(4) = True: ma.item(5) = 1@ Else ma.item(4) = False
    

    
    fieldlist.AddKey fl, ma
    fieldlist.sValue = -1
    Loop
    ' prepare record accumulator
    For i = 0 To fieldlist.count - 1
    fieldlist.Index = i
    fieldlist.Done = True
    If fieldlist.sValue = -1 Then
   '' Set ma = fieldlist.ValueObj
    acc = acc + 1
    End If
    Next i
    If acc > 0 Then
    Set pppp1 = New tuple
    pppp1.PushDim (acc), (0)
    pppp1.PushEnd

    j = 0
    For i = 0 To fieldlist.count - 1
    fieldlist.Index = i
    fieldlist.Done = True
    If fieldlist.sValue = -1 Then pppp1.item(j) = i: j = j + 1
    Next i
    End If
    Set pppp.item(17) = pppp1  ' fields indexes
    Set pppp2 = New tuple
    pppp2.PushDim (fieldlist.count), (0)
    pppp2.PushEnd
    Set pppp.item(18) = pppp2  ' accumulator
    pppp.item(21) = 0 ' no records yet
    Else
        ' not accepted, the name not exist
        Exit Function
End If


AddFields_ = True

End Function
Friend Sub RemoveFields_(ParamList As mStiva2)


End Sub
Friend Function AddIndexes_(ParamList As mStiva2) As Boolean
Dim Tablename, fieldname, pppp As tuple, fieldlist As FastCollection, fl As String
Dim indexField As FastCollection, i As Long, fieldindex As Long, IndexList As FastCollection
Dim originalname As String, vv, ordertype As Long, Temp As tuple
Dim ord()
If ParamList.count > 0 Then Tablename = ParamList.PopVal
If Tablename = vbNullString Then Exit Function
Tablename = myUcase(Tablename, True)
If BaseTable.Find(Tablename) Then
    Set pppp = BaseTable.ValueObj
    Set fieldlist = pppp.item(1)
    
    
    
    i = 1
    Do While i <= ParamList.count
    fl = myUcase(myUcase(ParamList.StackItem(i), True))
    If ParamList.StackItem(i + 1) >= 0 Then
    i = i + 2
    If Not fieldlist.ExistKey(fl) Then
        MyEr "field not exist", "�� ����� ��� �������"
        Exit Function
    Else
        If fieldlist.sValue <> -1 Then
            MyEr "field deleted, can't be used", "�� ����� ���� ���������, ��� ������ �� ��������������"
            Exit Function
        End If
        Set Temp = fieldlist.ValueObj
        If Temp.item(1) = 12 Then
            MyEr "Memo field can't be used for key", "�� ����� ����� �������� ��� ������ �� �������������� ��� ������"
            Exit Function
        End If
        
    End If
    End If
    Loop
    
    ' so now we pass all errors
    ' clear the current indexes
    Set indexField = New FastCollection
    
    Set IndexList = New FastCollection
    
    
    
    If ParamList.count \ 2 > 0 Then
    ReDim ord(ParamList.count \ 2 - 1)
    i = 0
    Do While ParamList.count <> 0

        originalname = ParamList.PopVal
        
        ordertype = ParamList.PopVal
        fl = myUcase(originalname, True)
        If fieldlist.ExistKey(fl) Then
            Set Temp = fieldlist.ValueObj
            If Temp.item(1) < 10 Then
                If ordertype < 2 Then
                ordertype = (ordertype And 3) + 2
                Else
                ordertype = 3 - ordertype   '+ 2
                End If
            ElseIf ordertype > 1 Then
               ' ordertype = 3 - ordertype
            End If
            If ordertype < 0 Then ordertype = 0
            If ordertype > 3 Then ordertype = 3
            indexField.AddKey fl, originalname + " " + LTrim(str$(ordertype))
            indexField.sValue = fieldlist.Index
            ord(i) = ordertype
            i = i + 1
        End If
    Loop
    If UBound(ord()) = 0 Then
    If ord(0) > 1 Then
    pppp.item(5) = 1
    IndexList.FeedSCol2 ord()
    Else
    pppp.item(5) = 1 - ord(0)
    End If
    Else
     pppp.item(5) = 1
    IndexList.FeedSCol2 ord()
    End If
    Set pppp.item(3) = indexField
    Set pppp.item(4) = IndexList
    If pppp.item(21) > 0 Then ReIndex pppp
    Else
    pppp.item(5) = 1
    Set pppp.item(3) = indexField
    Set pppp.item(4) = IndexList
    End If
Else
        ' not accepted, the name not exist
        Exit Function
End If
AddIndexes_ = True
End Function
Friend Sub RemoveIndexes_(ParamList As mStiva2)

End Sub

Friend Sub PutRowInPlace_(ParamList As mStiva2)

End Sub
Friend Sub PutRowInNewPlace_(ParamList As mStiva2)

End Sub
Friend Sub DeleteRow_(ParamList As mStiva2)

End Sub
Friend Sub GetRow_(IndexNo As Long, ParamList As mStiva2)

End Sub

Private Sub Class_Initialize()
mInMemory = True
MakeGuid mGuid, 16
End Sub

Private Property Get InMemory() As Variant
InMemory = mInMemory
End Property

Private Property Let InMemory(ByVal vNewValue As Variant)
InMemory = mInMemory
End Property
Public Sub Close_()
' do something
'Stop
End Sub
Property Get tables() As Object
Set tables = BaseTable
End Property
Sub ReIndexTable(Tablename As String)
Dim pppp As tuple
If BaseTable.ExistKey(myUcase(Tablename, True)) Then
    Set pppp = BaseTable.ValueObj
    ReIndex pppp
End If
End Sub
Friend Sub ReIndex(pppp As tuple)
Dim RecordList As FastCollection, IndexList As FastCollection, indexField As FastCollection
Dim i&, j&, allkey$, acc As tuple, b() As Byte
Set RecordList = pppp.item(2)
Set indexField = pppp.item(3)
Set IndexList = pppp.item(4)
If IndexList.count <> 0 Then
        b() = IndexList.ExportScol3()
        Set pppp.item(4) = New FastCollection
        Set IndexList = pppp.item(4)
        If UBound(b()) > 0 Then
        IndexList.FeedSCol3 b
        End If
        IndexList.NumericSort = True
End If
If pppp.item(21) = 0 Then Exit Sub
RecordList.Done = True
For i& = 0 To RecordList.count - 1
    RecordList.Index = i&
    If RecordList.sValue = 0 Then
    Set acc = RecordList.ValueObj
    allkey$ = vbNullString
    For j& = 0 To indexField.count - 1
        indexField.Index = j&
        indexField.Done = True
        If Len(allkey$) = 0 Then
        allkey$ = indexField.Normalize(acc.item(indexField.sValue))
        Else
        allkey$ = allkey$ + ChrW(1) + indexField.Normalize(acc.item(indexField.sValue))
        End If
    Next j&
    If IndexList.ExistKey(allkey$) Then
        IndexList.Value = i&
    Else
        IndexList.AddKey allkey$, i&
    End If
    End If
Next i&
If IndexList.count > 0 Then

        If pppp.item(5) = 2 Then
            IndexList.SortDes
        Else
            IndexList.Sort
        End If

End If
End Sub
Property Get LastRead(ByVal Tablename As String) As Variant
Dim pppp As tuple
Tablename = myUcase(Tablename, True)
If BaseTable.ExistKey(Tablename) Then
Set pppp = BaseTable.ValueObj
LastRead = pppp.item(22) + 1

Else
LastRead = -100

End If
End Property

