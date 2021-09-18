Attribute VB_Name = "mdlIDispatch"
' ************************************************************************
' Copyright:    All rights reserved.  � 2004
' Project:      AsyncServer
' Module:       mdlIDispatch
' Original Author:       james b tollan
' Changed by George Karras
' Change TLB to take care named arguments
'
    Const DISPATCH_METHOD = 1
    Const DISPATCH_PROPERTYGET = 2
    Const DISPATCH_PROPERTYPUT = 4
    Const DISPATCH_PROPERTYPUTREF = 8
    Const DISPID_UNKNOWN = -1
    Const DISPID_VALUE = 0
    Const DISPID_PROPERTYPUT = -3
    Const DISPID_NEWENUM = -4
    Const DISPID_EVALUATE = -5
    Const DISPID_CONSTRUCTOR = -6
    Const DISPID_DESTRUCTOR = -7
    Const DISPID_COLLECT = -8
Option Explicit
Enum cbnCallTypes
    VbLet = DISPATCH_PROPERTYPUT
    VbGet = DISPATCH_PROPERTYGET
    VbSet = DISPATCH_PROPERTYPUTREF
    VbMethod = DISPATCH_METHOD
    VbNext = DISPID_NEWENUM
End Enum
' Maybe need this http://support2.microsoft.com/kb/2870467/
'To update oleaut32
Private Declare Sub VariantCopy Lib "OleAut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
Private KnownProp As FastCollection
Private Init As Boolean
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Public Function FindDISPID(pobjTarget As Object, ByVal pstrProcName As Variant) As Long

    Dim IDsp        As IDispatch.IDispatchM2000
    Dim riid        As IDispatch.IID
    Dim dispid      As Long

    Dim lngRet      As Long
    FindDISPID = -1
    If pobjTarget Is Nothing Then Exit Function

    Dim a$(0 To 0), arrdispid(0 To 0) As Long, myptr() As Long
    ReDim myptr(0 To 0)
    myptr(0) = StrPtr(pstrProcName)
    
 Set IDsp = pobjTarget
 If Not getone(Typename(pobjTarget) & "." & pstrProcName, dispid) Then
      lngRet = IDsp.GetIDsOfNames(riid, myptr(0), 1&, Clid, arrdispid(0))
     
      If lngRet = 0 Then dispid = arrdispid(0): PushOne Typename(pobjTarget) & "." & pstrProcName, dispid
      
      Else
      lngRet = 0
End If
If lngRet = 0 Then FindDISPID = dispid
End Function
Public Sub ShutEnabledGuiM2000(Optional all As Boolean = False)
Dim x As Form, bb As Boolean

Do
For Each x In Forms
bb = True
If TypeOf x Is GuiM2000 Then
    If x.enabled Then bb = False: x.CloseNow: bb = False: Exit For
    
End If
Next x

Loop Until bb Or Not all

End Sub

Public Function CallByNameFixParamArray _
    (pobjTarget As Object, _
    ByVal pstrProcName As Variant, _
    ByVal CallType As cbnCallTypes, _
     pArgs(), pargs2() As String, items As Long, Optional robj As Object, Optional fixnamearg As Long = 0, Optional center2mouse As Boolean = False, Optional pUnk) As Variant


    ' pobjTarget    :   Class or form object that contains the procedure/property
    ' pstrProcName  :   Name of the procedure or property
    ' CallType      :   vbLet/vbGet/vbSet/vbMethod
    ' pArgs()       :   Param Array of parameters required for methode/property
    ' New by George
     ' pargs2() the names of arguments
     ' fixnamearg = the number of named arguments
    Dim myform As Form
    Dim IDsp        As IDispatch.IDispatchM2000
    Dim riid        As IDispatch.IID
    Dim params      As IDispatch.DISPPARAMS
    Dim Excep       As IDispatch.EXCEPINFO
    ' Do not remove TLB because those types
    ' are also defined in stdole
    Dim dispid      As Long
    Dim lngArgErr   As Long
    Dim VarRet      As Variant
    Dim varArr()    As Variant
    Dim varDISPID() As Long
    Dim lngRet      As Long
    Dim lngLoop     As Long
    Dim lngMax      As Long
Dim myptr() As Long
Dim mm As GuiM2000
Dim mmm As mArray
    ' Get IDispatch from object
    Set IDsp = pobjTarget

    ' Get DISPIP from pstrProcName
    If fixnamearg = 0 Then
        ReDim varDISPID(0 To 0)
If Not getone(Typename$(pobjTarget) & "." & pstrProcName, dispid) Then
            ReDim myptr(0 To 0)
            myptr(0) = StrPtr(pstrProcName)
            lngRet = IDsp.GetIDsOfNames(riid, myptr(0), 1&, Clid, varDISPID(0))
            
            If lngRet = 0 Then dispid = varDISPID(0): PushOne Typename$(pobjTarget) & "." & pstrProcName, dispid
            Else
            lngRet = 0
End If
Else
         ReDim myptr(0 To fixnamearg)
            myptr(0) = StrPtr(pstrProcName)
            For lngLoop = 1 To fixnamearg
            myptr(lngLoop) = StrPtr(pargs2(lngLoop))
            Next lngLoop
                ReDim varDISPID(0 To fixnamearg)
            lngRet = IDsp.GetIDsOfNames(riid, myptr(0), fixnamearg + 1, Clid, varDISPID(0))
 dispid = varDISPID(0)
End If
    If lngRet = 0 Then
passhere:
        If items > 0 Or fixnamearg > 0 Then
                ReDim varArr(0 To items - 1 + fixnamearg)
               
                For lngLoop = 0 To items - 1 + fixnamearg
                If Not MyIsNumericPointer(pArgs(lngLoop)) Then
                If IsNull(pArgs(lngLoop)) Then
                SwapVariant varArr(fixnamearg + items - 1 - lngLoop), pArgs(lngLoop)
                ElseIf TypeOf pArgs(lngLoop) Is mArray Then
                    If Typename(pArgs(lngLoop).refArray) = "Long" Then
                    Set mmm = pArgs(lngLoop)
                    mmm.ExportArrayNow
                    End If
                    SwapVariant varArr(fixnamearg + items - 1 - lngLoop), pArgs(lngLoop).refArray
                ElseIf TypeOf pArgs(lngLoop) Is MemBlock Then
                    varArr(fixnamearg + items - 1 - lngLoop) = pArgs(lngLoop).ExportToByte
                Else
                    SwapVariant varArr(fixnamearg + items - 1 - lngLoop), pArgs(lngLoop)
                End If
                Else
                    SwapVariant varArr(fixnamearg + items - 1 - lngLoop), pArgs(lngLoop)
                    End If
                Next
              With params
                    .cArgs = items + fixnamearg
                    .rgPointerToVariantArray = VarPtr(varArr(0))
                 If CallType = VbLet Or CallType = VbSet Or fixnamearg > 0 Then
                
        If fixnamearg = 0 Then
                ReDim varDISPID(0 To 0)
                 varDISPID(0) = DISPID_PROPERTYPUT
                   .cNamedArgs = 1
                 Else
                  .cNamedArgs = fixnamearg

      For lngLoop = 0 To fixnamearg - 1
      varDISPID(lngLoop) = varDISPID(fixnamearg - lngLoop)
                    
                Next

                   
                End If
                .rgPointerToDISPIDNamedArgs = VarPtr(varDISPID(0))
                
                Else
                .cNamedArgs = 0
                .rgPointerToDISPIDNamedArgs = 0
              End If
                End With
                If lngRet = -1 Then GoTo JUMPHERE
Else
    With params
        .cArgs = 0
        .cNamedArgs = 0
    End With
End If

        ' Invoke method/property
If LastErNum = 0 Then
        lngRet = IDsp.Invoke(dispid, riid, 0, CallType, params, VarRet, Excep, lngArgErr)
End If
If LastErNum <> 0 Then GoTo exithere
If lngRet <> 0 Then
    If lngRet = DISP_E_EXCEPTION Then
        ' CallByName pobjTarget, pstrProcName, VbMethod
        MyEr GetBStrFromPtr(Excep.StrPtrDescription, False), GetBStrFromPtr(Excep.StrPtrDescription, False)
        GoTo exithere
        Err.Raise Excep.wCode
    ElseIf Typename$(pobjTarget) = "GuiM2000" Then
JUMPHERE:
            On Error GoTo exithere
            lngRet = 0
            If UCase(pstrProcName) = "HIDE" Then
                
                Set mm = pobjTarget
                If mm.Enablecontrol Then
                    mm.TrueVisible = False
                Else
                If Not mm.Minimized Then
                    mm.VisibleOldState = True
                    mm.Visible = True
                    mm.MinimizeON
                End If
                End If
               Set mm = Nothing
                'CallByName pobjTarget, "MyHide", VbMethod
            ElseIf UCase(pstrProcName) = "SHOW" Then

                If pobjTarget.Quit Then MyEr "Form unloaded, use declare again using declare A new Form", "� ����� ��� ���� ��������, ������������� ��� ����� � ��� �����": Exit Function
                CallByName pobjTarget, "ShowTaskBar", VbMethod
                CallByName pobjTarget, "ShowmeALl", VbMethod
                If items = 0 Then
                    CallByName pobjTarget, pstrProcName, VbMethod, 0, GiveForm()
                    Set myform = pobjTarget
                    MoveFormToOtherMonitorOnly myform
               ElseIf items = 2 Then
                    If varArr(0) <> 0 Then
                    GoTo conthere
                    Else
                    CallByName pobjTarget, pstrProcName, VbMethod, 0, varArr(1)

                   
                    Set myform = pobjTarget
                    MoveFormToOtherMonitorOnly myform
                    pobjTarget.Modal = 0
                    pobjTarget.Modal = 0
                    End If
               ElseIf varArr(0) = 0 Then
                    CallByName pobjTarget, pstrProcName, VbMethod, 0, GiveForm()
                    Set myform = pobjTarget
                    MoveFormToOtherMonitorOnly myform
                    pobjTarget.Modal = 0
                    pobjTarget.Modal = 0
               Else
conthere:
                   Dim oldmoldid As Double, mycodeid As Double
                   oldmoldid = Modalid
                   mycodeid = Rnd * 1000000
                   pobjTarget.Modal = mycodeid
                   Dim x As Form, z As Form, zz As Form
                   Set zz = Screen.ActiveForm
                   If zz.Name = "Form3" Then
                   Set zz = zz.lastform
                   End If
                   If Not pobjTarget.IamPopUp Then
                        For Each x In Forms
                            If x.Visible And x.Name = "GuiM2000" Then
                                If Not x Is pobjTarget Then
                                    If x.Enablecontrol Then
                                        x.Modal = mycodeid
                                        x.Enablecontrol = False
                                    End If
                                End If
                            End If
                        Next x
                    End If
                    If pobjTarget.NeverShow Then
                    Modalid = mycodeid
                    'If Not pobjTarget.IamPopUp Then
                    ' change this
                    'pobjTarget.Title = pobjTarget.Title
                    pobjTarget.ShowTaskBar
                    'End If
                    If items = 2 Then
                    CallByName pobjTarget, pstrProcName, VbMethod, 0, varArr(1)
                    Else
                    CallByName pobjTarget, pstrProcName, VbMethod, 0, GiveForm()
                    End If
                    Set myform = pobjTarget
                    MoveFormToOtherMonitorOnly myform, center2mouse
                    pobjTarget.Refresh
                    Dim handlepopup As Boolean
                    If Not Screen.ActiveForm Is Nothing Then
                        If TypeOf Screen.ActiveForm Is GuiM2000 Then
                            If Not handlepopup Then
                                If Screen.ActiveForm.PopUpMenuVal Or Screen.ActiveForm.IamPopUp Then
                                    handlepopup = True
                                End If
                            End If
                        End If
                    End If
                    Form3.WindowState = 0
                    Do While Modalid <> 0 And (pobjTarget.Visible Or pobjTarget.VisibleOldState)
                        mywait Basestack1, 1, True
                        Sleep 1
                        If pobjTarget.Visible Then
                            If Not Screen.ActiveForm Is Nothing Then
                                If TypeOf Screen.ActiveForm Is GuiM2000 Then
                                    If Not handlepopup Then
                                        If Screen.ActiveForm.PopUpMenuVal Or Screen.ActiveForm.IamPopUp Then
                                            handlepopup = True
                                        End If
                                    ElseIf GetForegroundWindow <> Screen.ActiveForm.hWnd Then
                                        handlepopup = False
                                        If Screen.ActiveForm.PopUpMenuVal Or Screen.ActiveForm.IamPopUp Then
                                        If Screen.ActiveForm.Visible Then
                                            Screen.ActiveForm.Visible = False
                                            AppNoFocus = True
                                            End If
                                            
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If ExTarget Or LastErNum <> 0 Then Exit Do
                    Loop
                    Modalid = mycodeid
                Else
                    Modalid = mycodeid
                End If
                Set z = Nothing
                For Each x In Forms
                    If x.Visible And x.Name = "GuiM2000" Then
                        x.TestModal mycodeid
                        If x.Enablecontrol Then Set z = x
                    End If
                Next x
                If Not zz Is Nothing Then Set z = zz
                If Typename(z) = "GuiM2000" Then
                If z.Modal = Modalid Then
                
                Else
                    z.ShowmeALL
                   If z.Visible Then z.SetFocus
                    End If
                    Set z = Nothing
                    
                ElseIf Not z Is Nothing Then
                    If z.Visible Then z.SetFocus
                End If
                Modalid = oldmoldid
            End If
    ElseIf items = 0 Then
        CallByName pobjTarget, pstrProcName, VbMethod
    Else
        'CallByName pobjTarget, pstrProcName, VbMethod, varArr()
        Select Case items
        Case 1
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(0)
        Case 2
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(1), varArr(0)
        Case 3
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(2), varArr(1), varArr(0)
        Case 4
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(3), varArr(2), varArr(1), varArr(0)
        Case 5
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(4), varArr(3), varArr(2), varArr(1), varArr(0)
        Case 6
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(5), varArr(4), varArr(3), varArr(2), varArr(1), varArr(0)
        Case 7
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(6), varArr(5), varArr(4), varArr(3), varArr(2), varArr(1), varArr(0)
        Case 8
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(7), varArr(6), varArr(5), varArr(4), varArr(3), varArr(2), varArr(1), varArr(0)
        Case 9
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(8), varArr(7), varArr(6), varArr(5), varArr(4), varArr(3), varArr(2), varArr(1), varArr(0)
        Case 10
            CallByName pobjTarget, pstrProcName, VbMethod, varArr(9), varArr(8), varArr(7), varArr(6), varArr(5), varArr(4), varArr(3), varArr(2), varArr(1), varArr(0)
        Case Else
            Err.Raise -2147352567
        End Select
    End If
Else
    Err.Raise lngRet
End If
End If
    Else

        Err.Raise lngRet
    End If
    Dim where As Long
    If items > 0 Then
    ' Fill parameters arrays. The array must be
    ' filled in reverse order.
        For lngLoop = 0 To items - 1 + fixnamearg
            where = fixnamearg + items - 1 - lngLoop
            If VariantIsRef(VarPtr(varArr(where))) Then
                If Not MyIsNumericPointer(pArgs(lngLoop)) Then
                    If Not MyIsNumericPointer(varArr(where)) Then
                        If VarType(varArr(where)) = 8204 Then
                            VarByRefClean VarPtr(varArr(where))
                            Set mmm = pArgs(lngLoop)
                            mmm.FixArray
                            Set mmm = Nothing
                        Else
                            VarByRefCleanRef VarPtr(varArr(where))
                            SwapVariant varArr(where), pArgs(lngLoop)
                        End If
                    Else
                        VarByRefCleanRef VarPtr(varArr(where))
                        SwapVariant varArr(where), pArgs(lngLoop)
                    End If
                Else
                    VarByRefClean VarPtr(varArr(where))
                    If pArgs(lngLoop) = vbEmpty Then
                        SwapVariant varArr(where), pArgs(lngLoop)
                    End If
                End If
            Else
                SwapVariant varArr(where), pArgs(lngLoop)
            End If
            Next
    End If
    On Error Resume Next

    Set IDsp = Nothing
    If IsObject(VarRet) Then
            Set robj = VarRet
         VarRet = CLng(0)
End If
On Error GoTo there
If TypeOf VarRet Is IUnknown Then
Set robj = New mHandler
If UCase(pstrProcName) = "_NEWENUM" Then
robj.ConstructEnumerator VarRet
Else
MyEr "cant use this object", "��� ����� �� �������� ���� �� �����������"
End If
VarRet = CLng(0)
End If
there:
Err.Clear
CallByNameFixParamArray = VarRet
Exit Function
exithere:
    If Err.Number <> 0 Then CallByNameFixParamArray = VarRet
Err.Clear
    If items > 0 Then
                ' Fill parameters arrays. The array must be
                ' filled in reverse order.
                For lngLoop = 0 To items - 1 + fixnamearg
                    SwapVariant varArr(fixnamearg + items - 1 - lngLoop), pArgs(lngLoop)
                Next
    End If
End Function


Public Function ReadOneParameter(pobjTarget As Object, dispid As Long, ERrR$, VarRet As Variant) As Boolean
    
    Dim CallType As cbnCallTypes
    
    CallType = VbGet
    Dim IDsp        As IDispatch.IDispatchM2000
    Dim riid        As IDispatch.IID
    Dim params      As IDispatch.DISPPARAMS
    Dim Excep       As IDispatch.EXCEPINFO
    ' Do not remove TLB because those types
    ' are also defined in stdole
        Dim lngArgErr   As Long
    Dim varArr()    As Variant

    Dim lngRet      As Long
    Dim lngLoop     As Long
    Dim lngMax      As Long

    ' Get IDispatch from object
    Set IDsp = pobjTarget

    ' WE HAVE DISPIP

    If lngRet = 0 And False Then
       ' wrong
      
                ReDim varArr(0 To 0)
                varArr(0) = True
                With params
                    .cArgs = 1
                    .rgPointerToVariantArray = VarPtr(varArr(0))
                                    Dim aa As Long
        
                aa = DISPID_VALUE
                .cNamedArgs = 1
                .rgPointerToDISPIDNamedArgs = VarPtr(aa)
                End With
        End If

        ' Invoke method/property
        Err.Clear
       On Error Resume Next
        lngRet = IDsp.Invoke(dispid, riid, 0, CallType, params, VarRet, Excep, lngArgErr)
If Err > 0 Then
ERrR$ = Err.Description
Exit Function
Else
        If lngRet <> 0 Then
            If lngRet = DISP_E_EXCEPTION Then
             ERrR$ = Str$(Excep.wCode)
            Else
              ERrR$ = Str$(lngRet)
            End If
            Exit Function
        End If
  End If
    On Error Resume Next

    Set IDsp = Nothing
   'If IsObject(VarRet) Then

    'Set ReadOneParameter = VarRet
    'Else
    'ReadOneParameter = VarRet
    'End If
ReadOneParameter = Err = 0
  ''  If Err.Number <> 0 Then ReadOneParameter = varRet
Err.Clear
End Function
Public Function ReadOneIndexParameter(pobjTarget As Object, dispid As Long, ERrR$, ThisIndex As Variant, Optional useset As Boolean = False, Optional ByPass As Boolean) As Variant
    
    Dim CallType As cbnCallTypes
    
    If useset Then
    CallType = VbSet
    Else
    CallType = VbGet
    End If
    Dim IDsp        As IDispatch.IDispatchM2000
    Dim riid        As IDispatch.IID
    Dim params      As IDispatch.DISPPARAMS
    Dim Excep       As IDispatch.EXCEPINFO
    ' Do not remove TLB because those types
    ' are also defined in stdole
        Dim lngArgErr   As Long
    Dim VarRet      As Variant
    Dim varArr()    As Variant

    Dim lngRet      As Long
    Dim lngLoop     As Long
    Dim lngMax      As Long

    ' Get IDispatch from object
    Set IDsp = pobjTarget
    Dim aa As Long, i As Integer, k As Integer
    aa = DISPID_VALUE
    ' WE HAVE DISPIP
    If VarType(ThisIndex) = 8204 Then
                 ReDim varArr(0 To UBound(ThisIndex))
                 k = 0
                 For i = UBound(ThisIndex) - 1 To 0 Step -1
                    varArr(k) = ThisIndex(i)
                    k = k + 1
                 Next
                With params
                    .cArgs = k
                    .rgPointerToVariantArray = VarPtr(varArr(0))
                    
                    .cNamedArgs = 0
                     .rgPointerToDISPIDNamedArgs = VarPtr(aa)
               End With
               
    Else
                 ReDim varArr(0 To 0)
                varArr(0) = ThisIndex
                
                With params
                    .cArgs = 1
                    .rgPointerToVariantArray = VarPtr(varArr(0))
                    'Dim aa As Long
                    'aa = DISPID_VALUE
                    .cNamedArgs = 1
                     .rgPointerToDISPIDNamedArgs = VarPtr(aa)
               End With
  End If

  
        Err.Clear
        On Error Resume Next
        lngRet = IDsp.Invoke(dispid, riid, 0, CallType, params, VarRet, Excep, lngArgErr)
If Err > 0 Then
ERrR$ = Err.Description
Exit Function
Else
        If lngRet <> 0 Then
            If lngRet = DISP_E_EXCEPTION Then
             ERrR$ = Str$(Excep.wCode)
            Else
              ERrR$ = Str$(lngRet)
            End If
            Exit Function
        End If
  End If
    On Error Resume Next

    Set IDsp = Nothing
    If IsObject(VarRet) Then
        If ByPass Then
            Set pobjTarget = VarRet
            ReadOneIndexParameter = 0
            ByPass = False
        Else
            Set ReadOneIndexParameter = VarRet

        End If
    Else
    ReadOneIndexParameter = VarRet
    End If

  ''  If Err.Number <> 0 Then ReadOneParameter = varRet
Err.Clear
End Function
Public Sub ChangeOneParameter(pobjTarget As Object, dispid As Long, val1, ERrR$)
    
    Dim CallType As cbnCallTypes
    
    CallType = VbLet
    Dim IDsp        As IDispatch.IDispatchM2000
    Dim riid        As IDispatch.IID
    Dim params      As IDispatch.DISPPARAMS
    Dim Excep       As IDispatch.EXCEPINFO
    ' Do not remove TLB because those types
    ' are also defined in stdole
        Dim lngArgErr   As Long
    Dim VarRet      As Variant
    Dim varArr()    As Variant

    Dim lngRet      As Long
    Dim lngLoop     As Long
    Dim lngMax      As Long

    ' Get IDispatch from object
    Set IDsp = pobjTarget

    ' WE HAVE DISPIP

    If lngRet = 0 Then
       
      
                ReDim varArr(0 To 0)
                varArr(0) = val1
                With params
                    .cArgs = 1
                    .rgPointerToVariantArray = VarPtr(varArr(0))
                                    Dim aa As Long
        
                aa = DISPID_PROPERTYPUT
                .cNamedArgs = 1
                .rgPointerToDISPIDNamedArgs = VarPtr(aa)
                End With
        End If

        ' Invoke method/property
        
        lngRet = IDsp.Invoke(dispid, riid, 0, CallType, params, VarRet, Excep, lngArgErr)

        If lngRet <> 0 Then
            If lngRet = DISP_E_EXCEPTION Then
             ERrR$ = Str$(Excep.wCode)
            Else
              ERrR$ = Str$(lngRet)
            End If
            Exit Sub
        End If
    
    
    

    Set IDsp = Nothing
    
End Sub
Public Sub ChangeOneIndexParameter(pobjTarget As Object, dispid As Long, val1, ERrR$, ThisIndex As Variant)
' not only one;;;;
    Dim CallType As cbnCallTypes
    
    CallType = VbLet
    Dim IDsp        As IDispatch.IDispatchM2000
    Dim riid        As IDispatch.IID
    Dim params      As IDispatch.DISPPARAMS
    Dim Excep       As IDispatch.EXCEPINFO
    ' Do not remove TLB because those types
    ' are also defined in stdole
        Dim lngArgErr   As Long
    Dim VarRet      As Variant
    Dim varArr()    As Variant

    Dim lngRet      As Long
    Dim lngLoop     As Long
    Dim lngMax      As Long

    ' Get IDispatch from object
    Set IDsp = pobjTarget

    ' WE HAVE DISPIP
    Dim aa As Long, i As Integer, k As Integer
    aa = DISPID_PROPERTYPUT
        If VarType(ThisIndex) = 8204 Then
                 ReDim varArr(0 To UBound(ThisIndex) + 1)
                 k = 1
                 For i = UBound(ThisIndex) - 1 To 0 Step -1
                    varArr(k) = ThisIndex(i)
                    k = k + 1
                 Next
                 varArr(0) = val1
                With params
                    .cArgs = k
                    .rgPointerToVariantArray = VarPtr(varArr(0))
                    .cNamedArgs = 1
                     .rgPointerToDISPIDNamedArgs = VarPtr(aa)
               End With
    
       
      Else
                ReDim varArr(0 To 1)
                varArr(1) = ThisIndex
                varArr(0) = val1
                With params
                    .cArgs = 2
                    .rgPointerToVariantArray = VarPtr(varArr(0))
                
                .cNamedArgs = 1
                .rgPointerToDISPIDNamedArgs = VarPtr(aa)
                End With
        End If

        ' Invoke method/property
        
        lngRet = IDsp.Invoke(dispid, riid, 0, CallType, params, VarRet, Excep, lngArgErr)

        If lngRet <> 0 Then
            If lngRet = DISP_E_EXCEPTION Then
             ERrR$ = Str$(Excep.wCode)
            Else
              ERrR$ = Str$(lngRet)
            End If
            Exit Sub
        End If
    
    
    

    Set IDsp = Nothing
    
End Sub

Private Sub PushOne(KnownPropName As String, ByVal v As Long)
On Error Resume Next
If Not KnownProp.Find(LCase(KnownPropName)) Then KnownProp.AddKey LCase$(KnownPropName)
KnownProp.Value = v

End Sub
Private Function getone(KnownPropName As String, this As Long) As Boolean
On Error Resume Next
Dim v As Long
InitMe
If KnownProp.Find(LCase$(KnownPropName)) Then
getone = True: this = KnownProp.Value
End If
End Function

Private Sub InitMe()
If Init Then Exit Sub
Set KnownProp = New FastCollection
' from this collection we never delete items.
Init = True
End Sub
Public Function MakeObjectFromString(obj As Variant, objstr As String) As Object
Dim o As Object, strvar, varg(), obj1 As Object, varg2() As String
strvar = objstr
Set o = obj
CallByNameFixParamArray o, strvar, VbGet, varg(), varg2(), 0, obj1
Set MakeObjectFromString = obj1
End Function



