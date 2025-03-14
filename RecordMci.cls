VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecordMci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type WAVEINCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname(31)         As Integer
    dwFormats           As Long
    wChannels           As Integer
    wReserved1          As Integer
End Type
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long 'Get the error message of the mcidevice if any
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long 'Send command strings to the mci device
Private Declare Function mciSendString2 Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal id As Long, ByVal something1 As Long, ByVal hwndCallback As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveInGetID Lib "winmm.dll" ( _
                         ByVal hWaveIn As Long, _
                         ByRef lpuDeviceID As Long) As Long
Private Declare Function waveInGetDevCaps Lib "winmm.dll" _
                         Alias "waveInGetDevCapsW" ( _
                         ByVal uDeviceID As Long, _
                         ByRef lpCaps As WAVEINCAPS, _
                         ByVal USize As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32" _
                         Alias "lstrcpynW" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any, _
                         ByVal iMaxLength As Long) As Long

Private Data As String * 128 ' Used to store our return data
Public Error As Long ' Used to store our error message
Public Height As Long
Public Width As Long
Public Length As Long
Private bt As Long
Private ch As Long
Private Terminate0 As String
Private Terminate1 As String
Private mFilename  As String ' Used to store our file
Public Property Let FileName(n As String)
Dim ss$
ss$ = ExtractPath(n, True)
If ss$ = "" Then
    n = mcd & ExtractName(n)
End If
mFilename = n
End Property
Public Property Get FileName() As String
FileName = mFilename
End Property
Public Sub Rec_Initialize()
If Terminate0 <> "" Then
Error = mciSendString2(Terminate0, 0, 0, 0)
If Terminate1 <> "" Then Error = mciSendString2(Terminate1, 0, 0, 0)
End If
Error = mciSendString2("open new type waveaudio alias mycapture", 0, 0, 0)
Error = mciSendString2("open new type waveaudio alias waveLevel", 0, 0, 0)
Error = mciSendString2("set mycapture time format ms", 0, 0, 0)
If Error <= 0 Then Terminate0 = "close mycapture": Terminate1 = "close wavelevel"
End Sub
Public Sub GeneralMCI(OpenString As String, CloseString As String)
Error = mciSendString2(OpenString, 0&, 0, 0)
If Error = 0 Then Terminate0 = CloseString
End Sub
Public Sub oneMCI(NowString As String)
Error = mciSendString2(NowString, 0&, 0, 0)
End Sub
Private Function NeedByteFactor() As Long
If bt = 0 Then Bit8
If ch = 0 Then Mono
NeedByteFactor = bt * ch
If NeedByteFactor = 0 Then
NeedByteFactor = 1
End If
End Function

Public Sub Bit8()

 Error = mciSendString2("set mycapture bitspersample 8", 0&, 0, 0)
 If Error <= 0 Then bt = 1
End Sub
Public Sub Bit16()
 Error = mciSendString2("set mycapture bitspersample 16", 0&, 0, 0)
  If Error <= 0 Then bt = 2
End Sub
Public Sub QualityRadio()

  Error = mciSendString2("set mycapture samplespersec 11025 BYTESPERSEC " & CStr(11025 * NeedByteFactor), 0&, 0, 0)
  
End Sub
Public Sub QualityAny(qrec As Double)
 Error = mciSendString2("set mycapture samplespersec " & (CLng(qrec)) & " BYTESPERSEC " & (CLng(qrec) * NeedByteFactor), 0&, 0, 0)
End Sub
Public Sub Mono()
Error = mciSendString2("set mycapture channels 1", 0&, 0, 0)
 If Error <= 0 Then ch = 1
End Sub
Public Sub Stereo()
Error = mciSendString2("set mycapture channels 2", 0&, 0, 0)
If Error <= 0 Then ch = 2
End Sub
Public Sub RecFast()
Bit8
QualityRadio
Mono
Capture
' stop
' save and play
' or play
End Sub
Public Sub Capture(Optional ByVal ins As Boolean = False)
If ins Then
Error = mciSendString2("record mycapture insert", 0, 0, 0)
Else
Error = mciSendString2("record mycapture", 0, 0, 0)
End If
End Sub
Public Sub ReCapture()
Error = mciSendString2("record mycapture overwrite", 0, 0, 0)
End Sub
Sub CutRecordMs(recStart As Double, recEnd As Double)
Error = mciSendString2("set mycapture time format ms", 0, 0, 0)
Error = mciSendString2("delete mycapture from " & (recStart) & " to " & (recEnd), 0&, 0, 0)

End Sub
Public Function getPositionInMS() As Long
    'Get the position of the movie in milliseconds
    Error = mciSendString2("set mycapture time format ms", 0, 0, 0)
    Error = mciSendString("status mycapture position", Data, 128, 0)
    getPositionInMS = val(StrPart(Data))
End Function
Public Function getLevels() As Long
If Terminate1 <> "" Then
Error = mciSendString("status waveLevel level", Data, 128, 0)
getLevels = val(StrPart(Data))
End If
End Function
Public Function getLengthInMS() As Long
    'Get the position of the movie in milliseconds
    Error = mciSendString2("set mycapture time format ms", 0, 0, 0)
    Error = mciSendString("status mycapture length", Data, 128, 0)
    getLengthInMS = val(StrPart(Data))
End Function
Sub recStop()
Error = mciSendString2("stop mycapture", 0, 0, 0)
End Sub
Sub recPlay()
Dim again As Boolean
If Terminate1 <> "" Then Error = mciSendString2(Terminate1, 0, 0, 0): again = True
Error = mciSendString2("set mycapture time format ms", 0, 0, 0)
Error = mciSendString2("play mycapture from 0", 0, 0, 0)
If again Then Error = mciSendString2("open new type waveaudio alias waveLevel", 0, 0, 0)

End Sub
Sub recPlayFromMs(recPos As Double)
recPos = CLng(recPos)
Error = mciSendString2("set mycapture time format ms", 0, 0, 0)
Error = mciSendString2("play mycapture from " & CStr(recPos), 0, 0, 0)
End Sub
Public Function setVolume(Value As Long)
 Error = mciSendString2("setaudio mycapture volume to " & CStr(Value), 0, 0, 0)

End Function
Public Function isRecPlaying() As Boolean
    Dim isPlaying As String
    Error = mciSendString("status mycapture mode", Data, 128, 0)
    MyDoEvents
    isPlaying = Left(StrPart(Data), 7)
    If isPlaying = "playing" And Not NOEXECUTION Then
        isRecPlaying = True
    Else
        isRecPlaying = False
    End If
End Function
Sub Save()
If mFilename <> "" Then
Error = mciSendString2("save mycapture " & """" + mFilename + """", 0&, 0, 0)
End If
End Sub
Sub AnyMci(CommandString As String)
Error = mciSendString2(CommandString, 0&, 0, 0)

End Sub
Function AnyMciValue(CommandString As String)
 Error = mciSendString(CommandString, Data, 128, 0)
    AnyMciValue = val(StrPart(Data))
End Function
Public Function checkError() As String
    'A very useful function for getting any errors
    'associated with the mci device
    checkError = space$(255)
    mciGetErrorString Error, checkError, Len(checkError)
End Function
Sub SaveAs(fnm As String)

Error = mciSendString2("save mycapture " & """" + fnm + """", 0&, 0, 0)
If Error <= 0 Then
mFilename = fnm
End If

End Sub
Public Property Get HaveMic() As Boolean
HaveMic = waveInGetNumDevs()
End Property

Private Sub Class_Terminate()
If Terminate0 <> "" Then
Error = mciSendString2(Terminate0, 0, 0, 0)
If Terminate1 <> "" Then Error = mciSendString2(Terminate1, 0, 0, 0)
End If
End Sub
Private Function StrPart(sBuffer As String)
StrPart = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
End Function
