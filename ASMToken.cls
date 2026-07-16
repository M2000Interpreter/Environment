VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASMToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TType    As TokenType
Public Content  As String
Public Value    As Long
Public Line     As Long
Public Position As Long
Public Section  As String


Private Sub Class_Initialize()
    TType = TokenUnknown
    Content = ""
    Value = -1
    Line = 0
    Position = 0
End Sub
