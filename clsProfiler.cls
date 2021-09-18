VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProfiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MARK1 As Currency
Private Declare Function QueryPerformanceCounter Lib "KERNEL32" _
   (lpPerformanceCount As Currency) _
   As Boolean
Private Declare Function QueryPerformanceFrequency Lib "KERNEL32" _
   (lpPerformanceCount As Currency) _
   As Boolean
Private CountsPerSecond As Currency

Public Sub MARKONE()
Dim CTR As Currency
QueryPerformanceCounter CTR
MARK1 = CTR
End Sub
Public Function MARKTWO() As Double
Dim MARK2 As Currency
If QueryPerformanceCounter(MARK2) Then
If CountsPerSecond = 0 Then QueryPerformanceFrequency CountsPerSecond
MARKTWO = ((MARK2 - MARK1) / CountsPerSecond) * 1000
End If
End Function
Public Property Get getAPIOverhead() As Currency
   Dim a As Currency, b As Currency, c As Currency
   If QueryPerformanceCounter(a) Then
      QueryPerformanceCounter b
      
      If CountsPerSecond = 0 Then QueryPerformanceFrequency CountsPerSecond
      c = b - a
      If c = 0 Then
         getAPIOverhead = 0
      Else
      
         getAPIOverhead = c '
      End If
   Else
      getAPIOverhead = -1
     
   End If
End Property

