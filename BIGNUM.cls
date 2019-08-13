VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIGNUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PLUS = 1
Private Const MINUS = -1

Public infinity As Boolean
Public signbit As Integer
Public lastdigit As Integer
Public dijits As String


Public Function toString() As String
    Dim result As String
    result = StrReverse(Mid$(dijits, 1, lastdigit))
    Dim i As Integer
    For i = 1 To Len(result)
        Mid$(result, i, 1) = Chr$(Asc(Mid$(result, i, 1)) + &H30)
    Next
    toString = IIf(signbit = MINUS, "-", "") & result
End Function

Public Function Digits(item As Integer, Optional value As Integer = -1) As Integer
    Dim result As String
    If value = -1 Then
        'assume asking for element item of Digits string
        Digits = Asc(Mid$(dijits, item, 1))
    Else
        'assume giving value for item
        Mid$(dijits, item, 1) = Chr$(value)
        Digits = value
    End If
End Function

Public Function SetPrecision(Optional nPrec As Integer = 100)
    If nPrec < Len(dijits) Then
        dijits = Left$(dijits, nPrec)
    Else
        dijits = dijits & String$(nPrec - Len(dijits), Chr$(0))
    End If
    SetPrecision = Len(dijits)
End Function

Public Function Clone()
    Dim a As New BIGNUM
    a.signbit = signbit
    a.lastdigit = lastdigit
    a.dijits = dijits
    Set Clone = a
End Function

Private Sub Class_Initialize()
    infinity = False
    signbit = 1
    lastdigit = 0
    dijits = String$(1000, Chr$(0))
End Sub
