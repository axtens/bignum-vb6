VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIGNUMLIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PLUS = 1
Const MINUS = -1
Private prec As Integer
Private memo As New Scripting.Dictionary

Public lfPosInf As Double
Public lfNegInf As Double


Private Function ismemo(key As String) As Boolean
    Dim o As Object
    ismemo = memo.Exists(key)
End Function

Private Function getmemo(key As String) As Variant
    Dim i As Integer
    Select Case VarType(memo.item(key))
    Case vbObject
        Set getmemo = memo.item(key)
    Case Else
        getmemo = memo.item(key)
    End Select
End Function

Private Sub setmemo(key As String, b As Variant)
    If Not ismemo(key) Then
        memo.Add key, b
    End If
End Sub

Public Function Max(a As Integer, b As Integer) As Integer
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function
Public Function toString(n As BIGNUM) As String
    toString = n.toString()
End Function

Public Function toBignum(a As Variant) As BIGNUM
    Dim c As New BIGNUM
    Dim t As Integer
    Select Case VarType(a)
    Case vbString
        Set c = toBignum(0)
        If a <> vbNullString Then
            
            If Left$(a, 1) = "-" Then
                t = 1
                c.signbit = MINUS
            Else
                t = 0
                c.signbit = PLUS
            End If
            c.lastdigit = 0
            Dim i As Integer
            For i = Len(a) To 1 Step -1
                c.lastdigit = c.lastdigit + 1
                c.Digits c.lastdigit, Mid$(a, i, 1)
            Next
        End If
        Set toBignum = c
        
    Case vbInteger, vbLong
        t = Abs(a)
        If t = 0 Then
            c.lastdigit = 1
        Else
            While t > 0
                c.lastdigit = c.lastdigit + 1
                c.Digits c.lastdigit, t Mod 10
                t = t \ 10
            Wend
        End If
        
    Case Else
    End Select
    Set toBignum = c
End Function

Public Function intToBignum(s As Long, n As BIGNUM) As BIGNUM
    n.signbit = IIf(s >= 0, PLUS, MINUS)
    n.lastdigit = 0
    Dim t As Long
    t = Abs(s)
    While t > 0
        n.lastdigit = n.lastdigit + 1
        n.Digits n.lastdigit, t Mod 10
        t = t \ 10
    Wend
    Set intToBignum = n.Clone()
End Function

Public Function Precision(Optional nPrec As Integer = 100)
    prec = nPrec
    Precision = prec
End Function

Public Function Add(a As BIGNUM, b As BIGNUM) As BIGNUM
    
    Dim aa As String
    aa = a.toString()
    Dim bb As String
    bb = b.toString()
    
    If aa = "0" Then
        Set Add = b.Clone()
        Exit Function
    End If
    
    If bb = "0" Then
        Set Add = a.Clone()
        Exit Function
    End If
       
    Dim sKey As String
    sKey = aa & "+" & bb
    If ismemo(sKey) Then
        Set Add = getmemo(sKey)
        Exit Function
    End If
    
    Dim c As New BIGNUM
    
    
    If a.signbit = b.signbit Then
        c.signbit = a.signbit
    Else
        If a.signbit = MINUS Then
            a.signbit = PLUS
            Set c = Subtract(b, a)
            a.signbit = MINUS
        Else
            b.signbit = PLUS
            Set c = Subtract(a, b)
            b.signbit = MINUS
        End If
        Set Add = c
        Exit Function
    End If
    
    c.lastdigit = Max(a.lastdigit, b.lastdigit) + 1
    Dim i As Integer
    Dim carry As Integer
    Dim temp As Integer
    For i = 1 To c.lastdigit
        temp = carry + a.Digits(i) + b.Digits(i)
        c.Digits i, temp Mod 10
        carry = temp \ 10
    Next
    Set c = zeroJustify(c)
    setmemo sKey, c
    Set Add = c
End Function

Public Function Subtract(a As BIGNUM, b As BIGNUM) As BIGNUM
    Dim aa As String
    aa = a.toString()
    Dim bb As String
    bb = b.toString()
    
    If aa = "0" Then
        b.signbit = MINUS
        Set Subtract = b.Clone()
        Exit Function
    End If
    
    If bb = "0" Then
        Set Subtract = a.Clone()
        Exit Function
    End If
       
    Dim sKey As String
    sKey = aa & "-" & bb
    If ismemo(sKey) Then
        Set Subtract = getmemo(sKey)
        Exit Function
    End If
    
    
    Dim borrow As Integer
    Dim v As Integer
    Dim i As Integer
    Dim c As New BIGNUM
    If a.signbit = MINUS Or b.signbit = MINUS Then
        b.signbit = -1 * b.signbit
        Set c = Add(a, b)
        b.signbit = -1 * b.signbit
        Set Subtract = c
        Exit Function
    End If
    If Compare(a, b) = PLUS Then
        Set c = Subtract(b, a)
        c.signbit = MINUS
        Set Subtract = c
        Exit Function
    End If
    
    c.lastdigit = Max(a.lastdigit, b.lastdigit)
    borrow = 0
    
    For i = 1 To c.lastdigit
        v = a.Digits(i) - borrow - b.Digits(i)
        If a.Digits(i) > 0 Then
            borrow = 0
        End If
        If v < 0 Then
            v = v + 10
            borrow = 1
        End If
        c.Digits i, v Mod 10
    Next
    Set c = zeroJustify(c)
    setmemo sKey, c
    Set Subtract = c
End Function

Function Compare(a As BIGNUM, b As BIGNUM) As Integer
    Dim aa As String
    aa = a.toString()
    Dim bb As String
    bb = b.toString()
    
    Dim sKey As String
    sKey = aa & "$" & bb
    If ismemo(sKey) Then
        Compare = getmemo(sKey)
        Exit Function
    End If
    
    
    Dim res As Integer
    Dim i As Integer
    res = 0
    If a.signbit = MINUS And b.signbit = PLUS Then
        res = PLUS
    Else
        If a.signbit = PLUS And b.signbit = MINUS Then
            res = MINUS
        Else
            If b.lastdigit > a.lastdigit Then
                res = PLUS * a.signbit
            Else
                If a.lastdigit > b.lastdigit Then
                    res = MINUS * a.signbit
                Else
                    For i = a.lastdigit To 1 Step -1
                        If a.Digits(i) > b.Digits(i) Then
                            res = MINUS * a.signbit
                            Exit For
                        End If
                        If b.Digits(i) > a.Digits(i) Then
                            res = PLUS * a.signbit
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    End If
    setmemo sKey, res
    Compare = res
End Function

Public Function zeroJustify(n As BIGNUM) As BIGNUM
    While n.lastdigit > 1 And n.Digits(n.lastdigit) = 0
        n.lastdigit = n.lastdigit - 1
    Wend
    If n.lastdigit = 1 And n.Digits(1) = 0 Then
        n.signbit = PLUS
    End If
    Set zeroJustify = n
End Function

Public Function digitShift(n As BIGNUM, d As Integer) As BIGNUM
    Dim i As Integer
    If n.lastdigit = 1 And n.Digits(1) = 0 Then
        Set digitShift = n
        Exit Function
    End If
    For i = n.lastdigit To 1 Step -1
        n.Digits i + d, n.Digits(i)
    Next
    For i = 1 To d
        n.Digits i, 0
    Next
    n.lastdigit = n.lastdigit + d
    Set digitShift = n
End Function

Public Function Multiply(a As BIGNUM, b As BIGNUM) As BIGNUM
    Dim aa As String
    aa = a.toString()
    Dim bb As String
    bb = b.toString()
    
    If aa = "0" Then
        Set Multiply = a.Clone()
        Exit Function
    End If
    
    If bb = "0" Then
        Set Multiply = b.Clone()
        Exit Function
    End If
       
    Dim sKey As String
    sKey = aa & "*" & bb
    If ismemo(sKey) Then
        Set Multiply = getmemo(sKey)
        Exit Function
    End If
    
    Dim row As BIGNUM
    Dim tmp As BIGNUM
    Dim c As BIGNUM
    Set c = toBignum(0)
    Set row = a.Clone()
    Dim i As Integer
    Dim j As Integer
    For i = 1 To b.lastdigit
        For j = 1 To b.Digits(i)
            Set tmp = Add(c, row)
            Set c = tmp.Clone()
        Next
        Set row = digitShift(row, 1)
    Next
    c.signbit = a.signbit * b.signbit
    Set c = zeroJustify(c)
    setmemo sKey, c
    Set Multiply = c
End Function

Function Divide(a As BIGNUM, b As BIGNUM) As BIGNUM
    Dim aa As String
    aa = a.toString()
    Dim bb As String
    bb = b.toString()
    
    Dim c As BIGNUM
    
    If aa = "0" Then
        Set Divide = a.Clone()
        Exit Function
    End If
    
    If bb = "0" Then
        Set c = New BIGNUM
        c.infinity = True
        Set Divide = c.Clone()
        Exit Function
    End If
       
    Dim sKey As String
    sKey = aa & "/" & bb
    If ismemo(sKey) Then
        Set Divide = getmemo(sKey)
        Exit Function
    End If
    
    Dim asign As Integer
    Dim bsign As Integer
    Dim i As Integer
    Dim j As Integer
    Set c = toBignum(0)
    
    c.signbit = a.signbit * b.signbit
    asign = a.signbit
    bsign = b.signbit
    a.signbit = PLUS
    b.signbit = PLUS
    Dim row As New BIGNUM
    Dim tmp As New BIGNUM
    c.lastdigit = a.lastdigit
    For i = a.lastdigit To 1 Step -1
        Set row = digitShift(row, 1)
        row.Digits 1, a.Digits(i)
        c.Digits i, 0
        While Compare(row, b) <> PLUS
            c.Digits i, c.Digits(i) + 1
            Set tmp = Subtract(row, b)
            Set row = tmp.Clone()
        Wend
    Next
    Set c = zeroJustify(c)
    a.signbit = asign
    b.signbit = bsign
    setmemo sKey, c
    Set Divide = c
End Function

Public Function Power(a As BIGNUM, n As Long) As BIGNUM
    Dim aa As String
    aa = a.toString()
    
    If aa = "0" Then
        Set Power = a.Clone()
        Exit Function
    End If
        
    Dim sKey As String
    sKey = aa & "^" & CStr(n)
    If ismemo(sKey) Then
        Set Power = getmemo(sKey)
        Exit Function
    End If
    
    Dim c As BIGNUM
    Dim r As BIGNUM
    If n = 0 Then
        Set c = toBignum(1)
    Else
        Set c = a.Clone()
        Set r = toBignum(0)
        Dim i As Long
        For i = 2 To n
            Set r = Multiply(a, c)
            Set c = r.Clone()
        Next
    End If
    Set Power = c
End Function

Public Function Modulus(a As BIGNUM, b As BIGNUM) As BIGNUM
    Dim aa As String
    aa = a.toString()
    Dim bb As String
    bb = b.toString()
    
    Dim c As BIGNUM
    
    If aa = "0" Then
        Set Modulus = a.Clone()
        Exit Function
    End If
    
    If bb = "0" Then
        c.infinity = True
        Set Modulus = c.Clone()
        Exit Function
    End If
    
    Dim sKey As String
    sKey = aa & "%" & bb
    If ismemo(sKey) Then
        Set Modulus = getmemo(sKey)
        Exit Function
    End If
    
    Set c = toBignum(0)
    Dim comp As Integer
    comp = Compare(a, b)
    If comp = MINUS Then
        Dim divi As BIGNUM
        Dim mult As BIGNUM
        Dim subt As BIGNUM
        Set divi = Divide(a, b)
        Set mult = Multiply(divi, b)
        Set subt = Subtract(a, mult)
        Set c = subt.Clone()
    Else
        If comp = PLUS Then
            Set c = a.Clone()
        Else
            Set c = toBignum(0)
        End If
    End If
    setmemo sKey, c
    Set Modulus = c.Clone()
End Function

Public Function bignumToBase(a As BIGNUM, baseDigits As String) As String
    Dim base As BIGNUM
    Set base = toBignum(Len(baseDigits))
    Dim zero As BIGNUM
    Set zero = toBignum(0)
    Dim temp As BIGNUM
    Set temp = a.Clone()
    Dim c As String
    c = vbNullString
    While Compare(zero, temp) = PLUS
        Dim offs As BIGNUM
        Set offs = Modulus(temp, base)
        Dim bts As String
        bts = offs.toString()
        Dim i As Integer
        i = Val(bts)
        c = Mid$(baseDigits, i + 1, 1) & c
        Set temp = Divide(temp, base)
    Wend
    bignumToBase = c
End Function

Public Function baseToBignum(basedNumber As String, baseDigits As String) As BIGNUM
    Dim base As BIGNUM
    Set base = toBignum(Len(baseDigits))
    Dim b As BIGNUM
    Set b = toBignum(0)
    Dim c As String
    Dim i As Long
    Dim j As Long
    j = 0
    For i = Len(basedNumber) To 1 Step -1
        c = Mid$(basedNumber, i)
        Dim pwr As BIGNUM
        Set pwr = Power(base, j)
        Dim ioc As BIGNUM
        Set ioc = toBignum(InStr(baseDigits, c) - 1)
        Dim mul As BIGNUM
        Set mul = Multiply(pwr, ioc)
        Set b = Add(b, mul)
        j = j + 1
    Next
    Set baseToBignum = b.Clone()
End Function

'    -- base_to_bignum : function (based, numset) {
'      -- var base = this.toBignum(numset.length);
'      -- var bignum = this.toBignum(0);
'      -- var c = 0;
'      -- for (var i = based.length - 1, j = 0; i >= 0; i--, j++) {
'        -- c = based.substr(i, 1);
'        -- var pwr = this.power_bignum(base, j);
'        -- var ioc = this.toBignum(numset.indexOf(c));
'        -- var mul;
'        -- if (this.compare_bignum(ioc, pwr) === PLUS) {
'          -- mul = this.multiply_bignum(pwr, ioc);
'        -- } else {
'          -- mul = this.multiply_bignum(ioc, pwr);
'        -- }
'        -- bignum = this.add_bignum(bignum, mul);
'      -- }
'      -- return bignum;
'    -- }
'
Private Sub Class_Initialize()
    On Error Resume Next
    lfPosInf = 1 / 0
    On Error GoTo 0
    lfNegInf = -lfPosInf
End Sub
