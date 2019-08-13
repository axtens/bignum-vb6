Attribute VB_Name = "Module1"
Sub main()
    Dim b As New BIGNUM
    Dim L As New BIGNUMLIB
    Set b = L.toBignum("255")
        
    Dim hex As String
    hex = L.bignumToBase(b, "0123456789ABCDEF")
    Debug.Print hex
    Set b = L.baseToBignum("1DEADBEEF", "0123456789ABCDEF")
    
    'Set b = l.toBignum("255")
    Debug.Print "prec=" & L.Precision
    
    Debug.Print "b", b.toString()
    Debug.Print "b", L.toString(b)
    Dim a As Object
    Set a = L.toBignum(1)
    Dim c As New BIGNUM
    Set c = L.Add(a, b)
    Debug.Print a.toString() & " + " & b.toString() & " = "
    Debug.Print c.toString()
    
    L.intToBignum 0, c
    Set c = L.Subtract(a, b)
    Debug.Print a.toString() & " - " & b.toString() & " = "
    Debug.Print c.toString()
    
    L.intToBignum 0, c
    Set c = L.Multiply(a, b)
    Debug.Print a.toString() & " * " & b.toString() & " = "
    Debug.Print c.toString()
    
    Dim x As BIGNUM
    Dim y As BIGNUM
    Dim z As BIGNUM
    Set x = L.toBignum(4)
    Set y = L.toBignum(2)
    Set z = L.Divide(x, y)
    Debug.Print x.toString() & "\" & y.toString() & "=";
    Debug.Print z.toString()

    Dim i As Long
    Dim j As Long
    For i = 1 To 12
        For j = 1 To 12
            Set x = L.toBignum(i)
            Set y = L.toBignum(j)
            Set z = L.Multiply(x, y)
            Debug.Print z.toString(),
        Next
        Debug.Print
    Next
    For i = 1 To 12
        For j = 1 To 12
            Set x = L.toBignum(i)
            Set y = L.toBignum(j)
            Set z = L.Modulus(x, y)
            Debug.Print x.toString() & "%" & y.toString() & "=" & z.toString()
        Next
        Debug.Print
    Next
    Dim TWO As BIGNUM
    Set TWO = L.toBignum(2)
    Dim sTWO As String
    sTWO = TWO.toString()
    For j = 1 To 1000
        Set z = L.Power(TWO, j)
        Debug.Print sTWO & "^" & CStr(j) & "=" & z.toString()
    Next
    
End Sub
