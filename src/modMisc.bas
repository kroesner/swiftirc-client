Attribute VB_Name = "modMisc"
Option Explicit

Public Function swiftFormatTime(timestamp As Date, theFormat As String) As String
    Dim count As Long
    Dim char As String
    Dim nextChar As String
    Dim formatStr As String
    
    For count = 1 To Len(theFormat)
        char = Mid$(theFormat, count, 1)
        
        If count < Len(theFormat) Then
            nextChar = LCase$(Mid$(theFormat, count + 1, 1))
        Else
            nextChar = vbNullString
        End If
        
        If char = "h" Or char = "H" Then
            formatStr = "h"
            
            If nextChar = "h" Then
                formatStr = formatStr & "h"
                count = count + 1
            End If
            
            swiftFormatTime = swiftFormatTime & format(timestamp, formatStr)
        ElseIf char = "m" Or char = "M" Then
            If nextChar = "m" Then
                count = count + 1
            End If
            
            swiftFormatTime = swiftFormatTime & format(timestamp, "nn")
        ElseIf char = "s" Or char = "S" Then
            If nextChar = "s" Then
                count = count + 1
            End If
            
            swiftFormatTime = swiftFormatTime & format(timestamp, "ss")
        Else
            swiftFormatTime = swiftFormatTime & char
        End If
    Next count
End Function

Public Function encrypt(pass As String, text As String) As String
    Dim aes As New cRijndael
    Dim plain() As Byte
    Dim crypt() As Byte
    Dim realPass() As Byte
    
    realPass = StrConv(pass, vbFromUnicode)
    ReDim Preserve realPass(31)
    
    aes.SetCipherKey realPass, 256
    
    plain = StrConv(text, vbFromUnicode)
    
    aes.ArrayEncrypt plain, crypt, 0

    encrypt = HexDisplay(crypt, UBound(crypt) + 1, 16)
End Function

Public Function decrypt(pass As String, text As String) As String
    Dim aes As New cRijndael
    
    Dim crypt() As Byte
    Dim plain() As Byte
    Dim realPass() As Byte
    
    realPass = StrConv(pass, vbFromUnicode)
    ReDim Preserve realPass(31)
    
    aes.SetCipherKey realPass, 256
    
    If HexDisplayRev(text, crypt) = 0 Then
        decrypt = text
        Exit Function
    End If
    
    aes.ArrayDecrypt plain, crypt, 0
    decrypt = StrConv(plain, vbUnicode)
End Function

'Returns a String containing Hex values of data(0 ... n-1) in groups of k
Private Function HexDisplay(data() As Byte, n As Long, k As Long) As String
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim data2() As Byte

    If LBound(data) = 0 Then
        ReDim data2(n * 4 - 1 + ((n - 1) \ k) * 4)
        j = 0
        For i = 0 To n - 1
            If i Mod k = 0 Then
                If i <> 0 Then
                    data2(j) = 32
                    data2(j + 2) = 32
                    j = j + 4
                End If
            End If
            c = data(i) \ 16&
            If c < 10 Then
                data2(j) = c + 48     ' "0"..."9"
            Else
                data2(j) = c + 55     ' "A"..."F"
            End If
            c = data(i) And 15&
            If c < 10 Then
                data2(j + 2) = c + 48 ' "0"..."9"
            Else
                data2(j + 2) = c + 55 ' "A"..."F"
            End If
            j = j + 4
        Next i
Debug.Assert j = UBound(data2) + 1
        HexDisplay = data2
    End If

End Function


'Reverse of HexDisplay.  Given a String containing Hex values, convert to byte array data()
'Returns number of bytes n in data(0 ... n-1)
Private Function HexDisplayRev(TheString As String, data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d As Long
    Dim n As Long
    Dim data2() As Byte

    n = 2 * Len(TheString)
    data2 = TheString

    ReDim data(n \ 4 - 1)

    d = 0
    i = 0
    j = 0
    Do While j < n
        c = data2(j)
        Select Case c
        Case 48 To 57    '"0" ... "9"
            If d = 0 Then   'high
                d = c
            Else            'low
                data(i) = (c - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70   '"A" ... "F"
            If d = 0 Then   'high
                d = c - 7
            Else            'low
                data(i) = (c - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102  '"a" ... "f"
            If d = 0 Then   'high
                d = c - 39
            Else            'low
                data(i) = (c - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase data
    Else
        ReDim Preserve data(n - 1)
    End If
    HexDisplayRev = n

End Function
