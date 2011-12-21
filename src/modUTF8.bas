Attribute VB_Name = "modUTF8"
Option Explicit

Private Const U8_TWO_OCTETS As Byte = &HC0&
Private Const U8_THREE_OCTETS As Byte = &HE0&
Private Const U8_FOUR_OCTETS As Byte = &HF0&

Private Const U8_ASCII_CEILING As Long = &HFF&
Private Const U8_TWO_OCTET_CEILING As Long = &H800&
Private Const U8_THREE_OCTET_CEILING As Long = &H10000
Private Const U8_FOUR_OCTET_CEILING As Long = &H110000

Public Function UTF8Encode(Source As String) As String
    Dim result As String
    Dim charCode As Long
    Dim count As Long
    Dim temp As String
    
    For count = 1 To Len(Source)
        temp = vbNullString
        charCode = AscW(Mid$(Source, count, 1))
        
        If charCode < U8_ASCII_CEILING Then
            result = result & ChrW$(charCode)
        ElseIf charCode < U8_TWO_OCTET_CEILING Then
            temp = Chr$(&H80 + (charCode And &H3F))
            charCode = charCode \ &H40
            temp = Chr$(&HC0 + (charCode And &H1F)) & temp
        ElseIf charCode < U8_THREE_OCTET_CEILING Then
            temp = Chr$(&H80 + (charCode And &H3F))
            charCode = charCode \ &H40
            temp = Chr$(&H80 + (charCode And &H3F)) & temp
            charCode = charCode \ &H40
            temp = Chr$(&HE0 + (charCode And &HF)) & temp
        ElseIf charCode < U8_FOUR_OCTET_CEILING Then
            temp = Chr$(&H80 + (charCode And &H3F))
            charCode = charCode \ &H40
            temp = Chr$(&H80 + (charCode And &H3F)) & temp
            charCode = charCode \ &H40
            temp = Chr$(&H80 + (charCode And &H3F)) & temp
            charCode = charCode \ &H40
            temp = Chr$(&HF0 + (charCode And &H7)) & temp
        Else
            result = result & ChrW$(charCode)
        End If
        
        result = result & temp
    Next count
    
    UTF8Encode = result
End Function

Public Function UTF8Decode(utf8 As String) As String
    Dim bytes() As Byte
    Dim count As Long
    Dim result As String
    
    Dim octet1 As Byte
    Dim octet2 As Byte
    Dim octet3 As Byte
    Dim octet4 As Byte
    
    Dim charCode As Long
    
   On Error GoTo UTF8Decode_Error

    bytes = StrConv(utf8, vbFromUnicode)
    
    For count = 0 To UBound(bytes)
        If bytes(count) < &HC0 Then
            result = result & Chr$(bytes(count))
        ElseIf (bytes(count) And U8_FOUR_OCTETS) = U8_FOUR_OCTETS Then
            If count > UBound(bytes) - 3 Then
                result = result & Chr$(bytes(count))
                GoTo loopend
            End If
        
            octet1 = bytes(count)
            count = count + 1
            
            octet2 = bytes(count)
            
            If (octet2 And &H80) = 0 Or (octet2 And &H40) <> 0 Then
                result = result & Chr$(octet1) & Chr$(octet2)
                GoTo loopend
            End If
            
            count = count + 1
            octet3 = bytes(count)
            
            If (octet3 And &H80) = 0 Or (octet3 And &H40) <> 0 Then
                result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3)
                GoTo loopend
            End If
            
            count = count + 1
            octet4 = bytes(count)
            
            If (octet4 And &H80) = 0 Or (octet4 And &H40) <> 0 Then
                result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3) & Chr$(octet4)
                GoTo loopend
            End If
            
            charCode = ((octet1 And &H7) * &H1000000) + ((octet2 And &H3F) * &H1000&) + ((octet3 And &H3F) * &H40&) + (octet4 And &H3F)
            
            If charCode <= 65535 Then
                result = ChrW$(charCode)
            Else
                result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3) & Chr$(octet4)
            End If
        ElseIf (bytes(count) And U8_THREE_OCTETS) = U8_THREE_OCTETS Then
            If count > UBound(bytes) - 2 Then
                result = result & Chr$(bytes(count))
                GoTo loopend
            End If
        
            octet1 = bytes(count)
            count = count + 1
            
            octet2 = bytes(count)
            
            If (octet2 And &H80) = 0 Or (octet2 And &H40) <> 0 Then
                result = result & Chr$(octet1) & Chr$(octet2)
                GoTo loopend
            End If
            
            count = count + 1
            
            octet3 = bytes(count)
            
            If (octet3 And &H80) = 0 Or (octet3 And &H40) <> 0 Then
                result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3)
                GoTo loopend
            End If
            
            charCode = ((octet1 And &HF) * &H1000&) + ((octet2 And &H3F) * &H40&) + (octet3 And &H3F)
            
            If charCode <= 65535 Then
                result = result & ChrW$(charCode)
            Else
                result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3)
            End If
        ElseIf (bytes(count) And U8_TWO_OCTETS) = U8_TWO_OCTETS Then
            If count > UBound(bytes) - 1 Then
                result = result & Chr$(bytes(count))
                GoTo loopend
            End If
        
            octet1 = bytes(count)
            count = count + 1
            
            octet2 = bytes(count)
            
            If (octet2 And &H80) = 0 Or (octet2 And &H40) <> 0 Then
                result = result & Chr$(octet1) & Chr$(octet2)
                GoTo loopend
            End If
             
            charCode = ((octet1 And &H1F) * &H40) + (octet2 And &H3F)
            
            If charCode <= 65535 Then
                result = result & ChrW$(charCode)
            Else
                result = result & Chr$(octet1) & Chr$(octet2)
            End If
        End If
        
loopend:
    Next count
    
    UTF8Decode = result

   On Error GoTo 0
   Exit Function

UTF8Decode_Error:
    handleError "UTF8Decode", Err.Number, Err.Description, Erl, utf8
End Function

