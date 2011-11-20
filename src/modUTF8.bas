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
          
10        For count = 1 To Len(Source)
20            temp = vbNullString
30            charCode = AscW(Mid$(Source, count, 1))
              
40            If charCode < U8_ASCII_CEILING Then
50                result = result & ChrW$(charCode)
60            ElseIf charCode < U8_TWO_OCTET_CEILING Then
70                temp = Chr$(&H80 + (charCode And &H3F))
80                charCode = charCode \ &H40
90                temp = Chr$(&HC0 + (charCode And &H1F)) & temp
100           ElseIf charCode < U8_THREE_OCTET_CEILING Then
110               temp = Chr$(&H80 + (charCode And &H3F))
120               charCode = charCode \ &H40
130               temp = Chr$(&H80 + (charCode And &H3F)) & temp
140               charCode = charCode \ &H40
150               temp = Chr$(&HE0 + (charCode And &HF)) & temp
160           ElseIf charCode < U8_FOUR_OCTET_CEILING Then
170               temp = Chr$(&H80 + (charCode And &H3F))
180               charCode = charCode \ &H40
190               temp = Chr$(&H80 + (charCode And &H3F)) & temp
200               charCode = charCode \ &H40
210               temp = Chr$(&H80 + (charCode And &H3F)) & temp
220               charCode = charCode \ &H40
230               temp = Chr$(&HF0 + (charCode And &H7)) & temp
240           Else
250               result = result & ChrW$(charCode)
260           End If
              
270           result = result & temp
280       Next count
          
290       UTF8Encode = result
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
          
10       On Error GoTo UTF8Decode_Error

20        bytes = StrConv(utf8, vbFromUnicode)
          
30        For count = 0 To UBound(bytes)
40            If bytes(count) < &HC0 Then
50                result = result & Chr$(bytes(count))
60            ElseIf (bytes(count) And U8_FOUR_OCTETS) = U8_FOUR_OCTETS Then
70                If count > UBound(bytes) - 3 Then
80                    result = result & Chr$(bytes(count))
90                    GoTo loopend
100               End If
              
110               octet1 = bytes(count)
120               count = count + 1
                  
130               octet2 = bytes(count)
                  
140               If (octet2 And &H80) = 0 Or (octet2 And &H40) <> 0 Then
150                   result = result & Chr$(octet1) & Chr$(octet2)
160                   GoTo loopend
170               End If
                  
180               count = count + 1
190               octet3 = bytes(count)
                  
200               If (octet3 And &H80) = 0 Or (octet3 And &H40) <> 0 Then
210                   result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3)
220                   GoTo loopend
230               End If
                  
240               count = count + 1
250               octet4 = bytes(count)
                  
260               If (octet4 And &H80) = 0 Or (octet4 And &H40) <> 0 Then
270                   result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3) & Chr$(octet4)
280                   GoTo loopend
290               End If
                  
300               charCode = ((octet1 And &H7) * &H1000000) + ((octet2 And &H3F) * &H1000&) + ((octet3 And &H3F) * &H40&) + (octet4 And &H3F)
                  
310               If charCode <= 65535 Then
320                   result = ChrW$(charCode)
330               Else
340                   result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3) & Chr$(octet4)
350               End If
360           ElseIf (bytes(count) And U8_THREE_OCTETS) = U8_THREE_OCTETS Then
370               If count > UBound(bytes) - 2 Then
380                   result = result & Chr$(bytes(count))
390                   GoTo loopend
400               End If
              
410               octet1 = bytes(count)
420               count = count + 1
                  
430               octet2 = bytes(count)
                  
440               If (octet2 And &H80) = 0 Or (octet2 And &H40) <> 0 Then
450                   result = result & Chr$(octet1) & Chr$(octet2)
460                   GoTo loopend
470               End If
                  
480               count = count + 1
                  
490               octet3 = bytes(count)
                  
500               If (octet3 And &H80) = 0 Or (octet3 And &H40) <> 0 Then
510                   result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3)
520                   GoTo loopend
530               End If
                  
540               charCode = ((octet1 And &HF) * &H1000&) + ((octet2 And &H3F) * &H40&) + (octet3 And &H3F)
                  
550               If charCode <= 65535 Then
560                   result = result & ChrW$(charCode)
570               Else
580                   result = result & Chr$(octet1) & Chr$(octet2) & Chr$(octet3)
590               End If
600           ElseIf (bytes(count) And U8_TWO_OCTETS) = U8_TWO_OCTETS Then
610               If count > UBound(bytes) - 1 Then
620                   result = result & Chr$(bytes(count))
630                   GoTo loopend
640               End If
              
650               octet1 = bytes(count)
660               count = count + 1
                  
670               octet2 = bytes(count)
                  
680               If (octet2 And &H80) = 0 Or (octet2 And &H40) <> 0 Then
690                   result = result & Chr$(octet1) & Chr$(octet2)
700                   GoTo loopend
710               End If
                   
720               charCode = ((octet1 And &H1F) * &H40) + (octet2 And &H3F)
                  
730               If charCode <= 65535 Then
740                   result = result & ChrW$(charCode)
750               Else
760                   result = result & Chr$(octet1) & Chr$(octet2)
770               End If
780           End If
              
loopend:
790       Next count
          
800       UTF8Decode = result

810      On Error GoTo 0
820      Exit Function

UTF8Decode_Error:
830       handleError "UTF8Decode", Err.Number, Err.Description, Erl, utf8
End Function

