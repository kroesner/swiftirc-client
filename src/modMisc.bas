Attribute VB_Name = "modMisc"
Option Explicit

Public Function stripFormattingCodes(ByRef text As String) As String
          Dim wChar As Long
          Dim count As Long
          Dim output As String
          
10        For count = 1 To Len(text)
20            wChar = AscW(Mid$(text, count, 1))
              
30            Select Case wChar
                  Case 2
40                Case 4
50                Case 15
60                Case 22
70                Case 31
80                Case 3
90                    If count = Len(text) Then
100                       Exit For
110                   End If

120                   count = findColourCodeEnd(text, count + 1) - 1
130               Case Else
140                   output = output & ChrW$(wChar)
150           End Select
160       Next count
          
170       stripFormattingCodes = output
End Function

Private Function findColourCodeEnd(ByRef text As String, start As Integer) As Long
          Dim colourCount As Integer
          Dim digits As Byte
          Dim currentColour As Byte
          Dim hasColour As Boolean
          
          Dim wChar As Integer

10        For colourCount = start To Len(text)
20            wChar = AscW(Mid$(text, colourCount, 1))
              
30            If wChar > 47 And wChar < 58 Then
40                If digits = 0 Then
50                    hasColour = True
60                    digits = 1
70                    start = start + 1
80                ElseIf digits = 1 Then
90                    digits = 2
100                   start = start + 1
110               Else
120                   Exit For
130               End If
140           ElseIf wChar = AscW(",") Then
150               If Not hasColour Then
160                   Exit For
170               End If
                  
180               hasColour = False
190               digits = 0
200               start = start + 1
210           Else
220               Exit For
230           End If
240       Next colourCount
          
250       findColourCodeEnd = start
End Function

'Case insensitive wildcard matching
Public Function swiftMatch(ByVal pattern As String, ByVal text As String) As Boolean
          Dim char As Long
          Dim tempText As String
          Dim tempText2 As String
          
10        If LenB(pattern) = 0 Then
20            Exit Function
30        End If
          
40        pattern = LCase$(pattern)
50        tempText = LCase$(text)
          
60        Do While LenB(pattern) <> 0
70            char = AscW(left$(pattern, 1))
80            pattern = Mid$(pattern, 2)
              
90            If char = 92 Then '\
100               If left$(tempText, 1) <> left$(pattern, 1) Then
110                   Exit Function
120               End If
                  
130               pattern = Mid$(pattern, 2)
140               tempText = Mid$(tempText, 2)
150           ElseIf char = 63 Then '?
160               If LenB(tempText) = 0 Then
170                   Exit Function
180               End If
                  
190               tempText = Mid$(tempText, 2)
200           ElseIf char = 42 Then '*
210               If LenB(pattern) = 0 Then
220                   swiftMatch = True
230                   Exit Function
240               End If
                  
250               tempText2 = tempText
                  
260               Do While LenB(tempText2) <> 0
270                   If left$(tempText2, 1) = left$(pattern, 1) Then
280                       If swiftMatch(pattern, tempText2) Then
290                           swiftMatch = True
300                           Exit Function
310                       End If
320                   End If
                      
330                   tempText2 = Mid$(tempText2, 2)
340               Loop
350           Else
360               If LenB(tempText) = 0 Then
370                   Exit Function
380               End If
                  
390               If AscW(left$(tempText, 1)) <> char Then
400                   Exit Function
410               End If
                  
420               tempText = Mid$(tempText, 2)
430           End If
440       Loop
          
450       If LenB(tempText) = 0 Then
460           swiftMatch = True
470       End If
End Function

Public Sub saveIgnoreFile()
10        ignoreManager.saveIgnoreList g_userPath & "swiftirc_ignore_list.xml"
End Sub

Public Function sanitizeFilename(filename As String) As String
          Dim count As Long
          Dim char As String
          Dim output As String
          
          Const DISALLOWED_FILENAME_CHARACTERS = "\/:*?""<>|"
          Const DISALLOWED_CHARACTER_SUBSTITUTE = "_"
          
10        For count = 1 To Len(filename)
20            char = Mid$(filename, count, 1)
              
30            If InStr(DISALLOWED_FILENAME_CHARACTERS, char) <> 0 Then
40                output = output & DISALLOWED_CHARACTER_SUBSTITUTE
50            Else
60                output = output & char
70            End If
80        Next count
          
90        sanitizeFilename = output
End Function

Public Function swiftFormatTime(timestamp As Date, theFormat As String) As String
          Dim count As Long
          Dim char As String
          Dim nextChar As String
          Dim formatStr As String
          
10        For count = 1 To Len(theFormat)
20            char = Mid$(theFormat, count, 1)
              
30            If count < Len(theFormat) Then
40                nextChar = LCase$(Mid$(theFormat, count + 1, 1))
50            Else
60                nextChar = vbNullString
70            End If
              
80            If char = "h" Or char = "H" Then
90                formatStr = "h"
                  
100               If nextChar = "h" Then
110                   formatStr = formatStr & "h"
120                   count = count + 1
130               End If
                  
140               swiftFormatTime = swiftFormatTime & format(timestamp, formatStr)
150           ElseIf char = "m" Or char = "M" Then
160               If nextChar = "m" Then
170                   count = count + 1
180               End If
                  
190               swiftFormatTime = swiftFormatTime & format(timestamp, "nn")
200           ElseIf char = "s" Or char = "S" Then
210               If nextChar = "s" Then
220                   count = count + 1
230               End If
                  
240               swiftFormatTime = swiftFormatTime & format(timestamp, "ss")
250           Else
260               swiftFormatTime = swiftFormatTime & char
270           End If
280       Next count
End Function

Public Function encrypt(pass As String, text As String) As String
          Dim aes As New cRijndael
          Dim plain() As Byte
          Dim crypt() As Byte
          Dim realPass() As Byte
          
10        realPass = StrConv(pass, vbFromUnicode)
20        ReDim Preserve realPass(31)
          
30        aes.SetCipherKey realPass, 256
          
40        plain = StrConv(text, vbFromUnicode)
          
50        aes.ArrayEncrypt plain, crypt, 0

60        encrypt = HexDisplay(crypt, UBound(crypt) + 1, 16)
End Function

Public Function decrypt(pass As String, text As String) As String
          Dim aes As New cRijndael
          
          Dim crypt() As Byte
          Dim plain() As Byte
          Dim realPass() As Byte
          
10        realPass = StrConv(pass, vbFromUnicode)
20        ReDim Preserve realPass(31)
          
30        aes.SetCipherKey realPass, 256
          
40        If HexDisplayRev(text, crypt) = 0 Then
50            decrypt = text
60            Exit Function
70        End If
          
80        aes.ArrayDecrypt plain, crypt, 0
90        decrypt = StrConv(plain, vbUnicode)
End Function

'Returns a String containing Hex values of data(0 ... n-1) in groups of k
Private Function HexDisplay(data() As Byte, n As Long, k As Long) As String
          Dim i As Long
          Dim j As Long
          Dim c As Long
          Dim data2() As Byte

10        If LBound(data) = 0 Then
20            ReDim data2(n * 4 - 1 + ((n - 1) \ k) * 4)
30            j = 0
40            For i = 0 To n - 1
50                If i Mod k = 0 Then
60                    If i <> 0 Then
70                        data2(j) = 32
80                        data2(j + 2) = 32
90                        j = j + 4
100                   End If
110               End If
120               c = data(i) \ 16&
130               If c < 10 Then
140                   data2(j) = c + 48     ' "0"..."9"
150               Else
160                   data2(j) = c + 55     ' "A"..."F"
170               End If
180               c = data(i) And 15&
190               If c < 10 Then
200                   data2(j + 2) = c + 48 ' "0"..."9"
210               Else
220                   data2(j + 2) = c + 55 ' "A"..."F"
230               End If
240               j = j + 4
250           Next i
260   Debug.Assert j = UBound(data2) + 1
270           HexDisplay = data2
280       End If

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

10        n = 2 * Len(TheString)
20        data2 = TheString

30        ReDim data(n \ 4 - 1)

40        d = 0
50        i = 0
60        j = 0
70        Do While j < n
80            c = data2(j)
90            Select Case c
              Case 48 To 57    '"0" ... "9"
100               If d = 0 Then   'high
110                   d = c
120               Else            'low
130                   data(i) = (c - 48) Or ((d - 48) * 16&)
140                   i = i + 1
150                   d = 0
160               End If
170           Case 65 To 70   '"A" ... "F"
180               If d = 0 Then   'high
190                   d = c - 7
200               Else            'low
210                   data(i) = (c - 55) Or ((d - 48) * 16&)
220                   i = i + 1
230                   d = 0
240               End If
250           Case 97 To 102  '"a" ... "f"
260               If d = 0 Then   'high
270                   d = c - 39
280               Else            'low
290                   data(i) = (c - 87) Or ((d - 48) * 16&)
300                   i = i + 1
310                   d = 0
320               End If
330           End Select
340           j = j + 2
350       Loop
360       n = i
370       If n = 0 Then
380           Erase data
390       Else
400           ReDim Preserve data(n - 1)
410       End If
420       HexDisplayRev = n

End Function
