Attribute VB_Name = "modMisc"
Option Explicit

Private optionsForm As frmOptions
Private optionsListeningClients As New cArrayList

Public Sub registerForOptionsUpdates(client As SwiftIrcClient)
    optionsListeningClients.Add client
End Sub

Public Sub unregisterForOptionsUpdates(client As SwiftIrcClient)
    Dim count As Long

    For count = optionsListeningClients.count To 1 Step -1
        If optionsListeningClients.item(count) Is client Then
            optionsListeningClients.Remove count
        End If
    Next count
End Sub

Public Function isOptionsFormParent(parent As SwiftIrcClient)
    If optionsForm Is Nothing Then
        Exit Function
    End If
    
    If optionsForm.parent Is parent Then
        isOptionsFormParent = True
    End If
End Function

Public Sub openOptionsDialog(parent As swiftIrc.SwiftIrcClient)
    If Not optionsForm Is Nothing Then
        optionsForm.Show vbModeless, parent
        Exit Sub
    End If
    
    Set optionsForm = New frmOptions
    optionsForm.client = parent
    optionsForm.Show vbModeless, parent
End Sub

Public Sub closeOptionsDialog()
    If optionsForm Is Nothing Then
        Exit Sub
    End If
    
    optionsForm.Hide
    Unload optionsForm
    Set optionsForm = Nothing
End Sub

Public Sub saveAllSettings()
    optionsForm.saveSettings
    serverProfiles.saveProfiles
    settings.saveSettings
    
    applyAllSettings
End Sub

Private Sub applyAllSettings()
    Dim count As Long
    Dim client As SwiftIrcClient
    
    For count = 1 To optionsListeningClients.count
        Set client = optionsListeningClients.item(count)
        
        client.coloursUpdated
        client.refreshFontSettings
        client.refreshSwitchbarSettings
    Next count
End Sub

Public Function combinePath(folderPath As String, filename As String)
    If right$(folderPath, 1) = "\" Or right$(folderPath, 1) = "/" Then
        combinePath = folderPath & filename
    Else
        combinePath = folderPath & "\" & filename
    End If
End Function

Public Function stripFormattingCodes(ByRef text As String) As String
    Dim wChar As Long
    Dim count As Long
    Dim output As String
    
    For count = 1 To Len(text)
        wChar = AscW(Mid$(text, count, 1))
        
        Select Case wChar
            Case 2
            Case 4
            Case 15
            Case 22
            Case 31
            Case 3
                If count = Len(text) Then
                    Exit For
                End If

                count = findColourCodeEnd(text, count + 1) - 1
            Case Else
                output = output & ChrW$(wChar)
        End Select
    Next count
    
    stripFormattingCodes = output
End Function

Private Function findColourCodeEnd(ByRef text As String, start As Integer) As Long
    Dim colourCount As Integer
    Dim digits As Byte
    Dim currentColour As Byte
    Dim hasColour As Boolean
    
    Dim wChar As Integer

    For colourCount = start To Len(text)
        wChar = AscW(Mid$(text, colourCount, 1))
        
        If wChar > 47 And wChar < 58 Then
            If digits = 0 Then
                hasColour = True
                digits = 1
                start = start + 1
            ElseIf digits = 1 Then
                digits = 2
                start = start + 1
            Else
                Exit For
            End If
        ElseIf wChar = AscW(",") Then
            If Not hasColour Then
                Exit For
            End If
            
            hasColour = False
            digits = 0
            start = start + 1
        Else
            Exit For
        End If
    Next colourCount
    
    findColourCodeEnd = start
End Function

'Case insensitive wildcard matching
Public Function swiftMatch(ByVal pattern As String, ByVal text As String) As Boolean
    Dim char As Long
    Dim tempText As String
    Dim tempText2 As String
    
    If LenB(pattern) = 0 Then
        Exit Function
    End If
    
    pattern = LCase$(pattern)
    tempText = LCase$(text)
    
    Do While LenB(pattern) <> 0
        char = AscW(left$(pattern, 1))
        pattern = Mid$(pattern, 2)
        
        If char = 92 Then '\
            If left$(tempText, 1) <> left$(pattern, 1) Then
                Exit Function
            End If
            
            pattern = Mid$(pattern, 2)
            tempText = Mid$(tempText, 2)
        ElseIf char = 63 Then '?
            If LenB(tempText) = 0 Then
                Exit Function
            End If
            
            tempText = Mid$(tempText, 2)
        ElseIf char = 42 Then '*
            If LenB(pattern) = 0 Then
                swiftMatch = True
                Exit Function
            End If
            
            tempText2 = tempText
            
            Do While LenB(tempText2) <> 0
                If left$(tempText2, 1) = left$(pattern, 1) Then
                    If swiftMatch(pattern, tempText2) Then
                        swiftMatch = True
                        Exit Function
                    End If
                End If
                
                tempText2 = Mid$(tempText2, 2)
            Loop
        Else
            If LenB(tempText) = 0 Then
                Exit Function
            End If
            
            If AscW(left$(tempText, 1)) <> char Then
                Exit Function
            End If
            
            tempText = Mid$(tempText, 2)
        End If
    Loop
    
    If LenB(tempText) = 0 Then
        swiftMatch = True
    End If
End Function

Public Function sanitizeFilename(filename As String) As String
    Dim count As Long
    Dim char As String
    Dim output As String
    
    Const DISALLOWED_FILENAME_CHARACTERS = "\/:*?""<>|"
    Const DISALLOWED_CHARACTER_SUBSTITUTE = "_"
    
    For count = 1 To Len(filename)
        char = Mid$(filename, count, 1)
        
        If InStr(DISALLOWED_FILENAME_CHARACTERS, char) <> 0 Then
            output = output & DISALLOWED_CHARACTER_SUBSTITUTE
        Else
            output = output & char
        End If
    Next count
    
    sanitizeFilename = output
End Function

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
