Private Sub CommandButton4_Click()
    Dim columns(26) As String
    Dim final As String
    Dim i As Integer
    
    For i = 0 To 25
        columns(i) = setColumn(i)
    Next i
    
    Dim linhaStart As Integer
    Dim linhaEnd As Integer
    
    Dim colunaStart As String
    Dim colunaEnd As String
    
    If IsNumeric(Range("E15")) = True Or IsNumeric(Range("E16")) = True Then
        ActiveSheet.Range("E18").Value = "Inválido"
    Else
        colunaStart = ActiveSheet.Range("E15").Value
        colunaEnd = ActiveSheet.Range("E16").Value
    End If
    If IsNumeric(Range("C15")) = False Or IsNumeric(Range("C16")) = False Then
        ActiveSheet.Range("B18").Value = "Inválido"
    Else
        linhaStart = ActiveSheet.Range("C15").Value
        linhaEnd = ActiveSheet.Range("C16").Value
        If linhaStart < linhaEnd Then
            For i = 1 To 25
                ActiveSheet.Range("A" & i).Value = ""
            Next i

            Dim posInit As Integer
            Dim posFinale As Integer

            posInit = -1
            posFinale = -1
            
            For i = 0 To 25
                If columns(i) = colunaStart Then
                    posInit = i
                End If
                If columns(i) = colunaEnd Then
                    posFinale = i
                End If
            Next i
            
            If posFinale = -1 Or posInit = -1 Then
                ActiveSheet.Range("E18").Value = "Inválido"
            Else
                i = 0
                For linhaStart = linhaStart To linhaEnd
                    i = i + 1
                    final = "("
                    Dim j As Integer
                    For j = posInit To posFinale
                        If IsNumeric(Range(columns(j) & linhaStart)) = True Then
                            final = final & Range(columns(j) & linhaStart).Value
                        Else
                            final = final & "'" & Range(columns(j) & linhaStart).Value & "'"
                        End If
                        If j = posFinale Then
                            final = final & ")"
                            ActiveSheet.Range("A" & i).Value = final
                        Else
                            final = final & ","
                        End If
                    Next j
                Next linhaStart
            End If
        End If
    End If
End Sub

Function setColumn(i As Integer) As String
    Dim columns(26) As String
    
    columns(0) = "A"
    columns(1) = "B"
    columns(2) = "C"
    columns(3) = "D"
    columns(4) = "E"
    columns(5) = "F"
    columns(6) = "G"
    columns(7) = "H"
    columns(8) = "I"
    columns(9) = "J"
    columns(10) = "K"
    columns(11) = "L"
    columns(12) = "M"
    columns(13) = "N"
    columns(14) = "O"
    columns(15) = "P"
    columns(16) = "Q"
    columns(17) = "R"
    columns(18) = "S"
    columns(19) = "T"
    columns(20) = "U"
    columns(21) = "V"
    columns(22) = "W"
    columns(23) = "X"
    columns(24) = "Y"
    columns(25) = "Z"
    setColumn = columns(i)
End Function