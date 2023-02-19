Attribute VB_Name = "extractNumbers"
Sub extractNumbers()
    Dim inputString As String
    Dim outputString As String
    Dim i As Integer
    
    'Get the input string from cell A1
    inputString = Range("A1").Value
    
    'Loop through each character in the input string
    For i = 1 To Len(inputString)
        'Check if the character is a number
        If IsNumeric(Mid(inputString, i, 1)) Then
            'If the character is a number, add it to the output string
            outputString = outputString & Mid(inputString, i, 1)
        End If
    Next i
    
    'Write the output string to cell B1
    Range("B1").Value = outputString
End Sub


