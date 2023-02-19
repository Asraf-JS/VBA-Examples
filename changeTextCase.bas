Attribute VB_Name = "changeTextCase"
Sub changeTextCase()
    Dim inputText As String
    Dim outputText As String
    
    'Get the input text from cell A1
    inputText = Range("A1").Value
    
    'Check if the input text is in all caps
    If inputText = UCase(inputText) Then
        'If the input text is in all caps, convert it to all lowercase
        outputText = LCase(inputText)
    'Check if the input text is in title case
    ElseIf StrConv(inputText, vbProperCase) = inputText Then
        'If the input text is in title case, convert it to all lowercase
        outputText = UCase(inputText)
    Else
        'If the input text is in all lowercase or mixed case, convert it to title case
        outputText = StrConv(inputText, vbProperCase)
    End If
    
    'Write the output text to cell B1
    Range("A1").Value = outputText
End Sub

