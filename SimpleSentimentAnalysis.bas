Attribute VB_Name = "SimpleSentimentAnalysis"
Sub SimpleSentimentAnalysis()
    'Get the text to analyze from cell A1
    Dim text As String
    text = Range("A1").Value
    
    'Create a dictionary of positive and negative words
    Dim positiveWords As Variant
    positiveWords = Array("good", "great", "excellent", "wonderful", "fantastic")
    
    Dim negativeWords As Variant
    negativeWords = Array("bad", "poor", "terrible", "awful", "disappointing")
    
    'Tokenize the text and count the number of positive and negative words
    Dim tokens As Variant
    tokens = Split(LCase(text), " ")
    
    Dim positiveCount As Integer
    positiveCount = 0
    
    Dim negativeCount As Integer
    negativeCount = 0
    
    Dim i As Integer
    For i = LBound(tokens) To UBound(tokens)
        If UBound(Filter(positiveWords, tokens(i))) > -1 Then
            positiveCount = positiveCount + 1
        End If
        
        If UBound(Filter(negativeWords, tokens(i))) > -1 Then
            negativeCount = negativeCount + 1
        End If
    Next i
    
    'Calculate the sentiment score as the difference between positive and negative counts
    Dim sentimentScore As Integer
    sentimentScore = positiveCount - negativeCount
    
    'Display the sentiment score in cell B1
    Range("B1").Value = sentimentScore
End Sub

