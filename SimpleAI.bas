Attribute VB_Name = "SimpleAI"
Sub SimpleAI()
    'Set up the decision tree
    Dim tree As Variant
    tree = Array(Array("humidity", "high", "windy", "no"), _
                 Array("humidity", "high", "not windy", "yes"), _
                 Array("humidity", "normal", "windy", "no"), _
                 Array("humidity", "normal", "not windy", "yes"))
    
    'Get the input values from the user
    Dim humidity As String
    Dim windy As String
    humidity = InputBox("What is the humidity level?")
    windy = InputBox("Is it windy?")
    
    'Traverse the decision tree to make a prediction
    Dim i As Integer
    For i = LBound(tree) To UBound(tree)
        If tree(i)(0) = "humidity" And tree(i)(1) = humidity And tree(i)(3) = windy Then
            MsgBox "The prediction is: " & tree(i)(3)
            Exit Sub
        End If
    Next i
    
    'Handle the case where the input values do not match the decision tree
    MsgBox "Invalid input values"
End Sub


