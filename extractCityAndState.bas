Attribute VB_Name = "extractCityAndState"
Sub extractCityAndState()
    Dim fullAddress As String
    Dim city As String
    Dim state As String
    Dim addressParts() As String
    Dim i As Integer
    
    'Get the full address from cell A1
    fullAddress = Range("A1").Value
    
    'Split the address into its component parts
    addressParts = Split(fullAddress, ",")
    
    'Create an array of possible state keywords
    Dim stateKeywords()
    stateKeywords = Array("Kedah", "Kelantan", "Terengganu", "Pulau Pinang", "Perlis", "Perak", "Selangor", "Negeri Sembilan", "Melaka", "Johor", "Pahang", "Sabah", "Sarawak", "Kuala Lumpur", "Labuan", "Putrajaya")
    
    'Loop through the address parts and look for the state keyword
    For i = 0 To UBound(addressParts)
        For j = 0 To UBound(stateKeywords)
            If InStr(1, addressParts(i), stateKeywords(j), vbTextCompare) > 0 Then
                'If the state keyword is found, extract the city and state
                city = Trim(addressParts(i - 1))
                state = Trim(addressParts(i))
                Exit For
            End If
        Next j
    Next i
    
    'Write the city and state to cells B1 and C1
    Range("B1").Value = city
    Range("C1").Value = state
End Sub

