Sub WebScrape()

    Dim HTMLDoc As HTMLDocument
    Dim oBrowser As InternetExplorer
    Dim oHTML_Element As IHTMLElement
    
    Dim i As Single
    Dim data As Variant
    ReDim data(0 To 1000000)
    Dim CN As String
    CN = "text-muted"
    
    Set oBrowser = New InternetExplorer
    oBrowser.Silent = False
    
    oBrowser.navigate "https://allform.tech"
    oBrowser.Visible = False
    
    Do
    
    Loop Until oBrowser.readyState = READYSTATE_COMPLETE
    Set HTMLDoc = oBrowser.document
    
    Application.Wait DateAdd("s", 3, Now)
    
    i = 0
    
    For Each oHTML_Element In HTMLDoc.all
        'If oHTML_Element.tagName = "H3" Then
            data(i) = oHTML_Element.outerHTML
            'data(i) = oHTML_Element.t
            i = i + 1
        'End If
    Next
    
    ReDim Preserve data(0 To i - 1)

    
    
    Dim myFile As String, textToFile As String, cellValue As Variant, placeHolderString As String
    myFile = "C:\Users\krist\Desktop\test.html"
    textToFile = ""
    For i = 0 To UBound(data)
        'placeHolderString = Replace(data(i), """", "'")
        'textToFile = textToFile & placeHolderString & vbNewLine
        textToFile = textToFile & data(i) & vbNewLine
    Next i
    
    Open myFile For Output As #1
    
    Write #1, textToFile
    Close #1
End Sub
