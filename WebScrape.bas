Sub WebScrape()

    Dim HTMLDoc As HTMLDocument
    Dim oBrowser As InternetExplorer
    Dim oHTML_Element As IHTMLElement
    
    Dim i As Single
    Dim data As Variant
    ReDim data(0 To 10000)
    Dim CN As String
    CN = "text-muted"
    
    Set oBrowser = New InternetExplorer
    oBrowser.Silent = False
    
    oBrowser.navigate "https://allform.tech/"
    oBrowser.Visible = False
    
    Do
    
    Loop Until oBrowser.readyState = READYSTATE_COMPLETE
    Set HTMLDoc = oBrowser.document
    
    Application.Wait DateAdd("s", 3, Now)
    
    i = 0
    
    For Each oHTML_Element In HTMLDoc.getElementsByClassName(CN)
        'If oHTML_Element.className = "text-muted mb-0" Then
            data(i) = oHTML_Element.innerHTML
            i = i + 1
        'End If
    Next
    
    ReDim Preserve data(0 To i - 1)

End Sub
