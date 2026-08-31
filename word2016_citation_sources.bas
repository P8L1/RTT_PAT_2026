Attribute VB_Name = "WordCitationSources"
Option Explicit

Public Sub AddCitationSources()
    On Error GoTo Failed

    AddOneSource "StatSAQLFS2026", "DocumentFromInternetSite", "Statistics South Africa", _
        "Quarterly Labour Force Survey (QLFS), Quarter 2: 2026", "Statistics South Africa", _
        "2026", "8", "11", "https://www.statssa.gov.za/publications/P0211/P02112ndQuarter2026.pdf", "2026", "8", "20"

    AddOneSource "UberDeliverSA", "InternetSite", "Uber Technologies Inc.", _
        "Deliver with Uber Eats - South Africa", "Uber", _
        "", "", "", "https://www.uber.com/za/en/deliver/", "2026", "8", "20"

    AddOneSource "MrDDriverPartner", "InternetSite", "Takealot Group / Mr D", _
        "Mr D - Become a Driver Partner", "Takealot Group / Mr D", _
        "", "", "", "https://takealotgroup.com/mr-d/", "2026", "8", "20"

    AddOneSource "FairworkSA2023", "DocumentFromInternetSite", "Fairwork", _
        "Fairwork South Africa Ratings 2023: Advocating for Safety and Dignity in the Platform Economy", "Fairwork", _
        "2023", "", "", "https://fairwork.oii.ox.ac.uk/wp-content/uploads/sites/17/2024/01/Fairwork_South_Africa_Report_2023_UP-2.pdf", "2026", "8", "20"

    AddOneSource "ILODigitalEconomy2023", "DocumentFromInternetSite", "International Labour Organization", _
        "Skills Demand and Supply in South Africa's Digital Economy: A Focus on Youth Not in Employment, Education or Training", "International Labour Organization", _
        "2023", "11", "8", "https://www.ilo.org/sites/default/files/wcmsp5/groups/public/%40africa/%40ro-abidjan/documents/genericdocument/wcms_901442.pdf", "2026", "8", "20"

    AddOneSource "UberHelpSignup", "InternetSite", "Uber Help", _
        "What are the steps to sign up?", "Uber Help", _
        "", "", "", "https://help.uber.com/driving-and-delivering/article/signup-steps?nodeId=ec52e0b4-1004-43eb-babe-6ef6ccda95e8", "2026", "8", "20"

    MsgBox "Done. The sources are now available under References > Insert Citation.", vbInformation
    Exit Sub

Failed:
    MsgBox "Could not add the sources." & vbCrLf & vbCrLf & _
        "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical
End Sub

Private Sub AddOneSource(ByVal tag As String, ByVal sourceType As String, ByVal corporateAuthor As String, _
    ByVal title As String, ByVal siteTitle As String, ByVal pubYear As String, _
    ByVal pubMonth As String, ByVal pubDay As String, ByVal sourceUrl As String, _
    ByVal accessYear As String, ByVal accessMonth As String, ByVal accessDay As String)

    Dim xml As String

    If SourceAlreadyExists(tag) Then Exit Sub

    xml = "<b:Source xmlns:b=""http://schemas.microsoft.com/office/word/2004/10/bibliography"">"
    xml = xml & XmlTag("Tag", tag)
    xml = xml & XmlTag("SourceType", sourceType)
    xml = xml & "<b:Author><b:Author>" & XmlTag("Corporate", corporateAuthor) & "</b:Author></b:Author>"
    xml = xml & XmlTag("Title", title)
    xml = xml & XmlTag("InternetSiteTitle", siteTitle)
    xml = xml & XmlTag("Year", pubYear)
    xml = xml & XmlTag("Month", pubMonth)
    xml = xml & XmlTag("Day", pubDay)
    xml = xml & XmlTag("YearAccessed", accessYear)
    xml = xml & XmlTag("MonthAccessed", accessMonth)
    xml = xml & XmlTag("DayAccessed", accessDay)
    xml = xml & XmlTag("URL", sourceUrl)
    xml = xml & "</b:Source>"

    ActiveDocument.Bibliography.Sources.Add xml
End Sub

Private Function SourceAlreadyExists(ByVal tag As String) As Boolean
    Dim src As Object

    For Each src In ActiveDocument.Bibliography.Sources
        If StrComp(CStr(src.Tag), tag, vbTextCompare) = 0 Then
            SourceAlreadyExists = True
            Exit Function
        End If
    Next src
End Function

Private Function XmlTag(ByVal tagName As String, ByVal value As String) As String
    XmlTag = "<b:" & tagName & ">" & XmlEscape(value) & "</b:" & tagName & ">"
End Function

Private Function XmlEscape(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, Chr$(34), "&quot;")
    value = Replace(value, "'", "&apos;")
    XmlEscape = value
End Function
