Option Explicit

Private Const BIB_NS As String = _
    "http://schemas.microsoft.com/office/word/2004/10/bibliography"

Public Sub AddSourcesToCurrentDocument()
    AddSource "StatSAQLFS2026", "DocumentFromInternetSite", _
        "Statistics South Africa", _
        "Quarterly Labour Force Survey (QLFS), Quarter 2: 2026", _
        "Statistics South Africa", _
        "2026", "8", "11", _
        "https://www.statssa.gov.za/publications/P0211/P02112ndQuarter2026.pdf", _
        "2026", "8", "20"

    AddSource "UberDeliverSA", "InternetSite", _
        "Uber Technologies Inc.", _
        "Deliver with Uber Eats - South Africa", _
        "Uber", _
        "", "", "", _
        "https://www.uber.com/za/en/deliver/", _
        "2026", "8", "20"

    AddSource "MrDDriverPartner", "InternetSite", _
        "Takealot Group / Mr D", _
        "Mr D - Become a Driver Partner", _
        "Takealot Group / Mr D", _
        "", "", "", _
        "https://takealotgroup.com/mr-d/", _
        "2026", "8", "20"

    AddSource "FairworkSA2023", "DocumentFromInternetSite", _
        "Fairwork", _
        "Fairwork South Africa Ratings 2023: Advocating for Safety and Dignity in the Platform Economy", _
        "Fairwork", _
        "2023", "", "", _
        "https://fairwork.oii.ox.ac.uk/wp-content/uploads/sites/17/2024/01/Fairwork_South_Africa_Report_2023_UP-2.pdf", _
        "2026", "8", "20"

    AddSource "ILODigitalEconomy2023", "DocumentFromInternetSite", _
        "International Labour Organization", _
        "Skills Demand and Supply in South Africa's Digital Economy: A Focus on Youth Not in Employment, Education or Training", _
        "International Labour Organization", _
        "2023", "11", "8", _
        "https://www.ilo.org/sites/default/files/wcmsp5/groups/public/%40africa/%40ro-abidjan/documents/genericdocument/wcms_901442.pdf", _
        "2026", "8", "20"

    AddSource "UberHelpSignup", "InternetSite", _
        "Uber Help", _
        "What are the steps to sign up?", _
        "Uber Help", _
        "", "", "", _
        "https://help.uber.com/driving-and-delivering/article/signup-steps?nodeId=ec52e0b4-1004-43eb-babe-6ef6ccda95e8", _
        "2026", "8", "20"

    MsgBox "Sources added. Use References -> Insert Citation.", vbInformation
End Sub

Private Sub AddSource( _
    ByVal tag As String, _
    ByVal sourceType As String, _
    ByVal author As String, _
    ByVal title As String, _
    ByVal siteTitle As String, _
    ByVal pubYear As String, _
    ByVal pubMonth As String, _
    ByVal pubDay As String, _
    ByVal url As String, _
    ByVal accessYear As String, _
    ByVal accessMonth As String, _
    ByVal accessDay As String)

    Dim src As Source
    Dim xml As String

    For Each src In ActiveDocument.Bibliography.Sources
        If StrComp(src.Tag, tag, vbTextCompare) = 0 Then Exit Sub
    Next src

    xml = "<b:Source xmlns:b=""" & BIB_NS & """>" & _
          E("Tag", tag) & _
          E("SourceType", sourceType) & _
          "<b:Author><b:Author>" & E("Corporate", author) & "</b:Author></b:Author>" & _
          E("Title", title) & _
          E("InternetSiteTitle", siteTitle)

    If pubYear <> "" Then xml = xml & E("Year", pubYear)
    If pubMonth <> "" Then xml = xml & E("Month", pubMonth)
    If pubDay <> "" Then xml = xml & E("Day", pubDay)

    xml = xml & _
          E("URL", url) & _
          E("YearAccessed", accessYear) & _
          E("MonthAccessed", accessMonth) & _
          E("DayAccessed", accessDay) & _
          "</b:Source>"

    ActiveDocument.Bibliography.Sources.Add xml
End Sub

Private Function E(ByVal name As String, ByVal value As String) As String
    E = "<b:" & name & ">" & X(value) & "</b:" & name & ">"
End Function

Private Function X(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, """", "&quot;")
    value = Replace(value, "'", "&apos;")
    X = value
End Function
