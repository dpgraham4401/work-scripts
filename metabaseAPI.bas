Attribute VB_Name = "metabase"
'-------------------------------------------------------------------
' RCRAquery API
'-------------------------------------------------------------------
Public Function metaSession() As String
    Dim authUrl, Body, userName, psswd As String
    Dim objHTTP, resJson As Object
    
    authUrl = "https://rcraquery.epa.gov/metabase/api/session"
    userName = InputBox("metbase username", "Metabase Authorization", "rcraquery username (email)")
    psswd = InputBox("metbase Password", "Metabase Authorization", "rcraquery password")
    Body = "{""username"": """ & userName & """, ""password"": """ & psswd & """}"
    Set objHTTP = CreateObject("MSXML2.serverXMLHTTP")
    objHTTP.Open "POST", authUrl, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.Send Body
    metaSession = objHTTP.responseText
End Function

Sub getToken()
    Dim token, res As String
    Dim resJson As Object
    Dim expDate, curDate As Date
    
    curDate = Now
    If Range("TOKEN_EXP").Value > curDate Then
        Debug.Print "Token session Good"
        Exit Sub
    Else
        Debug.Print "Token updating"
        res = metaSession()
        Set resJson = ParseJSON(res)
        token = resJson("obj.id")
        Range("META_TOKEN").Value = token
        expDate = DateAdd("d", 14, curDate)
        Range("TOKEN_EXP").Value = expDate
        Debug.Print "Token updated"
    End If
End Sub
