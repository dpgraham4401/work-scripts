Attribute VB_Name = "eManAPI"
'-------------------------------------------------------------------
' RCRAInfo/e-Manifest API
'-------------------------------------------------------------------

Sub eManAuth()
'-------------------------------------------------------------------
' GET and return RCRAInfo Auth token as string
'-------------------------------------------------------------------
    Dim authUrl, baseUrl As String
    Dim res As String
    Dim objHTTP, resJson As Object
    baseUrl = "https://rcrainfopreprod.epa.gov/rcrainfo/rest/api/v1/"
    
    authUrl = baseUrl & "auth/" & Range("API_ID") & "/" & Range("API_Key")
    Set objHTTP = CreateObject("MSXML2.serverXMLHTTP")
    
    objHTTP.Open "GET", authUrl, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.Send
    res = objHTTP.responseText
    Set resJson = ParseJSON(res)
    'Debug.Print ListPaths(resJson)
    Range("TOKEN").Value = resJson("obj.token")
    Range("EXPIRATION").Value = resJson("obj.expiration")

End Sub

