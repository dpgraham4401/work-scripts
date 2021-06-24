Attribute VB_Name = "eManAPI1"
'-------------------------------------------------------------------
' EXCELlent haztrak
' example VBA functions to integrate with e-Manifest
'-------------------------------------------------------------------

Public Function getToken() As String
'-------------------------------------------------------------------
' GET and return RCRAInfo Auth token as string
'-------------------------------------------------------------------
    Dim authUrl, res As String
    Dim winHttpReq, resJson As Object
    
    authUrl = "https://rcrainfopreprod.epa.gov/rcrainfo/rest/api/v1/auth/" & Range("API_ID") & "/" & Range("API_Key")
    Set winHttpReq = CreateObject("msxml2.xmlhttp.6.0")
    
    winHttpReq.Open "GET", authUrl, False
    winHttpReq.Send
    res = winHttpReq.responseText
    Set resJson = ParseJSON(res)
    'Debug.Print resJson("obj.token")
    getToken = resJson("obj.token")
End Function

Public Function eManGet(endPoint As String)
'-------------------------------------------------------------------
' Basis function for RCRAInfo GET request
'-------------------------------------------------------------------
    Dim baseUrl, url, res, token As String
    Dim winHttpReq As Object
    
    baseUrl = "https://rcrainfopreprod.epa.gov/rcrainfo/rest/api/v1/"
    url = baseUrl & endPoint
    token = getToken
    token = "Bearer " & token
    
    Set winHttpReq = CreateObject("msxml2.xmlhttp.6.0")
    winHttpReq.Open "GET", url, False
    winHttpReq.setRequestHeader "Accept", "application/json"
    winHttpReq.setRequestHeader "Authorization", token
    winHttpReq.Send
    res = winHttpReq.responseText
    
    'Debug.Print res
End Function

Sub testGet()
'
' Testing area
'
    Dim testVar As String
    testVar = "site-exists/" & Range("Site_ID")
    eManGet (testVar)
End Sub
