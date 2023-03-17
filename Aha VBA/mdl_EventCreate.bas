Attribute VB_Name = "mdl_EventCreate"
Option Explicit


Function SendHTTPEvent(inSubject As String, inBody As String, ByVal inReferenceNum As String, ByVal UserFullName As String, ByVal inEventDate As String) As String
'create a new event on a project(epic)
'return the id of that event so it can be linked then to the project

    Dim oHTTP As Object
    Dim sURL As String

    Dim sBody As String
    Dim workspace As String
    'Dim UserFullName As String
    
    Dim EventObject As Object
    Dim cleanedJson$, apiKey$, Bearer$
    Dim sCustomFields$, sresponse$, recordType
    
    SendHTTPEvent = "Error"
    If Left(inReferenceNum, 10) = "ETPROJECTS" Then
        If Left(inReferenceNum, 12) = "ETPROJECTS-E" Then
            recordType = "epic"
        Else
            recordType = "feature"
        End If
    Else
        SendHTTPEvent = "Invalid Project or Task Number provided"
        GoTo errH:
    End If
    
    
'    UserFullName = Me.cbo_Assigned.value
    workspace = "6649329259590338967" ' Enablement Team
    
    sURL = "https://optum.aha.io/api/v1/products/" & workspace & "/custom_objects/et_events/records"
    apiKey = environ("ETL_Aha_API_Key")
    Bearer = "Bearer " + apiKey
    
    
    Set oHTTP = CreateObject("MSXML2.XMLHTTP")
    oHTTP.Open "POST", sURL, False
    'oHTTP.Open "GET", sURL, False
    oHTTP.setRequestHeader "Content-Type", "application/json"
    oHTTP.setRequestHeader "Accept", "application/json"
    oHTTP.setRequestHeader "Authorization", Bearer

    sCustomFields = "{""event"" : """ & inSubject & _
    """, ""et_events_assigned_to"":""" & UserFullName & _
    """, ""event_date"":""" & Format(DateValue(inEventDate), "YYYY-MM-DD") & _
    """, ""et_events_notes"":""" & inBody & _
    """}"
    
    sBody = "{""custom_object_record"":{""custom_fields"":" & sCustomFields & "}}"

    Debug.Print (sBody)
    
    oHTTP.Send sBody
    
    cleanedJson = CleanJson(oHTTP.responsetext)
    Set EventObject = JsonConverter.ParseJson(cleanedJson)
    

    SendHTTPEvent = EventObject("custom_object_record")("id")
'    If Left(tasknum, 10) = "ETPROJECTS" Then
'        SendHTTPEvent = "OK"
'
'        If UpdateEmailSubject(tasknum) <> "OK" Then
'            SendHTTPEvent = "OK - task created but email Subject not updated " & sresponse
'        End If
'    Else
'        SendHTTPEvent = "Error: " & sresponse
'    End If
    'Debug.Print (sresponse)
    
errH:
End Function


Function GetRecord(ByVal apiKey As String, ByVal recordType As String, ByVal recordID As String) As Object
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    Dim cleanedJson As String
    url = "https://optum.aha.io/api/v1/" & recordType & "s/" & recordID & "?fields=custom_object_links,reference_num"
    
    With http
        .Open "GET", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .Send
        
        cleanedJson = CleanJson(.responsetext)
        Debug.Print "Retrieved " & recordType & " " & recordID & vbCrLf & cleanedJson
        Dim jsonObj As Object
        Set jsonObj = JsonConverter.ParseJson(cleanedJson)
        Set GetRecord = jsonObj
    End With
End Function

Function UpdateRecord(ByVal apiKey As String, ByVal recordType As String, ByVal recordID As String, ByVal updatedData As String) As Object
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    Dim cleanedJson As String
    Dim jsonObj As Object
    url = "https://optum.aha.io/api/v1/" & recordType & "s/" & recordID
 
    Debug.Print url & vbCrLf & updatedData
    
    
    With http
        .Open "PUT", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .Send updatedData
        
        cleanedJson = CleanJson(.responsetext)
        Set jsonObj = JsonConverter.ParseJson(cleanedJson)
        Set UpdateRecord = jsonObj
        
    End With
End Function

Function UpdateCustomObjectLinks(ByVal recordID As String, ByVal newCustomObjectLink As String) As String
'record id should be ETPROJECTS-1234 for a task(feature) or ETPROJECTS-E-123 for a project(epic)
'anything else will return a handled error
'newCustomObjectLink is the ID of the custom object(et_events) to add to the custom_object_links node of task or project
'it's important to note the HTTP verb is PUT because POST won't work - see Aha API docs

    Dim apiKey As String
    Dim recordType As String
    
    Dim myList$
    
    Dim record As Object

    Dim updatedCustomObjectLinks As String
    Dim ConfirmedCustomObjectLinks As String
    Dim updatedData As String
    Dim Result$, ReturnedLinks$

    apiKey = environ("ETL_Aha_API_Key")
    If Left(recordID, 10) = "ETPROJECTS" Then
        If Left(recordID, 12) = "ETPROJECTS-E" Then
            recordType = "epic"
        Else
            recordType = "feature"
        End If
    Else
        UpdateCustomObjectLinks = "Invalid Project or Task Number provided"
        GoTo errH:
    End If
    
    'recordID = "ETPROJECTS-3825"
    'recordType = "epic"
    'recordId = "ETPROJECTS-E-1008"
    
'    newCustomObjectLink = "7211169429270925974"
    
    
' Get the record details
    Set record = GetRecord(apiKey, recordType, recordID)
    
'Get Links
    If recordType = "feature" Then
        existingCustomObjectLinks = GetExistingLinks(record("feature"))
    Else
        existingCustomObjectLinks = GetExistingLinks(record("epic"))
    End If
    
    If existingCustomObjectLinks = "[" Then
        updatedCustomObjectLinks = existingCustomObjectLinks & """" & CStr(newCustomObjectLink) & """"
    Else
        updatedCustomObjectLinks = existingCustomObjectLinks & "," & """" & CStr(newCustomObjectLink) & """"
    End If
    updatedCustomObjectLinks = updatedCustomObjectLinks & "]"
'***************TESTING***************
'    recordType = "epic"
'    recordId = "ETPROJECTS-E-1008"
'***************TESTING***************
    sCustomFields = "{""events"" :" & updatedCustomObjectLinks & _
    "}"
    'Debug.Print sCustomFields
    
    updatedData = "{""" & recordType & """:{""custom_object_links"":" & sCustomFields & "}}"
    Debug.Print updatedData

    
' Update the record with the new custom object links
    Set record = UpdateRecord(apiKey, recordType, recordID, updatedData)
    If recordType = "feature" Then
        ConfirmedCustomObjectLinks = GetExistingLinks(record("feature"))
    Else
        ConfirmedCustomObjectLinks = GetExistingLinks(record("epic"))
    End If
    
    'check if every item in both
    Debug.Print "sent " & updatedCustomObjectLinks
    Debug.Print "received back " & ConfirmedCustomObjectLinks
errH:
    If Err.Description <> "" Then
        UpdateCustomObjectLinks = UpdateCustomObjectLinks & " | " & Err.Description
    End If
End Function

Function GetExistingLinks(ByVal inObject As Dictionary) As String
    Dim existingCustomObjectLinks As String
    On Error GoTo errH:

    existingCustomObjectLinks = "["
    'if no existing links will error out to end
    For Each customrecordlink In inObject("custom_object_links")
        If customrecordlink("key") = "events" Then
            For Each ID In customrecordlink("record_ids")
                
                If existingCustomObjectLinks <> "[" Then
                    existingCustomObjectLinks = existingCustomObjectLinks & ","
                End If
                existingCustomObjectLinks = existingCustomObjectLinks & """" & CStr(ID) & """"
            Next ID
            
'Add the new event link

            Debug.Print updatedCustomObjectLinks

        End If
    Next customrecordlink
    
    
errH:
    GetExistingLinks = existingCustomObjectLinks
End Function

' *******************  Text Functions **************

Function CleanJson(jsonStr As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "((?:https?:\/\/|mailto:)[^""\s]+)"
    End With
    
    CleanJson = regex.Replace(jsonStr, AddressOf EscapeDoubleQuotes)
End Function

Function EscapeDoubleQuotes(match As String, pos As Long, src As String) As String
    EscapeDoubleQuotes = Replace(match, """", "\""")
End Function

Dim cleanedJson As String
cleanedJson = CleanJson(jsonStr)

Dim jsonData As Variant
Set jsonData = JsonConverter.ParseJson(cleanedJson)




'Function GetObject(ByVal inObjType As String, apiKey As String, ByVal ObjectId As String) As Object
'    Dim jsonStr As String
'    Dim jsonObj As Object
'    Dim oHTTP As Object
'    Dim sURL As String
'    Dim sSubject As String
'    Dim Bearer$, Result$
'    Dim sBody As String
'
'    sURL = "https://optum.aha.io/api/v1/" & inObjType & "/" & ObjectId & "?fields=custom_object_links"
'    apiKey = environ("ETL_Aha_API_Key")
'    Bearer = "Bearer " + apiKey
'
'    Set oHTTP = CreateObject("MSXML2.XMLHTTP")
'    oHTTP.Open "GET", sURL, False
'    oHTTP.setRequestHeader "Content-Type", "application/json"
'    oHTTP.setRequestHeader "Accept", "application/json"
'    oHTTP.setRequestHeader "Authorization", Bearer
'
'    oHTTP.Send sBody
'
'    jsonStr = oHTTP.responseText
'    Set jsonObj = JsonConverter.ParseJson(jsonStr)
'    Set GetObject = jsonObj
'
'End Function
'
'Function UpdateObject(ByVal inObjType As String, ByVal apiKey As String, ByVal ObjectId As String, ByVal updatedData As String) As String
'    Dim http As New MSXML2.XMLHTTP60
'    Dim url As String
'    url = "https://myCompany.aha.io/api/v1/" & inObjType & "/" & ObjectId
'
'    With http
'        .Open "PATCH", url, False
'        .setRequestHeader "Content-Type", "application/json"
'        .setRequestHeader "Authorization", "Bearer " & apiKey
'        .Send updatedData
'        UpdateObject = JSON.Parse(.responseText)
'    End With
'End Function
'
