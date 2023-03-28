Attribute VB_Name = "mdl_EventCreate"
'Option Explicit


Function SendHTTPEvent(inSubject As String, inBody As String, ByVal inReferenceNum As String, ByVal UserFullName As String, ByVal inEventDate As String, ByVal inAssignOwner As String) As String
'create a new event on a project(epic)
'return the id of that event so it can be linked then to the project

    Dim oHTTP As Object
    Dim sURL As String

    Dim sBody As String
    Dim workspace As String
    'Dim UserFullName As String
    
    Dim EventObject As Object
    Dim cleanedJson$, apiKey$, Bearer$
    Dim sCustomFields$, sresponse$, recordType$
    Dim sAttendees$
    
    SendHTTPEvent = "Error"
    On Error GoTo errH:
    
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
    """, ""et_events_assigned_to"":""" & inAssignOwner & _
    """, ""event_date"":""" & Format(DateValue(inEventDate), "YYYY-MM-DD") & _
    """, ""et_events_notes"":""" & inBody & _
    """}"
    '""", ""attendees"":" & "{""email_value"":[" & inAttendees & "]}" & _

    sBody = "{""custom_object_record"":{""custom_fields"":" & sCustomFields & "}}"


    sBody = "{""custom_object_record"":{""created_by_user"":""" & GetUserFNameLName() & """," & _
    """custom_fields"":" & sCustomFields & "}}"


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
If Err.Description <> "" Then
    SendHTTPEvent = SendHTTPEvent & " | " & Err.Description

End If
End Function


Function GetRecord(ByVal apiKey As String, ByVal recordType As String, ByVal recordID As String) As Object
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    Dim cleanedJson As String
    url = "https://optum.aha.io/api/v1/" & recordType & "s/" & recordID & "?fields=custom_object_links,reference_num"
    Debug.Print url
    
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
    ' updates links to the event from project or task

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    Dim cleanedJson As String
    Dim jsonObj As Object
    'Changed url to return only links and ref num - exclude description
    url = "https://optum.aha.io/api/v1/" & recordType & "s/" & recordID & "?fields=custom_object_links,reference_num"
 
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

'Function AddEventAttendees(ByVal inEventID As String, ByVal inListOfAttendeeIDs As String) As String
Sub AddEventAttendees()
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim url As String
    Dim inEventID$, inListOfAttendeeIDs$, sBody$
    Dim cleanedJson As String
    Dim jsonObj As Object
    Dim apiKey$
    
    apiKey = environ("ETL_Aha_API_Key")
    inEventID = "7212308318086166150"
    url = "https://optum.aha.io/api/v1/custom_object_records/" & inEventID
'    inListOfAttendeeIDs = "[""7036045937940797544" & _
'                    """, ""6723137271886703697" & _
'                    """, ""6501697300510190517" & _
'                    """, ""6945094516201562459" & _
'                """]"
    
'    inListOfAttendeeIDs = "[7036045937940797544" & _
'                    ",6723137271886703697" & _
'                    ",6501697300510190517" & _
'                    ",6945094516201562459" & _
'                "]"
'
                
    inListOfAttendeeIDs = "[""7036045937940797544""]"
    
    sBody = "{""custom_object_record"":{""custom_object_links"":{""attendees"":" & inListOfAttendeeIDs & "}}}"
    Debug.Print url & vbCrLf & sBody
    
    
    With http
        .Open "PUT", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .Send sBody

        cleanedJson = CleanJson(.responsetext)
        'Set jsonObj = JsonConverter.ParseJson(cleanedJson)
        Debug.Print cleanedJson

    End With


End Sub

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
    Dim existingCustomObjectLinks As String
    Dim sCustomFields$
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
    
'now check the new linkes include the old and the new link
    If recordType = "feature" Then
        ConfirmedCustomObjectLinks = GetExistingLinks(record("feature"))
    Else
        ConfirmedCustomObjectLinks = GetExistingLinks(record("epic"))
    End If
    ConfirmedCustomObjectLinks = ConfirmedCustomObjectLinks & "]"
    'check if every item in both
    
    If CompareLists(Replace(Replace(updatedCustomObjectLinks, "[", ""), "]", ""), Replace(Replace(ConfirmedCustomObjectLinks, "[", ""), "]", "")) = "OK" Then
        If UpdateAppointmentSubject(recordID) <> "OK" Then
            MsgBox "OK - event created & linked but Appointment subject not updated "
        End If
    Else
        MsgBox "An issue has occured linking the event to your project/task. Please take a screenshot & email to your admin. updatedCustomObjectLinks:" & updatedCustomObjectLinks & " ConfirmedCustomObjectLinks: " & ConfirmedCustomObjectLinks
    End If
errH:
    If Err.Description <> "" Then
        UpdateCustomObjectLinks = UpdateCustomObjectLinks & " | " & Err.Description
    Else
        UpdateCustomObjectLinks = "OK"
    End If
End Function




Function CompareLists(strList1 As String, strList2 As String) As String
'CompareLists("123,456,789", "123,456,789,789") should return OK regardless of dupes

CompareLists = "Error"
    Dim list1() As String
    Dim list2() As String
    Dim dict1 As Object
    Dim dict2 As Object
    Dim elem As Variant
On Error GoTo errH
    ' Convert string parameters to arrays
    list1 = Split(strList1, ",")
    list2 = Split(strList2, ",")
    
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    
    ' Add all elements from list1 to dictionary 1
    For Each elem In list1
        If Not dict1.Exists(elem) Then
            dict1.Add elem, 0
        End If
    Next elem
    
    ' Add all elements from list2 to dictionary 2
    For Each elem In list2
        If Not dict2.Exists(elem) Then
            dict2.Add elem, 0
        End If
    Next elem
    
    ' Check if both dictionaries have the same keys (elements)
    If dict1.Count <> dict2.Count Then
        CompareLists = "Not OK"
        Exit Function
    End If
    
     For Each elem In dict1.Keys()
         If Not dict2.Exists(elem) Then
             CompareLists = "Not OK"
             Exit Function
         End If
        
         ' Check if the count of each element is the same in both dictionaries (lists)
         If dict1.Item(elem) <> dict2.Item(elem) Then
             CompareLists = "Not OK"
             Exit Function
         End If
     Next elem
     
     CompareLists = "OK"
     Exit Function
errH:
End Function





Function GetExistingLinks(ByVal inObject As Dictionary) As String
    Dim existingCustomObjectLinks As String
    On Error GoTo errH:
'Dim customrecordlink As Object
'Dim ID As Object

existingCustomObjectLinks = "Error"
    existingCustomObjectLinks = "["
    'if no existing links will error out to end
    'If customrecordlink <> Null Then
        For Each customrecordlink In inObject("custom_object_links")
            If customrecordlink("key") = "events" Then
                For Each ID In customrecordlink("record_ids")
                    
                    If existingCustomObjectLinks <> "[" Then
                        existingCustomObjectLinks = existingCustomObjectLinks & ","
                    End If
                    existingCustomObjectLinks = existingCustomObjectLinks & """" & CStr(ID) & """"
                Next ID
                
    'Add the new event link
    
                Debug.Print existingCustomObjectLinks
                
                '****** don't close brackets here - needs to be closed from calling funciton
                '****** after adding the new link to the list
                'existingCustomObjectLinks = existingCustomObjectLinks & "]"
    
            End If
        Next customrecordlink
    'End If
    
    
errH:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
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



Sub GetAhaData()
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    'Set up the HTTP request with necessary headers
    'httpRequest.Open "GET", "https://optum.aha.io/oauth/authorize"
    httpRequest.Open "POST", "https://optum.aha.io/oauth/token"
    'httpRequest.Open "GET", "https://optum.aha.io/features/ETPROJECTS-3829", False
'    httpRequest.setRequestHeader "Authorization", "Bearer YOUR_ACCESS_TOKEN"
    
    'Send the HTTP request and get response
    httpRequest.Send
    Dim X$
    X = httpRequest.responsetext
    Debug.Print X
    
End Sub



Sub GetAccessToken()
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    Dim X$, Y$
    
    'Set up the HTTP request with necessary headers and data
    httpRequest.Open "GET", "https://optum.aha.io/oauth/authorize"
    httpRequest.Send
    X = httpRequest.responsetext
    
    httpRequest.Open "POST", "https://optum.aha.io/oauth/token", False
    httpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    Dim postData As String
    postData = "client_id=0f9149b40054fee0afe938db1fd003536f788f6d42b5608ad4455d72251f9de5&client_secret=120f4b5c1ecc5549aae6f4ccf7ed9dbc815487d3280d815ac7ec9ebb6fdc6768&code=10dbffca01ce6985868b4cee84e0444f5bcdda104b60a13038c1d74b72e6797f&grant_type=authorization_code&redirect_uri=http://lvh.me/app_callback.html"
    
    'Send the HTTP request and get response
    httpRequest.Send postData
    Debug.Print httpRequest.responsetext
    
End Sub


Sub InitiateAuthorizationRequest()
    Dim authUrl As String
    authUrl = "https://optum.aha.io/oauth/authorize" & _
              "?client_id=0f9149b40054fee0afe938db1fd003536f788f6d42b5608ad4455d72251f9de5" & _
              "&redirect_uri=https://optum.aha.io/oauth2/callback.html" & _
              "&response_type=access_token"
    
    'Open a web browser window and navigate to the authorization URL
    Shell "cmd /c start " & authUrl
    
End Sub
Sub test()
MsgBox CheckOutlookCalendar
End Sub


Function CheckOutlookCalendar() As Double
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    Dim olItems As Outlook.Items
    Dim olAppt As Outlook.AppointmentItem
    Dim olStart As Date
    Dim olEnd As Date
    Dim TotalHours As Double
    
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderCalendar)
    Set olItems = olFolder.Items
    
    olStart = Date + 7 - Weekday(Date, vbMonday) + TimeValue("08:00:00")
    olEnd = olStart + 7 + TimeValue("17:00:00")
    
    TotalHours = 0
    
    For Each olAppt In olItems
        If olAppt.Start >= olStart And olAppt.Start <= olEnd Then
            TotalHours = TotalHours + (olAppt.Duration / 60)
        End If
    Next olAppt
    
    CheckOutlookCalendar = TotalHours
End Function


Sub FindAvailableSlots()
    Dim objItem As Object
    Dim olMail As MailItem
    
    Dim objAppointment As Outlook.AppointmentItem
        
    
    Set objItem = Application.ActiveInspector.CurrentItem
    Set objAppointment = objItem
    
    MsgBox FindAvailableMeetingTimes(objAppointment, 30, False)


    
End Sub


Public Function FindAvailableMeetingTimes(MeetingItem As Outlook.AppointmentItem, _
                                          MeetingDuration As Long, _
                                          ConsiderTentative As Boolean) As Collection
' Outlook.MeetingItem
    Dim FreeBusyResults As Collection
    Set FreeBusyResults = New Collection
    
    
    Dim i As Integer
    
    Dim Recipients As Outlook.Recipients
    Set Recipients = MeetingItem.Recipients
    
    Dim StartTime As Date
    StartTime = MeetingItem.Start
    
    Dim EndTime As Date
    EndTime = DateAdd("d", 7, StartTime)
    
    Dim TimeInterval As Date
    TimeInterval = DateAdd("n", MeetingDuration, 0)
    
    Dim CurrentTime As Date
    CurrentTime = StartTime
    
    Dim AllAttendeesFree As Boolean
    AllAttendeesFree = True
    
    Debug.Print Recipients.Count
    
    Dim FreeBusyStatus(20) As String
    i = -1
    While CurrentTime < EndTime
        For Each Recipient In Recipients
            i = i + 1
        'this returns 1 char for each period from midnight on the date provided for 1 month
        'e.g. if you pass 30 mins you get a string with around 1394 chars
        'if you pass 60 mins it's half the length
        Debug.Print Recipient.DisplayType
            FreeBusyStatus(i) = Recipient.FreeBusy(CurrentTime, MeetingDuration, True)
            
'            If (FreeBusyStatus = olFree) Or (ConsiderTentative And FreeBusyStatus = olTentative) Then
'                AllAttendeesFree = AllAttendeesFree And True
'            Else
'                AllAttendeesFree = False
'                Exit For
'            End If
        Next Recipient
        
        If AllAttendeesFree Then
            FreeBusyResults.Add CurrentTime
        End If
        
        AllAttendeesFree = True
        CurrentTime = DateAdd("n", MeetingDuration, CurrentTime)
    Wend
    
    Set FindAvailableMeetingTimes = FreeBusyResults
End Function


Sub test222()
Dim l(2) As String
l(0) = "test"




For Each Item In l
    Debug.Print Item
Next Item
End Sub



