Attribute VB_Name = "mdl_Aha"
Sub CreateEvent()
Dim myFrm As New frmNewTask
myFrm.cmd_CreateTask.Caption = "Create Event"
myFrm.Caption = "Create a new event"
myFrm.lblDate = "Event Date"
myFrm.DTPicker1.value = GetAppointmentDate
myFrm.Show


End Sub



Sub CreateTask()

frmNewTask.Show

End Sub

Sub AddCommentToTask()
frmComment.Show
End Sub



Function GetObjectSubject() As String
    Dim objItem As Object
    Dim olMail As MailItem
    Dim objMeeting As Outlook.MeetingItem
    Dim objAppointment As Outlook.AppointmentItem
    
    Dim sSubject As String
    GetObjectSubject = ""
    
    
    Set objItem = Application.ActiveInspector.CurrentItem

    If objItem.Class = olMeetingRequest Then
        Set objMeeting = objItem
    ElseIf objItem.Class = 26 Then 'appointment
        Set objAppointment = objItem
        sSubject = objAppointment.Subject
    ElseIf objItem.Class = 43 Then
        Set olMail = objItem
        If TypeOf olMail Is Outlook.MailItem Then
            sSubject = olMail.Subject
            
        Else
            GetObjectSubject = ""
        End If
    End If

GetObjectSubject = sSubject
    
    
End Function


Function GetObjectBody(Optional ByRef AttendeeEmails As Variant) As String
    Dim objItem As Object
    Dim olMail As MailItem
    Dim objMeeting As Outlook.MeetingItem
    Dim objAppointment As Outlook.AppointmentItem
    
    Dim sBody As String
    Dim sAttendees As String
    Dim cEmails As New Collection, AttendeeEmails$
    
    sAttendees = ListAttendees(cEmails)
    For Each e In cEmails
        
        Debug.Print e
    Next e
    
    
    GetObjectBody = "Something went wrong"
    
    Set objItem = Application.ActiveInspector.CurrentItem
    If objItem.Class = olMeetingRequest Then
        Set objMeeting = objItem
        sBody = sAttendees & vbCrLf & objMeeting.Body
    ElseIf objItem.Class = 26 Then 'appointment
        Set objAppointment = objItem
        sBody = "Attendees: " & sAttendees & vbCrLf & vbCrLf & objAppointment.Body
    ElseIf objItem.Class = 43 Then
        Set olMail = objItem
        If TypeOf olMail Is Outlook.MailItem Then
            sBody = olMail.Body
            'Debug.Print "Email subject: " & sSubject
            
        Else
            GetObjectBody = ""
        End If
    End If
GetObjectBody = SanitizeEmailBody(sBody)
End Function



Sub HTTPGetProjects()
    Dim oHTTP As Object
    Dim sURL As String
    Dim sBody As String
    Dim oJSON As Object
    Dim UserFullName As String
    'workspace = "6877550836684669702" ' workspace admin 4
    workspace = "6649329259590338967" ' Enablement Team
    
    'sURL = "https://optum.aha.io/api/v1/products/" & workspace & "/epics"
    sURL = "https://optum.aha.io/api/v1/bookmarks/custom_pivots//6799654134455411518"

    apiKey = environ("ETL_Aha_API_Key")
    Bearer = "Bearer " + apiKey
    
    UserFullName = GetUserFNameLName()
    
    Set oHTTP = CreateObject("MSXML2.XMLHTTP")
    'oHTTP.Open "POST", sURL, False
    oHTTP.Open "GET", sURL, False
    oHTTP.setRequestHeader "Content-Type", "application/json"
    oHTTP.setRequestHeader "Accept", "application/json"
    oHTTP.setRequestHeader "Authorization", Bearer

    oHTTP.Send
    
    sresponse = oHTTP.responsetext

    Set oJSON = JsonConverter.ParseJson(sresponse)
    
End Sub


Function SanitizeEmailBody(sText As String) As String
    ' Replace any invalid characters with spaces
    SanitizeEmailBody = Replace(sText, Chr(34), " ") ' remove double quotes
    SanitizeEmailBody = Replace(SanitizeEmailBody, Chr(39), " ") ' remove single quotes
    SanitizeEmailBody = Replace(SanitizeEmailBody, Chr(92), " ") ' remove backslashes
    SanitizeEmailBody = Replace(SanitizeEmailBody, Chr(10), " ") ' remove line breaks
    SanitizeEmailBody = Replace(SanitizeEmailBody, Chr(13), " ") ' remove carriage returns
    SanitizeEmailBody = RemoveNonASCII(SanitizeEmailBody)
End Function



Function GetUserFNameLName() As String
Dim fName$
Dim iPos%

fName = Application.Session.CurrentUser
iPos = InStr(1, fName, ",")
fName = Mid(fName, iPos + 2, 20) & " " & Left(fName, iPos - 1)
GetUserFNameLName = fName
'
'    Dim oWMI As Object
'    Dim oUserAccounts As Object
'    Dim oUserAccount As Object
'    Dim sUserName As String
'    Dim sFirstName As String
'    Dim sLastName As String
'
'    Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
'    Set oUserAccounts = oWMI.ExecQuery("SELECT * FROM Win32_UserAccount WHERE Name='" & Environ("USERNAME") & "'")
'
'    If oUserAccounts.Count > 0 Then
'        Set oUserAccount = oUserAccounts.ItemIndex(0)
'
'        sUserName = oUserAccount.Name
'        sFirstName = oUserAccount.FullName
'        sLastName = Split(sFirstName, " ")(UBound(Split(sFirstName, " ")))
'        sFirstName = Split(sFirstName, " ")(0)
'
'        GetUserFNameLName = sLastName & " " & Replace(sFirstName, ",", "")
'
'    Else
'        GetUserFNameLName = "Unknown User"
'    End If
End Function



Function RemoveNonASCII(str As String) As String
    Dim i As Integer
    Dim newStr As String
    For i = 1 To Len(str)
        If Asc(Mid(str, i, 1)) < 256 Then
            newStr = newStr & Mid(str, i, 1)
        End If
    Next i
    RemoveNonASCII = newStr
End Function


Function UpdateEmailSubject(ByVal inTaskNum) As String
UpdateEmailSubject = "Error"

Dim olApp As Outlook.Application
Dim olMail As Outlook.MailItem
Dim strNewSubject As String
Dim FeaturePos%


' Set the new subject you want to update the email with
strNewSubject = " | https://optum.aha.io/features/" & inTaskNum

' Get the Outlook Application object
Set olApp = New Outlook.Application

' Check if there is a selected email

Set olMail = Application.ActiveInspector.CurrentItem
' Check if the selected item is a MailItem
If TypeOf olMail Is Outlook.MailItem Then
    
    'Check if already contains task number
    FeaturePos = InStr(1, olMail.Subject, "https://optum.aha.io/features/")
    If FeaturePos = 0 Then
        ' Update the subject of the email
        olMail.Subject = olMail.Subject & strNewSubject
        olMail.Save
    End If
    ' Save the changes
Else
    MsgBox "Only available on emails", vbOKOnly
End If


UpdateEmailSubject = "OK"
' Clean up
Set olMail = Nothing
Set olApp = Nothing
End Function

Function GetTaskNumFromSubject() As String
Dim Result As String
Dim feature As String
Dim FeaturePos%

GetTaskNumFromSubject = "ETPROJECTS-"

sSubject = GetObjectSubject()
FeaturePos = InStr(1, sSubject, "https://optum.aha.io/features/")

If FeaturePos > 0 Then
    GetTaskNumFromSubject = Mid(sSubject, FeaturePos + 30, 15)
Else
    GetTaskNumFromSubject = InputBox("Please enter the task to update", "No task number found in email subject", "ETPROJECTS-", vbOKCancel)

End If


End Function


Function AddCommentToFeature(ByVal inFeature As String, ByVal InComment) As String
    AddCommentToFeature = "Error"
    Dim jsonStr As String
    Dim jsonObj As Object
    Dim oHTTP As Object
    Dim sURL As String
    Dim sSubject As String
    Dim sBody As String
    
    sURL = "https://optum.aha.io/api/v1/features/" & inFeature & "/comments"
        
    apiKey = environ("ETL_Aha_API_Key")
    Bearer = "Bearer " + apiKey
    
    
    Set oHTTP = CreateObject("MSXML2.XMLHTTP")
    oHTTP.Open "POST", sURL, False
    
    oHTTP.setRequestHeader "Content-Type", "application/json"
    oHTTP.setRequestHeader "Accept", "application/json"
    oHTTP.setRequestHeader "Authorization", Bearer
    
    sBody = "{""comment"":{""body"":""" & InComment & """}}"
    
        
    'sBody = "{""name"":""" & sSubject & """," & _
    '"""description"":""" & sBody & """, " & _
    '"""assigned_to_user"":""" & UserFullName & """, " & _
    '"""epic_reference_num"":""" & epic_reference_num & """, " & _
    '"""custom_fields"":" & sCustomFields & "," & _
    '"""release"":" & sRelease & "," & _
    '"""epic"":" & sMasterFeature & "}"
    Debug.Print (sBody)
    
    oHTTP.Send sBody
    
    sresponse = oHTTP.responsetext
    
    If InStr(1, Left(sresponse, 20), "error") > 0 Then
        AddCommentToFeature = "Error: " & sresponse
    Else
        If UpdateEmailSubject(inFeature) = "OK" Then
            AddCommentToFeature = "OK"
        Else
            AddCommentToFeature = "Comment not added " & sresponse
        End If
    End If
    Debug.Print (sresponse)
End Function


Function WriteToFile(ByVal inFilePath As String, ByVal inValToWrite As String) As String
WriteToFile = "Error"

' Create a new FileSystemObject
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")

' Create a new TextStream for writing
Dim ts As Object
Set ts = fs.CreateTextFile("C:\myfolder\myfile.txt", True)

' Write a string to the file
ts.WriteLine "Hello, world!"

' Close the TextStream
ts.Close

End Function

Sub TEST()


End Sub

Sub RefreshEpics()
Dim oHTTP As Object
Dim jsonStr As String
Dim jsonObj As Object
Dim epic As Object
Dim refNum As String
Dim EpicsFilePath As String
Dim Result As String
Dim pageNum%, total_pages%
Result = "Error"
EpicsFilePath = environ("USERPROFILE") & "\Documents\myEpics"

Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")

Dim ts As Object
Set ts = fs.CreateTextFile(EpicsFilePath, True)
On Error GoTo errH:

apiKey = environ("ETL_Aha_API_Key")
Bearer = "Bearer " + apiKey

Set oHTTP = CreateObject("MSXML2.XMLHTTP")

pageNum = 0
total_pages = 1

While pageNum < total_pages
    pageNum = pageNum + 1
    
    sURL = "https://optum.aha.io/api/v1/products/6649329259590338967/epics?per_page=200&page=" & pageNum & "&fields=name,reference_num,workflow_status,assigned_to_user"
    oHTTP.Open "GET", sURL, False
    oHTTP.setRequestHeader "Content-Type", "application/json"
    oHTTP.setRequestHeader "Accept", "application/json"
    oHTTP.setRequestHeader "Authorization", Bearer
    
    oHTTP.Send
    
    jsonStr = oHTTP.responsetext
    
        
    Dim myCollection As New Collection
    
    Set jsonObj = JsonConverter.ParseJson(jsonStr)
    total_pages = jsonObj("pagination")("total_pages")
    
    For Each epic In jsonObj("epics")
        'If epic("assigned_to_user")("name") = inName Then
        If epic("workflow_status")("name") <> "Cancelled" And epic("workflow_status")("name") <> "Archive" And epic("workflow_status")("name") <> "On hold" Then
            refNum = epic("assigned_to_user")("name") & "|" & epic("workflow_status")("name") & "|" & epic("reference_num") & "|" & epic("name")
            ts.WriteLine refNum
    
            
            'Me.ListBox1.AddItem (epic("reference_num") & vbTab & epic("workflow_status")("name") & vbTab & epic("name"))
            'Debug.Print refNum
        End If
    Next epic
Wend


Result = "Epics OK"
errH:
    ts.Close
    MsgBox Result
End Sub



Sub RefreshReleases()

Dim oHTTP As Object
Dim jsonStr As String
Dim jsonObj As Object
Dim epic As Object
Dim refNum As String
Dim ReleasesFilePath As String
Dim Result As String
Dim pageNum%, total_pages%
Result = "Error"
ReleasesFilePath = environ("USERPROFILE") & "\Documents\myReleases"

Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")

Dim ts As Object
Set ts = fs.CreateTextFile(ReleasesFilePath, True)
On Error GoTo errH:

apiKey = environ("ETL_Aha_API_Key")
Bearer = "Bearer " + apiKey

Set oHTTP = CreateObject("MSXML2.XMLHTTP")

pageNum = 0
total_pages = 1

While pageNum < total_pages
    pageNum = pageNum + 1
    
    sURL = "https://optum.aha.io/api/v1/products/6649329259590338967/releases?per_page=200&page=" & pageNum
    oHTTP.Open "GET", sURL, False
    oHTTP.setRequestHeader "Content-Type", "application/json"
    oHTTP.setRequestHeader "Accept", "application/json"
    oHTTP.setRequestHeader "Authorization", Bearer
    
    oHTTP.Send
    
    jsonStr = oHTTP.responsetext
    
        
    Dim myCollection As New Collection
    
    Set jsonObj = JsonConverter.ParseJson(jsonStr)
    total_pages = jsonObj("pagination")("total_pages")
    
    For Each r In jsonObj("releases")
        'If epic("assigned_to_user")("name") = inName Then
        If Year(r("release_date")) = Year(Now()) Then
            refNum = r("release_date") & "|" & r("reference_num") & "|" & r("name")
            ts.WriteLine refNum
    
            
            'Me.ListBox1.AddItem (epic("reference_num") & vbTab & epic("workflow_status")("name") & vbTab & epic("name"))
            Debug.Print refNum
        End If
    Next r
Wend


Result = "Releases OK"
errH:
    ts.Close
    MsgBox Result
End Sub
Function GetAppointmentDate() As String
    Dim objApp As Outlook.Application
    Dim objItem As Object
    Dim objMeeting As Outlook.MeetingItem
    Dim objAppointment As Outlook.AppointmentItem
    GetAppointmentDate = "Something went wrong with GetAppointmentDate"
    On Error GoTo errH:
    'Set objApp = CreateObject("Outlook.Application")
    'Set objItem = objApp.ActiveExplorer.Selection(1)
    
    Set objItem = Application.ActiveInspector.CurrentItem
    If objItem.Class = olMeetingRequest Then
        Set objMeeting = objItem
        GetAppointmentDate = objMeeting.Start
    ElseIf objItem.Class = 26 Then
        Set objAppointment = objItem
        GetAppointmentDate = objAppointment.Start
        
    End If
    
    Set objApp = Nothing
    Set objItem = Nothing
    Set objMeeting = Nothing
    Set objAppointment = Nothing
    
errH:
    If Err.Description <> "" Then
        GetAppointmentDate = GetAppointmentDate & Err.Description
    End If
End Function


Function ListAttendees(Optional ByRef inEmails As Variant) As String

    Dim objApp As Outlook.Application
    Dim objItem As Object
    Dim objMeeting As Outlook.MeetingItem
    Dim objAppointment As Outlook.AppointmentItem
    Dim strAttendees As String, strEmails As String
    Dim objAttendee As Outlook.Recipient
    ListAttendees = "Something went wrong getting attendees"
    On Error GoTo errH:
    'Set objApp = CreateObject("Outlook.Application")
    'Set objItem = objApp.ActiveExplorer.Selection(1)
    
    Set objItem = Application.ActiveInspector.CurrentItem
    If objItem.Class = olMeetingRequest Then
        Set objMeeting = objItem
        For Each objAttendee In objMeeting.Recipients
            If objAttendee.Type = olRequired Then
                strAttendees = strAttendees & objAttendee.name & "; "
            End If
        Next
        
        strAttendees = Left(strAttendees, Len(strAttendees) - 2)
    ElseIf objItem.Class = 26 Then
        Set objAppointment = objItem
        For Each objAttendee In objAppointment.Recipients
        Debug.Print objAttendee.Address
        
        If objAttendee.Type = olRequired Or objAttendee.Type = olOptional Then
            If objAttendee.MeetingResponseStatus = olResponseAccepted Then
                strAttendees = strAttendees & objAttendee.name & "; "
                If IsMissing(inEmails) <> True Then
                    inEmails.Add objAttendee.AddressEntry.GetExchangeUser.PrimarySmtpAddress & ","
                End If
            End If
        ElseIf objAttendee.Type = 0 Or objAttendee.Type = 1 Then ' organizer
            strAttendees = strAttendees & objAttendee.name & "; "
            
        End If
        Next
        If Len(strAttendees) > 0 Then
            strAttendees = Left(strAttendees, Len(strAttendees) - 2)
        End If
    End If
    
        

        'Me.txtBody = "Attendees: " & strAttendees
        'MsgBox "The required attendees are: " & strAttendees
        

    
    Set objApp = Nothing
    Set objItem = Nothing
    Set objMeeting = Nothing
    
errH:
    If Err.Description = "" Then
        ListAttendees = strAttendees
    Else
        ListAttendees = ListAttendees & Err.Description
    End If
    
End Function


Sub LinkEventToFeature()

Dim oHTTP As Object
Dim sURL As String
Dim sBody As String
Dim EventID$, FeatureNum$, sCustomFields$

FeatureNum = "ETPROJECTS-3825"
EventID = "7211169429270925974"
'EventID = "7208093333413471625"
'SendHTTPEvent = "Error"

sURL = "https://optum.aha.io/api/v1/features/" & FeatureNum
Debug.Print sURL
apiKey = environ("ETL_Aha_API_Key")
Bearer = "Bearer " + apiKey


Set oHTTP = CreateObject("MSXML2.XMLHTTP")
oHTTP.Open "PUT", sURL, False
oHTTP.setRequestHeader "Content-Type", "application/json"
oHTTP.setRequestHeader "Accept", "application/json"
oHTTP.setRequestHeader "Authorization", Bearer

sCustomFields = "{""events"" : [""" & EventID & _
"""]}"

sCustomFields = "{""events"" : [""7211040793514501966"",""7211221623170278696"",""7211169429270925974" & _
"""]}"
Debug.Print sCustomFields

sBody = "{""feature"":{""custom_object_links"":" & sCustomFields & "}}"
'sBody = "{""custom_object_links"":" & sCustomFields & "}"

Debug.Print (sBody)

oHTTP.Send sBody

sresponse = oHTTP.responsetext
Debug.Print sresponse

End Sub
