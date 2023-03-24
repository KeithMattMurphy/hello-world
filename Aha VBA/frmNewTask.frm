VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewTask 
   Caption         =   "Create a new task"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11625
   OleObjectBlob   =   "frmNewTask.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNewTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbo_Assigned_Change()
Dim Result$
Me.cboSetOwner.value = Me.cbo_Assigned.value
Result = GetEpics(Me.cbo_Assigned.value, cbo_ProjectStatus.value)
End Sub

Private Sub cbo_ProjectStatus_Change()
Dim Result$
Result = GetEpics(Me.cbo_Assigned.value, cbo_ProjectStatus.value)
End Sub



Private Sub cmd_CreateTask_Click() ' also creates event
Dim sSubject As String
Dim sBody As String
Dim Result As String

'DTPicker1 = DateValue(DateAdd("d", 7, Now()))

        If Me.cbo_Assigned.value <> "" Then
            If Me.lst_Projects.value <> "" Then
                projnum = Left(Me.lst_Projects.value, InStr(1, Me.lst_Projects.value, "|") - 1)
                If Left(projnum, 10) = "ETPROJECTS" Then
                    sBody = ReplaceCarriageReturns(Me.txtBody)
                    If Me.Caption = "Create a new event" Then
                        
                        'SendHTTPEvent returns the ref. number of the new event
                        
                        Result = SendHTTPEvent(Me.txtSubject, sBody, projnum, Me.cbo_Assigned.value, Me.DTPicker1.value, Me.cboSetOwner)
                        If Left(Result, 5) = "Error" Then
                            MsgBox "Error: " & Result
                        Else
                        'new event number has to be linked to the proj/task it's created for
                            Result = UpdateCustomObjectLinks(projnum, Result)
                        End If
                        
                        
                        Unload Me
                    Else
                        If InStr(1, GetObjectSubject, "ETPROJECTS") = 0 Then ' if no already assigned
                            Result = SendHTTPPost(Me.txtSubject, sBody, projnum, Me.cboSetOwner)
                        Else
                            MsgBox "It looks like this email is already assigned to a Task - check the subject line", vbCritical
                        End If
                    End If
                    If Left(Result, 2) = "OK" Then
                        Unload Me
                    Else
                        MsgBox "Error: " & Result
                    End If
                Else
                    MsgBox "Invalid Project selected"
                End If
            Else
                MsgBox "Please select a project"
            End If
        Else
            MsgBox "Please select an assignee"
        End If


End Sub

Private Sub cmd_refresh_Click()
RefreshEpics
RefreshReleases

Dim Result$
Result = GetEpics(Me.cbo_Assigned.value, cbo_ProjectStatus.value)

End Sub




Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Frame2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim searchStr$, Result$
searchStr = InputBox("Enter search criteria")
Result = SearchEpics(searchStr)
MsgBox Result
End Sub

Private Sub lst_Projects_Click()


End Sub

Private Sub lst_Projects_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim Command$, sURL$
If IsNull(Me.lst_Projects.value) = False Then
    sURL = "https://optum.aha.io/epics/" & Left(Me.lst_Projects.value, InStr(1, Me.lst_Projects.value, "|") - 1)
    Command = "cmd /c start " & sURL
    Shell Command, vbHide
End If
End Sub

Private Sub txtAttendees_Change()

End Sub

Private Sub txtBody_Change()

End Sub

Private Sub UserForm_Initialize()
Dim CurrentUser As String
CurrentUser = GetUserFNameLName()

With cbo_Assigned
    .AddItem "Akshay Goyal"
    .AddItem "Angelica Cox"
    .AddItem "Barry Lavin"
    .AddItem "Ioana Digennaro"
    .AddItem "Jennifer Poos"
    .AddItem "Jennifer Vercauteren"
    .AddItem "Jogil Jose"
    .AddItem "Keith Murphy"
    .AddItem "Lindsay Schagrin"
    .AddItem "Rachael Lowe"
    .AddItem "Seth Riley"
    .AddItem "Stephen Quinn"

    .AddItem CurrentUser
    .value = CurrentUser
End With

With cboSetOwner
    .AddItem "Akshay Goyal"
    .AddItem "Angelica Cox"
    .AddItem "Barry Lavin"
    .AddItem "Ioana Digennaro"
    .AddItem "Jennifer Poos"
    .AddItem "Jennifer Vercauteren"
    .AddItem "Jogil Jose"
    .AddItem "Keith Murphy"
    .AddItem "Lindsay Schagrin"
    .AddItem "Rachael Lowe"
    .AddItem "Seth Riley"
    .AddItem "Stephen Quinn"

    .AddItem CurrentUser
    .value = CurrentUser
End With


With Me.cbo_ProjectStatus
    .AddItem "Not Started"
    .AddItem "Next"
    .AddItem "Recurring"
    .AddItem "In progress"
    .AddItem "Monitoring"
    '.AddItem "Complete"
    '.AddItem "On hold"
    '.AddItem "Cancelled"
    '.AddItem "Archive"
    .value = "In progress"
End With





    Dim cEmails As New Collection
    
    sAttendees = "Attendees: " & ListAttendees(cEmails)
    If sAttendees <> "Attendees: " Then
        For Each e In cEmails
            Me.txtAttendees = """" & e & """" & ", " & Me.txtAttendees
        Next e
        Me.txtBody = sAttendees & Left(GetObjectBody(), 500)
    Else
        Me.txtBody = Left(GetObjectBody(), 500)
    End If
    
Me.txtSubject = GetObjectSubject()



End Sub


Function SendHTTPPost(inSubject As String, inBody As String, ByVal inProjNum As String, ByVal inAssignOwner As String) As String
    Dim oHTTP As Object
    Dim sURL As String

    Dim sBody As String
    Dim workspace As String
    Dim UserFullName As String
    Dim epic_reference_num As String
    Dim ReleaseNum As String
    
    SendHTTPPost = "Error"
    'epic_reference_num = "ETPROJECTS-E-972"
    epic_reference_num = inProjNum
    
    'ReleaseNum = "ETPROJECTS-R-70"
    ReleaseNum = GetRelease(Me.DTPicker1.value)
    If ReleaseNum = "Error" Then
        ReleaseNum = "CRDROLL-R-14"
    End If
    ReleaseNum = InputBox("Confirm add to release " & ReleaseNum, "Confirm Release to add to ", ReleaseNum)
    
    
    

    'UserFullName = GetUserFNameLName()
    UserFullName = Me.cbo_Assigned.value
    
    Dim enteam_support_type As String
    Dim enteam_preferred_delivery_due_date As String
    Dim enteam_priority As String
    Dim enteam_expected_benefit As String
    Dim enteam_request_type As String
    Dim enteam_type As String
    Dim mandatorydiscretionary_need As String
    Dim value_stream_ets As String
    

    enteam_support_type = "Create New"
    enteam_preferred_delivery_due_date = Format(DateValue(Me.DTPicker1.value), "YYYY-MM-DD")
    enteam_priority = "Medium – Important, but non critical request"
    enteam_expected_benefit = "Meet client / enterprise requirement"
    enteam_request_type = "Reporting"
    enteam_type = "UHC"
    mandatorydiscretionary_need = "Unknown"
    value_stream_ets = "all"


    'workspace = "6925895405996373081" ' Commercial Portfolio - Prepay - Team 2
    workspace = "6649329259590338967" ' Enablement Team
    'https://optum.aha.io/api/v1/products/6649329259590338967/epics?fields=name,reference_num,workflow_status,assigned_to_user
    
'    sURL = "https://optum.aha.io/api/v1/products/" & workspace & "/features"
    sURL = "https://optum.aha.io/api/v1/releases/" & ReleaseNum & "/features"
    apiKey = environ("ETL_Aha_API_Key")
    Bearer = "Bearer " + apiKey
    
    
    Set oHTTP = CreateObject("MSXML2.XMLHTTP")
    oHTTP.Open "POST", sURL, False
    'oHTTP.Open "GET", sURL, False
    oHTTP.setRequestHeader "Content-Type", "application/json"
    oHTTP.setRequestHeader "Accept", "application/json"
    oHTTP.setRequestHeader "Authorization", Bearer

    sCustomFields = "{""enteam_support_type"" : """ & enteam_support_type & _
    """, ""enteam_preferred_delivery_due_date"":""" & enteam_preferred_delivery_due_date & _
    """, ""enteam_priority"":""" & enteam_priority & _
    """, ""enteam_expected_benefit"":""" & enteam_expected_benefit & _
    """, ""enteam_request_type"":""" & enteam_request_type & _
    """, ""enteam_type"":""" & enteam_type & _
    """, ""mandatorydiscretionary_need"":""" & mandatorydiscretionary_need & _
    """, ""value_stream_ets"":""" & value_stream_ets & _
    """}"
    
    
    sBody = "{""feature"":{""name"":""" & inSubject & """," & _
    """created_by_user"":""" & GetUserFNameLName() & """," & _
    """description"":""" & inBody & """, " & _
    """assigned_to_user"":""" & inAssignOwner & """, " & _
    """epic"":""" & epic_reference_num & """, " & _
    """custom_fields"":" & sCustomFields & "}}" ' & "," & _

    Debug.Print (sBody)
    
    oHTTP.Send sBody
    
    sresponse = oHTTP.responsetext
    Debug.Print sresponse

    tasknum = Mid(sresponse, InStr(1, sresponse, "reference_num") + 16, 15)
    If Left(tasknum, 10) = "ETPROJECTS" Then
        SendHTTPPost = "OK"
        
        If UpdateEmailSubject(tasknum) <> "OK" Then
            SendHTTPPost = "OK - task created but email Subject not updated " & sresponse
        End If
    Else
        SendHTTPPost = "Error: " & sresponse
    End If
    Debug.Print (sresponse)
    
    
End Function

Function GetCurrUserFromFile() As String

CurrentUser = GetUserFNameLName()

End Function

Function GetEpics(ByVal inUser As String, ByVal inStatus As String) As String
GetEpics = "Error"

Dim EpicsFilePath  As String
Dim fileHandle As Integer
Dim currentLine As String
Me.lst_Projects.Clear
EpicsFilePath = environ("USERPROFILE") & "\Documents\myEpics"

' Open the file for input
fileHandle = FreeFile()
Open EpicsFilePath For Input As #fileHandle

' Read the file line by line
Do While Not EOF(fileHandle)
    Line Input #fileHandle, currentLine
    pos1 = InStr(1, currentLine, "|")
    pos2 = InStr(pos1 + 1, currentLine, "|")
    pos3 = InStr(pos2, currentLine, "|")
    If pos1 > 0 Then
        If Mid(currentLine, 1, pos1 - 1) = inUser Then
            If Mid(currentLine, pos1 + 1, pos2 - pos1 - 1) = inStatus Then
                With Me.lst_Projects
                    .AddItem Mid(currentLine, pos2 + 1, 100)
                End With
                'Debug.Print currentLine
            End If
        End If
    End If
    
    
Loop



errh:
If Err.Description <> "" Then
    MsgBox Err.Description
End If
Close #fileHandle
End Function

Function SearchEpics(ByVal inSearchStr As String) As String
SearchEpics = "Error"

Dim EpicsFilePath  As String
Dim fileHandle As Integer
Dim currentLine As String
Me.lst_Projects.Clear
EpicsFilePath = environ("USERPROFILE") & "\Documents\myEpics"

' Open the file for input
fileHandle = FreeFile()
Open EpicsFilePath For Input As #fileHandle

' Read the file line by line
Do While Not EOF(fileHandle)
    Line Input #fileHandle, currentLine
    
    pos1 = InStr(1, currentLine, "|")
    pos2 = InStr(pos1 + 1, currentLine, "|")
'    pos3 = InStr(pos2, currentLine, "|")
'    If pos1 > 0 Then
'        If Mid(currentLine, 1, pos1 - 1) = inUser Then
'            If Mid(currentLine, pos1 + 1, pos2 - pos1 - 1) = inStatus Then
            If InStr(1, UCase(currentLine), UCase(inSearchStr)) > 0 Then
                With Me.lst_Projects
                    .AddItem Mid(currentLine, pos2 + 1, 100)
                End With
                'Debug.Print currentLine
            End If
'            End If
'        End If
'    End If
    
    
Loop



errh:
If Err.Description <> "" Then
    MsgBox Err.Description
Else
    SearchEpics = "OK"
End If
Close #fileHandle
End Function

Function GetRelease(ByVal inDate As Date) As String
GetReleases = "Error"

Dim EpicsFilePath  As String
Dim fileHandle As Integer
Dim currentLine As String
Me.lst_Projects.Clear
ReleasesFilePath = environ("USERPROFILE") & "\Documents\myReleases"

' Open the file for input
fileHandle = FreeFile()
Open ReleasesFilePath For Input As #fileHandle

' Read the file line by line
Do While Not EOF(fileHandle)
    Line Input #fileHandle, currentLine
    pos1 = InStr(1, currentLine, "|")
    pos2 = InStr(pos1 + 1, currentLine, "|")
    rdate = Left(currentLine, pos1 - 1)
    If Month(rdate) = Month(inDate) Then
        GetRelease = Mid(currentLine, pos1 + 1, pos2 - pos1 - 1)
        GoTo Done:
    End If
    
'    Debug.Print currentLine
    
    
Loop

Done:
GetReleases = "OK"

errh:
Close #fileHandle
End Function




