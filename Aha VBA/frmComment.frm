VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmComment 
   Caption         =   "Add a comment to a task"
   ClientHeight    =   5460
   ClientLeft      =   210
   ClientTop       =   795
   ClientWidth     =   15015
   OleObjectBlob   =   "frmComment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_Click()

End Sub

Private Sub UserForm_Initialize()
Dim sTask As String
Dim sBody As String

sTask = GetTaskNumFromSubject()

sBody = SanitizeEmailBody(Left(GetObjectBody(), 800))

Me.txtTask = sTask
Me.txtComment = sBody
    

End Sub

Private Sub cmdComment_Click()


Result = AddCommentToFeature(Me.txtTask, ReplaceCarriageReturns(Me.txtComment))
If Result = "OK" Then
    MsgBox "OK"
    Unload Me
Else
    MsgBox "Failed: " & Result
End If

End Sub




