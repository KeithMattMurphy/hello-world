Attribute VB_Name = "mdl_environ"
Private Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" _
    Alias "SetEnvironmentVariableA" (ByVal lpName As String, _
    ByVal lpValue As String) As Long

Public Function mySetEnvironmentVariable(name As String, value As String) As Boolean
    Dim ret As Long
    ret = SetEnvironmentVariable(name, value)
    If ret = 0 Then
        mySetEnvironmentVariable = False
    Else
        mySetEnvironmentVariable = True
    End If
End Function



