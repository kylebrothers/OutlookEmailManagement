Attribute VB_Name = "Module1"
Public CurrentFolder As String

Sub Conference1()
    AssignFolder ("1 - Conference Talks and Work Travel")
End Sub

Sub Ethics2()
    AssignFolder ("2 - Ethics")
End Sub

Sub Event3()
    AssignFolder ("3 - Event Planning and Other Service")
End Sub

Sub Grants4()
    AssignFolder ("4 - Grants and Funding")
End Sub

Sub Mentoring5()
    AssignFolder ("5 - Educational Activities")
End Sub

Sub Pediatrics6()
    AssignFolder ("6 - Pediatrics")
End Sub

Sub Personal7()
    AssignFolder ("7 - Personal")
End Sub

Sub Publications8()
    AssignFolder ("8 - Publications and Journals")
End Sub

Sub Research9()
    AssignFolder ("9 - Research Projects")
End Sub

Sub AssignFolder(ByVal StartingFolder As String)
    CurrentFolder = StartingFolder
    Project1.AssignFolderForm.LoadSubFolders
    Call Project1.AssignFolderForm.LoadStorageToColumn("FolderHistory")
    Call Project1.AssignFolderForm.SetFormColor
    Project1.AssignFolderForm.txtSelect.Text = ""
    Project1.AssignFolderForm.MatchSubjects
    Project1.AssignFolderForm.txtSelect.SetFocus
    Project1.AssignFolderForm.Show
    Project1.ThisOutlookSession.Application_Startup
    'Project1.ThisOutlookSession.HASSUpdates
End Sub

