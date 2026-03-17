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
End Sub

Public Function GetRulesStorage() As StorageItem
    Dim Session As Outlook.NameSpace
    Dim Folder As Outlook.Folder
    Dim SubFolder As Outlook.Folder
    
    Set Session = Application.Session
    For Each Folder In Session.Folders
        If Folder.Name = "kyle.brothers@louisville.edu" Then
            For Each SubFolder In Folder.Folders
                If SubFolder.Name = "Inbox" Then
                    Set GetRulesStorage = SubFolder.GetStorage("RulesStorage", olIdentifyBySubject)
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Public Function SenderHasRule(ByVal SenderAddress As String) As Boolean
    Dim myStorage As StorageItem
    Dim ParseString As String
    Dim TargetArray As Variant
    Dim SplitArray As Variant
    Dim i As Long
    
    SenderHasRule = False
    
    Set myStorage = GetRulesStorage()
    If myStorage Is Nothing Then Exit Function
    
    ParseString = myStorage.Body
    If ParseString = "" Then Exit Function
    
    TargetArray = Split(ParseString, "::")
    For i = LBound(TargetArray) To UBound(TargetArray)
        If TargetArray(i) <> "" Then
            SplitArray = Split(TargetArray(i), "|")
            If UBound(SplitArray) >= 5 Then
                If SplitArray(0) = "SENDERDELETE" Then
                    If LCase(SplitArray(1)) = LCase(SenderAddress) Then
                        SenderHasRule = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function
