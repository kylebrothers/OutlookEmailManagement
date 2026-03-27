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
                If SplitArray(0) = "SENDERDELETE" Or _
                   SplitArray(0) = "SENDERIMMEDIATE" Or _
                   SplitArray(0) = "SENDERFOLDER" Or _
                   SplitArray(0) = "TOPICDELETE" Or _
                   SplitArray(0) = "TOPICIMMEDIATEDELETE" Then
                    If LCase(SplitArray(1)) = LCase(SenderAddress) Then
                        SenderHasRule = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Public Sub CountBySender()
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim senderCounts As Scripting.Dictionary
    Dim i As Long, j As Long
    Dim addr As String
    Dim keys() As String
    Dim vals() As Long
    Dim tempK As String
    Dim tempV As Long
    Dim output As String
    Dim n As Long
    Dim exUser As Outlook.ExchangeUser

    Dim Session As Outlook.NameSpace
    Dim Folder As Outlook.Folder
    Dim SubFolder As Outlook.Folder
    Set Session = Application.Session
    For Each Folder In Session.Folders
        If Folder.Name = "kyle.brothers@louisville.edu" Then
            For Each SubFolder In Folder.Folders
                If SubFolder.Name = "Inbox" Then
                    Set myFolder = SubFolder
                    Exit For
                End If
            Next
            Exit For
        End If
    Next

    If myFolder Is Nothing Then
        MsgBox "Inbox not found."
        Exit Sub
    End If

    Set senderCounts = New Scripting.Dictionary
    Set myItems = myFolder.Items

    For i = 1 To myItems.Count
        Set myItem = myItems.Item(i)
        If myItem.Class = olMail Then
            If myItem.SenderEmailType = "EX" Then
                Set exUser = myItem.Sender.GetExchangeUser()
                If Not exUser Is Nothing Then
                    addr = LCase(exUser.PrimarySmtpAddress)
                Else
                    addr = LCase(myItem.SenderEmailAddress)
                End If
            Else
                addr = LCase(myItem.SenderEmailAddress)
            End If

            If senderCounts.Exists(addr) Then
                senderCounts(addr) = senderCounts(addr) + 1
            Else
                senderCounts.Add addr, 1
            End If
        End If
    Next i

    n = senderCounts.Count
    ReDim keys(0 To n - 1)
    ReDim vals(0 To n - 1)
    For i = 0 To n - 1
        keys(i) = senderCounts.keys()(i)
        vals(i) = senderCounts.Items()(i)
    Next i

    For j = 0 To n - 2
        For i = 0 To n - 2
            If vals(i) < vals(i + 1) Then
                tempV = vals(i): vals(i) = vals(i + 1): vals(i + 1) = tempV
                tempK = keys(i): keys(i) = keys(i + 1): keys(i + 1) = tempK
            End If
        Next i
    Next j

    output = "Sender counts in Inbox (" & n & " senders, " & myFolder.Items.Count & " total):" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        Dim line As String
        line = vals(i) & vbTab & keys(i)
        output = output & line & vbCrLf
        Debug.Print line
    Next i

    MsgBox output
End Sub

Public Sub CountByConversation()
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim convCounts As Scripting.Dictionary
    Dim i As Long, j As Long
    Dim Topic As String
    Dim keys() As String
    Dim vals() As Long
    Dim tempK As String
    Dim tempV As Long
    Dim output As String
    Dim n As Long

    Dim Session As Outlook.NameSpace
    Dim Folder As Outlook.Folder
    Dim SubFolder As Outlook.Folder
    Set Session = Application.Session
    For Each Folder In Session.Folders
        If Folder.Name = "kyle.brothers@louisville.edu" Then
            For Each SubFolder In Folder.Folders
                If SubFolder.Name = "Inbox" Then
                    Set myFolder = SubFolder
                    Exit For
                End If
            Next
            Exit For
        End If
    Next

    If myFolder Is Nothing Then
        MsgBox "Inbox not found."
        Exit Sub
    End If

    Set convCounts = New Scripting.Dictionary
    Set myItems = myFolder.Items

    For i = 1 To myItems.Count
        Set myItem = myItems.Item(i)
        If myItem.Class = olMail Then
            Topic = myItem.ConversationTopic
            If Topic = "" Then Topic = "(no subject)"
            If convCounts.Exists(Topic) Then
                convCounts(Topic) = convCounts(Topic) + 1
            Else
                convCounts.Add Topic, 1
            End If
        End If
    Next i

    n = convCounts.Count
    ReDim keys(0 To n - 1)
    ReDim vals(0 To n - 1)
    For i = 0 To n - 1
        keys(i) = convCounts.keys()(i)
        vals(i) = convCounts.Items()(i)
    Next i

    For j = 0 To n - 2
        For i = 0 To n - 2
            If vals(i) < vals(i + 1) Then
                tempV = vals(i): vals(i) = vals(i + 1): vals(i + 1) = tempV
                tempK = keys(i): keys(i) = keys(i + 1): keys(i + 1) = tempK
            End If
        Next i
    Next j

    output = "Conversation counts in Inbox (" & n & " conversations, " & myFolder.Items.Count & " total):" & vbCrLf & vbCrLf
    For i = 0 To n - 1
        If vals(i) > 1 Then
            Dim line As String
            line = vals(i) & vbTab & keys(i)
            output = output & line & vbCrLf
            Debug.Print line
        Else
            Exit For
        End If
    Next i

    MsgBox output
End Sub
