Public Actions As Scripting.Dictionary

Private Sub lbActions_Click()

End Sub

Private Sub txtSelect_Change()
On Error GoTo On_Error
    If Int(txtSelect.Value) < FoldersList.ListCount Then
        FoldersList.Selected(Int(txtSelect.Value)) = True
    Else
        txtSelect.Value = Left(txtSelect.Value, Len(txtSelect.Value) - 1)
    End If

Exiting:
    Exit Sub

On_Error:
    txtSelect.Value = ""
End Sub

Private Sub txtSelect_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim TargetFolder As String
        Dim i As Long
        With Me.FoldersList
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    TargetFolder = .List(i, 1)
                    MoveSelectedMessages (TargetFolder)
                    Exit For
                End If
            Next
        End With
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    ElseIf KeyCode = vbKeyF1 Then
        Dim ReturnValue As Integer
        ReturnValue = MsgBox("Do you want to delete the FolderHistory body?", vbOKCancel)
        If ReturnValue = 1 Then
            Call DeleteFolderHistory
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call DisplayFolderHistory
    ElseIf KeyCode = vbKeyF3 Then
        If ThisOutlookSession.SaveColumns Then
            ThisOutlookSession.SaveColumns = False
            Call Project1.AssignFolderForm.SetFormColor
        Else
            ThisOutlookSession.SaveColumns = True
            Call Project1.AssignFolderForm.SetFormColor
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    txtSelect.Text = ""
    txtSelect.SetFocus
End Sub

Private Sub UserForm_Activate()
    Call SetManageRulesButtonColor
End Sub

Sub LoadSubFolders()
On Error GoTo On_Error
    
    FoldersList.Clear
    
    Dim Session As Outlook.NameSpace
    Dim Folders As Outlook.Folders
    Dim Folder As Outlook.Folder
    Dim SubFolders As Outlook.Folders
    Dim SubFolder As Outlook.Folder
    Dim SubSubFolders As Outlook.Folders
    Dim SubSubFolder As Outlook.Folder
    
    Set Session = Application.Session
    
    Set Folders = Session.Folders
    For Each Folder In Folders
        If Folder.Name = "kyle.brothers@louisville.edu" Then
            Set SubFolders = Folder.Folders
            For Each SubFolder In SubFolders
                If SubFolder.Name = Module1.CurrentFolder Then
                    Set SubSubFolders = SubFolder.Folders
                    For Each SubSubFolder In SubSubFolders
                        If Not Left(SubSubFolder.Name, 1) = "*" Then
                            With FoldersList
                                .AddItem
                                .List(.ListCount - 1, 1) = SubSubFolder.Name
                            End With
                        End If
                    Next
                    SortListbox
                    NumberRows
                    Exit For
                End If
            Next
        End If
    Next

Exiting:
        Set Session = Nothing
        Exit Sub
On_Error:
    MsgBox "error=" & Err.Number & " " & Err.Description
    Resume Exiting

End Sub

Sub MatchSubjects()
    Dim objItem As Outlook.MailItem
    Dim CurrentSelectedSubject As String
    Dim CurrentListSubject As String
    Dim CurrentListFolder As String
    Dim i As Long
    Dim j As Long
    Dim key As Variant
    
    For Each objItem In Application.ActiveExplorer.Selection
            If Not objItem.Subject = "" Then
                CurrentSelectedSubject = CleanSubject(objItem.Subject)
                
                For Each key In Actions.keys
                    CurrentListSubject = key
                    
                    If CurrentListSubject = CurrentSelectedSubject Then
                        CurrentListFolder = Actions.Item(CurrentListSubject)
                        For j = 0 To (Me.FoldersList.ListCount - 1)
                            If CurrentListFolder = Me.FoldersList.List(j, 1) Then
                                Me.FoldersList.Selected(j) = True
                                Me.txtSelect = Me.FoldersList.List(j, 0)
                                Exit Sub
                            End If
                       Next j
                    End If
                    
                Next
                
            End If
            
    Next
End Sub

Sub SortListbox()
    Dim i As Long
    Dim j As Long
    Dim temp As Variant
       
    With Me.FoldersList
        For j = 0 To FoldersList.ListCount - 2
            For i = 0 To FoldersList.ListCount - 2
                If UCase(.List(i, 1)) > UCase(.List(i + 1, 1)) Then
                    temp = .List(i, 1)
                    .List(i, 1) = .List(i + 1, 1)
                    .List(i + 1, 1) = temp
                End If
            Next i
        Next j
    End With
End Sub

Sub CleanLbActions()
    Dim i As Long
    Dim j As Long
    Dim temp0 As Variant
    Dim temp1 As Variant
       
    With Me.lbActions
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If UCase(.List(i, 0)) > UCase(.List(i + 1, 0)) Then
                    temp0 = .List(i, 0)
                    temp1 = .List(i, 1)
                    .List(i, 0) = .List(i + 1, 0)
                    .List(i, 1) = .List(i + 1, 1)
                    .List(i + 1, 0) = temp0
                    .List(i + 1, 1) = temp1
                End If
            Next i
        Next j
        
        For i = (.ListCount - 1) To 1 Step -1
            If .List(i, 0) = .List(i - 1, 0) Then
                .RemoveItem (i)
            End If
        Next
    End With
End Sub

Sub NumberRows()
    Dim i As Long
    For i = 0 To (FoldersList.ListCount - 1)
        FoldersList.List(i, 0) = i
    Next i
End Sub

Sub MoveSelectedMessages(ByVal TargetFolder As String)
    Dim moveToFolder As Outlook.Folder
    Dim objItem As Outlook.MailItem
    
    Dim Session As Outlook.NameSpace
    Dim Folders As Outlook.Folders
    Dim Folder As Outlook.Folder
    Dim SubFolders As Outlook.Folders
    Dim SubFolder As Outlook.Folder
    Dim SubSubFolders As Outlook.Folders
    Dim SubSubFolder As Outlook.Folder
    Dim i As Long
    
    Set Session = Application.Session
    
    Set Folders = Session.Folders
    For Each Folder In Folders
        If Folder.Name = "kyle.brothers@louisville.edu" Then
            Set SubFolders = Folder.Folders
            For Each SubFolder In SubFolders
                If SubFolder.Name = Module1.CurrentFolder Then
                    Set SubSubFolders = SubFolder.Folders
                    For Each SubSubFolder In SubSubFolders
                        If SubSubFolder.Name = TargetFolder Then
                            Set moveToFolder = SubSubFolder
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        End If
    Next
    
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox ("No item selected")
    Else
        For Each objItem In Application.ActiveExplorer.Selection
            With Actions
                .Item(CleanSubject(objItem.Subject)) = moveToFolder
            End With
            
            objItem.UnRead = False
            objItem.TaskCompletedDate = Date
            objItem.Move moveToFolder
        Next
        
        Call ThisOutlookSession.RunRules
        
        Unload Me
        Exit Sub
    End If
End Sub

Sub SaveColumnToStorage(ByVal StorageSubject As String)
    Dim Session As Outlook.NameSpace
    Dim Folders As Outlook.Folders
    Dim Folder As Outlook.Folder
    Dim SubFolders As Outlook.Folders
    Dim SubFolder As Outlook.Folder
    Dim myStorage As StorageItem
    Dim TempBody As String
    Dim key As Variant
     
    If ThisOutlookSession.SaveColumns Then
         
        Set Session = Application.Session
        
        Set Folders = Session.Folders
        For Each Folder In Folders
            If Folder.Name = "kyle.brothers@louisville.edu" Then
                Set SubFolders = Folder.Folders
                For Each SubFolder In SubFolders
                    If SubFolder.Name = "Inbox" Then
                        Set myStorage = SubFolder.GetStorage(StorageSubject, olIdentifyBySubject)
                        Exit For
                    End If
                Next
                Exit For
            End If
        Next
        
        If Not myStorage Is Nothing Then
            myStorage.Body = ""
            For Each key In Actions.keys
                If Not TempBody = "" Then
                    TempBody = TempBody + "::"
                End If
                TempBody = TempBody + key + ":" + Actions(key)
            Next
        End If
            
        myStorage.Body = TempBody
        myStorage.Save
    
    End If
    
End Sub

Sub DeleteFolderHistory()
    Dim Session As Outlook.NameSpace
    Dim Folders As Outlook.Folders
    Dim Folder As Outlook.Folder
    Dim SubFolders As Outlook.Folders
    Dim SubFolder As Outlook.Folder
    Dim myStorage As StorageItem
    Dim StorageSubject As String
    
    StorageSubject = "FolderHistory"
    
    Set Session = Application.Session
    
    Set Folders = Session.Folders
    For Each Folder In Folders
        If Folder.Name = "kyle.brothers@louisville.edu" Then
            Set SubFolders = Folder.Folders
            For Each SubFolder In SubFolders
                If SubFolder.Name = "Inbox" Then
                    Set myStorage = SubFolder.GetStorage(StorageSubject, olIdentifyBySubject)
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    MsgBox (myStorage.Body)
    
    If Not myStorage Is Nothing Then
        myStorage.Body = ""
    End If
        
    myStorage.Save
End Sub

Sub DisplayFolderHistory()
    Dim Session As Outlook.NameSpace
    Dim Folders As Outlook.Folders
    Dim Folder As Outlook.Folder
    Dim SubFolders As Outlook.Folders
    Dim SubFolder As Outlook.Folder
    Dim myStorage As StorageItem
    Dim StorageSubject As String
    
    StorageSubject = "FolderHistory"
    
    Set Session = Application.Session
    
    Set Folders = Session.Folders
    For Each Folder In Folders
        If Folder.Name = "kyle.brothers@louisville.edu" Then
            Set SubFolders = Folder.Folders
            For Each SubFolder In SubFolders
                If SubFolder.Name = "Inbox" Then
                    Set myStorage = SubFolder.GetStorage(StorageSubject, olIdentifyBySubject)
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    MsgBox ("Character length:" + Str(Len(myStorage.Body)) + "   " + myStorage.Body)
    
End Sub

Sub LoadStorageToColumn(ByVal StorageSubject As String)
    Dim Session As Outlook.NameSpace
    Dim Folders As Outlook.Folders
    Dim Folder As Outlook.Folder
    Dim SubFolders As Outlook.Folders
    Dim SubFolder As Outlook.Folder
    Dim myStorage As StorageItem
    Dim ParseString As String
    Dim TargetArray As Variant
    Dim SplitArray As Variant
    Dim i As Long
    
    Set Session = Application.Session
    
    Set Folders = Session.Folders
    Set Actions = New Scripting.Dictionary
        
    For Each Folder In Folders
        If Folder.Name = "kyle.brothers@louisville.edu" Then
            Set SubFolders = Folder.Folders
            For Each SubFolder In SubFolders
                If SubFolder.Name = "Inbox" Then
                    Set myStorage = SubFolder.GetStorage(StorageSubject, olIdentifyBySubject)
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    If Not myStorage Is Nothing Then
        ParseString = myStorage.Body
        TargetArray = Split(ParseString, "::")
        For i = LBound(TargetArray) To UBound(TargetArray)
            SplitArray = Split(TargetArray(i), ":")
            With Actions
                .Item(SplitArray(0)) = SplitArray(1)
            End With
        Next
    End If
    
End Sub

Public Sub SetFormColor()
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim NumberOfItems As Integer
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    
    NumberOfItems = myFolder.Items.Count
    
    If NumberOfItems > 40 Then
        If ThisOutlookSession.SaveColumns Then
            AssignFolderForm.BackColor = RGB(220, 20, 60)
        Else
            AssignFolderForm.BackColor = RGB(255, 153, 153)
        End If
    Else
        If ThisOutlookSession.SaveColumns Then
            AssignFolderForm.BackColor = RGB(208, 248, 208)
        Else
            AssignFolderForm.BackColor = RGB(144, 238, 144)
        End If
    End If
    
End Sub

Public Sub SetManageRulesButtonColor()
    Dim objItem As Object
    Dim SenderAddress As String
    
    On Error Resume Next
    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    
    If objItem Is Nothing Then
        Me.btnManageRules.BackColor = &H8000000F
        Exit Sub
    End If
    
    If Not objItem.Class = olMail Then
        Me.btnManageRules.BackColor = &H8000000F
        Exit Sub
    End If
    
    SenderAddress = objItem.SenderEmailAddress
    
    If Module1.SenderHasRule(SenderAddress) Then
        Me.btnManageRules.BackColor = RGB(180, 0, 0)
        Me.btnManageRules.ForeColor = RGB(255, 255, 255)
    Else
        Me.btnManageRules.BackColor = &H8000000F
        Me.btnManageRules.ForeColor = &H80000012
    End If
End Sub

Private Sub btnManageRules_Click()
    Unload Me
    ManageRulesForm.Show
End Sub

Function CleanSubject(ByVal OriginalSubject As String) As String
    Dim TempOutput As String
    TempOutput = UCase(Replace(OriginalSubject, " ", ""))
    TempOutput = Replace(TempOutput, "-", "")
    TempOutput = Replace(TempOutput, "RE", "")
    TempOutput = Replace(TempOutput, "FWD", "")
    TempOutput = Replace(TempOutput, "FW", "")
    TempOutput = Replace(TempOutput, "1", "")
    TempOutput = Replace(TempOutput, "2", "")
    TempOutput = Replace(TempOutput, "3", "")
    TempOutput = Replace(TempOutput, "4", "")
    TempOutput = Replace(TempOutput, "5", "")
    TempOutput = Replace(TempOutput, "6", "")
    TempOutput = Replace(TempOutput, "7", "")
    TempOutput = Replace(TempOutput, "8", "")
    TempOutput = Replace(TempOutput, "9", "")
    TempOutput = Replace(TempOutput, "0", "")
    TempOutput = Replace(TempOutput, ",", "")
    TempOutput = Replace(TempOutput, ":", "")
    CleanSubject = TempOutput
End Function
