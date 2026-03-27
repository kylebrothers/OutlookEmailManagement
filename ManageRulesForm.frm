Option Explicit

' =====================================================================
' INITIALIZE
' =====================================================================

Private Sub UserForm_Initialize()
    Call InitializeLayout
    Call InitializeData
End Sub

Private Sub InitializeLayout()
    
    ' === FORM ===
    With Me
        .Width = 420
        .Height = 520
        .Caption = "Manage Rules"
        .BackColor = RGB(240, 240, 240)
    End With
    
    ' === RULE TYPE LABEL ===
    With Me.lblRuleType
        .Caption = "Rule Type"
        .Left = 10
        .Top = 12
        .Width = 80
        .Height = 20
        .BackColor = RGB(240, 240, 240)
    End With
    
    ' === RULE TYPE DROPDOWN ===
    With Me.cboRuleType
        .Left = 100
        .Top = 10
        .Width = 200
        .Height = 20
    End With
    
    ' === PARAMETER LABELS AND TEXTBOXES ===
    Dim i As Integer
    For i = 1 To 5
        With Me.Controls("lblP" & i)
            .Left = 10
            .Top = 10 + (i * 30)
            .Width = 80
            .Height = 20
            .BackColor = RGB(240, 240, 240)
        End With
        
        With Me.Controls("txtP" & i)
            .Left = 100
            .Top = 10 + (i * 30)
            .Width = 290
            .Height = 20
        End With
    Next i
    
    ' === FOLDER PICKER BUTTON ===
    With Me.btnPickFolder
        .Caption = "Pick Folder"
        .Left = 100
        .Top = 10 + (3 * 30)
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    
    ' === ADD BUTTON ===
    With Me.btnAdd
        .Caption = "Add Rule"
        .Left = 10
        .Top = 210
        .Width = 80
        .Height = 25
    End With
    
    ' === DELETE BUTTON ===
    With Me.btnDelete
        .Caption = "Delete Rule"
        .Left = 100
        .Top = 210
        .Width = 80
        .Height = 25
    End With
    
    ' === CLOSE BUTTON ===
    With Me.btnClose
        .Caption = "Close"
        .Left = 190
        .Top = 210
        .Width = 80
        .Height = 25
    End With
    
    ' === RULES LIST LABEL ===
    With Me.lblRules
        .Caption = "Current Rules"
        .Left = 10
        .Top = 250
        .Width = 100
        .Height = 20
        .BackColor = RGB(240, 240, 240)
    End With
    
    ' === RULES LISTBOX ===
    With Me.lbRules
        .Left = 10
        .Top = 270
        .Width = 380
        .Height = 180
        .ColumnCount = 3
        .ColumnWidths = "120;200;50"
    End With
    
End Sub

Private Sub InitializeData()
    With Me.cboRuleType
        .Clear
        .AddItem "SENDERDELETE"
        .AddItem "SENDERIMMEDIATE"
        .AddItem "SENDERFOLDER"
        .AddItem "TOPICDELETE"
        .AddItem "TOPICIMMEDIATEDELETE"
        .ListIndex = 1
    End With
    Call LoadRules
    Call PopulateParameters
    Call HighlightMatchingRule
End Sub

' =====================================================================
' RULE TYPE CHANGE - DISPATCHER
' =====================================================================

Private Sub cboRuleType_Change()
    Call ClearInputs
    Call PopulateParameters
    Call SetPickFolderVisibility
End Sub

Private Sub PopulateParameters()
    Call SetParameterLabels
    Select Case Me.cboRuleType.Text
        Case "SENDERDELETE"
            Call PopulateForSenderDelete
        Case "SENDERIMMEDIATE"
            Call PopulateForSenderImmediate
        Case "SENDERFOLDER"
            Call PopulateForSenderFolder
        Case "TOPICDELETE"
            Call PopulateForTopicDelete
        Case "TOPICIMMEDIATEDELETE"
            Call PopulateForTopicImmediateDelete
    End Select
End Sub

' =====================================================================
' PARAMETER LABELS - DISPATCHER
' =====================================================================

Private Sub SetParameterLabels()
    Select Case Me.cboRuleType.Text
        Case "SENDERDELETE"
            Me.lblP1.Caption = "Sender Email"
            Me.lblP2.Caption = "Days (default 30)"
            Me.lblP3.Caption = "P3"
            Me.lblP4.Caption = "P4"
            Me.lblP5.Caption = "P5"
        Case "SENDERIMMEDIATE"
            Me.lblP1.Caption = "Sender Email"
            Me.lblP2.Caption = "P2"
            Me.lblP3.Caption = "P3"
            Me.lblP4.Caption = "P4"
            Me.lblP5.Caption = "P5"
        Case "SENDERFOLDER"
            Me.lblP1.Caption = "Sender Email"
            Me.lblP2.Caption = "Days (default 30)"
            Me.lblP3.Caption = "Folder EntryID"
            Me.lblP4.Caption = "Folder Path"
            Me.lblP5.Caption = "P5"
        Case "TOPICDELETE"
            Me.lblP1.Caption = "Sender Email"
            Me.lblP2.Caption = "Topic"
            Me.lblP3.Caption = "Days (default 30)"
            Me.lblP4.Caption = "P4"
            Me.lblP5.Caption = "P5"
        Case "TOPICIMMEDIATEDELETE"
            Me.lblP1.Caption = "Sender Email"
            Me.lblP2.Caption = "Topic"
            Me.lblP3.Caption = "P3"
            Me.lblP4.Caption = "P4"
            Me.lblP5.Caption = "P5"
    End Select
End Sub

' =====================================================================
' RULE-SPECIFIC POPULATION LOGIC
' =====================================================================

Private Sub PopulateForSenderDelete()
    Dim objItem As Object
    On Error Resume Next
    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    If Not objItem Is Nothing Then
        If objItem.Class = olMail Then
            Me.txtP1.Text = objItem.SenderEmailAddress
            Me.txtP2.Text = "30"
        End If
    End If
End Sub

Private Sub PopulateForSenderImmediate()
    Dim objItem As Object
    On Error Resume Next
    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    If Not objItem Is Nothing Then
        If objItem.Class = olMail Then
            Me.txtP1.Text = objItem.SenderEmailAddress
        End If
    End If
End Sub

Private Sub PopulateForSenderFolder()
    Dim objItem As Object
    On Error Resume Next
    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    If Not objItem Is Nothing Then
        If objItem.Class = olMail Then
            Me.txtP1.Text = objItem.SenderEmailAddress
            Me.txtP2.Text = "30"
        End If
    End If
End Sub

Private Sub PopulateForTopicDelete()
    Dim objItem As Object
    On Error Resume Next
    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    If Not objItem Is Nothing Then
        If objItem.Class = olMail Then
            Me.txtP1.Text = objItem.SenderEmailAddress
            Me.txtP2.Text = objItem.ConversationTopic
            Me.txtP3.Text = "30"
        End If
    End If
End Sub

Private Sub PopulateForTopicImmediateDelete()
    Dim objItem As Object
    On Error Resume Next
    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    If Not objItem Is Nothing Then
        If objItem.Class = olMail Then
            Me.txtP1.Text = objItem.SenderEmailAddress
            Me.txtP2.Text = objItem.ConversationTopic
        End If
    End If
End Sub

' =====================================================================
' FOLDER PICKER
' =====================================================================

Private Sub SetPickFolderVisibility()
    If Me.cboRuleType.Text = "SENDERFOLDER" Then
        Me.btnPickFolder.Visible = True
        Me.txtP3.Visible = False
        Me.txtP4.Visible = False
    Else
        Me.btnPickFolder.Visible = False
        Me.txtP3.Visible = True
        Me.txtP4.Visible = True
    End If
End Sub

Private Sub btnPickFolder_Click()
    Dim PickedFolder As Outlook.Folder
    Set PickedFolder = Application.Session.PickFolder()
    If Not PickedFolder Is Nothing Then
        Me.txtP3.Text = PickedFolder.EntryID
        Me.txtP4.Text = PickedFolder.FolderPath
    End If
End Sub

' =====================================================================
' ADD / DELETE / CLOSE
' =====================================================================

Private Sub btnAdd_Click()
    If Me.cboRuleType.Text = "" Then
        MsgBox "Please select a rule type."
        Exit Sub
    End If
    If Me.txtP1.Text = "" Then
        MsgBox "Please enter a sender email address."
        Exit Sub
    End If

    Dim RuleType As String
    RuleType = Me.cboRuleType.Text

    ' === DUPLICATE CHECKING ===
    If RuleType = "SENDERDELETE" Or RuleType = "SENDERIMMEDIATE" Or RuleType = "SENDERFOLDER" Then
        ' These rule types require the sender to have no rules at all
        If SenderAlreadyHasRule(Me.txtP1.Text) Then
            MsgBox "A rule already exists for sender: " & Me.txtP1.Text & ". Delete the existing rule before adding a new one."
            Exit Sub
        End If
    ElseIf RuleType = "TOPICDELETE" Or RuleType = "TOPICIMMEDIATEDELETE" Then
        ' Topic rules block if sender has a SENDERDELETE or SENDERIMMEDIATE rule
        If SenderHasBroadRule(Me.txtP1.Text) Then
            MsgBox "A sender-level rule (SENDERDELETE or SENDERIMMEDIATE) already exists for: " & Me.txtP1.Text & ". Delete it before adding a topic rule."
            Exit Sub
        End If
        ' Topic rules also block on exact sender+topic duplicate
        If Me.txtP2.Text = "" Then
            MsgBox "Please enter a topic."
            Exit Sub
        End If
        If TopicRuleAlreadyExists(Me.txtP1.Text, Me.txtP2.Text) Then
            MsgBox "A topic rule already exists for sender: " & Me.txtP1.Text & " with topic: " & Me.txtP2.Text & "."
            Exit Sub
        End If
    End If

    ' === RULE-SPECIFIC VALIDATION ===
    If RuleType = "SENDERDELETE" And Me.txtP2.Text = "" Then Me.txtP2.Text = "30"
    If RuleType = "SENDERFOLDER" Then
        If Me.txtP2.Text = "" Then Me.txtP2.Text = "30"
        If Me.txtP3.Text = "" Then
            MsgBox "Please pick a destination folder."
            Exit Sub
        End If
    End If
    If RuleType = "TOPICDELETE" And Me.txtP3.Text = "" Then Me.txtP3.Text = "30"

    Dim NewRecord As String
    NewRecord = RuleType & "|" & _
                Me.txtP1.Text & "|" & _
                Me.txtP2.Text & "|" & _
                Me.txtP3.Text & "|" & _
                Me.txtP4.Text & "|" & _
                Me.txtP5.Text

    Call SaveRule(NewRecord)
    Call LoadRules
    Call HighlightMatchingRule
    Call ClearInputs
    Call PopulateParameters
    Call SetPickFolderVisibility
End Sub

Private Sub btnDelete_Click()
    If Me.lbRules.ListIndex = -1 Then
        MsgBox "Please select a rule to delete."
        Exit Sub
    End If
    
    Dim SelectedType As String
    Dim SelectedP1 As String
    SelectedType = Me.lbRules.List(Me.lbRules.ListIndex, 0)
    SelectedP1 = Me.lbRules.List(Me.lbRules.ListIndex, 1)
    
    Dim myStorage As StorageItem
    Set myStorage = GetRulesStorage()
    If myStorage Is Nothing Then Exit Sub
    
    Dim ParseString As String
    Dim TargetArray As Variant
    Dim SplitArray As Variant
    Dim NewBody As String
    Dim i As Long
    
    ParseString = myStorage.Body
    TargetArray = Split(ParseString, "::")
    NewBody = ""
    
    For i = LBound(TargetArray) To UBound(TargetArray)
        If TargetArray(i) <> "" Then
            SplitArray = Split(TargetArray(i), "|")
            If Not (SplitArray(0) = SelectedType And SplitArray(1) = SelectedP1) Then
                If NewBody <> "" Then NewBody = NewBody & "::"
                NewBody = NewBody & TargetArray(i)
            End If
        End If
    Next i
    
    myStorage.Body = NewBody
    myStorage.Save
    Call LoadRules
    Call HighlightMatchingRule
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' =====================================================================
' STORAGE
' =====================================================================

Private Sub SaveRule(ByVal NewRecord As String)
    Dim myStorage As StorageItem
    Set myStorage = GetRulesStorage()
    If myStorage Is Nothing Then Exit Sub
    
    Dim CurrentBody As String
    CurrentBody = myStorage.Body
    If CurrentBody <> "" Then CurrentBody = CurrentBody & "::"
    CurrentBody = CurrentBody & NewRecord
    
    myStorage.Body = CurrentBody
    myStorage.Save
End Sub

Private Sub LoadRules()
    Dim myStorage As StorageItem
    Dim ParseString As String
    Dim TargetArray As Variant
    Dim SplitArray As Variant
    Dim i As Long
    
    Me.lbRules.Clear
    
    Set myStorage = GetRulesStorage()
    If myStorage Is Nothing Then Exit Sub
    
    ParseString = myStorage.Body
    If ParseString = "" Then Exit Sub
    
    TargetArray = Split(ParseString, "::")
    For i = LBound(TargetArray) To UBound(TargetArray)
        If TargetArray(i) <> "" Then
            SplitArray = Split(TargetArray(i), "|")
            If UBound(SplitArray) >= 5 Then
                Me.lbRules.AddItem
                Me.lbRules.List(Me.lbRules.ListCount - 1, 0) = SplitArray(0)
                Me.lbRules.List(Me.lbRules.ListCount - 1, 1) = SplitArray(1)
                Select Case SplitArray(0)
                    Case "TOPICDELETE", "TOPICIMMEDIATEDELETE"
                        Me.lbRules.List(Me.lbRules.ListCount - 1, 2) = SplitArray(2)
                    Case Else
                        Me.lbRules.List(Me.lbRules.ListCount - 1, 2) = SplitArray(2)
                End Select
            End If
        End If
    Next i
End Sub

Private Function GetRulesStorage() As StorageItem
    Set GetRulesStorage = Module1.GetRulesStorage()
End Function

' =====================================================================
' HIGHLIGHT MATCHING RULE
' =====================================================================

Private Sub HighlightMatchingRule()
    Dim objItem As Object
    Dim SenderAddress As String
    Dim i As Long
    
    On Error Resume Next
    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    If objItem Is Nothing Then Exit Sub
    If Not objItem.Class = olMail Then Exit Sub
    
    SenderAddress = LCase(objItem.SenderEmailAddress)
    
    For i = 0 To Me.lbRules.ListCount - 1
        If LCase(Me.lbRules.List(i, 1)) = SenderAddress Then
            Me.lbRules.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

' =====================================================================
' DUPLICATE / CONSTRAINT CHECKING
' =====================================================================

Private Function SenderAlreadyHasRule(ByVal SenderAddress As String) As Boolean
    Dim i As Long
    SenderAlreadyHasRule = False
    For i = 0 To Me.lbRules.ListCount - 1
        If LCase(Me.lbRules.List(i, 1)) = LCase(SenderAddress) Then
            SenderAlreadyHasRule = True
            Exit Function
        End If
    Next i
End Function

Private Function SenderHasBroadRule(ByVal SenderAddress As String) As Boolean
    Dim myStorage As StorageItem
    Dim ParseString As String
    Dim TargetArray As Variant
    Dim SplitArray As Variant
    Dim i As Long
    
    SenderHasBroadRule = False
    Set myStorage = GetRulesStorage()
    If myStorage Is Nothing Then Exit Function
    ParseString = myStorage.Body
    If ParseString = "" Then Exit Function
    
    TargetArray = Split(ParseString, "::")
    For i = LBound(TargetArray) To UBound(TargetArray)
        If TargetArray(i) <> "" Then
            SplitArray = Split(TargetArray(i), "|")
            If UBound(SplitArray) >= 5 Then
                If SplitArray(0) = "SENDERDELETE" Or SplitArray(0) = "SENDERIMMEDIATE" Then
                    If LCase(SplitArray(1)) = LCase(SenderAddress) Then
                        SenderHasBroadRule = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Function TopicRuleAlreadyExists(ByVal SenderAddress As String, ByVal Topic As String) As Boolean
    Dim myStorage As StorageItem
    Dim ParseString As String
    Dim TargetArray As Variant
    Dim SplitArray As Variant
    Dim i As Long
    
    TopicRuleAlreadyExists = False
    Set myStorage = GetRulesStorage()
    If myStorage Is Nothing Then Exit Function
    ParseString = myStorage.Body
    If ParseString = "" Then Exit Function
    
    TargetArray = Split(ParseString, "::")
    For i = LBound(TargetArray) To UBound(TargetArray)
        If TargetArray(i) <> "" Then
            SplitArray = Split(TargetArray(i), "|")
            If UBound(SplitArray) >= 5 Then
                If SplitArray(0) = "TOPICDELETE" Or SplitArray(0) = "TOPICIMMEDIATEDELETE" Then
                    If LCase(SplitArray(1)) = LCase(SenderAddress) Then
                        If LCase(SplitArray(2)) = LCase(Topic) Then
                            TopicRuleAlreadyExists = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

' =====================================================================
' UTILITIES
' =====================================================================

Private Sub ClearInputs()
    Dim i As Integer
    For i = 1 To 5
        Me.Controls("txtP" & i).Text = ""
    Next i
End Sub

