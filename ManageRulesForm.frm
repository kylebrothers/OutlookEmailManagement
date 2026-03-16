VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageRulesForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3900
   OleObjectBlob   =   "ManageRulesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManageRulesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Version 5#
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageRulesForm
   Caption = "Manage Rules"
   ClientHeight = 9702
   ClientLeft = 36
   ClientTop = 384
   ClientWidth = 6000
   StartUpPosition = 3    'Windows Default
End
Attribute VB_Name = "ManageRulesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
        .ColumnWidths = "100;220;50"
    End With
    
End Sub

Private Sub InitializeData()
    With Me.cboRuleType
        .Clear
        .AddItem "SENDERDELETE"
        .ListIndex = 0
    End With
    Call LoadRules
    Call PopulateParameters
End Sub

' =====================================================================
' RULE TYPE CHANGE - DISPATCHER
' =====================================================================

Private Sub cboRuleType_Change()
    Call ClearInputs
    Call PopulateParameters
End Sub

Private Sub PopulateParameters()
    Call SetParameterLabels
    Select Case Me.cboRuleType.Text
        Case "SENDERDELETE"
            Call PopulateForSenderDelete
        ' Future rule types added here:
        ' Case "SUBJECTARCHIVE"
        '     Call PopulateForSubjectArchive
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
        ' Future rule types added here:
        ' Case "SUBJECTARCHIVE"
        '     Me.lblP1.Caption = "..."
        '     Me.lblP2.Caption = "..."
        '     Me.lblP3.Caption = "..."
        '     Me.lblP4.Caption = "..."
        '     Me.lblP5.Caption = "..."
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

' =====================================================================
' ADD / DELETE / CLOSE
' =====================================================================

Private Sub btnAdd_Click()
    If Me.cboRuleType.Text = "" Then
        MsgBox "Please select a rule type."
        Exit Sub
    End If
    If Me.txtP1.Text = "" Then
        MsgBox "Please enter a value for Parameter 1."
        Exit Sub
    End If
    
    ' For SENDERDELETE, default P2 to 30 if left blank
    If Me.cboRuleType.Text = "SENDERDELETE" And Me.txtP2.Text = "" Then
        Me.txtP2.Text = "30"
    End If
    
    Dim NewRecord As String
    NewRecord = Me.cboRuleType.Text & ":" & _
                Me.txtP1.Text & ":" & _
                Me.txtP2.Text & ":" & _
                Me.txtP3.Text & ":" & _
                Me.txtP4.Text & ":" & _
                Me.txtP5.Text
    
    Call SaveRule(NewRecord)
    Call LoadRules
    Call ClearInputs
    Call PopulateParameters
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
            SplitArray = Split(TargetArray(i), ":")
            If Not (SplitArray(0) = SelectedType And SplitArray(1) = SelectedP1) Then
                If NewBody <> "" Then NewBody = NewBody & "::"
                NewBody = NewBody & TargetArray(i)
            End If
        End If
    Next i
    
    myStorage.Body = NewBody
    myStorage.Save
    Call LoadRules
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
            SplitArray = Split(TargetArray(i), ":")
            If UBound(SplitArray) >= 5 Then
                Me.lbRules.AddItem
                Me.lbRules.List(Me.lbRules.ListCount - 1, 0) = SplitArray(0) ' RuleType
                Me.lbRules.List(Me.lbRules.ListCount - 1, 1) = SplitArray(1) ' P1
                Me.lbRules.List(Me.lbRules.ListCount - 1, 2) = SplitArray(2) ' P2
            End If
        End If
    Next i
End Sub

Private Function GetRulesStorage() As StorageItem
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

' =====================================================================
' UTILITIES
' =====================================================================

Private Sub ClearInputs()
    Dim i As Integer
    For i = 1 To 5
        Me.Controls("txtP" & i).Text = ""
    Next i
End Sub


