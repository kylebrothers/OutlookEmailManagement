Option Explicit
Private WithEvents oExpl As Explorer
Private WithEvents oApp As Application
Public WithEvents oItem As MailItem
Public SaveColumns As Boolean

Public Sub Application_Startup()
   Set oApp = Application
   Set oExpl = Application.ActiveExplorer
   SaveColumns = False
End Sub

Private Sub oExpl_InlineResponse(ByVal Item As Object)
   Dim oResponse As MailItem
   Set oResponse = Item

  Dim RemoveThis As VBA.Collection
  Dim Recipients As Outlook.Recipients
  Dim R As Outlook.Recipient
  Dim i&, y&
  Set RemoveThis = New VBA.Collection

  RemoveThis.Add "kyle.brothers@louisville.edu"
  RemoveThis.Add "kbbrot01@louisville.edu"
  RemoveThis.Add "kyle.brothers@vanderbilt.edu"
  RemoveThis.Add "kyle.b.brothers@gmail.com"
  
  Set Recipients = oResponse.Recipients
  For i = Recipients.Count To 1 Step -1
    Set R = Recipients.Item(i)
    For y = 1 To RemoveThis.Count
      If LCase$(R.Address) = LCase$(RemoveThis(y)) Then
        Recipients.Remove i
        Exit For
      End If
    Next
  Next
       
  Set oItem = Nothing
End Sub

Private Sub oExpl_SelectionChange()
   On Error Resume Next
   Set oItem = oExpl.Selection.Item(1)
End Sub

Public Sub HASSUpdates()
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim NumberOfItems As Integer
    Dim NumberOfUnreadItems As Integer
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    
    NumberOfItems = myFolder.Items.Count
    NumberOfUnreadItems = myFolder.UnReadItemCount

    Dim Body As String
    Dim CurrentFriendlyName As String

    CurrentFriendlyName = "Total E-mails"
    Body = "{""state"": """ & CStr(NumberOfItems) & """, ""attributes"": {""friendly_name"": """ & CurrentFriendlyName & """}}"

    CurrentFriendlyName = "Unread E-mails"
    Body = "{""state"": """ & CStr(NumberOfUnreadItems) & """, ""attributes"": {""friendly_name"": """ & CurrentFriendlyName & """}}"
End Sub

Public Sub HASSPost(ByVal SensorName As String, ByVal JSON As String)
    Dim sURL As String
    Dim oHttp As Object
    Set oHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    sURL = "http://www.brothersbrothers.net:8123/api/states/sensor." & SensorName & "?api_password=smonYARChR"
    oHttp.Open "POST", sURL, False
    oHttp.setRequestHeader "Content-type", "application/json"
    oHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    oHttp.Send JSON
    Set oHttp = Nothing
End Sub

Public Sub RunRules()
    Dim myNameSpace As Outlook.NameSpace
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim myStorage As StorageItem
    Dim ParseString As String
    Dim TargetArray As Variant
    Dim SplitArray As Variant
    Dim i As Long

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    Set myStorage = Module1.GetRulesStorage()
    If myStorage Is Nothing Then Exit Sub
    ParseString = myStorage.Body
    If ParseString = "" Then Exit Sub

    TargetArray = Split(ParseString, "::")

    Dim j As Long
    For j = myItems.Count To 1 Step -1
        Set myItem = myItems.Item(j)
        If myItem.Class = olMail Then
            If myItem.FlagStatus = olNoFlag Then
                For i = LBound(TargetArray) To UBound(TargetArray)
                    If TargetArray(i) <> "" Then
                        SplitArray = Split(TargetArray(i), "|")
                        If UBound(SplitArray) >= 5 Then
                            Select Case SplitArray(0)
                                Case "SENDERDELETE"
                                    Dim ThresholdDays As Long
                                    ThresholdDays = CLng(SplitArray(2))
                                    If LCase(myItem.SenderEmailAddress) = LCase(SplitArray(1)) Then
                                        If myItem.ReceivedTime < Now - ThresholdDays Then
                                            myItem.Delete
                                            Exit For
                                        End If
                                    End If
                                Case "SENDERIMMEDIATE"
                                    If LCase(myItem.SenderEmailAddress) = LCase(SplitArray(1)) Then
                                        myItem.Delete
                                        Exit For
                                    End If
                                Case "SENDERFOLDER"
                                    Dim FolderThresholdDays As Long
                                    FolderThresholdDays = CLng(SplitArray(2))
                                    If LCase(myItem.SenderEmailAddress) = LCase(SplitArray(1)) Then
                                        If myItem.ReceivedTime < Now - FolderThresholdDays Then
                                            On Error Resume Next
                                            Dim DestFolder As Outlook.Folder
                                            Set DestFolder = Application.Session.GetFolderFromID(SplitArray(3))
                                            On Error GoTo 0
                                            If Not DestFolder Is Nothing Then
                                                myItem.Move DestFolder
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Case "TOPICIMMEDIATEDELETE"
                                    If LCase(myItem.SenderEmailAddress) = LCase(SplitArray(1)) Then
                                        If LCase(myItem.ConversationTopic) = LCase(SplitArray(2)) Then
                                            myItem.Delete
                                            Exit For
                                        End If
                                    End If
                                Case "TOPICDELETE"
                                    Dim TopicThresholdDays As Long
                                    TopicThresholdDays = CLng(SplitArray(3))
                                    If LCase(myItem.SenderEmailAddress) = LCase(SplitArray(1)) Then
                                        If LCase(myItem.ConversationTopic) = LCase(SplitArray(2)) Then
                                            If myItem.ReceivedTime < Now - TopicThresholdDays Then
                                                myItem.Delete
                                                Exit For
                                            End If
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                Next i
            End If
        End If
    Next j
End Sub
