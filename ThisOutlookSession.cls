VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private WithEvents oExpl As Explorer
Attribute oExpl.VB_VarHelpID = -1
Private WithEvents oApp As Application
Attribute oApp.VB_VarHelpID = -1
Public WithEvents oItem As MailItem
Attribute oItem.VB_VarHelpID = -1
Public SaveColumns As Boolean

Public Sub Application_Startup()
   Set oApp = Application
   Set oExpl = Application.ActiveExplorer
   SaveColumns = False
End Sub
  
'Private Sub oApp_NewMail()sCallMe.HASSUpdates
'End Sub
  
  
Private Sub oExpl_InlineResponse(ByVal Item As Object)
   Dim oResponse As MailItem
   Set oResponse = Item

' add the fields here
  Dim RemoveThis As VBA.Collection
  Dim Recipients As Outlook.Recipients
  Dim R As Outlook.Recipient
  Dim i&, y&
  Set RemoveThis = New VBA.Collection

  ' here add addresses
  RemoveThis.Add "kyle.brothers@louisville.edu"
  RemoveThis.Add "kbbrot01@louisville.edu"
  RemoveThis.Add "kyle.brothers@vanderbilt.edu"
  RemoveThis.Add "kyle.b.brothers@gmail.com"
  
  'MsgBox ("Captured Event")
  
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
'Count the e-mails in the inbox
Dim myNameSpace As Outlook.NameSpace
Dim myFolder As Outlook.Folder
Dim NumberOfItems As Integer
Dim NumberOfUnreadItems As Integer
    
Set myNameSpace = Application.GetNamespace("MAPI")
Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
    
NumberOfItems = myFolder.Items.Count
NumberOfUnreadItems = myFolder.UnReadItemCount

'Call the post sub
Dim Body As String
Dim Number As Integer
Dim CurrentFriendlyName As String

'Total E-mails
CurrentFriendlyName = "Total E-mails"
Body = "{""state"": """ & CStr(NumberOfItems) & """, ""attributes"": {""friendly_name"": """ & CurrentFriendlyName & """}}"
'Call HASSPost("totalemails", Body)

'Unread E-mails
CurrentFriendlyName = "Unread E-mails"
Body = "{""state"": """ & CStr(NumberOfUnreadItems) & """, ""attributes"": {""friendly_name"": """ & CurrentFriendlyName & """}}"
'Call HASSPost("unreademails", Body)

End Sub

Public Sub HASSPost(ByVal SensorName As String, ByVal JSON As String)
    Dim sURL As String, sHTML As String, sAllPosts As String
    Dim oHttp As Object
    Dim blWSExists As Boolean
    Set oHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    sURL = "http://www.brothersbrothers.net:8123/api/states/sensor." & SensorName & "?api_password=smonYARChR"
    oHttp.Open "POST", sURL, False
    oHttp.setRequestHeader "Content-type", "application/json"
    oHttp.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    oHttp.Send JSON
    'sHTML = oHttp.ResponseText
    
    Set oHttp = Nothing
End Sub
  
' ReplyAll
'Private Sub oItem_ReplyAll(ByVal Response As Object, Cancel As Boolean)
        
'   Cancel = True
'   bDiscardEvents = True
   
'   Dim oResponse As MailItem
   
'   Set oResponse = oItem.ReplyAll
    
' add the fields here
'  Dim RemoveThis As VBA.Collection
'  Dim Recipients As Outlook.Recipients
'  Dim R As Outlook.Recipient
'  Dim i&, y&
'  Set RemoveThis = New VBA.Collection

  ' here add addresses
'  RemoveThis.Add "kyle.brothers@louisville.edu"
'  RemoveThis.Add "kbbrot01@louisville.edu"
'  RemoveThis.Add "kyle.brothers@vanderbilt.edu"
'  RemoveThis.Add "kyle.b.brothers@gmail.com"
  
'  Set Recipients = oResponse.Recipients
'  For i = Recipients.Count To 1 Step -1
'    Set R = Recipients.Item(i)

'    For y = 1 To RemoveThis.Count
'     If LCase$(R.Address) = LCase$(RemoveThis(y)) Then
'        Recipients.Remove i
'        Exit For
'      End If
'    Next
'  Next


'   oResponse.Display
       
'   bDiscardEvents = False
'Set oItem = Nothing
'End Sub






