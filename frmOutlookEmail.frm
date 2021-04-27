VERSION 5.00
Begin VB.Form frmOutlookEmail 
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   3870
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton cmdSaveAndExit 
         Caption         =   "55"
         Height          =   495
         Index           =   2
         Left            =   105
         TabIndex        =   9
         Top             =   960
         Width           =   1740
      End
      Begin VB.CommandButton cmdSaveAndExit 
         Caption         =   "55"
         Height          =   495
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   1560
         Width           =   1740
      End
      Begin VB.CommandButton cmdSaveAndExit 
         Caption         =   "55"
         Height          =   495
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   360
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "33"
      Height          =   2175
      Left            =   150
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
      Begin VB.OptionButton optSelect 
         Caption         =   "Option1"
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Option1"
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   780
         Width           =   3015
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "Option1"
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label lblReadEmail 
      Caption         =   "88"
      Height          =   375
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label lblReadEmail 
      Caption         =   "88"
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmOutlookEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DataBaseFileName As String '// This is MS Access DataBase Neme
Dim optSelecteIndex As Integer '// Store the selected option button index
Dim strUserName As String
'//Varible Define for the Exchange Server
Dim oSession As Object
Dim oFolder As Object
Dim oMessages As Object
Dim oMsg As Object
Dim oAttachments As Object
'//Varible Define To Pass The Paramater in Function
Dim strEmailToAdd As String
Dim strEmailFromAdd As String
Dim strEmailBody As String
Dim strEmailSubject As String
Dim EmailRecivedDate As Date
Dim EmailSendDate As Date
Dim I As Integer
Private Sub cmdSaveAndExit_Click(Index As Integer)
   On Error Resume Next
 Dim bolTemp As Boolean
 Dim bolTemp1 As Boolean
 
 Dim varMsg As Variant
 Dim varStyle As Variant
 Dim varTitle As Variant
 Dim varResponse As Variant
 Dim DirPath As String '// Check if the Attachment directory is alrady exit or not
 Dim DirName As String

 Select Case Index
 
 Case 0:
  '// Read Email Details From Outlook and insert into the database
    If optSelecteIndex = 0 Then
        bolTemp = True '// Only used to create attachment folder one time only
        bolTemp1 = True '// For NewEmailFolder and Email Folder Only
       '//Open your Exchange server and logon on it
        Set oSession = CreateObject("MAPI.Session")
        oSession.Logon , , False, False, 0 'Use the existing Outlook session.
        
        Set oFolder = oSession.Inbox
        Set oMessages = oFolder.Messages
        
        Set oMsg = oMessages.GetFirst
        Screen.MousePointer = vbHourglass ' the DeliverNow method could take a while right?
        oSession.DeliverNow ' this now gets all mail services sent and delivered just as the menu option, tools/deliver now/all services does
        '
        While Not oMsg Is Nothing ' If the message collection was empty, oMsg should be equal to the Nothing object, not "Is" operator takes object inputs
                  If oMsg.Unread = False Then
                    strEmailToAdd = oMsg.Sender.Address
                    strEmailFromAdd = oMsg.Sender.Address
                    strEmailBody = Trim(Replace(oMsg.Text, "'", """"))
                    strEmailSubject = Trim(Replace(oMsg.Subject, "'", "''"))
                    EmailRecivedDate = oMsg.TimeReceived
                    EmailSendDate = oMsg.TimeSent
                    
                    Call OpenDataBaseAndStoreDetailsY(strEmailFromAdd, strEmailToAdd, strEmailSubject, strEmailBody, EmailRecivedDate, EmailSendDate)
                           
                    
                    '// This will store all attechment files into the spacified path
                    If oMsg.Attachments.Count > 0 Then
                    '// Check if the directory name Attachment is alrady exit in application path
                        DirPath = App.Path & "\" & "Attachments"
                        DirName = Dir(DirPath, vbDirectory)   ' Retrieve the first entry.
                    
                        If DirName <> "" And bolTemp = True Then
                        
                            varMsg = Cap11   ' Define message.
                            varStyle = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
                            varTitle = App.Title    ' Define title.
                            ' Display message.
                            varResponse = MsgBox(varMsg, varStyle, varTitle)
                            If varResponse = vbYes Then   ' User chose Yes.
                                'Creates a new directory
                                MkDir App.Path & "\" & "NewAttachments"
                                bolTemp1 = False
                                Set oAttachments = oMsg.Attachments
                                For I = 1 To oAttachments.Count
                                    If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" Then
                                    oAttachments(I).WriteToFile (App.Path & "\NewAttachments\" & oAttachments(I))
                                    End If
                                Next I  ' Perform some action.
                            Else   ' User chose No.
                                    Kill App.Path & "\" & "Attachments\*.*"
                                Set oAttachments = oMsg.Attachments
                                    For I = 1 To oAttachments.Count
                                        If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" Then
                                        oAttachments(I).WriteToFile (App.Path & "\Attachments\" & oAttachments(I))
                                        End If
                                    Next I
                             End If
                             
                            bolTemp = False
                      Else
                            If bolTemp = True Then
                            'Creates a new directory
                                MkDir App.Path & "\" & "Attachments"
                                bolTemp = False
                            End If  ' Perform some action.
                            Set oAttachments = oMsg.Attachments
                               For I = 1 To oAttachments.Count
                                    If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" And oAttachments(I) <> "Untitled Attachment" Then
                                    If bolTemp1 = True Then
                                        oAttachments(I).WriteToFile (App.Path & "\Attachments\" & oAttachments(I))
                                    Else
                                        oAttachments(I).WriteToFile (App.Path & "\NewAttachments\" & oAttachments(I))
                                    End If
                                    End If
                                Next I
                    End If
                        
                End If
            End If
            Set oMsg = oMessages.GetNext
        Wend
        Screen.MousePointer = vbDefault ' delivery is now complete so let them know
    'Explicitly release objects.
        oSession.Logoff
        Set oSession = Nothing
    End If
    
    If optSelecteIndex = 1 Then
        bolTemp = True '// Only used to create attachment folder one time only
        bolTemp1 = True '// For NewEmailFolder and Email Folder Only
        '//Open your Exchange server and logon on it
        Set oSession = CreateObject("MAPI.Session")
        oSession.Logon , , False, False, 0 'Use the existing Outlook session.
        
        Set oFolder = oSession.Inbox
        Set oMessages = oFolder.Messages
        
        Set oMsg = oMessages.GetFirst
        Screen.MousePointer = vbHourglass ' the DeliverNow method could take a while right?
        oSession.DeliverNow ' this now gets all mail services sent and delivered just as the menu option, tools/deliver now/all services does
        '
        While Not oMsg Is Nothing ' If the message collection was empty, oMsg should be equal to the Nothing object, not "Is" operator takes object inputs
                  If oMsg.Unread = True Then
                    strEmailToAdd = oMsg.Sender.Address
                    strEmailFromAdd = oMsg.Sender.Address
                    strEmailBody = Trim(Replace(oMsg.Text, "'", """"))
                    strEmailSubject = Trim(Replace(oMsg.Subject, "'", "''"))
                    EmailRecivedDate = oMsg.TimeReceived
                    EmailSendDate = oMsg.TimeSent
                    
                    Call OpenDataBaseAndStoreDetailsY(strEmailFromAdd, strEmailToAdd, strEmailSubject, strEmailBody, EmailRecivedDate, EmailSendDate)
                           
                    
                    '// This will store all attechment files into the spacified path
                    If oMsg.Attachments.Count > 0 Then
                    '// Check if the directory name Attachment is alrady exit in application path
                        DirPath = App.Path & "\" & "Attachments"
                        DirName = Dir(DirPath, vbDirectory)   ' Retrieve the first entry.
                    
                        If DirName <> "" And bolTemp = True Then
                        
                            varMsg = Cap11   ' Define message.
                            varStyle = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
                            varTitle = App.Title    ' Define title.
                            ' Display message.
                            varResponse = MsgBox(varMsg, varStyle, varTitle)
                            If varResponse = vbYes Then   ' User chose Yes.
                                'Creates a new directory
                                MkDir App.Path & "\" & "NewAttachments"
                                bolTemp1 = False
                                Set oAttachments = oMsg.Attachments
                                For I = 1 To oAttachments.Count
                                    If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" Then
                                    oAttachments(I).WriteToFile (App.Path & "\NewAttachments\" & oAttachments(I))
                                    End If
                                Next I  ' Perform some action.
                            Else   ' User chose No.
                                    Kill App.Path & "\" & "Attachments\*.*"
                                Set oAttachments = oMsg.Attachments
                                    For I = 1 To oAttachments.Count
                                        If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" Then
                                        oAttachments(I).WriteToFile (App.Path & "\Attachments\" & oAttachments(I))
                                        End If
                                    Next I
                             End If
                             
                            bolTemp = False
                      Else
                            If bolTemp = True Then
                            'Creates a new directory
                                MkDir App.Path & "\" & "Attachments"
                                bolTemp = False
                            End If  ' Perform some action.
                            Set oAttachments = oMsg.Attachments
                               For I = 1 To oAttachments.Count
                                    If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" And oAttachments(I) <> "Untitled Attachment" Then
                                    If bolTemp1 = True Then
                                        oAttachments(I).WriteToFile (App.Path & "\Attachments\" & oAttachments(I))
                                    Else
                                        oAttachments(I).WriteToFile (App.Path & "\NewAttachments\" & oAttachments(I))
                                    End If
                                    End If
                                Next I
                    End If
                        
                End If
                '//***You can Uncomment this 2 line code if you want to set unread messages
                '//***to read messages After store in database
                'oMsg.Unread = False '// Please set the status of the unread email as false
                'oMsg.Update '//Make sure we won't read same message twice
                '//*******
            End If
            Set oMsg = oMessages.GetNext
        Wend
        Screen.MousePointer = vbDefault ' delivery is now complete so let them know
    'Explicitly release objects.
        oSession.Logoff
        Set oSession = Nothing
  End If
    If optSelecteIndex = 2 Then
        bolTemp = True '// Only used to create attachment folder one time only
        bolTemp1 = True '// For NewEmailFolder and Email Folder Only
       '//Open your Exchange server and logon on it
        Set oSession = CreateObject("MAPI.Session")
        oSession.Logon , , False, False, 0 'Use the existing Outlook session.
        
        Set oFolder = oSession.Inbox
        Set oMessages = oFolder.Messages
        
        Set oMsg = oMessages.GetFirst
        Screen.MousePointer = vbHourglass ' the DeliverNow method could take a while right?
        oSession.DeliverNow ' this now gets all mail services sent and delivered just as the menu option, tools/deliver now/all services does
        '
        While Not oMsg Is Nothing ' If the message collection was empty, oMsg should be equal to the Nothing object, not "Is" operator takes object inputs
                    strEmailToAdd = oMsg.Sender.Address
                    strEmailFromAdd = oMsg.Sender.Address
                    strEmailBody = Trim(Replace(oMsg.Text, "'", """"))
                    strEmailSubject = Trim(Replace(oMsg.Subject, "'", "''"))
                    EmailRecivedDate = oMsg.TimeReceived
                    EmailSendDate = oMsg.TimeSent
                    
                    Call OpenDataBaseAndStoreDetailsY(strEmailFromAdd, strEmailToAdd, strEmailSubject, strEmailBody, EmailRecivedDate, EmailSendDate)
                           
                    
                    '// This will store all attechment files into the spacified path
                    If oMsg.Attachments.Count > 0 Then
                    '// Check if the directory name Attachment is alrady exit in application path
                        DirPath = App.Path & "\" & "Attachments"
                        DirName = Dir(DirPath, vbDirectory)   ' Retrieve the first entry.
                    
                        If DirName <> "" And bolTemp = True Then
                        
                            varMsg = Cap11   ' Define message.
                            varStyle = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
                            varTitle = App.Title    ' Define title.
                            ' Display message.
                            varResponse = MsgBox(varMsg, varStyle, varTitle)
                            If varResponse = vbYes Then   ' User chose Yes.
                                'Creates a new directory
                                MkDir App.Path & "\" & "NewAttachments"
                                bolTemp1 = False
                                Set oAttachments = oMsg.Attachments
                                For I = 1 To oAttachments.Count
                                    If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" Then
                                    oAttachments(I).WriteToFile (App.Path & "\NewAttachments\" & oAttachments(I))
                                    End If
                                Next I  ' Perform some action.
                            Else   ' User chose No.
                                    Kill App.Path & "\" & "Attachments\*.*"
                                Set oAttachments = oMsg.Attachments
                                    For I = 1 To oAttachments.Count
                                        If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" Then
                                        oAttachments(I).WriteToFile (App.Path & "\Attachments\" & oAttachments(I))
                                        End If
                                    Next I
                             End If
                             
                            bolTemp = False
                      Else
                            If bolTemp = True Then
                            'Creates a new directory
                                MkDir App.Path & "\" & "Attachments"
                                bolTemp = False
                            End If  ' Perform some action.
                            Set oAttachments = oMsg.Attachments
                               For I = 1 To oAttachments.Count
                                    If Not IsNull(oAttachments(I)) And oAttachments(I) <> "" And oAttachments(I) <> "Untitled Attachment" Then
                                    If bolTemp1 = True Then
                                        oAttachments(I).WriteToFile (App.Path & "\Attachments\" & oAttachments(I))
                                    Else
                                        oAttachments(I).WriteToFile (App.Path & "\NewAttachments\" & oAttachments(I))
                                    End If
                                    End If
                                Next I
                    End If
                        
                End If
            Set oMsg = oMessages.GetNext
        Wend
        Screen.MousePointer = vbDefault ' delivery is now complete so let them know
    'Explicitly release objects.
        oSession.Logoff
        Set oSession = Nothing
    End If
 Case 1:
    Unload Me '// Unload Form.
    End '// End Application
 Case 2:
    strEmailToAdd = InputBox(Cap12, Cap1)
    strEmailSubject = InputBox(Cap13, Cap1)
    strEmailBody = InputBox(Cap14, Cap1)
    Call SendMail(strEmailToAdd, strEmailBody, strEmailSubject)
 End Select
End Sub
Private Sub Form_Load()
      On Error Resume Next
    '// Define Captions Only
    frmOutlookEmail.Caption = Cap1 '// Form Caption
    frmOutlookEmail.BorderStyle = 1
    Frame1.Caption = Cap6 '// For Frame Caption
    lblReadEmail(0).AutoSize = True
    lblReadEmail(0).ForeColor = vbRed
    lblReadEmail(1).ForeColor = vbRed
    lblReadEmail(0).Caption = Cap7 '// For Lable Caption
    lblReadEmail(1).Caption = Cap8 '// For Lable Caption
    optSelect(0).Caption = Cap2 '// Option Button caption
    optSelect(1).Caption = Cap3 '// Option Button caption
    optSelect(2).Caption = Cap4 '// Option Button caption
    cmdSaveAndExit(0).Caption = Cap5 '// Command Button caption
    cmdSaveAndExit(1).Caption = Cap10 '// Command Button caption
    cmdSaveAndExit(2).Caption = Cap9 '// Command Button caption
    DataBaseFileName = Trim("EMailDetails.mdb") '// Database File Name
    
    Call CreateAccessDatabaseX '// Create MicroSoft Access Database
    
End Sub

Sub SendMail(tTo As String, tBody As String, tSubject As String)
    On Error Resume Next
'***********************
'Description: Uses the outlook object to create and send a mail using the passed
 'parameters of tsubject,tbody and tTo
'***********************
    Dim osendFolder As Object
    Dim osendMessages As Object
    Dim osendMsg As Object
    Dim osendRcpt As Object
    
    '//Open your Exchange server and logon on it
    Set oSession = CreateObject("MAPI.Session")
    oSession.Logon , , False, False, 0 'Use the existing Outlook session.

    Set osendFolder = oSession.Outbox
    Set osendMessages = osendFolder.Messages
    Set osendMsg = osendMessages.Add
    osendMsg.Subject = tSubject
    osendMsg.Text = tBody
    Set osendRcpt = osendMsg.Recipients
    osendRcpt.Add , "SMTP:" & tTo
    osendRcpt.Resolve
    osendMsg.Send
    
    
    Set osendRcpt = Nothing
    Set osendMsg = Nothing
    Set osendMessages = Nothing
    Set osendFolder = Nothing
    oSession.Logoff
    Set oSession = Nothing
End Sub
Sub CreateAccessDatabaseX()
      On Error Resume Next
'//This function creates the database into the curent application path
    Dim objDAO As Object
    Dim objdbsNew As Object
    Dim objTableDef As Object
    Dim objIndex As Object
    Dim objFields As Object
    
    Set objDAO = CreateObject("DAO.DBEngine.35")
    
    ' Make sure there isn't already a file with the name of
    ' the new database.
    If Dir(App.Path & "\" & DataBaseFileName) <> "" Then Kill (App.Path & "\" & DataBaseFileName)

    ' Create a new encrypted database with the specified
    ' collating order.
     Set objdbsNew = objDAO.CreateDatabase(App.Path & "\" & DataBaseFileName, ";LANGID=0x0409;CP=1252;COUNTRY=0")
     Set objTableDef = objdbsNew.CreateTableDef("EMailDetails")
     
     Set objIndex = objTableDef.CreateIndex("RefIDX")
            objIndex.Primary = True 'Set primary key
            objIndex.Fields.Append objIndex.CreateField("SrNo") 'Trust me, this works, weird as it is.
            objTableDef.Indexes.Append objIndex
            
     Set objFields = objTableDef.CreateField("SrNo", 4, 4)
            objFields.Attributes = 49 'Make it a counter.
            objTableDef.Fields.Append objFields
     Set objFields = Nothing
     
     Set objFields = objTableDef.CreateField("From Address", 10, 255)
            objTableDef.Fields.Append objFields
     Set objFields = Nothing

     Set objFields = objTableDef.CreateField("To Address", 10, 255)
            objTableDef.Fields.Append objFields
     Set objFields = Nothing

     Set objFields = objTableDef.CreateField("Subject", 10, 255)
            objTableDef.Fields.Append objFields
     Set objFields = Nothing

     Set objFields = objTableDef.CreateField("Body", 12, 65535)
            objTableDef.Fields.Append objFields
     Set objFields = Nothing

     Set objFields = objTableDef.CreateField("Received DateTime", 8, 8)
            objTableDef.Fields.Append objFields
     Set objFields = Nothing
     
     Set objFields = objTableDef.CreateField("Send DateTime", 8, 8)
            objTableDef.Fields.Append objFields
     Set objFields = Nothing
     
            objdbsNew.TableDefs.Append objTableDef
            
        objdbsNew.Close
        Set objFields = Nothing
        Set objIndex = Nothing
        Set objTableDef = Nothing
        Set objdbsNew = Nothing
        Set objDAO = Nothing
        
End Sub

Private Sub optSelect_Click(Index As Integer)
    optSelecteIndex = Index
End Sub

Private Sub OpenDataBaseAndStoreDetailsY(strFromAddress As String, strToAddress As String, strSubject As String, strBody As String, RecivedDataTime As Date, SendDateTime As Date)
       On Error Resume Next
'// This will Open DataBase And Store the Values InTo It
    Dim objDAO As Object
    Dim objdbsNew As Object
    Dim rs As Object
    Set objDAO = CreateObject("DAO.DBEngine.35")
    Set objdbsNew = objDAO.OpenDatabase(App.Path & "\" & DataBaseFileName)
   
    Set rs = objdbsNew.OpenRecordset("EMailDetails")
    rs.AddNew
    
    If Not IsNull(strFromAddress) Then
        rs![From Address] = strFromAddress
    Else
        rs![From Address] = Null
    End If
    
    If Not IsNull(strToAddress) Then
        rs![To Address] = strToAddress
    Else
        rs![To Address] = Null
    End If
    
    If Not IsNull(strSubject) And Trim(strSubject) <> "" Then
        rs![Subject] = strSubject
    Else
        rs![Subject] = Null
    End If
    
    If Not IsNull(strBody) And Trim(strBody) <> "" Then
        rs![Body] = strBody
    Else
        rs![Body] = Null
    End If
    If Not IsNull(RecivedDataTime) Then
        rs![Received DateTime] = RecivedDataTime
    Else
       rs![Received DateTime] = Null
    End If
    If Not IsNull(SendDateTime) Then
        rs![Send DateTime] = SendDateTime
    Else
        rs![Send DateTime] = Null
    End If
    rs.Update

    rs.Close
    Set rs = Nothing
    objdbsNew.Close
    Set objdbsNew = Nothing
    Set objDAO = Nothing

End Sub
