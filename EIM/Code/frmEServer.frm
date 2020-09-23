VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Tracker - Inbox"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "frmEServer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeleteMarked 
      Caption         =   "&Delete Marked"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   3660
      Width           =   1395
   End
   Begin VB.CommandButton cmdSendMessage 
      Height          =   615
      Left            =   60
      Picture         =   "frmEServer.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Send Message"
      Top             =   60
      Width           =   675
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "&Hide"
      Default         =   -1  'True
      Height          =   375
      Left            =   4740
      TabIndex        =   4
      Top             =   3660
      Width           =   1395
   End
   Begin VB.CheckBox chkSound 
      Caption         =   "Enable Sound"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4140
      TabIndex        =   3
      Top             =   420
      Width           =   1335
   End
   Begin VB.CheckBox chkPopup 
      Caption         =   "Enable Popup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4140
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
   Begin MSWinsockLib.Winsock sckSYS 
      Left            =   60
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSComctlLib.ListView lvwInbox 
      Height          =   2895
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgTray"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Machine"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Subject"
         Object.Width           =   4454
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Message"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Address"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Timer NewMessage 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   480
      Top             =   4080
   End
   Begin MSComctlLib.ImageList imgTray 
      Left            =   900
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEServer.frx":044E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEServer.frx":08A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEServer.frx":0CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEServer.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEServer.frx":15A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblUSER 
      BackStyle       =   0  'Transparent
      Caption         =   "mschmidt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   180
      Width           =   1935
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuShowInbox 
         Caption         =   "Inbox"
      End
   End
End
Attribute VB_Name = "frmEServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myTrayIcon As New clsTray
Dim bOn As Boolean
Dim tIcon As Integer
Dim itmX As ListItem
Dim NewMsgIndex As Integer


'====================================
'   Minimize To Tray
'====================================
Private Sub cmdDeleteMarked_Click()
Dim X, Y As Integer
Dim Marked As Boolean

Do
    Marked = False ' We have no marked boxes...
    For X = 1 To lvwInbox.ListItems.Count
        If lvwInbox.ListItems(X).Checked = True Then
            Marked = True ' We have a marked box...
            Y = X         ' Set index to that box for delete...
        End If
    Next X
    If Marked = True Then lvwInbox.ListItems.Remove (Y)
Loop Until Marked = False


End Sub


'====================================
'   Minimize To Tray
'====================================
Private Sub cmdMinimize_Click()
Me.Hide
End Sub


'====================================
'   Send Message
'====================================
Private Sub cmdSendMessage_Click()
    SendEMessage
End Sub

'====================================
'   Form Load
'====================================
Private Sub Form_Load()

 ' Initialize Tray Icon
 tIcon = 1
 myTrayIcon.RemoveIcon Me
 myTrayIcon.ShowIcon Me
 myTrayIcon.ChangeIcon Me, imgTray.ListImages(tIcon)
 myTrayIcon.ChangeToolTip Me, "List Tracker"
 Me.Hide
 lblUSER = GetUser
 LoadMessages

 
End Sub


'====================================
'   Read Message (Pass Index)
'====================================
Private Sub ReadMessage(Index As Integer)
Dim eRead As New frmERead
On Error Resume Next
    ' Initialize Read Form
    eRead.txtDATE = lvwInbox.ListItems(Index).Text & " " & lvwInbox.ListItems(Index).SubItems(1)
    eRead.txtFROM = lvwInbox.ListItems(Index).SubItems(2)
    eRead.txtSUBJECT = lvwInbox.ListItems(Index).SubItems(3)
    eRead.txtMessage = lvwInbox.ListItems(Index).SubItems(4)
    eRead.Show
    
    ' Mark message as read
    lvwInbox.ListItems(Index).SmallIcon = 4

End Sub


'====================================
'   Double Click Message
'====================================
Private Sub lvwInbox_DblClick()
Dim eRead As New frmERead
    ' If listview not empty...
    If lvwInbox.ListItems.Count <> 0 Then ReadMessage lvwInbox.SelectedItem.Index
End Sub


'====================================
'   Click Sort
'====================================
Private Sub lvwInbox_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ' When a ColumnHeader object is clicked, the ListView control is
   ' sorted by the subitems of that column.
   ' Set the SortKey to the Index of the ColumnHeader - 1
   lvwInbox.SortKey = ColumnHeader.Index - 1
   ' Set Sorted to True to sort the list.
   lvwInbox.Sorted = True
End Sub


'====================================
'   Unload Event
'====================================
Private Sub Form_Unload(Cancel As Integer)
Dim Msg, Response   ' Declare variables.

    ' Set Message And Display
    Msg = ("Are you sure you wish to quit?" & vbCrLf & _
    "You will no longer receive messages!")
    Response = MsgBox(Msg, vbExclamation + vbYesNo)
    
    ' Exit & Save or Cancel
    If Response = vbYes Then
        SaveMessages              ' Save Messages
        myTrayIcon.RemoveIcon Me  ' Remove Icon
    Else
        Cancel = -1
    End If
   
End Sub


'====================================
'   Click Sort
'====================================
Private Sub mnuShowInbox_Click()
    Me.Show
End Sub


'====================================
'   sckYS Data Arrival
'====================================
Private Sub sckSYS_DataArrival(ByVal bytesTotal As Long)
Dim pData As String
Dim pCom  As String

    
    sckSYS.GetData pData                 ' Pull Packet From Buffer
    Debug.Print pData                    ' Debug Print Data
    pCom = Word(pData, 1, Chr(1))        ' Parse Com
    pData = DelWord(pData, 1, Chr(1), 1) ' Parse Data
 
    Select Case pCom
     Case "020": '-------------------------------------- Incomming Message
                 AddMessage (pData)
     Case "030": '-------------------------------------- Hello Packet
                 ReplyHello (pData)
    End Select
    
End Sub

'====================================
'   Reply Hello
'====================================
Private Sub ReplyHello(pData As String)
Dim pAddress As String
Dim pPort    As String
Dim pPacket  As String

    pPacket = "040" & Chr(1) & lblUSER.Caption & _
                      Chr(1) & sckSYS.LocalIP & Chr(1)

    pAddress = Word(pData, 1, Chr(1))
    pPort = Word(pData, 2, Chr(1))
    
    UDPPACKET sckSYS, pPort, pAddress, pPacket
End Sub

'====================================
'   Add Message (Inbox)
'====================================
Private Sub AddMessage(pData As String)
Dim pMachine As String
Dim pAddress As String
Dim pMessage As String
Dim pTime    As String
Dim pSubject As String
Dim pPort    As String
Dim pDate    As String

    ' Parse Data
    pAddress = Word(pData, 1, Chr(1)) ' Remote Address
    pMachine = Word(pData, 2, Chr(1)) ' Remote Machine
    pPort = Word(pData, 3, Chr(1))    ' Remote Port
    pSubject = Word(pData, 4, Chr(1)) ' List ID
    pMessage = Word(pData, 5, Chr(1)) ' Message

    ' Add Message Data To Inbox
    Set itmX = lvwInbox.ListItems.Add(, , Time, , 2)
    itmX.SubItems(1) = Date
    itmX.SubItems(2) = pMachine
    itmX.SubItems(3) = pSubject
    itmX.SubItems(4) = pMessage
    itmX.SubItems(5) = pAddress
    
    ' Confirm Received
    UDPPACKET sckSYS, pPort, pAddress, "010" & Chr(1) ' Send Packet
    NewMessage = True
    MessageAlert ' Beep or popup
    SaveMessages ' Save All Messages

End Sub


'====================================
'   Tray Click
'====================================
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '7680   ' MouseMove
    '7695   ' Left MouseDown
    '7710   ' Left MouseUp
    '7725   ' Left DoubleClick
    '7740   ' Right MouseDown
    '7755   ' Right MouseUp
    '7770   ' Right DoubleClick
If myTrayIcon.bRunningInTray Then           'Check to see if form is in the system tray
    Select Case X                           'If it is, use X to get message value
        Case 7725:  ' Double Click - Stop Blinking Pull Up Server
                    If NewMessage = True Then
                     ReadMessage (NewMsgIndex)
                    Else
                     Me.Show
                    End If
        Case 7755:  PopupMenu Me.mnuTray, vbPopupMenuRightButton  'Show a menubar
    End Select
End If
End Sub


'====================================
'   Message Alert
'====================================
Private Sub MessageAlert()
    If chkSound.Value = Checked Then Beep
    If chkPopup.Value = Checked Then ReadMessage (lvwInbox.ListItems.Count)
End Sub


'====================================
'   Incoming Message (Timer)
'====================================
' When new message arrives, we must
' blink in tray, when user reads
' a message, we must decide if we
' continue to blink based on the
' unread messages which are 'new'
'
' If we find no unread messages
' then we stop blinking.
'====================================
Private Sub NewMessage_Timer()
Dim X      As Integer
Dim Unread As Boolean

    ' Default False
    Unread = False
    
    ' Loop Through Inbox Find Unread
    For X = 1 To lvwInbox.ListItems.Count
        If lvwInbox.ListItems(X).SmallIcon <> 4 Then
            If Unread = False Then NewMsgIndex = X
            ' The above line grabs the first message in the list
            ' we know this because unread is still false...
            Unread = True
        End If
    Next
    
    ' If Unread Messages Blink, Else No Blink
    
    If Unread = True Then
          tIcon = 5
          
          If bOn = True Then
            tIcon = 3 ' blank
            bOn = False
          Else
          
          bOn = True
    End If
        myTrayIcon.ChangeIcon Me, imgTray.ListImages(tIcon)
    Else
        NewMessage = False
        tIcon = 1
        myTrayIcon.ChangeIcon Me, imgTray.ListImages(tIcon)
    End If
    
End Sub

Private Sub SaveMessages()
Dim X As Integer
Dim rKey As String ' KeyName
Dim rDat As String ' MessageData


    ' Key (Users Name)
    rKey = lblUSER.Caption
    
    ' Save Default Settings
    SaveSetting App.ProductName, rKey & "EIM", "Popup", chkPopup.Value
    SaveSetting App.ProductName, rKey & "EIM", "Sound", chkSound.Value
    
    ' Delete Previous Key (Messages)
    SaveSetting App.ProductName, rKey, "INI", "DELETEME"
    DeleteSetting App.ProductName, rKey
    ' Loop through messages and save
    For X = 1 To lvwInbox.ListItems.Count
     ' Save Message In Packet Form To String
     rDat = lvwInbox.ListItems(X).Text & Chr(1) & _
            lvwInbox.ListItems(X).SubItems(1) & Chr(1) & _
            lvwInbox.ListItems(X).SubItems(2) & Chr(1) & _
            lvwInbox.ListItems(X).SubItems(3) & Chr(1) & _
            lvwInbox.ListItems(X).SubItems(4) & Chr(1) & _
            lvwInbox.ListItems(X).SubItems(5) & Chr(1) & _
            lvwInbox.ListItems(X).SmallIcon & Chr(1)
     SaveSetting App.ProductName, rKey, X, rDat
    Next
    
End Sub

Private Sub LoadMessages()
Dim X As Integer
Dim rKey As String
Dim rDat As String
Dim rAry As Variant
    Dim lTime    As String
    Dim lDate    As String
    Dim lMachine As String
    Dim lSubject As String
    Dim lMessage As String
    Dim lAddress As String
    Dim lIcon    As Integer
On Error GoTo ErrorHandler
' The problem I get here, is rARY is a variant,
' where registry items are stored. If the registry
' is empty, then we want to exit the application.
' We can't do a comparison on rAry though because
' if the registry is empty or full, the data type
' is different:
' RegEmpty) rAry = Variant (Uknown)
' RegStuff) rAry = 2-Dim Array
' And you can't compare the two by saying:
' if rAry = "" or vbNull (etc) because you can't
' compare a 2-dim array to those types!

    ' Clear ListBox
    lvwInbox.ListItems.Clear

    ' Key (Users Name)
    rKey = lblUSER.Caption

    ' Load Defaults
    chkPopup.Value = GetSetting(App.ProductName, rKey & "EIM", "Popup", False)
    chkSound.Value = GetSetting(App.ProductName, rKey & "EIM", "Sound", False)

    ' Pull Registry
    rAry = GetAllSettings(App.ProductName, rKey)

    ' Populate ListBox
    For X = LBound(rAry, 1) To UBound(rAry, 1)
      rDat = rAry(X, 1)
      lTime = Word(rDat, 1, Chr(1))
      lDate = Word(rDat, 2, Chr(1))
      lMachine = Word(rDat, 3, Chr(1))
      lSubject = Word(rDat, 4, Chr(1))
      lMessage = Word(rDat, 5, Chr(1))
      lAddress = Word(rDat, 6, Chr(1))
      lIcon = Word(rDat, 7, Chr(1))
       Set itmX = lvwInbox.ListItems.Add(, , lTime, , lIcon)
       itmX.SubItems(1) = lDate
       itmX.SubItems(2) = lMachine
       itmX.SubItems(3) = lSubject
       itmX.SubItems(4) = lMessage
       itmX.SubItems(5) = lAddress
    Next X

Exit Sub      ' Exit to avoid handler.
ErrorHandler:   ' Error-handling routine.
    If Err.Number = 13 Then Exit Sub
   Resume   ' Resume execution at same line
End Sub


