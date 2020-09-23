VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Tracker - Send Message"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   Icon            =   "frmEMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   540
      Top             =   3420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSComctlLib.ImageCombo imcNetwork 
      Height          =   330
      Left            =   1440
      TabIndex        =   7
      Top             =   60
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "imgNetwork"
   End
   Begin MSComctlLib.ImageList imgNetwork 
      Left            =   1380
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEMessage.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   960
      Top             =   3420
   End
   Begin VB.TextBox txtPORT 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "115"
      Top             =   3420
      Width           =   615
   End
   Begin MSWinsockLib.Winsock sckSM 
      Left            =   120
      Top             =   3420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton cmdSendIM 
      Caption         =   "&Send"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2160
      Width           =   1035
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmEMessage.frx":045E
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox txtLOCAL 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "10.0.0.240 Godzilla"
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtSUBJECT 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Text            =   "List Order"
      Top             =   420
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmEMessage.frx":046A
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblToFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   420
      Width           =   600
   End
End
Attribute VB_Name = "frmEMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MessageSentOK As Boolean
'=============================================================
' How To Call This Form
' Dim msgSend As New frmSendMessage
' msgSend.txtLISTID.Text = (DBASE LIST)
' msgSend.txtREMOTE.Text = (REMOTEIP & CHR(1) & REMOTENAME)
' msgSend.txtPort.Text = (PORT REMOTE LISTENS)
'=============================================================






'====================================
'   Build Packet And Send
'====================================
Private Sub cmdSendIM_Click()
Dim pPacket  As String
Dim pRemote  As String
Dim pAddress As String

    ' Have to have a valid address!
    If imcNetwork.Text = "" Then Exit Sub

    ' Build Packet
    ' mschmidt(10.0.0.0) parse to get ip...
    pAddress = Word(imcNetwork.Text, 2, " ")
    pPacket = "020" & Chr(1) & sckSM.LocalIP & Chr(1) & txtLOCAL.Text & Chr(1) & sckSM.LocalPort & Chr(1) ' Server Stuff
    pPacket = pPacket & txtSUBJECT.Text & Chr(1) & txtMessage.Text & Chr(1)                               ' Info Data

    ' Send Data
    MessageSentOK = False                            ' Default Not Sent
    UDPPACKET sckSM, txtPORT.Text, pAddress, pPacket ' Send Packet
    cmdSendIM.Enabled = False ' Disable Button
    TimeOut.Enabled = True    ' Wait For Response From Remote Confirming Message...

End Sub





'====================================
'   Form Load
'====================================
Private Sub Form_Load()

    txtLOCAL.Text = GetUser
    TimeOut.Interval = 5000   ' Second Timeout
    
    
End Sub


'====================================
'   Winsock Error (sckSM) (sckListen)
'====================================
Private Sub sckListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 MsgBox "Error #" & Number & vbCrLf & Description, vbCritical, App.Title
End Sub

Private Sub sckSM_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 MsgBox "Error #" & Number & vbCrLf & Description, vbCritical, App.Title
End Sub


'====================================
'   Time Out (Timer)
'====================================
Private Sub TimeOut_Timer()
Dim pMessage As String

    pMessage = "Request timed out!" & vbCrLf & "Message not sent." & sckSM.LocalPort
    MsgBox pMessage, vbExclamation, App.Title
    Unload Me
    
End Sub


'====================================
'   SckSM Data Arrival
'====================================
Private Sub sckSM_DataArrival(ByVal bytesTotal As Long)
Dim pData As String
Dim pCom  As String
On Error GoTo ConReset
 
    sckSM.GetData pData   ' Pull Packet From Buffer
    pCom = Word(pData, 1, Chr(1)) ' Parse COM From Packet
    pData = DelWord(pData, 1, Chr(1), 1) ' Parse Data
    
 Select Case pCom
  Case "010": '-------------------------------------- Message Received
              TimeOut.Enabled = False: Unload Me
 End Select
 
 Exit Sub
 
ConReset:
    Select Case Err.Number
        Case 10054
            Resume Next
        Case Else
            MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description
            Resume Next
            
    End Select
    
End Sub

'====================================
'   Add User
'====================================
Private Sub AddNetwork(pData)
Dim pUser As String
Dim pAddr As String

pUser = Word(pData, 1, Chr(1))
pAddr = Word(pData, 2, Chr(1))
 imcNetwork.ComboItems.Add , , pUser & " " & pAddr & " ", 1
End Sub

'====================================
'   SckListen Data Arrival
'====================================
Private Sub sckListen_DataArrival(ByVal bytesTotal As Long)
Dim pData As String
Dim pCom  As String
 
    sckListen.GetData pData   ' Pull Packet From Buffer
    pCom = Word(pData, 1, Chr(1)) ' Parse COM From Packet
    pData = DelWord(pData, 1, Chr(1), 1) ' Parse Data
    
 Select Case pCom
  Case "040": '-------------------------------------- Add User
              AddNetwork (pData)
 End Select
    
End Sub
