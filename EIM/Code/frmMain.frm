VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EIM"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDATE 
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Michael Anthony Schmidt July 2000"
      Top             =   6300
      Width           =   5235
   End
   Begin VB.Frame Frame1 
      Caption         =   "Additional Information"
      Height          =   3795
      Left            =   0
      TabIndex        =   1
      Top             =   1500
      Width           =   5235
      Begin VB.Label Label10 
         Caption         =   "•"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   2880
         Width           =   195
      End
      Begin VB.Label Label9 
         Caption         =   $"frmMain.frx":0442
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   420
         TabIndex        =   11
         Top             =   2880
         Width           =   4740
      End
      Begin VB.Label Label8 
         Caption         =   $"frmMain.frx":051F
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   420
         TabIndex        =   9
         Top             =   2160
         Width           =   4740
      End
      Begin VB.Label Label6 
         Caption         =   "•"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   2160
         Width           =   195
      End
      Begin VB.Label Label5 
         Caption         =   $"frmMain.frx":05C3
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   420
         TabIndex        =   7
         Top             =   1440
         Width           =   4740
      End
      Begin VB.Label Label7 
         Caption         =   "•"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   1380
         Width           =   195
      End
      Begin VB.Label Label2 
         Caption         =   "See Load Form under frmEServer. Change lbluser to equivalent of winsock local host name."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   480
         TabIndex        =   5
         Top             =   900
         Width           =   4665
      End
      Begin VB.Label Label1 
         Caption         =   "•"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label3 
         Caption         =   $"frmMain.frx":0687
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   4680
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Message"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "frmMain.frx":0719
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   5295
   End
   Begin VB.Label Label11 
      Caption         =   $"frmMain.frx":0B5B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   5160
   End
   Begin VB.Label Label4 
      BackColor       =   &H00DBFDE3&
      BackStyle       =   0  'Transparent
      Caption         =   "Electronic Instant Messaging"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   180
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   3300
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   3300
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      X1              =   1200
      X2              =   3300
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line4 
      X1              =   1320
      X2              =   3300
      Y1              =   540
      Y2              =   540
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendEMessage ' Send Message
End Sub

Private Sub Form_Load()
StartEServer ' Start Server
End Sub

