VERSION 5.00
Begin VB.Form MP3Player 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mini MP3 Player"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3975
      Begin VB.ListBox List1 
         Height          =   3960
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton Command5 
         Caption         =   "Back"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "UnPause"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Pause"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Play"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4440
      Top             =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   585
   End
End
Attribute VB_Name = "MP3Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MP3Play.MP3Play (putanja)
End Sub

Private Sub Command2_Click()
MP3Play.StopMp3
End Sub

Private Sub Command3_Click()
MP3Play.PauseMp3
Command3.Visible = False
Command4.Visible = True
End Sub

Private Sub Command4_Click()
MP3Play.UnPauseMp3
Command3.Visible = True
Command4.Visible = False
End Sub

Private Sub Command5_Click()
MP3Play.Back
End Sub

Private Sub Form_Load()
Command4.Visible = False
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
MP3Play.MP3Stop
End Sub

Private Sub Timer1_Timer()
If MP3Play.IsPlaying = True Then
Label1 = "Playing... "
Else
Label1 = "Stoped..."
End If
End Sub
