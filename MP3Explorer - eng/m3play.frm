VERSION 5.00
Begin VB.Form Form9 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MP3 Explorer - Player"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6060
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   600
      Picture         =   "m3play.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Pauziraj pjesmu"
      Top             =   960
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   1080
      Picture         =   "m3play.frx":020A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Zaustavi pjesmu"
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   600
      Picture         =   "m3play.frx":0414
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Pauziraj pjesmu"
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   120
      Picture         =   "m3play.frx":061E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sviraj pjesmu"
      Top             =   960
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0080FF80&
      Height          =   735
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pauza As Boolean
Private Sub Command1_Click()
MP3Play.MP3Play (Label2)
Label4 = "Duration: " & MP3Play.TrackLength
End Sub

Private Sub Command2_Click()
Command4.Visible = True
MP3Play.PauseMp3
pauza = True
Command2.Visible = False
End Sub

Private Sub Command3_Click()
MP3Play.StopMp3
End Sub

Private Sub Command4_Click()
If pauza = True Then
MP3Play.UnPauseMp3
Command4.Visible = False
Command2.Visible = True
End If
End Sub

Private Sub Form_Load()
MP3Play.StopMp3
End Sub

Private Sub Form_Unload(Cancel As Integer)
MP3Play.StopMp3
End Sub
