VERSION 5.00
Begin VB.Form startup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   Icon            =   "startup.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   1380
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   240
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "e-mail: jpejakovic@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   3225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Author: Josip Pejakoviæ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   0
      Picture         =   "startup.frx":030A
      Top             =   0
      Width           =   4170
   End
End
Attribute VB_Name = "startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Static StartTime As Single
Dim TimeNow As Single

If StartTime = 0 Then

StartTime = Time

End If
TimeNow = Time
If Format(TimeNow - StartTime, "ss") >= 1 Then
Unload Me
Form1.Show , Me
Timer1.Enabled = False
End If
End Sub
