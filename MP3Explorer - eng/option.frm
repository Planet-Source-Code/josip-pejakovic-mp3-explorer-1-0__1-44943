VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   310
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path to WinAMP:"
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sFilter = "WINAMP" & vbNullChar & "winamp.exe"
aa = GetOpenFilePath(hWnd, sFilter, 0, "winamp", "", "Putanja do Winampa", sPath) = True
Text1 = sPath
End Sub

Private Sub Command2_Click()
If Text1 = "" Then
Unload Me
Else
Open App.path & "\conf.dat" For Output As #1
Print #1, Text1
Close #1
Unload Me
End If
End Sub
