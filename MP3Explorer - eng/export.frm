VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export list..."
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Options"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Export in HTML"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3735
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Export Artist && Song Title"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Export Song Title"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Export path && &filename"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
sFilter = "Text Files (*.txt)" & vbNullChar & "*.txt"
aa = GetSaveFilePath(hWnd, sFilter, 0, "txt", "", "", "Saving", sPath)
If aa = False Then Exit Sub
If Dir(sPath) <> "" Then Kill sPath


If Form4.Visible = True Then
Module1.Export_Path_Filename sPath, Form4.lv
MsgBox ("The list is saved in " & UCase(sPath) & vbCrLf & "You can open it with any text editor!"), vbInformation
Else
Module1.Export_Path_Filename sPath, Form1.lv
MsgBox ("The list is saved in " & UCase(sPath) & vbCrLf & "You can open it with any text editor!"), vbInformation
End If
End Sub

Private Sub Command3_Click()
sFilter = "Text Files (*.txt)" & vbNullChar & "*.txt"
aa = GetSaveFilePath(hWnd, sFilter, 0, "txt", "", "", "Saving", sPath)
If aa = False Then Exit Sub
If Dir(sPath) <> "" Then Kill sPath


If Form4.Visible = True Then
Module1.Export_Song_Title sPath, Form4.lv
MsgBox ("The list is saved in " & UCase(sPath) & vbCrLf & "You can open it with any text editor!"), vbInformation
Else
Module1.Export_Song_Title sPath, Form1.lv
MsgBox ("The list is saved in " & UCase(sPath) & vbCrLf & "You can open it with any text editor!"), vbInformation
End If

End Sub

Private Sub Command4_Click()
sFilter = "Text Files (*.txt)" & vbNullChar & "*.txt"
'sFilter = "Text Files (*.txt)" & vbNullChar & "*.txt" & vbNullChar & "MP3 Organizer (*.mpo)" & vbNullChar & "*.mpo"
aa = GetSaveFilePath(hWnd, sFilter, 0, "txt", "", "", "Saving", sPath)

'If ofn.nFilterIndex = 2 Then
'MsgBox ("2")
'End If

If aa = False Then Exit Sub
If Dir(sPath) <> "" Then Kill sPath

If Form4.Visible = True Then
Module1.Export_SongTitle_Artist sPath, Form4.lv
MsgBox ("The list is saved in " & UCase(sPath) & vbCrLf & "You can open it with any text editor!"), vbInformation
Else
Module1.Export_SongTitle_Artist sPath, Form1.lv
MsgBox ("The list is saved in " & UCase(sPath) & vbCrLf & "You can open it with any text editor!"), vbInformation
End If

'MsgBox ofn.nFilterIndex

End Sub

Private Sub Command5_Click()
Form11.Show , Me
End Sub
