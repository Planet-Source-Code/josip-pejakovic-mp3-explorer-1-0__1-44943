VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rename"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5010
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Rename"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New name:"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current filename:"
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1425
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2 = "" Then
MsgBox ("Type new name of file!"), vbCritical
Text2.SetFocus
Exit Sub
End If

potvrda = MsgBox("Do you want rename file " & Text1 & "?", vbYesNo)
If potvrda = vbYes Then
SetFileAttributes Form1.lv.SelectedItem.SubItems(1), &H80
novinaziv = trenutni_dir & "\" & Text2 + ".mp3"
MoveFile Form1.lv.SelectedItem.SubItems(1), novinaziv
Unload Me
Form1.skeniraj
editirano = True
End If

If potvrda = vbNo Then
Exit Sub
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
