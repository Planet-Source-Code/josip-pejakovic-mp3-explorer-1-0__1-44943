VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move file(s)..."
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5580
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Move"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copy"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4095
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2385
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Then
MsgBox ("You must select path where you will copy files!"), vbCritical
Text1.SetFocus
Exit Sub
End If
putanja = Text1

For ii2 = 0 To Form1.List1.ListCount - 1
If Dir(putanja & Form1.List2.List(ii2)) <> "" Then
Select Case MsgBox("File " & Form1.List2.List(ii2) & " already exist. Do you wan't to continue with copying??", vbYesNo)
Case vbYes
DeleteFile putanja & Form1.List2.List(ii2)
CopyFile Form1.List1.List(ii2), putanja & Form1.List2.List(ii2), 0
Case vbNo
End Select
Else
CopyFile Form1.List1.List(ii2), putanja & Form1.List2.List(ii2), 0
End If
Next ii2
End Sub
Private Sub Command2_Click()
Form1.List1.Clear
Form1.List2.Clear
Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Text1 = "" Then
MsgBox ("You must select path where you will move files!"), vbCritical
Text1.SetFocus
Exit Sub
End If
putanja = Text1

'MsgBox Form1.List1.ListCount

For mv = 0 To Form1.List1.ListCount - 1
'if Form1.lv.ListItems(mv).Checked = True Or Form1.lv.ListItems(mv).Selected = True Then
If Dir(putanja & Form1.List2.List(mv)) <> "" Then
Select Case MsgBox("File " & Form1.List2.List(mv) & " already exist. Do you wan't to continue with moving??", vbYesNo)
Case vbYes
MoveFile Form1.List1.List(mv), putanja & Form1.List2.List(mv)
Case vbNo
End Select
Else
MoveFile Form1.List1.List(mv), putanja & Form1.List2.List(mv)
End If
Next mv
haha:
 For i = 1 To Form1.lv.ListItems.Count
        If Form1.lv.ListItems.Item(i).Checked = True Or Form1.lv.ListItems(i).Selected = True Then
            Form1.lv.ListItems.Remove (i)
           GoTo haha:
            Exit For
        End If
    Next i

editirano = True
Unload Me
End Sub

Private Sub Dir1_Change()
If Len(Dir1.path) > 3 Then
p = Dir1.path & "\"
Text1 = p
Else
p = Dir1.path
Text1 = p
End If
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1
End Sub


