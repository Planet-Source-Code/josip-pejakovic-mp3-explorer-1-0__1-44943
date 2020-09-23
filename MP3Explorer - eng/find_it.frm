VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form12 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pretraživanje liste"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Find"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   3375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   6135
         Begin MSComctlLib.ListView lv 
            Height          =   3015
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   5318
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Find inside 'Author'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   1440
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Find inside 'Song Title"""
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Find whole word only"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find:"
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   405
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Text1 = "" Then
Text1.SetFocus
Exit Sub
End If

lv.ListItems.Clear

If Check1.Value = 1 Then
    For x = 1 To Form1.lv.ListItems.Count
        If Form1.lv.ListItems(x).Text = Text1 Then
        Set bleh = lv.ListItems.Add(, , Form1.lv.ListItems(x).Text)
        bleh.SubItems(1) = Form1.lv.ListItems(x).SubItems(1)
        bleh.SubItems(2) = Form1.lv.ListItems(x).SubItems(2)
        bleh.SubItems(3) = Form1.lv.ListItems(x).SubItems(3)
        bleh.SubItems(4) = Form1.lv.ListItems(x).SubItems(4)
        bleh.SubItems(5) = Form1.lv.ListItems(x).SubItems(5)
        End If
    Next x
Else
    For x = 1 To Form1.lv.ListItems.Count
        If InStr(Form1.lv.ListItems(x).Text, Text1) > 0 Then
        ffs = Mid$(Form1.lv.ListItems(x).Text, 1, Len(Form1.lv.ListItems(x).Text))
        Set bleh = lv.ListItems.Add(, , ffs)
        bleh.SubItems(1) = Form1.lv.ListItems(x).SubItems(1)
        bleh.SubItems(2) = Form1.lv.ListItems(x).SubItems(2)
        bleh.SubItems(3) = Form1.lv.ListItems(x).SubItems(3)
        bleh.SubItems(4) = Form1.lv.ListItems(x).SubItems(4)
        bleh.SubItems(5) = Form1.lv.ListItems(x).SubItems(5)
        End If
    Next x
End If

If Option1.Value = 1 And Check1.Value = 1 Then
    For x = 1 To Form1.lv.ListItems.Count
        If Form1.lv.ListItems(x).SubItems(2) = Text1 Then
        Set bleh = lv.ListItems.Add(, , Form1.lv.ListItems(x).Text)
        bleh.SubItems(1) = Form1.lv.ListItems(x).SubItems(1)
        bleh.SubItems(2) = Form1.lv.ListItems(x).SubItems(2)
        bleh.SubItems(3) = Form1.lv.ListItems(x).SubItems(3)
        bleh.SubItems(4) = Form1.lv.ListItems(x).SubItems(4)
        bleh.SubItems(5) = Form1.lv.ListItems(x).SubItems(5)
        End If
    Next x
Else
    For x = 1 To Form1.lv.ListItems.Count
        If InStr(Form1.lv.ListItems(x).SubItems(3), Text1) > 0 Then
        ffs = Mid$(Form1.lv.ListItems(x).SubItems(3), 1, Len(Form1.lv.ListItems(x).SubItems(3)))
        Set bleh = lv.ListItems.Add(, , ffs)
        bleh.SubItems(1) = Form1.lv.ListItems(x).SubItems(1)
        bleh.SubItems(2) = Form1.lv.ListItems(x).SubItems(2)
        bleh.SubItems(3) = Form1.lv.ListItems(x).SubItems(3)
        bleh.SubItems(4) = Form1.lv.ListItems(x).SubItems(4)
        bleh.SubItems(5) = Form1.lv.ListItems(x).SubItems(5)
        End If
    Next x
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
With lv
.View = lvwReport
.ColumnHeaders.Add , , "Filename"
.ColumnHeaders.Add , , "Path"
.ColumnHeaders.Add , , "Song Title"
.ColumnHeaders.Add , , "Author"
.ColumnHeaders.Add , , "Album"
.ColumnHeaders.Add , , "Duration"
End With
End Sub
