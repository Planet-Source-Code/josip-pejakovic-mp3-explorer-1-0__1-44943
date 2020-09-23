VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Favorite list..."
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10665
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin MSComctlLib.ListView lv 
         Height          =   3735
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   6588
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   3855
         Left            =   9000
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         Begin VB.CommandButton Command7 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Exit"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3360
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sor&t"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "O&pen"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Save"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Delete &All"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Delete"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Menu sorting 
      Caption         =   "sorting"
      Visible         =   0   'False
      Begin VB.Menu sort_filename 
         Caption         =   "Sort by filenname"
      End
      Begin VB.Menu sort_path 
         Caption         =   "Sort by path"
      End
      Begin VB.Menu sort_song_title 
         Caption         =   "Sort by song title"
      End
      Begin VB.Menu sort_artist 
         Caption         =   "Sort by artist"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim itms As Integer
    
With lv
For itms = .ListItems.Count To 1 Step -1
If .ListItems(itms).Selected Then
.ListItems.Remove (itms)
End If
Next itms
End With
End Sub

Private Sub Command2_Click()
lv.ListItems.Clear
lv.Refresh
End Sub

Private Sub Command3_Click()
sFilter = "MP3 Explorer Favorite List (*.mef)" & vbNullChar & "*.mef"
aa = GetSaveFilePath(hWnd, sFilter, 0, "mef", "", "", "Save...", sPath)
If aa = False Then Exit Sub
If Dir(sPath) <> "" Then Kill sPath
Module1.savefavorite sPath, lv
End Sub

Private Sub Command4_Click()
sFilter = "MP3 Explorer Favorite List (*.mef)" & vbNullChar & "*.mef"
bb = OpenSave.GetOpenFilePath(hWnd, sFilter, 0, "", "", "Otvaranje liste...", sPath)

If bb = False Then Exit Sub
'If Dir(sPath) <> "" Then Kill sPath

Module1.openfavorite sPath, lv
End Sub


Private Sub Command6_Click()
PopupMenu sorting
End Sub


Private Sub Command7_Click()
Form4.Hide
End Sub

Private Sub Form_Load()
With lv
.View = lvwReport
.ColumnHeaders.Add , , "Filename"
.ColumnHeaders.Add , , "Path" ', lv.Width ' vbCenter
.ColumnHeaders.Add , , "Song Title" ', lv.Width * 0.2, vbCenter
.ColumnHeaders.Add , , "Author/Artist"
.ColumnHeaders.Add , , "Album"
End With
End Sub

Sub sort_favorite(kolona As Integer)
lv.Sorted = True
lv.SortKey = kolona
lv.SortOrder = lvwAscending
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form4.Hide
End Sub

Private Sub sort_artist_Click()
sort_favorite (3)
End Sub

Private Sub sort_filename_Click()
sort_favorite (0)
End Sub

Private Sub sort_path_Click()
sort_favorite (1)
End Sub

Private Sub sort_song_title_Click()
sort_favorite (2)
End Sub
