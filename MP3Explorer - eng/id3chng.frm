VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ID3v1 Tag"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Exit"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Save"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4455
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   13
            Top             =   2160
            Width           =   3015
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   12
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   11
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   10
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   9
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comment:"
            Height          =   240
            Left            =   120
            TabIndex        =   8
            Top             =   2280
            Width           =   900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Year:"
            Height          =   240
            Left            =   120
            TabIndex        =   7
            Top             =   1800
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Album:"
            Height          =   240
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Artis:"
            Height          =   240
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Song title:"
            Height          =   240
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   210
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Path:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   360
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Tag As String * 3
Dim Songname As String * 30
Dim Artist As String * 30
Dim Album As String * 30
Dim Year As String * 4
Dim Comment As String * 30
Dim Genre As Byte

Tag = "TAG"
Songname = Text1.Text
Artist = Text2.Text
Album = Text3.Text
Year = Text4.Text
Comment = Text5.Text

atr = GetFileAttributes(Form1.lv.SelectedItem.SubItems(1))
If atr = FILE_ATTRIBUTE_ARCHIVE Or atr = FILE_ATTRIBUTE_HIDDEN Or atr = FILE_ATTRIBUTE_READONLY Or atr = FILE_ATTRIBUTE_SYSTEM Then
SetFileAttributes Form1.lv.SelectedItem.SubItems(1), FILE_ATTRIBUTE_NORMAL
End If

If Text1 = "" Then
Form1.lv.SelectedItem.SubItems(2) = Left(Form1.lv.SelectedItem, Len(Form1.lv.SelectedItem) - 4)
Else
Form1.lv.SelectedItem.SubItems(2) = Text1.Text
End If
Form1.lv.SelectedItem.SubItems(3) = Text2.Text
Form1.lv.SelectedItem.SubItems(4) = Text3.Text



Open Form1.lv.SelectedItem.SubItems(1) For Binary Access Write As #1
Seek #1, FileLen(Label2) - 127
Put #1, , Tag
Put #1, , Songname
Put #1, , Artist
Put #1, , Album
Put #1, , Year
Put #1, , Comment
Put #1, , Genre
Close #1

editirano = True
MsgBox ("ID3v1 Tags are saved in file!"), vbOKOnly
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
