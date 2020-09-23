VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MP3 Explorer"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList iml 
      Left            =   9120
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1002
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1116
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1432
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":174E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":20A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":338E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":41AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":482A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4CA2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   820
      ButtonWidth     =   820
      ButtonHeight    =   767
      Appearance      =   1
      Style           =   1
      ImageList       =   "iml"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newlist"
            Object.ToolTipText     =   "Create new list"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "openlist"
            Object.ToolTipText     =   "Open list"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "savelist"
            Object.ToolTipText     =   "Save list"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copyitem"
            Object.ToolTipText     =   "Copy file(s)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "moveitem"
            Object.ToolTipText     =   "Move file(s)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delitem"
            Object.ToolTipText     =   "Delete file(s)"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delitemlist"
            Object.ToolTipText     =   "Delete files from list"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "finditem"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "infoitem"
            Object.ToolTipText     =   "Change ID3 Tag"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "addfavorite"
            Object.ToolTipText     =   "Add file to the favorite list"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "viewfavorite"
            Object.ToolTipText     =   "View favorite list"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exportitem"
            Object.ToolTipText     =   "Export list"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "playitem"
            Object.ToolTipText     =   "Play file"
            ImageIndex      =   16
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "... play in MP3 Explorer player"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "...play in WinAMP"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ListBox List2 
      Height          =   1185
      Left            =   5160
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ListBox List1 
      Height          =   1185
      Left            =   1080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   6375
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   8895
      Begin MSComctlLib.ListView lv 
         Height          =   5175
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9128
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "size"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   6120
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "no of files"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   5880
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "path"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   5520
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Stop"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Walk through subdirecotires"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   6000
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Search"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5400
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4650
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dir..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   7080
      Width           =   450
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu delete_file 
         Caption         =   "&Delete File(s)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu delete_samo_lista 
         Caption         =   "Delete File(s) from &list"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu rename_file 
         Caption         =   "&Rename..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu copy_to 
         Caption         =   "&Copy File(s)..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu move_to 
         Caption         =   "Mo&ve File(s)..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu change_tag 
         Caption         =   "Chan&ge ID3 Tag"
         Shortcut        =   {F12}
      End
      Begin VB.Menu razv 
         Caption         =   "-"
      End
      Begin VB.Menu nova_lista 
         Caption         =   "&New list"
         Shortcut        =   ^N
      End
      Begin VB.Menu otvaranje_liste 
         Caption         =   "&Open list"
         Shortcut        =   ^O
      End
      Begin VB.Menu spremi_aktivnu 
         Caption         =   "Sav&e"
         Shortcut        =   ^S
      End
      Begin VB.Menu spremanje_liste 
         Caption         =   "Save list A&s..."
      End
      Begin VB.Menu razv1 
         Caption         =   "-"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu options 
         Caption         =   "&Options"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu lista 
      Caption         =   "&List"
      Begin VB.Menu lsep 
         Caption         =   "-"
      End
      Begin VB.Menu export_mp3 
         Caption         =   "&Export MP3 list"
         Shortcut        =   ^E
      End
      Begin VB.Menu favorite_list 
         Caption         =   "&Add file to the favorite list"
         Shortcut        =   ^D
      End
      Begin VB.Menu favorite_lista_pregled 
         Caption         =   "&View favorite list"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu sortiranje 
         Caption         =   "&Sort"
         Begin VB.Menu sort_kol_0 
            Caption         =   "... by filename"
         End
         Begin VB.Menu sort_kol_1 
            Caption         =   "... by path"
         End
         Begin VB.Menu sort_col_2 
            Caption         =   "... by song title"
         End
         Begin VB.Menu sort_kol_3 
            Caption         =   "... by song author"
         End
      End
      Begin VB.Menu sep51 
         Caption         =   "-"
      End
      Begin VB.Menu pretrazivanje_liste 
         Caption         =   "Find"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
   Begin VB.Menu popup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu play 
         Caption         =   "Play song"
         Begin VB.Menu ugradeni_player_sviraj 
            Caption         =   "... in MP3 Explorer player"
         End
         Begin VB.Menu sviraj_u_winamp 
            Caption         =   "... in WinAmp"
         End
      End
      Begin VB.Menu aha 
         Caption         =   "-"
      End
      Begin VB.Menu pop_delete 
         Caption         =   "Delete file(s)"
      End
      Begin VB.Menu popup_brisi_iz_liste 
         Caption         =   "Delete file(s) from list"
      End
      Begin VB.Menu pop_rename 
         Caption         =   "Rename"
      End
      Begin VB.Menu pop_copy 
         Caption         =   "Copy file(s)..."
      End
      Begin VB.Menu pop_move 
         Caption         =   "Move file(s)"
      End
      Begin VB.Menu pop_tag 
         Caption         =   "Change ID3 Tag"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu pop_favorite 
         Caption         =   "&Add file to the favorite list"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public x1, k, x
Public filesize

Private Sub about_Click()
Form7.Show , Me
End Sub

Private Sub change_tag_Click()
If odabrana_datoteka_putanja = "" Then
MsgBox ("Please select file from list!"), vbCritical
lv.SetFocus
Exit Sub
End If

Form6.Label2 = odabrana_datoteka_putanja
With Form6
.Text1 = Module1.GetTag(odabrana_datoteka_putanja)
.Text2 = nas
.Text3 = alb
.Text4 = god
.Text5 = komentari
End With
Form6.Show , Me
Label1.Visible = True
End Sub

Private Sub Command1_Click()
sortiranje_kolone (0)
prekini_proces = False
Command1.Enabled = False
Command2.Enabled = True
Call skeniraj

If lv.ListItems.Count = 0 Then
editirano = False
Else
editirano = True
End If
End Sub

Sub skeniraj()
   Dim SearchPath As String, FindStr As String
    Dim filesize As Double
    Dim NumFiles As Long, NumDirs As Long
    Screen.MousePointer = vbHourglass
    
    If otvorena_lista = False Then
    lv.ListItems.Clear
    Else
    End If
    
    SearchPath = Dir1.path
    FindStr = "*.mp3"
    filesize = FindFilesAPI(SearchPath, FindStr, NumFiles, NumDirs, Check1.Value)
    Screen.MousePointer = vbDefault

Label2 = ""
Label3 = ""
Label4 = ""
Label1 = UCase(Dir1)
For x = 1 To lv.ListItems.Count
Label2 = "Number of MP3 files: " & x
Next x
List1.Clear
List2.Clear
Command1.Enabled = True
Command2.Enabled = False
sortiranje.Enabled = True
Label5 = ""
veaf = 0
For x = 1 To lv.ListItems.Count
veaf = veaf + lv.ListItems(x).SubItems(6)
Next x
Label3 = "Total size: " & veaf & " MB"
End Sub
Sub obrisi()
For ll = 0 To List1.ListCount - 1
DeleteFile List1.List(ll)
Next ll
List1.Clear

 For i = 1 To lv.ListItems.Count
        If lv.ListItems.Item(i).Checked = True Then
            lv.ListItems.Remove (i)
            Call obrisi
            Exit For
        End If
    Next i


editirano = True
Exit Sub
End Sub

Private Sub Command2_Click()
prekini_proces = True
Command2.Enabled = False
Command1.Enabled = True
Label5 = ""
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Screen.MousePointer = vbDefault
End Sub






Private Sub copy_to_Click()
On Error GoTo err35600

X2 = ""
List1.Clear
For X2 = 1 To lv.ListItems.Count
If lv.ListItems(X2).Selected = True Or lv.ListItems(X2).Checked = True Then
List1.AddItem lv.ListItems.Item(X2).SubItems(1)
List2.AddItem lv.ListItems.Item(X2)
End If
Next X2

If List1.List(0) = "" Then
MsgBox ("Please select file!"), vbCritical
Exit Sub
End If

Form2.Show , Me
Form2.Caption = "Kopiranje..."
Form2.Command3.Visible = False
Form2.Command1.Visible = True

err35600:
If Err = 35600 Then
MsgBox ("Please select file!"), vbCritical
Exit Sub
End If
End Sub

Private Sub delete_file_Click()
On Error Resume Next
Dim it As ListItem

For x1 = 1 To lv.ListItems.Count
If lv.ListItems(x1).Checked = True Then
List1.AddItem lv.ListItems.Item(x1).SubItems(1) 'lv.SelectedItem
End If
Next x1

dane = MsgBox("Delete?", vbYesNo)
If dane = vbYes Then
'promjena atributa i brisanje fajlova
For ii = 1 To lv.ListItems.Count
If lv.ListItems.Item(ii).Checked = True Then
atr = GetFileAttributes(lv.ListItems.Item(ii).SubItems(1))
If atr = FILE_ATTRIBUTE_ARCHIVE Or atr = FILE_ATTRIBUTE_HIDDEN Or atr = FILE_ATTRIBUTE_READONLY Or atr = FILE_ATTRIBUTE_SYSTEM Then
SetFileAttributes lv.ListItems.Item(ii).SubItems(1), FILE_ATTRIBUTE_NORMAL
CloseHandle atr
Call obrisi
End If
If atr = FILE_ATTRIBUTE_NORMAL Then
CloseHandle atr
Call obrisi
Exit Sub
End If
End If
Next ii
End If

If dane = vbNo Then
Exit Sub
End If

End Sub

Private Sub delete_samo_lista_Click()
On Error Resume Next
For x1 = 1 To lv.ListItems.Count
If lv.ListItems(x1).Checked = True Then
Call obrisi2
End If
Next x1
End Sub

Sub obrisi2()
On Error Resume Next
For ll = 1 To lv.ListItems.Count
If lv.ListItems.Item(ll).Checked = True Then
lv.ListItems.Remove (ll)
End If
Next ll
editirano = True
Exit Sub
End Sub
Private Sub Dir1_Change()
'Text1 = Dir1.path
trenutni_dir = Dir1.path
End Sub

Private Sub Drive1_Change()
On Error GoTo err68:
Dir1.path = Drive1

err68:
If Err = 68 Then
MsgBox ("Device unavailable!"), vbCritical
Exit Sub
End If
End Sub

Private Sub exit_Click()
Call Form_Unload(1)
End Sub

Private Sub export_mp3_Click()
Form5.Show , Me
End Sub

Private Sub favorite_list_Click()
z = 0
For z = 1 To lv.ListItems.Count

If lv.ListItems(z).Selected = True Or lv.ListItems(z).Checked = True Then
Set bleh = Form4.lv.ListItems.Add(, , lv.ListItems(z).Text)
bleh.SubItems(1) = lv.ListItems(z).SubItems(1)
bleh.SubItems(2) = lv.ListItems(z).SubItems(2)
bleh.SubItems(3) = lv.ListItems(z).SubItems(3)
bleh.SubItems(4) = lv.ListItems(z).SubItems(4)
End If

Next z
lv.Refresh
z = 0
End Sub

Private Sub favorite_lista_pregled_Click()
Form4.Show , Me
End Sub

Private Sub Form_Load()
Dim hMenu As Long
    Const SC_SIZE = &HF000
    Const MF_BYCOMMAND = &H0
    hMenu = GetSystemMenu(hWnd, 0)
Call DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)

Label1 = "": Label2 = "": Label3 = ""
With lv
.View = lvwReport
.ColumnHeaders.Add , , "Filename"
.ColumnHeaders.Add , , "Path" ', lv.Width ' vbCenter
.ColumnHeaders.Add , , "Song Title" ', lv.Width * 0.2, vbCenter
.ColumnHeaders.Add , , "Artist"
.ColumnHeaders.Add , , "Album"
.ColumnHeaders.Add , , "Duration"
.ColumnHeaders.Add , , "Size (MB)"
End With
Label5 = ""


Form2.Command3.Visible = False
Command2.Enabled = False

sortiranje.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If editirano = True Then
Select Case MsgBox("You have made changes in list. Do you wan't save changes?", vbYesNo)
Case vbYes
Call spremi_aktivnu_Click
End
Case vbNo
End
End Select
Else
End
End If
End Sub

Private Sub lv_AfterLabelEdit(Cancel As Integer, NewString As String)
editirano = True
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
odabrana_datoteka = lv.SelectedItem
odabrana_datoteka_putanja = lv.SelectedItem.SubItems(1)
End Sub

'Private Sub lv_ItemClick(ByVal Item As ComctlLib.ListItem)
'Label1 = UCase(Item.SubItems(1))
'putanja = Item.SubItems(1)
'End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyUp Then
Label1 = UCase(lv.SelectedItem.SubItems(1))
Label4 = lv.SelectedItem
End If

If KeyAscii = vbKeyDown Then
Label1 = UCase(lv.SelectedItem.SubItems(1))
Label4 = lv.SelectedItem
End If
End Sub
Private Sub lv_LostFocus()
Label1 = ""
End Sub

Private Sub lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu popup
End If
End Sub

Private Sub move_to_Click()
On Error GoTo err35600

List1.Clear
List2.Clear
x3 = 0
For x3 = 1 To lv.ListItems.Count
If lv.ListItems(x3).Selected = True Or lv.ListItems(x3).Checked = True Then
'If lv.ListItems(x3).Checked = True Then
List1.AddItem lv.ListItems.Item(x3).SubItems(1)
List2.AddItem lv.ListItems.Item(x3)
End If
Next x3

If List1.List(0) = "" Then
MsgBox ("Please select file!"), vbCritical
Exit Sub
End If

Form2.Show , Me
Form2.Command3.Visible = True
Form2.Command1.Visible = False

err35600:
If Err = 35600 Then
MsgBox ("Please select file!!"), vbCritical
Exit Sub
End If
End Sub

Private Sub play_mp_Click()
putanja = Label1
'Form7.Show , Me
MP3Player.Show , Me
End Sub

Private Sub nova_lista_Click()
If editirano = True Then
Select Case MsgBox("You made changes in list. Do you wan't save changes?", vbYesNoCancel)
Case vbYes
Call spremi_aktivnu_Click
lv.ListItems.Clear
otvorena_lista = True
editirano = False
Label1 = "": Label2 = "": Label3 = ""
Case vbNo
lv.ListItems.Clear
otvorena_lista = True
editirano = False
Label1 = "": Label2 = "": Label3 = ""
Case vbCancel
Exit Sub
End Select
Else
lv.ListItems.Clear
End If
Frame2.Caption = ""
End Sub

Private Sub options_Click()
Form10.Show , Me
End Sub



Private Sub otvaranje_liste_Click()
If editirano = True Then
Select Case MsgBox("You made changes in list. Do you wan't save changes?", vbYesNoCancel)
Case vbYes
Call spremi_aktivnu_Click
Call otvori_listu
editirano = False
Label1 = "": Label2 = "": Label3 = ""
Case vbNo
editirano = False
Call otvori_listu
Label1 = "": Label2 = "": Label3 = ""
Case vbCancel
Exit Sub
End Select
Else
Call otvori_listu
End If

'If editirano = False Then
'Call otvori_listu
'End If


End Sub

Private Sub pop_copy_Click()
copy_to_Click
End Sub

Private Sub pop_delete_Click()
delete_file_Click
End Sub

Private Sub pop_favorite_Click()
favorite_list_Click
End Sub

Private Sub pop_move_Click()
move_to_Click
End Sub

Private Sub pop_rename_Click()
rename_file_Click
End Sub

Private Sub pop_tag_Click()
change_tag_Click
End Sub

Private Sub popup_brisi_iz_liste_Click()
Call delete_samo_lista_Click
End Sub

Private Sub pretrazivanje_liste_Click()
Form12.Show , Me
End Sub

Private Sub rename_file_Click()
If odabrana_datoteka = "" Then
MsgBox ("Please select file!"), vbCritical
lv.SetFocus
Exit Sub
End If

Form3.Text1 = odabrana_datoteka
Form3.Show , Me
End Sub

Sub provjera_selected()
For oo = 1 To lv.ListItems.Count
If lv.ListItems.Item(oo).Selected = False Then
MsgBox ("Please select file!"), vbCritical
Exit Sub
End If
Next oo
End Sub

Private Sub sort_col_2_Click()
sortiranje_kolone (2)
End Sub

Private Sub sort_kol_0_Click()
sortiranje_kolone (0)
End Sub

Sub sortiranje_kolone(kolona As Integer)
lv.Sorted = True
lv.SortKey = kolona
lv.SortOrder = lvwAscending
End Sub

Private Sub sort_kol_1_Click()
sortiranje_kolone (1)
End Sub

Private Sub sort_kol_3_Click()
sortiranje_kolone (3)
End Sub

Private Sub spremanje_liste_Click()
sFilter = "MP3 Explorer list" & vbNullChar & "*.ls"
aa = GetSaveFilePath(hWnd, sFilter, 0, "ls", "", "", "Saving list...", sPath)
If aa = False Then Exit Sub

Open sPath For Output As #1
For x = 1 To lv.ListItems.Count
Print #1, lv.ListItems(x) & "@" & lv.ListItems(x).SubItems(1) & "@" & lv.ListItems(x).SubItems(2) & "@" & lv.ListItems(x).SubItems(3) & "@" & lv.ListItems(x).SubItems(4) & "@" & lv.ListItems(x).SubItems(5) & "@" & lv.ListItems(x).SubItems(6)
Next x
Close #1
Frame2.Caption = sPath


End Sub

Private Sub spremi_aktivnu_Click()
If sPath = "" Then
sFilter = "MP3 Explorer list" & vbNullChar & "*.ls"
aa = GetSaveFilePath(hWnd, sFilter, 0, "ls", "", "", "Saving list...", sPath)
If aa = False Then Exit Sub
Open sPath For Output As #1
For x = 1 To lv.ListItems.Count
Print #1, lv.ListItems(x) & "@" & lv.ListItems(x).SubItems(1) & "@" & lv.ListItems(x).SubItems(2) & "@" & lv.ListItems(x).SubItems(3) & "@" & lv.ListItems(x).SubItems(4) & "@" & lv.ListItems(x).SubItems(5) & "@" & lv.ListItems(x).SubItems(6)
Next x
Close #1
Else
Open sPath For Output As #1
For x = 1 To lv.ListItems.Count
Print #1, lv.ListItems(x) & "@" & lv.ListItems(x).SubItems(1) & "@" & lv.ListItems(x).SubItems(2) & "@" & lv.ListItems(x).SubItems(3) & "@" & lv.ListItems(x).SubItems(4) & "@" & lv.ListItems(x).SubItems(5) & "@" & lv.ListItems(x).SubItems(6)
Next x
Close #1
End If
Frame2.Caption = sPath
End Sub

Sub otvori_listu()
sFilter = "MP3 Explorer list" & vbNullChar & "*.ls"
aa = GetOpenFilePath(hWnd, sFilter, 0, "", "", "", sPath) = True
If aa = False Then Exit Sub

Open sPath For Input As #1
lv.ListItems.Clear
Do Until EOF(1)
Line Input #1, st
fn = Split(st, "@")(0)
pth = Split(st, "@")(1)
npj = Split(st, "@")(2)
artn = Split(st, "@")(3)
albn = Split(st, "@")(4)
vrt = Split(st, "@")(5)
velfa = Split(st, "@")(6)
Set bleh = lv.ListItems.Add(, , fn)
bleh.SubItems(1) = pth
bleh.SubItems(2) = npj
bleh.SubItems(3) = artn
bleh.SubItems(4) = albn
bleh.SubItems(5) = vrt
bleh.SubItems(6) = velfa
Loop
Close #1

veaf = 0
For x = 1 To lv.ListItems.Count
veaf = veaf + lv.ListItems(x).SubItems(6)
Next x
Label2 = "Number of MP3 files: " & lv.ListItems.Count
Label3 = "Total size: " & veaf & " MB"

otvorena_lista = True
Frame2.Caption = sPath
End Sub

Private Sub sviraj_u_winamp_Click()
On Error GoTo obrerr53:
Dim putanja
Open App.path & "\conf.dat" For Input As #1
Line Input #1, putanja
Close #1
For x = 1 To lv.ListItems.Count
If lv.ListItems(x).Selected = True Then
dodpj = putanja & " /ADD " & Chr$(34) & lv.ListItems(x).SubItems(1) & Chr$(34)
Shell dodpj
End If
Next x

obrerr53:
If Err = 53 Then
Form10.Show , Me
Err.Clear
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case Is = 1
    Call nova_lista_Click
    Case Is = 2
    Call otvaranje_liste_Click
    Case Is = 3
    Call spremi_aktivnu_Click
    Case Is = 5
    Call copy_to_Click
    Case Is = 6
    Call move_to_Click
    Case Is = 7
    Call delete_file_Click
    Case Is = 8
    Call delete_samo_lista_Click
    Case Is = 10
    Call pretrazivanje_liste_Click
    Case Is = 11
    Call change_tag_Click
    Case Is = 12
    Call favorite_list_Click
    Case Is = 13
    Call favorite_lista_pregled_Click
    Case Is = 14
    export_mp3_Click
End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   Select Case ButtonMenu.Index
   Case 1
      Call ugradeni_player_sviraj_Click
   Case 2
      Call sviraj_u_winamp_Click
   End Select
End Sub
Private Sub ugradeni_player_sviraj_Click()
On Error GoTo obr91
Form9.Label2 = lv.SelectedItem.SubItems(1)
Form9.Label3 = lv.SelectedItem.SubItems(3) & " - " & lv.SelectedItem.SubItems(2)
Form9.Show , Me

obr91:
If Err = 91 Then
MsgBox ("Odaberite MP3 datoteku iz liste!")
Err.Clear
Exit Sub
End If
End Sub
