VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export in HTML"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   5175
      Left            =   5160
      TabIndex        =   21
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Make page"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4560
         Width           =   4815
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2. Step"
         Height          =   4095
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   4815
         Begin VB.ComboBox Combo2 
            Height          =   345
            Left            =   960
            TabIndex        =   2
            Top             =   1080
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Height          =   345
            Left            =   120
            TabIndex        =   1
            Top             =   600
            Width           =   4455
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   1920
            Width           =   4455
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            TabIndex        =   4
            Top             =   2760
            Width           =   4455
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   120
            TabIndex        =   5
            Top             =   3600
            Width           =   4455
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font size:"
            Height          =   225
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Chose font:"
            Height          =   225
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   945
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Page Title:"
            Height          =   225
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   870
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type your name:"
            Height          =   225
            Left            =   120
            TabIndex        =   34
            Top             =   2520
            Width           =   1350
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Your e-mail address:"
            Height          =   225
            Left            =   120
            TabIndex        =   33
            Top             =   3360
            Width           =   1725
         End
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1. Step"
         Height          =   4695
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4695
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sample:"
            Height          =   225
            Left            =   840
            TabIndex        =   41
            Top             =   2880
            Width           =   690
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AaBbCc"
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   840
            TabIndex        =   40
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Height          =   225
            Left            =   1680
            TabIndex        =   39
            Top             =   360
            Width           =   45
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Height          =   225
            Left            =   1080
            TabIndex        =   38
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label Label23 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label22 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   30
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label21 
            BackColor       =   &H00800000&
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   1800
            TabIndex        =   28
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FF0000&
            Height          =   255
            Left            =   3480
            TabIndex        =   27
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFFC0&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label17 
            BackColor       =   &H0000C000&
            Height          =   255
            Left            =   960
            TabIndex        =   25
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label16 
            BackColor       =   &H0080FF80&
            Height          =   255
            Left            =   1800
            TabIndex        =   24
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label15 
            BackColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   23
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   3480
            TabIndex        =   22
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text color:"
            Height          =   225
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   3480
            TabIndex        =   19
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   18
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label10 
            BackColor       =   &H0080FF80&
            Height          =   255
            Left            =   1800
            TabIndex        =   17
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000C000&
            Height          =   255
            Left            =   960
            TabIndex        =   16
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFC0&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FF0000&
            Height          =   255
            Left            =   3480
            TabIndex        =   14
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label5 
            BackColor       =   &H00800000&
            Height          =   255
            Left            =   2640
            TabIndex        =   12
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   11
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label3 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Background color:"
            Height          =   225
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1485
         End
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
If bojapozadine = "" Then
MsgBox ("Select background color!"), vbInformation
Exit Sub
End If

If bojaslova = "" Then
MsgBox ("Select text color!"), vbInformation
Exit Sub
End If

sFilter = "HTML" & vbNullChar & "*.html"
aa = GetSaveFilePath(hWnd, sFilter, 0, "html", "", "", "Saving list in HTML...", sPath)
If aa = False Then Exit Sub

Open sPath For Output As #1
Print #1, "<html>"
Print #1, "<head>"
Print #1, "<title>" & Text1 & "</title>"
Print #1, "<meta http-quiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=iso-8859-2" & Chr(34) & ">"
Print #1, "</head>"
Print #1, "<body bgcolor=" & Chr(34) & bojapozadine & Chr(34) & " text=" & Chr(34) & bojaslova & Chr(34) & ">"
Print #1, "<font face=" & Chr(34) & Combo1.Text & Chr(34) & " size=" & Chr(34) & Combo2.Text & Chr(34) & ">"

Print #1, "<hr>"
For sta = 1 To Form1.lv.ListItems.Count
Print #1, Form1.lv.ListItems(sta).SubItems(2) & " - " & Form1.lv.ListItems(sta).SubItems(3) & "<br>"
Next sta

Print #1, "<hr>"
Print #1, "Listu je kreirao: " & Text2 & "<br>"
Print #1, "email: " & "<a href=" & Chr(34) & "mailto:" & Text3 & Chr(34) & ">" & Text3 & "</a>"
Print #1, "</font>"
Print #1, "</body>"
Print #1, "</html>"
Close #1

MsgBox ("Done!"), vbInformation
End Sub



Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()

With Combo1
.AddItem "Arial, Helvetica, sans-serif"
.AddItem "Times New Roman, Times, serif"
.AddItem "Courier New, Courier, mono"
.AddItem "Verdana, Arial, Helvetica, sans-serif"
End With

With Combo2
.AddItem "1"
.AddItem "2"
.AddItem "3"
.AddItem "4"
End With

Combo1.Text = Combo1.List(0)
Combo2.Text = Combo2.List(0)
End Sub

Private Sub Label10_Click()
bojapozadine = "#00FF99"
Label30 = bojapozadine
Label30.BackColor = Label10.BackColor
Label1.BackColor = Label10.BackColor
End Sub

Private Sub Label11_Click()
bojapozadine = "#FFFF33"
Label30 = bojapozadine
Label30.BackColor = Label11.BackColor
Label1.BackColor = Label11.BackColor
End Sub

Private Sub Label12_Click()
bojapozadine = "#FFFFCC"
Label30 = bojapozadine
Label30.BackColor = Label12.BackColor
Label1.BackColor = Label12.BackColor
End Sub

Private Sub Label14_Click()
bojaslova = "#FFFFCC"
Label29 = bojaslova
Label29.BackColor = Label14.BackColor
Label1.ForeColor = Label14.BackColor
End Sub

Private Sub Label15_Click()
bojaslova = "#FFFF33"
Label29 = bojaslova
Label29.BackColor = Label15.BackColor
Label1.ForeColor = Label15.BackColor
End Sub

Private Sub Label16_Click()
bojaslova = "#00FF99"
Label29 = bojaslova
Label29.BackColor = Label16.BackColor
Label1.ForeColor = Label16.BackColor
End Sub

Private Sub Label17_Click()
bojaslova = "#00CC66"
Label29 = bojaslova
Label29.BackColor = Label17.BackColor
Label1.ForeColor = Label17.BackColor
End Sub

Private Sub Label18_Click()
bojaslova = "#99FFFF"
Label29 = bojaslova
Label29.BackColor = Label18.BackColor
Label1.ForeColor = Label18.BackColor
End Sub

Private Sub Label19_Click()
bojaslova = "#3300FF"
Label29 = bojaslova
Label29.BackColor = Label19.BackColor
Label1.ForeColor = Label19.BackColor
End Sub

Private Sub Label20_Click()
bojaslova = "#CCCCCC"
Label29 = bojaslova
Label29.BackColor = Label20.BackColor
Label1.ForeColor = Label20.BackColor
End Sub

Private Sub Label21_Click()
bojaslova = "#330099"
Label29 = bojaslova
Label29.BackColor = Label21.BackColor
Label1.ForeColor = Label21.BackColor
End Sub

Private Sub Label22_Click()
bojaslova = "#000000"
Label29 = bojaslova
Label29.BackColor = Label22.BackColor
Label1.ForeColor = Label22.BackColor
End Sub

Private Sub Label23_Click()
bojaslova = "#FF0000"
Label29 = bojaslova
Label29.BackColor = Label23.BackColor
Label1.ForeColor = Label23.BackColor
End Sub

Private Sub Label3_Click()
bojapozadine = "#FF0000"
Label30 = bojapozadine
Label30.BackColor = Label3.BackColor
Label1.BackColor = Label3.BackColor
End Sub

Private Sub Label4_Click()
bojapozadine = "#000000"
Label30 = bojapozadine
Label30.BackColor = Label4.BackColor
Label1.BackColor = Label4.BackColor
End Sub

Private Sub Label5_Click()
bojapozadine = "#330099"
Label30 = bojapozadine
Label30.BackColor = Label5.BackColor
Label1.BackColor = Label30.BackColor
End Sub

Private Sub Label6_Click()
bojapozadine = "#CCCCCC"
Label30 = bojapozadine
Label30.BackColor = Label6.BackColor
Label1.BackColor = Label6.BackColor
End Sub

Private Sub Label7_Click()
bojapozadine = "#3300FF"
Label30 = bojapozadine
Label30.BackColor = Label7.BackColor
Label1.BackColor = Label7.BackColor
End Sub

Private Sub Label8_Click()
bojapozadine = "#99FFFF"
Label30 = bojapozadine
Label30.BackColor = Label8.BackColor
Label1.BackColor = Label8.BackColor
End Sub

Private Sub Label9_Click()
bojapozadine = "#00CC66"
Label30 = bojapozadine
Label30.BackColor = Label9.BackColor
Label1.BackColor = Label9.BackColor
End Sub


