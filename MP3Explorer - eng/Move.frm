VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move File"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6180
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
   ScaleHeight     =   2385
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destionaton"
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Move"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
