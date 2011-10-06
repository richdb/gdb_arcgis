VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   3735
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   11145
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2577.963
   ScaleMode       =   0  'User
   ScaleWidth      =   10465.73
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   105
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1800
      ScaleWidth      =   10935
      TabIndex        =   12
      Top             =   105
      Width           =   10935
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Technische Informatie"
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   8400
      TabIndex        =   9
      Top             =   1995
      Width           =   2430
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "5692"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "Richard de Bruin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   210
         TabIndex        =   10
         Top             =   315
         Width           =   2010
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Informatie Cartografie"
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   5880
      TabIndex        =   6
      Top             =   1995
      Width           =   2430
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Alle Oldenbeuving"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   210
         TabIndex        =   8
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "5481"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   630
         Width           =   1485
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3150
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Informatie Metagegevens"
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   3360
      TabIndex        =   2
      Top             =   1995
      Width           =   2430
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "5894"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   5
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "Tineke Roodnat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   225
         Left            =   210
         TabIndex        =   4
         Top             =   315
         Width           =   1695
      End
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000005&
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   105
      TabIndex        =   0
      Top             =   2205
      Width           =   3150
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000005&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   1
      Top             =   2730
      Width           =   3150
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Versie " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub
