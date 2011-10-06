VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "ArcGIS - GDB"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   12960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   12960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   350
      Left            =   9870
      TabIndex        =   116
      Top             =   8085
      Width           =   1380
   End
   Begin VB.CommandButton SluitCmd 
      Caption         =   "&Sluiten"
      Height          =   350
      Left            =   11445
      TabIndex        =   115
      ToolTipText     =   "ArcGIS - GBI afsluiten"
      Top             =   8085
      Width           =   1400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PDF"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5775
      TabIndex        =   92
      Top             =   8085
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Zoeken"
      Height          =   350
      Left            =   11130
      TabIndex        =   91
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   350
      Left            =   4125
      TabIndex        =   90
      Top             =   360
      Width           =   6855
   End
   Begin VB.CommandButton ArcGISCmd 
      Caption         =   "Arc&GIS"
      Height          =   350
      Left            =   4095
      TabIndex        =   2
      ToolTipText     =   "Gegevens toevoegen aan ArcGIS"
      Top             =   8085
      Width           =   1380
   End
   Begin VB.CommandButton AllemaalCmd 
      Caption         =   "&Allemaal"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Alle meta-gegevens weergeven"
      Top             =   8085
      Width           =   1380
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   7635
      ItemData        =   "Form1.frx":014A
      Left            =   120
      List            =   "Form1.frx":014C
      TabIndex        =   0
      ToolTipText     =   "Dubbelklik op de naam voor het ophalen van de meta-gegevens"
      Top             =   360
      Width           =   3810
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7150
      Left            =   4125
      TabIndex        =   4
      Top             =   840
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   12621
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   499
      TabCaption(0)   =   "Algemeen"
      TabPicture(0)   =   "Form1.frx":014E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label34"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label33"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label31"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label32"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "AlgemeenTextbox10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "AlgemeenTextbox9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "AlgemeenTextbox8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "AlgemeenTextbox3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "AlgemeenTextbox4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "AlgemeenTextbox1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "AlgemeenTextbox7"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "AlgemeenTextbox6"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "AlgemeenTextbox5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "AlgemeenTextbox2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Inhoud"
      TabPicture(1)   =   "Form1.frx":016A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label43"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label42"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label41"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label38"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label37"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label36"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label35"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label11"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label10"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label9"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label7"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label39"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label40"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label54"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "InhoudTextbox13"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "InhoudTextbox12"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "InhoudTextbox9"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "InhoudTextbox8"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "InhoudTextbox7"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "InhoudTextbox6"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "InhoudTextbox5"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "InhoudTextbox4"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "InhoudTextbox3"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "InhoudTextbox2"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "InhoudTextbox1"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "InhoudTextbox10"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "InhoudTextbox11"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Text5"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).ControlCount=   29
      TabCaption(2)   =   "Specifiek"
      TabPicture(2)   =   "Form1.frx":0186
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SpecifiekTextbox13"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SpecifiekTextbox1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "SpecifiekTextbox2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "SpecifiekTextbox3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "SpecifiekTextbox4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "SpecifiekTextbox5"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "SpecifiekTextbox6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "SpecifiekTextbox7"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "SpecifiekTextbox8"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "SpecifiekTextbox9"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "SpecifiekTextbox10"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "SpecifiekTextbox11"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "SpecifiekTextbox12"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label51"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label12"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label13"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label14"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label15"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label16"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label44"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Label45"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Label46"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label47"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label48"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label49"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Label50"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "Items"
      TabPicture(3)   =   "Form1.frx":01A2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Text3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Text2"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "ItemsList1"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "ItemsTextbox1"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "ItemsTextbox2"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label55"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label53"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label52"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label20"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label19"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label18"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label17"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).ControlCount=   13
      TabCaption(4)   =   "Metadata"
      TabPicture(4)   =   "Form1.frx":01BE
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label29"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label28"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label27"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label26"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label25"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label24"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label23"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label22"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label30"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "MetadataTextbox7"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "MetadataTextbox6"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "MetadataTextbox5"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "MetadataTextbox4"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "MetadataTextbox3"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "MetadataTextbox2"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "MetadataTextbox1"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "MetadataTextbox8"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "MetadataTextbox9"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).ControlCount=   18
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   2730
         Width           =   6090
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   1695
         Left            =   -72720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   110
         Top             =   3400
         Width           =   6105
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   2950
         Width           =   6105
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   2500
         Width           =   6105
      End
      Begin VB.ListBox ItemsList1 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "Form1.frx":01DA
         Left            =   -72720
         List            =   "Form1.frx":01DC
         TabIndex        =   104
         Top             =   950
         Width           =   6105
      End
      Begin VB.TextBox SpecifiekTextbox13 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -69160
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   2750
         Width           =   2520
      End
      Begin VB.TextBox InhoudTextbox11 
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   -72720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   100
         Top             =   5625
         Width           =   6105
      End
      Begin VB.TextBox InhoudTextbox10 
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   -72720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   98
         Top             =   4965
         Width           =   6105
      End
      Begin VB.TextBox MetadataTextbox9 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   1850
         Width           =   6090
      End
      Begin VB.TextBox MetadataTextbox8 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   1400
         Width           =   6090
      End
      Begin VB.TextBox AlgemeenTextbox2 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   950
         Width           =   6090
      End
      Begin VB.TextBox AlgemeenTextbox5 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3250
         Width           =   2655
      End
      Begin VB.TextBox AlgemeenTextbox6 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   3700
         Width           =   2655
      End
      Begin VB.TextBox AlgemeenTextbox7 
         Appearance      =   0  'Flat
         Height          =   500
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   4150
         Width           =   6105
      End
      Begin VB.TextBox AlgemeenTextbox1 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   500
         Width           =   6090
      End
      Begin VB.TextBox AlgemeenTextbox4 
         Appearance      =   0  'Flat
         Height          =   500
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   2650
         Width           =   6105
      End
      Begin VB.TextBox InhoudTextbox1 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   500
         Width           =   6090
      End
      Begin VB.TextBox InhoudTextbox2 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   950
         Width           =   6090
      End
      Begin VB.TextBox InhoudTextbox3 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1400
         Width           =   6090
      End
      Begin VB.TextBox InhoudTextbox4 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1850
         Width           =   6090
      End
      Begin VB.TextBox InhoudTextbox5 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2300
         Width           =   6090
      End
      Begin VB.TextBox SpecifiekTextbox1 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   525
         Width           =   6090
      End
      Begin VB.TextBox SpecifiekTextbox2 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   950
         Width           =   6090
      End
      Begin VB.TextBox SpecifiekTextbox3 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1400
         Width           =   6090
      End
      Begin VB.TextBox SpecifiekTextbox4 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1850
         Width           =   6090
      End
      Begin VB.TextBox SpecifiekTextbox5 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2300
         Width           =   6090
      End
      Begin VB.TextBox ItemsTextbox1 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   500
         Width           =   6090
      End
      Begin VB.TextBox ItemsTextbox2 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2050
         Width           =   6090
      End
      Begin VB.TextBox MetadataTextbox1 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   500
         Width           =   6090
      End
      Begin VB.TextBox MetadataTextbox2 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   950
         Width           =   6090
      End
      Begin VB.TextBox MetadataTextbox3 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2300
         Width           =   6090
      End
      Begin VB.TextBox MetadataTextbox4 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2750
         Width           =   6090
      End
      Begin VB.TextBox MetadataTextbox5 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3200
         Width           =   6090
      End
      Begin VB.TextBox MetadataTextbox6 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3650
         Width           =   6090
      End
      Begin VB.TextBox MetadataTextbox7 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   4095
         Width           =   6090
      End
      Begin VB.TextBox AlgemeenTextbox3 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1400
         Width           =   6075
      End
      Begin VB.TextBox AlgemeenTextbox8 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4750
         Width           =   2535
      End
      Begin VB.TextBox AlgemeenTextbox9 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   5200
         Width           =   2535
      End
      Begin VB.TextBox AlgemeenTextbox10 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   5670
         Width           =   2535
      End
      Begin VB.TextBox InhoudTextbox6 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3165
         Width           =   6090
      End
      Begin VB.TextBox InhoudTextbox7 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3615
         Width           =   6090
      End
      Begin VB.TextBox InhoudTextbox8 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4065
         Width           =   6090
      End
      Begin VB.TextBox InhoudTextbox9 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4515
         Width           =   6090
      End
      Begin VB.TextBox InhoudTextbox12 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   6540
         Width           =   2415
      End
      Begin VB.TextBox InhoudTextbox13 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -69360
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   6540
         Width           =   2415
      End
      Begin VB.TextBox SpecifiekTextbox6 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2750
         Width           =   2520
      End
      Begin VB.TextBox SpecifiekTextbox7 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3200
         Width           =   6090
      End
      Begin VB.TextBox SpecifiekTextbox8 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3650
         Width           =   6090
      End
      Begin VB.TextBox SpecifiekTextbox9 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4100
         Width           =   1815
      End
      Begin VB.TextBox SpecifiekTextbox10 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -68460
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4100
         Width           =   1815
      End
      Begin VB.TextBox SpecifiekTextbox11 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4550
         Width           =   1815
      End
      Begin VB.TextBox SpecifiekTextbox12 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   -68460
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4550
         Width           =   1815
      End
      Begin VB.Label Label55 
         Caption         =   "Dubbelklik voor gegevens"
         ForeColor       =   &H80000001&
         Height          =   225
         Left            =   -74790
         TabIndex        =   114
         Top             =   1155
         Width           =   1905
      End
      Begin VB.Label Label54 
         Caption         =   "Veiligheidsrestricties:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   112
         Top             =   2770
         Width           =   1635
      End
      Begin VB.Label Label53 
         Caption         =   "Mogelijke waarden:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   111
         Top             =   3465
         Width           =   1695
      End
      Begin VB.Label Label52 
         Caption         =   "Eenheid:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   109
         Top             =   3003
         Width           =   1380
      End
      Begin VB.Label Label20 
         Caption         =   "Itemdefinitie:"
         Height          =   330
         Left            =   -74790
         TabIndex        =   107
         Top             =   2510
         Width           =   1590
      End
      Begin VB.Label Label19 
         Caption         =   "Itemnaam:"
         Height          =   330
         Left            =   -74790
         TabIndex        =   105
         Top             =   2060
         Width           =   1485
      End
      Begin VB.Label Label18 
         Caption         =   "Items:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   103
         Top             =   945
         Width           =   1380
      End
      Begin VB.Label Label51 
         Caption         =   "Geometrie:"
         Height          =   225
         Left            =   -70065
         TabIndex        =   101
         Top             =   2835
         Width           =   960
      End
      Begin VB.Label Label40 
         Caption         =   "Trefwoorden:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   99
         Top             =   5565
         Width           =   1380
      End
      Begin VB.Label Label39 
         Caption         =   "Contact leverancier:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   97
         Top             =   4935
         Width           =   1905
      End
      Begin VB.Label Label30 
         Caption         =   "Karakterset:"
         Height          =   255
         Left            =   -74790
         TabIndex        =   96
         Top             =   1898
         Width           =   1935
      End
      Begin VB.Label Label22 
         Caption         =   "Taal dataset:"
         Height          =   255
         Left            =   -74790
         TabIndex        =   94
         Top             =   1448
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Alternatieve titel:"
         Height          =   270
         Left            =   210
         TabIndex        =   88
         Top             =   990
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "Referentie datum:"
         Height          =   270
         Left            =   210
         TabIndex        =   87
         Top             =   3290
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Wijzigingsdatum:"
         Height          =   270
         Left            =   210
         TabIndex        =   86
         Top             =   3740
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Bronvermelding:"
         Height          =   270
         Left            =   210
         TabIndex        =   85
         Top             =   4200
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Titel dataset:"
         Height          =   270
         Left            =   210
         TabIndex        =   84
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Omschrijving dataset:"
         Height          =   270
         Left            =   210
         TabIndex        =   83
         Top             =   1365
         Width           =   1995
      End
      Begin VB.Label Label7 
         Caption         =   "Contactpersoon inhoud:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   82
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label Label8 
         Caption         =   "Beleidsterrein:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   81
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label Label9 
         Caption         =   "Team:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   80
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label10 
         Caption         =   "Thema:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   79
         Top             =   1890
         Width           =   1755
      End
      Begin VB.Label Label11 
         Caption         =   "Gebruiksbeperking:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   78
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label Label12 
         Caption         =   "Geografisch gebied:"
         Height          =   270
         Left            =   -74760
         TabIndex        =   77
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label Label13 
         Caption         =   "Ruimtelijk schema:"
         Height          =   270
         Left            =   -74760
         TabIndex        =   76
         Top             =   990
         Width           =   1635
      End
      Begin VB.Label Label14 
         Caption         =   "Aanvullende informatie:"
         Height          =   270
         Left            =   -74760
         TabIndex        =   75
         Top             =   1440
         Width           =   1785
      End
      Begin VB.Label Label15 
         Caption         =   "Layernaam:"
         Height          =   270
         Left            =   -74760
         TabIndex        =   74
         Top             =   1890
         Width           =   1650
      End
      Begin VB.Label Label16 
         Caption         =   "Fysieke locatie:"
         Height          =   270
         Left            =   -74760
         TabIndex        =   73
         Top             =   2340
         Width           =   1530
      End
      Begin VB.Label Label17 
         Caption         =   "Standaarditem:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   72
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label23 
         Caption         =   "Contactpersoon:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   71
         Top             =   540
         Width           =   1530
      End
      Begin VB.Label Label24 
         Caption         =   "Datum opbouw:"
         Height          =   270
         Left            =   -74760
         TabIndex        =   70
         Top             =   990
         Width           =   1395
      End
      Begin VB.Label Label25 
         Caption         =   "Metadatastandaard:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   69
         Top             =   2352
         Width           =   1710
      End
      Begin VB.Label Label26 
         Caption         =   "Versie:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   68
         Top             =   2790
         Width           =   1035
      End
      Begin VB.Label Label27 
         Caption         =   "Referentiesysteem:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   67
         Top             =   3240
         Width           =   1665
      End
      Begin VB.Label Label28 
         Caption         =   "Referentie organisatie:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   66
         Top             =   3690
         Width           =   2010
      End
      Begin VB.Label Label29 
         Caption         =   "Contactpersoon distributie:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   65
         Top             =   4135
         Width           =   1905
      End
      Begin VB.Label Label32 
         Caption         =   "Algemene opmerking:"
         Height          =   255
         Left            =   210
         TabIndex        =   64
         Top             =   2625
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "Opbouwmethode:"
         Height          =   255
         Left            =   210
         TabIndex        =   63
         Top             =   4798
         Width           =   1335
      End
      Begin VB.Label Label33 
         Caption         =   "Gebeurtenis:"
         Height          =   255
         Left            =   210
         TabIndex        =   62
         Top             =   5248
         Width           =   1455
      End
      Begin VB.Label Label34 
         Caption         =   "Status:"
         Height          =   255
         Left            =   210
         TabIndex        =   61
         Top             =   5718
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "Toegangsrestricties:"
         Height          =   255
         Left            =   -74790
         TabIndex        =   60
         Top             =   3225
         Width           =   1935
      End
      Begin VB.Label Label36 
         Caption         =   "Copyright:"
         Height          =   255
         Left            =   -74790
         TabIndex        =   59
         Top             =   3675
         Width           =   1575
      End
      Begin VB.Label Label37 
         Caption         =   "Herzienings frequentie:"
         Height          =   255
         Left            =   -74790
         TabIndex        =   58
         Top             =   4125
         Width           =   1935
      End
      Begin VB.Label Label38 
         Caption         =   "Toepassingsschaal:"
         Height          =   255
         Left            =   -74790
         TabIndex        =   57
         Top             =   4575
         Width           =   1815
      End
      Begin VB.Label Label41 
         Caption         =   "Temporele dekking:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   56
         Top             =   6300
         Width           =   1575
      End
      Begin VB.Label Label42 
         Caption         =   "Begin datum:"
         Height          =   255
         Left            =   -72720
         TabIndex        =   55
         Top             =   6300
         Width           =   1335
      End
      Begin VB.Label Label43 
         Caption         =   "Eind datum:"
         Height          =   255
         Left            =   -69360
         TabIndex        =   54
         Top             =   6300
         Width           =   1335
      End
      Begin VB.Label Label44 
         Caption         =   "Datatype:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   53
         Top             =   2798
         Width           =   1575
      End
      Begin VB.Label Label45 
         Caption         =   "Nauwkeurigheid:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   52
         Top             =   3248
         Width           =   1815
      End
      Begin VB.Label Label46 
         Caption         =   "Hiërarchieniveau:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   51
         Top             =   3698
         Width           =   1815
      End
      Begin VB.Label Label47 
         Caption         =   "Minimale x-coördinaat:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   50
         Top             =   4148
         Width           =   1815
      End
      Begin VB.Label Label48 
         Caption         =   "Maximale x-coördinaat:"
         Height          =   255
         Left            =   -70380
         TabIndex        =   49
         Top             =   4155
         Width           =   1695
      End
      Begin VB.Label Label49 
         Caption         =   "Minimale y-coördinaat:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   48
         Top             =   4598
         Width           =   1815
      End
      Begin VB.Label Label50 
         Caption         =   "Maximale y-coördinaat:"
         Height          =   255
         Left            =   -70380
         TabIndex        =   47
         Top             =   4598
         Width           =   1695
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Zoeken op trefwoord:"
      Height          =   255
      Left            =   4125
      TabIndex        =   89
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label LabelNamen 
      Caption         =   "Namen (516 van 516)"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programma:     ArcGIS - GDB
'Datum:         september 2002
'Datum:         februari 2007
'Auteurs:       Richard de Bruin
'Copyright:     Provincie Drenthe

'Versie 1.0.0:
'   ArcGIS plugin voor het doorzoeken van de GBI database
'

'Versie 2.0:
'   ArcGIS plugin voor het doorzoeken van de GBI database
'   Nu met gebruik van de Oracle Spatial en LayerFiles

Option Explicit
Private WithEvents db As ADODB.Connection
Attribute db.VB_VarHelpID = -1
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private WithEvents rsselect As ADODB.Recordset
Attribute rsselect.VB_VarHelpID = -1
Public dbdir As String
Public ladir As String
Public versie As String
Public totaalmeta As Integer
Dim m_pApp As IApplication
Public Property Let Application(ByRef pApp As IApplication)
    Set m_pApp = pApp
Exit Property
End Property
Function SelectRecord(SQL As String)
Dim waarde As String

Set rsselect = New ADODB.Recordset
rsselect.Open SQL, db, adOpenStatic, adLockOptimistic
waarde = ""

If Not rsselect.EOF Then
    If Not IsNull(rsselect.Fields(0)) Then
        waarde = rsselect.Fields(0)
    Else
        waarde = ""
    End If
    rsselect.Close
End If
SelectRecord = waarde
End Function
Function SelectRecordInt(SQL As String)
Dim waarde As Integer

Set rsselect = New ADODB.Recordset
rsselect.Open SQL, db, adOpenStatic, adLockOptimistic
waarde = ""

If Not rsselect.EOF Then
    If Not IsNull(rsselect.Fields(0)) Then
        waarde = rsselect.Fields(0)
    Else
        waarde = ""
    End If
    rsselect.Close
End If
SelectRecordInt = waarde
End Function
Sub initialiseer()
Dim regel As String
' Routine voor het inlezen van het ini bestand
On Error GoTo ErrorHandler

versie = "ArcGIS - GDB Versie" + Str$(App.Major) + "." + Str$(App.Minor) + "." + Str$(App.Revision)
Form1.Caption = versie
Open App.Path + "\arcgisgbi.ini" For Input As #1
    Do Until EOF(1)
        Line Input #1, regel
        If InStr(1, regel, "DBdir=") > 0 Then
            dbdir = Trim(Mid$(regel, InStr(1, regel, "=") + 1, Len(regel)))
            If Dir(dbdir, vbNormal) = "" Then
                MsgBox "De database bestaat niet!!"
            End If
        End If
        If InStr(1, regel, "LAdir=") > 0 Then
            ladir = Trim(Mid$(regel, InStr(1, regel, "=") + 1, Len(regel)))
            If Dir(ladir, vbNormal) = "" Then
                MsgBox "Geen pad ingevuld in het INI bestand!"
            End If
        End If
    Loop
    Close #1
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Sub vullenmetadatalijst()
Dim SQL, lijst, status As String
On Error GoTo ErrorHandler

SQL = "SELECT DATASET_TITEL FROM DATASET WHERE TYPE = 1 ORDER BY DATASET_TITEL;"

Form1.List1.Clear
Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

totaalmeta = 0

With rs
  If .RecordCount > 0 Then
    .MoveFirst
    Do
      lijst = .Fields("DATASET_TITEL")
      Form1.List1.AddItem lijst
      .MoveNext
      totaalmeta = totaalmeta + 1
    Loop Until .EOF
  End If
End With

rs.Close
LabelNamen.Caption = "Namen (" + Str(totaalmeta) + " van " + Str(totaalmeta) + ")"

Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Sub VullenListbox(SQL As String, waarde As Integer, list As ListBox)
Dim lijst As String

list.Clear
Set rsselect = New ADODB.Recordset
rsselect.Open SQL, db, adOpenStatic, adLockOptimistic
  
With rsselect
  If Not .EOF Then
        .MoveFirst
        Do
        If Not IsNull(rsselect.Fields(waarde)) Then
          lijst = rsselect.Fields(waarde) & ""
          list.AddItem (lijst)
        End If
        rsselect.MoveNext
        Loop Until .EOF
    End If
    .Close
End With
End Sub
Sub Opschonen()
AlgemeenTextbox1.Text = ""
AlgemeenTextbox2.Text = ""
AlgemeenTextbox3.Text = ""
AlgemeenTextbox4.Text = ""
AlgemeenTextbox5.Text = ""
AlgemeenTextbox6.Text = ""
AlgemeenTextbox7.Text = ""
AlgemeenTextbox8.Text = ""
AlgemeenTextbox9.Text = ""
AlgemeenTextbox10.Text = ""
InhoudTextbox1.Text = ""
InhoudTextbox2.Text = ""
InhoudTextbox3.Text = ""
InhoudTextbox4.Text = ""
InhoudTextbox5.Text = ""
Text5.Text = ""
InhoudTextbox6.Text = ""
InhoudTextbox7.Text = ""
InhoudTextbox8.Text = ""
InhoudTextbox9.Text = ""
InhoudTextbox10.Text = ""
InhoudTextbox11.Text = ""
InhoudTextbox12.Text = ""
InhoudTextbox13.Text = ""
SpecifiekTextbox1.Text = ""
SpecifiekTextbox2.Text = ""
SpecifiekTextbox3.Text = ""
SpecifiekTextbox4.Text = ""
SpecifiekTextbox5.Text = ""
SpecifiekTextbox6.Text = ""
SpecifiekTextbox7.Text = ""
SpecifiekTextbox8.Text = ""
SpecifiekTextbox9.Text = ""
SpecifiekTextbox10.Text = ""
SpecifiekTextbox11.Text = ""
SpecifiekTextbox12.Text = ""
SpecifiekTextbox13.Text = ""
ItemsTextbox1.Text = ""
ItemsList1.Clear
ItemsTextbox2.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
MetadataTextbox1.Text = ""
MetadataTextbox2.Text = ""
MetadataTextbox3.Text = ""
MetadataTextbox4.Text = ""
MetadataTextbox5.Text = ""
MetadataTextbox6.Text = ""
MetadataTextbox7.Text = ""
MetadataTextbox8.Text = ""
MetadataTextbox9.Text = ""
End Sub
Private Sub AllemaalCmd_Click()
List1.Clear
vullenmetadatalijst
End Sub
Private Sub ArcGISCmd_Click()
Dim lijst As String
Dim nummer As Integer
Dim waarde As String
Dim SQL As String
Dim naam As String
Dim pad As String
Dim soort As String
Dim nwsoort As String
Dim pPropertySet As IPropertySet
Dim pSdeFact As IWorkspaceFactory
Dim pMxDocument As IMxDocument
Dim pActiveView As IActiveView
Dim pFeatureClass As IFeatureClass
Dim pFeatureLayer As IFeatureLayer
Dim pFeatureWorkspace As IFeatureWorkspace
Dim env As IEnvelope
Dim punt As IPoint
Dim minX As Double
Dim minY As Double
Dim maxX As Double
Dim maxY As Double
Dim filePath As String
Dim pGxLayer As IGxLayer
Dim pGxFile As IGxFile
Dim pMap As IMap
Dim Application As IApplication
Dim status As String
Dim alt_titel As String
Dim map As String

On Error GoTo ErrorHandler

If List1.ListIndex = -1 Then
    MsgBox "Selecteer eerst een metadata object uit de lijst!", vbOKOnly, versie
    Exit Sub
End If

Form1.MousePointer = vbHourglass
nummer = List1.ListIndex
waarde = List1.list(nummer)

SQL = "SELECT NAAM FROM DATASET WHERE DATASET_TITEL = """ + waarde + """;"

naam = SelectRecord(SQL)

SQL = "SELECT FYSIEKE_LOCATIE FROM DATASET WHERE DATASET_TITEL = """ + waarde + """;"

status = SelectRecord(SQL)

map = ladir + status + "\" + naam + ".lyr"

If Dir(map) <> "" Then
    Set pMxDocument = m_pApp.Document
    Set pActiveView = pMxDocument.ActiveView
    
    Set env = pActiveView.Extent
    Set punt = env.LowerLeft
    minX = punt.x
    minY = punt.y
    
    Set punt = env.UpperRight
    maxX = punt.x
    maxY = punt.y
    
    filePath = ladir + status + "\" + naam + ".lyr"
   
    Set pGxLayer = New GxLayer
    Set pGxFile = pGxLayer
    pGxFile.Path = filePath
    
    Set pMap = pMxDocument.FocusMap
    pMap.AddLayer pGxLayer.Layer

    pActiveView.Extent.XMin = minX
    pActiveView.Extent.XMax = maxX
    pActiveView.Extent.YMin = minY
    pActiveView.Extent.YMax = maxY

    pMxDocument.UpdateContents
    pMxDocument.ActiveView.Refresh
    
    Form1.MousePointer = vbNormal
Else
    MsgBox "Er is geen opmaakbestand van de geselecteerde dataset aanwezig. De dataset wordt nu zonder opmaak uit de ruimtelijke database gehaald.", vbInformation, "Informatie"
    SQL = "SELECT ALT_TITEL FROM DATASET WHERE DATASET_TITEL = """ + waarde + """;"
    
    alt_titel = SelectRecord(SQL)
        
    If alt_titel <> "" Then
        Set pMxDocument = m_pApp.Document
        Set pActiveView = pMxDocument.ActiveView
        
        Set pPropertySet = New PropertySet
        pPropertySet.SetProperty "SERVER", "chios"
        pPropertySet.SetProperty "INSTANCE", "5151"
        pPropertySet.SetProperty "DATABASE", ""
        pPropertySet.SetProperty "USER", "gisuser"
        pPropertySet.SetProperty "PASSWORD", "zonnetje"
        pPropertySet.SetProperty "VERSION", "SDE.DEFAULT"
        
        Set env = pActiveView.Extent
        Set punt = env.LowerLeft
        minX = punt.x
        minY = punt.y
        Set punt = env.UpperRight
        maxX = punt.x
        maxY = punt.y
        
        Set pSdeFact = New SdeWorkspaceFactory
        Set pFeatureWorkspace = pSdeFact.Open(pPropertySet, Me.hWnd)
        Set pFeatureLayer = New FeatureLayer
        Set pFeatureClass = pFeatureWorkspace.OpenFeatureClass(alt_titel)
        Set pMap = pMxDocument.FocusMap
        
        Set pFeatureLayer = New FeatureLayer
        Set pFeatureLayer.FeatureClass = pFeatureClass
        
        pFeatureLayer.Name = pFeatureClass.AliasName
        pFeatureLayer.Visible = True
        pFeatureLayer.Selectable = True
        pMap.AddLayer pFeatureLayer
        
        pActiveView.Extent.XMin = minX
        pActiveView.Extent.XMax = maxX
        pActiveView.Extent.YMin = minY
        pActiveView.Extent.YMax = maxY
        
        pMxDocument.UpdateContents
        pMxDocument.ActiveView.Refresh
        Form1.MousePointer = vbNormal
        Exit Sub
    Else
        MsgBox "Gegevens van de dataset zijn niet aanwezig"
        Form1.MousePointer = vbNormal
        Exit Sub
    End If
End If
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
  Form1.MousePointer = vbNormal
End Sub
Private Sub Command1_Click()
Dim SQL As String
Dim lijst As String
Dim totaal As Integer
Dim totaalset As Integer

On Error GoTo ErrorHandler

Me.MousePointer = vbHourglass
totaalset = 0

If Text1.Text <> "" Then
    SQL = "SELECT DISTINCT DATASET.DATASET_TITEL " + _
    "FROM ((DATASET INNER JOIN MEMOTABEL ON DATASET.OMSCHRIJVING_CODE = MEMOTABEL.CODE) INNER JOIN TREFCODE ON DATASET.DATACODE = " + _
    "TREFCODE.DATACODE) INNER JOIN TREFTEXT ON TREFCODE.TREFCODE = TREFTEXT.TREFCODE " + _
    "WHERE instr(DATASET.DATASET_TITEL, '" + Text1.Text + "') OR instr(MEMOTABEL.TEKST, '" + Text1.Text + "') OR " + _
    "instr(TREFTEXT.TREFWOORD, '" + Text1.Text + "') AND DATASET.TYPE = 1 AND DATASET.VEILIGHEID = 'vrij toegankelijk' ORDER BY DATASET.DATASET_TITEL;"

 '    SQL = "SELECT DISTINCT DATASET.DATASET_TITEL FROM " + _
 '   "((TREFCODE INNER JOIN DATASET ON TREFCODE.DATACODE = DATASET.DATACODE) INNER JOIN TREFTEXT ON " + _
 '   "TREFCODE.TREFCODE = TREFTEXT.TREFCODE) INNER JOIN MEMOTABEL ON DATASET.DATACODE = MEMOTABEL.CODE " + _
 '   "WHERE (((DATASET.DATASET_TITEL) Like '%" + Text1.Text + "%')) " + _
 '   "OR (((MEMOTABEL.TEKST) Like '%" + Text1.Text + "%')) OR (((TREFTEXT.TREFWOORD) Like '%" + Text1.Text + "%')) AND DATASET.TYPE=1 ORDER BY DATASET.DATASET_TITEL;"

    Form1.List1.Clear
    
    Set rsselect = New ADODB.Recordset
    rsselect.Open SQL, db, adOpenStatic, adLockOptimistic
      
    With rsselect
      If Not .EOF Then
            .MoveFirst
            Do
            If Not IsNull(rsselect.Fields(0)) Then
              lijst = rsselect.Fields(0)
              Form1.List1.AddItem (lijst)
            End If
            rsselect.MoveNext
            totaalset = totaalset + 1
            Loop Until .EOF
        End If
        .Close
    End With
End If

LabelNamen.Caption = "Datasets (" + Str(totaalset) + " van " + Str(totaalmeta) + ")"
Me.MousePointer = vbNormal

Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub Command2_Click()
Dim naam, filename, datacode, dataset, SQL, temp_string, treftekst As String
Dim code_contact As String
Dim item_def As String
Dim temp_begin, temp_eind As String
Dim objPDF As New mjwPDF
Dim soort, x_tref As Integer
Dim x As Double
Dim x_pdf As Double
Dim tabel As String
Dim kolom As String
Dim lijst As String

On Error GoTo ErrorHandler

If AlgemeenTextbox1.Text <> "" Then
  Form1.MousePointer = vbHourglass
  objPDF.PDFFileName = App.Path & "\" + AlgemeenTextbox1.Text + ".pdf"
  objPDF.PDFLoadAfm = App.Path & "\Fonts"
  objPDF.PDFBeginDoc
  'Volledig uitgezoomd op pagina
  objPDF.PDFSetLayoutMode = LAYOUT_SINGLE
  objPDF.PDFFormatPage = FORMAT_A4
  objPDF.PDFOrientation = ORIENT_PORTRAIT
  objPDF.PDFSetUnit = UNIT_PT

  objPDF.PDFView = True

  objPDF.PDFSetDrawColor = "#E0E8F5"
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetDrawMode = DRAW_DRAW
  objPDF.PDFDrawEllipse -100, -90, 700, 1000
  
  objPDF.PDFSetFont FONT_ARIAL, 12, FONT_NORMAL
  objPDF.PDFSetDrawColor = vbWhite
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_CENTER
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell " ", 30, 0, objPDF.PDFGetPageWidth - 30, 125
  
  objPDF.PDFSetDrawColor = vbBlack
  objPDF.PDFDrawLineHor 25, 20, objPDF.PDFGetPageWidth - 30
  
  objPDF.PDFImage App.Path & "\pics\drenthe.jpg", 375, 45, 393 / 2, 58 / 2
  
  objPDF.PDFSetFont FONT_ARIAL, 18, FONT_BOLD
  objPDF.PDFSetDrawColor = "#0A94FC"
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_CENTER
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell " ", 0, 125, objPDF.PDFGetPageWidth + 20, 25
  
  objPDF.PDFSetFont FONT_ARIAL, 18, FONT_BOLD
  objPDF.PDFSetDrawColor = "#0A94FC"
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_CENTER
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell "Overzicht metagegevens", 0, 150, objPDF.PDFGetPageWidth + 20, 25
  
  objPDF.PDFSetFont FONT_ARIAL, 18, FONT_BOLD
  objPDF.PDFSetDrawColor = "#0A94FC"
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_CENTER
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell AlgemeenTextbox1.Text, 0, 175, objPDF.PDFGetPageWidth + 20, 25
  
  objPDF.PDFSetFont FONT_ARIAL, 18, FONT_BOLD
  objPDF.PDFSetDrawColor = "#0A94FC"
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_CENTER
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell " ", 0, 200, objPDF.PDFGetPageWidth + 20, 25
  
  objPDF.PDFSetFont FONT_ARIAL, 18, FONT_BOLD
  objPDF.PDFSetDrawColor = "#E0E8F5"
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_CENTER
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell " ", 0, 225, objPDF.PDFGetPageWidth + 20, 100
  
  objPDF.PDFSetFont FONT_ARIAL, 18, FONT_BOLD
  objPDF.PDFSetDrawColor = "#E0E8F5"
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_CENTER
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell " ", 0, objPDF.PDFGetPageHeight - 100, 200, 100
  
  objPDF.PDFSetDrawColor = vbBlack
  objPDF.PDFDrawLineHor 25, objPDF.PDFGetPageHeight - 40, objPDF.PDFGetPageWidth - 30
  
  objPDF.PDFSetTextColor = vbBlack
  objPDF.PDFSetFont FONT_ARIAL, 8, FONT_NORMAL
  objPDF.PDFTextOut "Geografische metagegevens | provincie Drenthe", 28, objPDF.PDFGetPageHeight - 31
  
  objPDF.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
  objPDF.PDFSetDrawColor = "#0A94FC"
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_RIGHT
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell "Versie " + Str(Date), 0, objPDF.PDFGetPageHeight - 130, objPDF.PDFGetPageWidth + 3, 30
  
  objPDF.PDFImage App.Path & "\pics\wapen.jpg", objPDF.PDFGetPageWidth - 75, objPDF.PDFGetPageHeight - 95, 271 / 4, 196 / 4
  
  objPDF.PDFSetFont FONT_ARIAL, 12, FONT_NORMAL
  objPDF.PDFSetDrawColor = "#007ADB"
  objPDF.PDFSetTextColor = vbWhite
  objPDF.PDFSetAlignement = ALIGN_CENTER
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  objPDF.PDFCell " ", 0, 0, 30, objPDF.PDFGetPageHeight
  
  objPDF.PDFEndPage
  
  objPDF.PDFNewPage
  
  objPDF.PDFSetDrawColor = vbBlack
  objPDF.PDFDrawLineHor 25, 20, objPDF.PDFGetPageWidth - 30
  
  objPDF.PDFSetDrawColor = vbBlack
  objPDF.PDFDrawLineHor 25, objPDF.PDFGetPageHeight - 40, objPDF.PDFGetPageWidth - 30
  
  objPDF.PDFSetTextColor = vbBlack
  objPDF.PDFSetFont FONT_ARIAL, 8, FONT_NORMAL
  objPDF.PDFTextOut "Geografische metagegevens | provincie Drenthe", 28, objPDF.PDFGetPageHeight - 31
   
  x = 50
       
  objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
  objPDF.PDFSetDrawColor = "#E0E8F5"
  objPDF.PDFSetTextColor = vbBlack
  objPDF.PDFSetAlignement = ALIGN_LEFT
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
  
  dataset = AlgemeenTextbox1.Text
  datacode = SelectRecord("SELECT DATACODE FROM DATASET WHERE DATASET_TITEL = '" + dataset + "'")
    
  objPDF.PDFCell "Titel dataset:", 30, x, 120, 14
  If Len(AlgemeenTextbox1.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox1.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Alternatieve titel:", 30, x, 120, 14
  If Len(AlgemeenTextbox2.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox2.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Omschrijving dataset:", 30, x, 120, 14
  If Len(AlgemeenTextbox3.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox3.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
          
  If Len(AlgemeenTextbox3.Text) < 92 Then
    x = x + 14
  End If
  If Len(AlgemeenTextbox3.Text) >= 93 And Len(AlgemeenTextbox3.Text) <= 184 Then
    x = x + 28
    objPDF.PDFCell " ", 30, x - 14, 120, 14
  End If
  If Len(AlgemeenTextbox3.Text) >= 185 And Len(AlgemeenTextbox3.Text) <= 276 Then
    x = x + 42
    objPDF.PDFCell " ", 30, x - 28, 120, 14
    objPDF.PDFCell " ", 30, x - 14, 120, 14
  End If
  If Len(AlgemeenTextbox3.Text) >= 277 And Len(AlgemeenTextbox3.Text) <= 368 Then
    x = x + 56
    objPDF.PDFCell " ", 30, x - 42, 120, 14
    objPDF.PDFCell " ", 30, x - 28, 120, 14
    objPDF.PDFCell " ", 30, x - 14, 120, 14
  End If
  If Len(AlgemeenTextbox3.Text) >= 369 And Len(AlgemeenTextbox3.Text) <= 460 Then
    x = x + 70
    objPDF.PDFCell " ", 30, x - 56, 120, 14
    objPDF.PDFCell " ", 30, x - 42, 120, 14
    objPDF.PDFCell " ", 30, x - 28, 120, 14
    objPDF.PDFCell " ", 30, x - 14, 120, 14
  End If
  If Len(AlgemeenTextbox3.Text) >= 461 And Len(AlgemeenTextbox3.Text) <= 552 Then
    x = x + 84
  End If
    
  objPDF.PDFCell "Algemene opmerking:", 30, x, 120, 14
  If Len(AlgemeenTextbox4.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox4.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Datum opbouw:", 30, x, 120, 14
  If Len(AlgemeenTextbox5.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox5.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Brondatum:", 30, x, 120, 14
  If Len(AlgemeenTextbox6.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox6.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Bronvermelding:", 30, x, 120, 14
  If Len(AlgemeenTextbox7.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox7.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  If Len(AlgemeenTextbox7.Text) < 92 Then
    x = x + 14
  End If
  If Len(AlgemeenTextbox7.Text) >= 93 And Len(AlgemeenTextbox7.Text) <= 184 Then
    x = x + 28
    objPDF.PDFCell " ", 30, x - 14, 120, 14
  End If
    
  objPDF.PDFCell "Opbouwmethode:", 30, x, 120, 14
  If Len(AlgemeenTextbox8.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox8.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Gebeurtenis:", 30, x, 120, 14
  If Len(AlgemeenTextbox9.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox9.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Status:", 30, x, 120, 14
  If Len(AlgemeenTextbox10.Text) > 0 Then
    objPDF.PDFCell AlgemeenTextbox10.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
        
  objPDF.PDFSetLineColor = vbBlack
  objPDF.PDFSetFill = True
  objPDF.PDFSetLineStyle = pPDF_SOLID
  objPDF.PDFSetLineWidth = 1
  objPDF.PDFSetDrawMode = DRAW_NORMAL
  objPDF.PDFDrawPolygon Array(27, 50, objPDF.PDFGetPageWidth - 33, 50, objPDF.PDFGetPageWidth - 33, x + 14, 27, x + 14)
    
  x = x + 42
  x_pdf = x
    
  objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
  objPDF.PDFSetDrawColor = "#E0E8F5"
  objPDF.PDFSetTextColor = vbBlack
  objPDF.PDFSetAlignement = ALIGN_LEFT
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
    
  objPDF.PDFCell "Contactpersoon inhoud:", 30, x, 120, 14
  If Len(InhoudTextbox1.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox1.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Beleidsterrein:", 30, x, 120, 14
  If Len(InhoudTextbox2.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox2.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Team:", 30, x, 120, 14
  If Len(InhoudTextbox3.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox3.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Thema:", 30, x, 120, 14
  If Len(InhoudTextbox4.Text) > 0 Then
      objPDF.PDFCell InhoudTextbox4.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Gebruiksbeperking:", 30, x, 120, 14
  If Len(InhoudTextbox5.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox5.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Toegangsrestricties:", 30, x, 120, 14
  If Len(InhoudTextbox6.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox6.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Copyright:", 30, x, 120, 14
  If Len(InhoudTextbox7.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox7.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Herzienings frequentie:", 30, x, 120, 14
  If Len(InhoudTextbox8.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox8.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Toepassingsschaal:", 30, x, 120, 14
  If Len(InhoudTextbox9.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox9.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Contact leverancier:", 30, x, 120, 14
  If Len(InhoudTextbox10.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox10.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  If Len(InhoudTextbox10.Text) < 92 Then
    x = x + 14
  End If
  If Len(InhoudTextbox10.Text) >= 93 And Len(InhoudTextbox10.Text) <= 184 Then
    x = x + 28
    objPDF.PDFCell " ", 30, x - 14, 120, 14
  End If
       
     
  objPDF.PDFCell "Trefwoorden:", 30, x, 120, 14
  If Len(InhoudTextbox11.Text) > 0 Then
    objPDF.PDFCell InhoudTextbox11.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  If Len(InhoudTextbox11.Text) < 92 Then
    x = x + 14
  End If
  If Len(InhoudTextbox11.Text) >= 93 And Len(InhoudTextbox11.Text) <= 184 Then
    x = x + 28
    objPDF.PDFCell " ", 30, x - 14, 120, 14
  End If
      
  objPDF.PDFCell "Temporele dekking:", 30, x, 120, 14
  If InhoudTextbox12.Text <> "" Then
    temp_begin = InhoudTextbox12.Text
  Else
    temp_begin = "niet ingevuld"
  End If
        
  If InhoudTextbox13.Text <> "" Then
    temp_eind = InhoudTextbox13.Text
  Else
    temp_eind = "niet ingevuld"
  End If
    
  objPDF.PDFCell "Begin datum: " + temp_begin + ", Eind datum: " + temp_eind, 150, x, objPDF.PDFGetPageWidth - 180, 14
  
  objPDF.PDFSetLineColor = vbBlack
  objPDF.PDFSetFill = True
  objPDF.PDFSetLineStyle = pPDF_SOLID
  objPDF.PDFSetLineWidth = 1
  objPDF.PDFSetDrawMode = DRAW_NORMAL
  objPDF.PDFDrawPolygon Array(27, x_pdf, objPDF.PDFGetPageWidth - 33, x_pdf, objPDF.PDFGetPageWidth - 33, x + 14, 27, x + 14)
    
  x = x + 42
  x_pdf = x
    
  objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
  objPDF.PDFSetDrawColor = "#E0E8F5"
  objPDF.PDFSetTextColor = vbBlack
  objPDF.PDFSetAlignement = ALIGN_LEFT
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
    
  objPDF.PDFCell "Geografisch gebied:", 30, x, 120, 14
  If Len(SpecifiekTextbox1.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox1.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Ruimtelijk schema:", 30, x, 120, 14
  If Len(SpecifiekTextbox2.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox2.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Aanvullende informatie:", 30, x, 120, 14
  If Len(SpecifiekTextbox3.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox3.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  If Len(SpecifiekTextbox3.Text) < 92 Then
    x = x + 14
  End If
  If Len(SpecifiekTextbox3.Text) >= 93 And Len(SpecifiekTextbox3.Text) <= 184 Then
    x = x + 28
  End If
    
  objPDF.PDFCell "Layernaam:", 30, x, 120, 14
  If Len(SpecifiekTextbox4.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox4.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Fysieke locatie:", 30, x, 120, 14
  If Len(SpecifiekTextbox5.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox5.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Datatype:", 30, x, 120, 14
  If Len(SpecifiekTextbox6.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox6.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Geometrie:", 30, x, 120, 14
  If Len(SpecifiekTextbox13.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox13.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Nauwkeurigheid:", 30, x, 120, 14
  If Len(SpecifiekTextbox7.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox7.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Hiërarchieniveau:", 30, x, 120, 14
  If Len(SpecifiekTextbox8.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox8.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Minimale x-coördinaat:", 30, x, 120, 14
  If Len(SpecifiekTextbox9.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox9.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell "niet ingevuld", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Maximale x-coördinaat:", 30, x, 120, 14
  If Len(SpecifiekTextbox10.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox10.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell "niet ingevuld", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Minimale y-coördinaat:", 30, x, 120, 14
  If Len(SpecifiekTextbox11.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox11.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell "niet ingevuld", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Maximale y-coördinaat:", 30, x, 120, 14
  If Len(SpecifiekTextbox12.Text) > 0 Then
    objPDF.PDFCell SpecifiekTextbox12.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell "niet ingevuld", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
    
  objPDF.PDFSetLineColor = vbBlack
  objPDF.PDFSetFill = True
  objPDF.PDFSetLineStyle = pPDF_SOLID
  objPDF.PDFSetLineWidth = 1
  objPDF.PDFSetDrawMode = DRAW_NORMAL
  objPDF.PDFDrawPolygon Array(27, x_pdf, objPDF.PDFGetPageWidth - 33, x_pdf, objPDF.PDFGetPageWidth - 33, x + 14, 27, x + 14)
  
    
  objPDF.PDFEndPage
  
  objPDF.PDFNewPage
  
  objPDF.PDFSetDrawColor = vbBlack
  objPDF.PDFDrawLineHor 25, 20, objPDF.PDFGetPageWidth - 30
  
  objPDF.PDFSetDrawColor = vbBlack
  objPDF.PDFDrawLineHor 25, objPDF.PDFGetPageHeight - 40, objPDF.PDFGetPageWidth - 30
  
  objPDF.PDFSetTextColor = vbBlack
  objPDF.PDFSetFont FONT_ARIAL, 8, FONT_NORMAL
  objPDF.PDFTextOut "Geografische metagegevens | provincie Drenthe", 28, objPDF.PDFGetPageHeight - 31
    
  x = 50
    
  objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
  objPDF.PDFSetDrawColor = "#E0E8F5"
  objPDF.PDFSetTextColor = vbBlack
  objPDF.PDFSetAlignement = ALIGN_LEFT
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
    
 
  objPDF.PDFCell "Standaarditem:", 30, x, 120, 14
  If Len(ItemsTextbox1.Text) > 0 Then
    objPDF.PDFCell ItemsTextbox1.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  If Not (ItemsTextbox1.Text = "") Then
    item_def = SelectRecord("SELECT ITEMDEFINITIE FROM ITEMS WHERE ITEMS.DATACODE = " + datacode + " AND ITEMS.ITEMNAAM = '" + ItemsTextbox1.Text + "'")
    objPDF.PDFCell "Definitie standaarditem:", 30, x, 120, 14
    objPDF.PDFCell item_def, 150, x, objPDF.PDFGetPageWidth - 180, 14
    x = x + 14
  End If
    
  objPDF.PDFCell "Items:", 30, x, 120, 14
  objPDF.PDFCell "Itemnaam", 150, x, 240, 14
  objPDF.PDFCell "Itemsdefinitie", 240, x, objPDF.PDFGetPageWidth - 270, 14
  x = x + 14
         
  Set rsselect = New ADODB.Recordset
    rsselect.Open "SELECT ITEMNAAM, ITEMDEFINITIE FROM ITEMS WHERE DATACODE=" + datacode + " ORDER BY VOLGNR", db
    If Not (rsselect.EOF) Then
      rsselect.MoveFirst
      Do
        objPDF.PDFCell rsselect.Fields(0), 150, x, 240, 14
        objPDF.PDFCell rsselect.Fields(1), 240, x, objPDF.PDFGetPageWidth - 270, 14
        x = x + 14
        rsselect.MoveNext
      Loop Until rsselect.EOF
      rsselect.Close
    End If
  
  objPDF.PDFSetLineColor = vbBlack
  objPDF.PDFSetFill = True
  objPDF.PDFSetLineStyle = pPDF_SOLID
  objPDF.PDFSetLineWidth = 1
  objPDF.PDFSetDrawMode = DRAW_NORMAL
  objPDF.PDFDrawPolygon Array(27, 50, objPDF.PDFGetPageWidth - 33, 50, objPDF.PDFGetPageWidth - 33, x, 27, x)
    
  x = x + 42
  x_pdf = x
    
  objPDF.PDFSetFont FONT_ARIAL, 10, FONT_NORMAL
  objPDF.PDFSetDrawColor = "#E0E8F5"
  objPDF.PDFSetTextColor = vbBlack
  objPDF.PDFSetAlignement = ALIGN_LEFT
  objPDF.PDFSetBorder = BORDER_NONE
  objPDF.PDFSetFill = True
    
  objPDF.PDFCell "Contact metadata:", 30, x, 120, 14
  If Len(ItemsTextbox1.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox1.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Datum opbouw:", 30, x, 120, 14
  If Len(MetadataTextbox2.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox2.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Taal dataset:", 30, x, 120, 14
  If Len(MetadataTextbox8.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox8.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Karakterset:", 30, x, 120, 14
  If Len(MetadataTextbox9.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox9.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Metadatastandaard:", 30, x, 120, 14
  If Len(MetadataTextbox3.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox3.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Versie:", 30, x, 120, 14
  If Len(MetadataTextbox4.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox4.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Referentiesysteem:", 30, x, 120, 14
  If Len(MetadataTextbox5.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox5.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Referentie organisatie:", 30, x, 120, 14
  If Len(MetadataTextbox6.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox6.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
    
  objPDF.PDFCell "Contact distributie:", 30, x, 120, 14
  If Len(MetadataTextbox7.Text) > 0 Then
    objPDF.PDFCell MetadataTextbox7.Text, 150, x, objPDF.PDFGetPageWidth - 180, 14
  Else
    objPDF.PDFCell " ", 150, x, objPDF.PDFGetPageWidth - 180, 14
  End If
  x = x + 14
  
  objPDF.PDFSetLineColor = vbBlack
  objPDF.PDFSetFill = True
  objPDF.PDFSetLineStyle = pPDF_SOLID
  objPDF.PDFSetLineWidth = 1
  objPDF.PDFSetDrawMode = DRAW_NORMAL
  objPDF.PDFDrawPolygon Array(27, x_pdf, objPDF.PDFGetPageWidth - 33, x_pdf, objPDF.PDFGetPageWidth - 33, x, 27, x)
    
  objPDF.PDFEndPage
  objPDF.PDFEndDoc
End If
Form1.MousePointer = vbNormal

Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
    "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
  Form1.MousePointer = vbNormal
End Sub
Private Sub Command3_Click()
frmAbout.Show vbModal
End Sub
Private Sub Form_Load()
initialiseer
Set db = New ADODB.Connection
db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0.;Persist Security Info=False;Data Source= " & dbdir
db.Open
vullenmetadatalijst
End Sub
Private Sub Form_Close()
If (db.State = adStateOpen) Then
  db.Close
End If
'Set Form1 = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
If db.State = adStateOpen Then
  db.Close
End If
Unload Me
'Set Form1 = Nothing
End Sub
Private Sub ItemsList1_DblClick()
Dim SQL, waarde, waarde1, dc As String


On Error GoTo ErrorHandler

waarde1 = List1.Text
dc = SelectRecord("SELECT DATACODE FROM DATASET WHERE DATASET_TITEL = '" + waarde1 + "'")

waarde = ItemsList1.Text

SQL = "SELECT a.VOLGNR, a.ITEMNAAM, a.ITEMDEFINITIE, a.EENHEID, b.TEKST FROM ITEMS a, MEMOTABEL b WHERE a.DOMEIN = b.CODE AND a.DATACODE = " + dc + " AND a.ITEMNAAM = '" + waarde + "'"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

If rs.EOF = False Then
  If IsNull(rs.Fields("ITEMNAAM")) Then
      ItemsTextbox2.Text = ""
  Else
      ItemsTextbox2.Text = rs.Fields("ITEMNAAM")
  End If
  If IsNull(rs.Fields("ITEMDEFINITIE")) Then
      Text2.Text = ""
  Else
      Text2.Text = rs.Fields("ITEMDEFINITIE")
  End If
  If IsNull(rs.Fields("EENHEID")) Then
      Text3.Text = ""
  Else
      Text3.Text = rs.Fields("EENHEID")
  End If
  If IsNull(rs.Fields("TEKST")) Then
      Text4.Text = ""
  Else
      Text4.Text = rs.Fields("TEKST")
  End If
Else
  MsgBox "Geen gegevens over item bekend.", vbInformation + vbOKOnly, "Melding"
End If
rs.Close

Exit Sub
ErrorHandler:
    MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
    Form1.MousePointer = vbNormal
End Sub
Private Sub List1_DblClick()
Dim nummer As Integer
Dim waarde As String
Dim trefwaarde As String
Dim SQL As String
Dim test As String
Dim I As Long
Dim laag As Integer
Dim teller As Integer
Dim tekst As String
Dim stditem As String
Dim datacode As String
Dim x As Integer
Dim metapersoon As String
Dim geoloket As String
Dim kolom As String
Dim lijst As String

On Error GoTo ErrorHandler

Form1.MousePointer = vbHourglass
nummer = List1.ListIndex
waarde = List1.list(nummer)

Opschonen

' Gegevens voor tabblad algemeen

SQL = "SELECT a.DATASET_TITEL, a.ALT_TITEL, b.TEKST, a.OPMERKING, a.OPBOUWDATUM, a.BRONDATUM, c.BRONVERMELDING, " + _
"c.OPBOUWMETHODE , a.ACTIE, a.STATUS FROM DATASET a, MEMOTABEL b, GEOGRAFISCH c WHERE c.DATACODE = a.DATACODE AND " + _
"a.OMSCHRIJVING_CODE = b.Code AND a.DATASET_TITEL = '" + waarde + "';"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

If IsNull(rs.Fields("DATASET_TITEL")) Then
    AlgemeenTextbox1.Text = ""
Else
    AlgemeenTextbox1.Text = rs.Fields("DATASET_TITEL")
End If

If IsNull(rs.Fields("ALT_TITEL")) Then
    AlgemeenTextbox2.Text = ""
Else
    AlgemeenTextbox2.Text = rs.Fields("ALT_TITEL")
End If

If IsNull(rs.Fields("TEKST")) Then
    AlgemeenTextbox3.Text = ""
Else
    AlgemeenTextbox3.Text = rs.Fields("TEKST")
End If

If IsNull(rs.Fields("OPMERKING")) Then
    AlgemeenTextbox4.Text = ""
Else
    AlgemeenTextbox4.Text = rs.Fields("OPMERKING")
End If

If IsNull(rs.Fields("BRONDATUM")) Then
    AlgemeenTextbox5.Text = ""
Else
    AlgemeenTextbox5.Text = rs.Fields("BRONDATUM")
End If

If IsNull(rs.Fields("OPBOUWDATUM")) Then
    AlgemeenTextbox6.Text = ""
Else
    AlgemeenTextbox6.Text = rs.Fields("OPBOUWDATUM")
End If

If IsNull(rs.Fields("BRONVERMELDING")) Then
    AlgemeenTextbox7.Text = ""
Else
    AlgemeenTextbox7.Text = rs.Fields("BRONVERMELDING")
End If

If IsNull(rs.Fields("OPBOUWMETHODE")) Then
    AlgemeenTextbox8.Text = ""
Else
    AlgemeenTextbox8.Text = rs.Fields("OPBOUWMETHODE")
End If

If IsNull(rs.Fields("ACTIE")) Then
    AlgemeenTextbox9.Text = ""
Else
    AlgemeenTextbox9.Text = rs.Fields("ACTIE")
End If

If IsNull(rs.Fields("STATUS")) Then
    AlgemeenTextbox10.Text = ""
Else
    AlgemeenTextbox10.Text = rs.Fields("STATUS")
End If

rs.Close

SQL = "SELECT a.DATACODE, a.BELEIDSVELD, a.VEILIGHEID, a.TEAM, a.THEMA, a.GEBRUIKSBEPERKING, a.JURIDISCH, a.COPYRIGHT, a.BIJHOUDING, b.SCHAAL, " + _
"a.CONTACT_LEVERANCIER, a.DEKKING_BEGIN, a.DEKKING_EIND, c.CONTACTPERSOON FROM DATASET a, GEOGRAFISCH b, CONTACT c WHERE a.DATACODE " + _
"= b.DATACODE AND a.CONTACTPERSOON = c.CONTACT_ID AND a.DATASET_TITEL = '" + waarde + "';"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

datacode = Val(rs.Fields("DATACODE"))

If IsNull(rs.Fields("CONTACTPERSOON")) Then
    InhoudTextbox1.Text = ""
Else
    InhoudTextbox1.Text = rs.Fields("CONTACTPERSOON")
End If

If IsNull(rs.Fields("BELEIDSVELD")) Then
    InhoudTextbox2.Text = ""
Else
    InhoudTextbox2.Text = rs.Fields("BELEIDSVELD")
End If

If IsNull(rs.Fields("TEAM")) Then
    InhoudTextbox3.Text = ""
Else
    InhoudTextbox3.Text = rs.Fields("TEAM")
End If

If IsNull(rs.Fields("THEMA")) Then
    InhoudTextbox4.Text = ""
Else
    InhoudTextbox4.Text = rs.Fields("THEMA")
End If

If IsNull(rs.Fields("GEBRUIKSBEPERKING")) Then
    InhoudTextbox5.Text = ""
Else
    InhoudTextbox5.Text = rs.Fields("GEBRUIKSBEPERKING")
End If

If IsNull(rs.Fields("VEILIGHEID")) Then
    Text5.Text = ""
Else
    Text5.Text = rs.Fields("VEILIGHEID")
End If

If IsNull(rs.Fields("JURIDISCH")) Then
    InhoudTextbox6.Text = ""
Else
    InhoudTextbox6.Text = rs.Fields("JURIDISCH")
End If

If IsNull(rs.Fields("COPYRIGHT")) Then
    InhoudTextbox7.Text = ""
Else
    InhoudTextbox7.Text = rs.Fields("COPYRIGHT")
End If

If IsNull(rs.Fields("BIJHOUDING")) Then
    InhoudTextbox8.Text = ""
Else
    InhoudTextbox8.Text = rs.Fields("BIJHOUDING")
End If

If IsNull(rs.Fields("SCHAAL")) Then
    InhoudTextbox9.Text = ""
Else
    InhoudTextbox9.Text = rs.Fields("SCHAAL")
End If

If IsNull(rs.Fields("CONTACT_LEVERANCIER")) Then
    InhoudTextbox10.Text = ""
Else
    InhoudTextbox10.Text = rs.Fields("CONTACT_LEVERANCIER")
End If

If IsNull(rs.Fields("DEKKING_BEGIN")) Then
    InhoudTextbox12.Text = ""
Else
    InhoudTextbox12.Text = rs.Fields("DEKKING_BEGIN")
End If

If IsNull(rs.Fields("DEKKING_EIND")) Then
    InhoudTextbox13.Text = ""
Else
    InhoudTextbox13.Text = rs.Fields("DEKKING_EIND")
End If

rs.Close

SQL = "SELECT a.TREFWOORD FROM TREFTEXT a, TREFCODE b WHERE " + _
   "b.TREFCODE = a.TREFCODE AND b.DATACODE = " + Str(datacode) + " ORDER BY a.TREFWOORD;"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

x = 0
With rs
If .RecordCount > 0 Then
    .MoveFirst
    Do
      If x = 0 Then
        InhoudTextbox11.Text = .Fields("TREFWOORD")
        x = x + 1
      Else
        InhoudTextbox11.Text = InhoudTextbox11.Text + ", " + .Fields("TREFWOORD")
      End If
      .MoveNext
    Loop Until .EOF
  End If
  .Close
End With

SQL = "SELECT b.DEELGEBIED, a.RSCHEMA, a.AANVUL_INFO, a.NAAM, a.FYSIEKE_LOCATIE, a.DATATYPE, " + _
"b.POS_NAUWKEURIGHEID , a.KWALITEIT_BESCH, b.GEOMETRIE " + _
"FROM DATASET a, GEOGRAFISCH b WHERE a.DATACODE = b.DATACODE AND a.DATASET_TITEL = '" + waarde + "';"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

If IsNull(rs.Fields("DEELGEBIED")) Then
    SpecifiekTextbox1.Text = ""
Else
    SpecifiekTextbox1.Text = rs.Fields("DEELGEBIED")
End If

If IsNull(rs.Fields("RSCHEMA")) Then
    SpecifiekTextbox2.Text = ""
Else
    SpecifiekTextbox2.Text = rs.Fields("RSCHEMA")
End If

If IsNull(rs.Fields("AANVUL_INFO")) Then
    SpecifiekTextbox3.Text = ""
Else
    SpecifiekTextbox3.Text = rs.Fields("AANVUL_INFO")
End If

If IsNull(rs.Fields("NAAM")) Then
    SpecifiekTextbox4.Text = ""
Else
    SpecifiekTextbox4.Text = rs.Fields("NAAM")
End If

If IsNull(rs.Fields("FYSIEKE_LOCATIE")) Then
    SpecifiekTextbox5.Text = ""
Else
    SpecifiekTextbox5.Text = rs.Fields("FYSIEKE_LOCATIE")
End If

If IsNull(rs.Fields("DATATYPE")) Then
    SpecifiekTextbox6.Text = ""
Else
    SpecifiekTextbox6.Text = rs.Fields("DATATYPE")
End If

If IsNull(rs.Fields("GEOMETRIE")) Then
    SpecifiekTextbox13.Text = ""
Else
    SpecifiekTextbox13.Text = rs.Fields("GEOMETRIE")
End If

If IsNull(rs.Fields("POS_NAUWKEURIGHEID")) Then
    SpecifiekTextbox7.Text = ""
Else
    SpecifiekTextbox7.Text = rs.Fields("POS_NAUWKEURIGHEID")
End If

If IsNull(rs.Fields("KWALITEIT_BESCH")) Then
    SpecifiekTextbox8.Text = ""
Else
    SpecifiekTextbox8.Text = rs.Fields("KWALITEIT_BESCH")
End If

If Not IsNull(rs.Fields("DEELGEBIED")) Then
  SpecifiekTextbox9.Text = SelectRecord("SELECT MIN_X FROM GEBIED WHERE GEBIED = '" + rs.Fields("DEELGEBIED") + "'")
  SpecifiekTextbox10.Text = SelectRecord("SELECT MAX_X FROM GEBIED WHERE GEBIED = '" + rs.Fields("DEELGEBIED") + "'")
  SpecifiekTextbox11.Text = SelectRecord("SELECT MIN_Y FROM GEBIED WHERE GEBIED = '" + rs.Fields("DEELGEBIED") + "'")
  SpecifiekTextbox12.Text = SelectRecord("SELECT MAX_Y FROM GEBIED WHERE GEBIED = '" + rs.Fields("DEELGEBIED") + "'")
Else
  SpecifiekTextbox9.Text = ""
  SpecifiekTextbox10.Text = ""
  SpecifiekTextbox11.Text = ""
  SpecifiekTextbox12.Text = ""
End If
  
rs.Close

SQL = "SELECT a.DATACODE, b.STD_ITEM FROM DATASET a, GEOGRAFISCH b WHERE " + _
"a.DATACODE = b.DATACODE AND a.DATASET_TITEL = '" + waarde + "';"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

x = 0
stditem = ""
With rs
If .RecordCount > 0 Then
    .MoveFirst
    Do
      datacode = .Fields("DATACODE") & ""
      stditem = .Fields("STD_ITEM") & ""
      .MoveNext
    Loop Until .EOF
  End If
  .Close
End With

If stditem = "" Then
    ItemsTextbox1.Text = ""
Else
    ItemsTextbox1.Text = stditem
End If

If stditem <> "" Then
VullenListbox "SELECT ITEMNAAM FROM ITEMS WHERE DATACODE = " + datacode + " ORDER BY VOLGNR", 0, ItemsList1
End If

SQL = "SELECT a.METAPERSOON, a.OPBOUWDATUM, a.METADATASTD, a.TAAL, a.KARAKTERSET, a.VERSIE_METASTD, a.CODE_REF, a.ORG_NAMESPACE, " + _
"a.GEOLOKET FROM DATASET a WHERE a.DATASET_TITEL = '" + waarde + "';"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

If IsNull(rs.Fields("OPBOUWDATUM")) Then
    MetadataTextbox2.Text = ""
Else
    MetadataTextbox2.Text = rs.Fields("OPBOUWDATUM")
End If

If IsNull(rs.Fields("TAAL")) Then
    MetadataTextbox8.Text = ""
Else
    MetadataTextbox8.Text = rs.Fields("TAAL")
End If

If IsNull(rs.Fields("KARAKTERSET")) Then
    MetadataTextbox9.Text = ""
Else
    MetadataTextbox9.Text = rs.Fields("KARAKTERSET")
End If

If IsNull(rs.Fields("METADATASTD")) Then
    MetadataTextbox3.Text = ""
Else
    MetadataTextbox3.Text = rs.Fields("METADATASTD")
End If

If IsNull(rs.Fields("VERSIE_METASTD")) Then
    MetadataTextbox4.Text = ""
Else
    MetadataTextbox4.Text = rs.Fields("VERSIE_METASTD")
End If

If IsNull(rs.Fields("CODE_REF")) Then
    MetadataTextbox5.Text = ""
Else
    MetadataTextbox5.Text = rs.Fields("CODE_REF")
End If
If IsNull(rs.Fields("ORG_NAMESPACE")) Then
    MetadataTextbox6.Text = ""
Else
    MetadataTextbox6.Text = rs.Fields("ORG_NAMESPACE")
End If

If Not IsNull(rs.Fields("METAPERSOON")) Then
    metapersoon = rs.Fields("METAPERSOON")
End If
If Not IsNull(rs.Fields("GEOLOKET")) Then
    geoloket = rs.Fields("GEOLOKET")
End If

rs.Close

SQL = "SELECT CONTACTPERSOON FROM CONTACT WHERE CONTACT_ID = " + metapersoon + ";"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

With rs
If .RecordCount > 0 Then
    .MoveFirst
    Do
      MetadataTextbox1.Text = .Fields("CONTACTPERSOON")
      .MoveNext
    Loop Until .EOF
  End If
  .Close
End With

SQL = "SELECT CONTACTPERSOON FROM CONTACT WHERE CONTACT_ID = " + geoloket + ";"

Set rs = New ADODB.Recordset
rs.Open SQL, db, adOpenStatic, adLockOptimistic

With rs
If .RecordCount > 0 Then
    .MoveFirst
    Do
      MetadataTextbox7.Text = .Fields("CONTACTPERSOON")
      .MoveNext
    Loop Until .EOF
  End If
  .Close
End With

Form1.MousePointer = vbNormal
Command2.Enabled = True

Exit Sub
ErrorHandler:
    MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
    Form1.MousePointer = vbNormal
End Sub
Private Sub SluitCmd_Click()
If db.State = adStateOpen Then
  db.Close
End If
Unload Me
'Set Form1 = Nothing
End Sub
Private Sub WordCmd_Click()
Dim w1 As Word.Application
Dim otable As Word.Table
Dim teller As Integer
Dim aantalrij As Integer
Dim nummer As Integer
Dim waarde As String
Dim I As Integer

On Error GoTo ErrorHandler

If List1.ListIndex = -1 Then
    MsgBox "Selecteer eerst een metadata object uit de lijst!", vbOKOnly, versie
    Exit Sub
End If

nummer = List1.ListIndex
waarde = List1.list(nummer)

Set w1 = New Word.Application
w1.Visible = True
w1.WindowState = wdWindowStateMaximize
w1.Documents.Add

If w1.Documents.Count < 1 Then
    MsgBox "Er zijn geen documenten open!"
    Exit Sub
End If

w1.Selection.Font.Size = 10
w1.Selection.Font.Bold = True
w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
w1.Selection.TypeText "Provincie Drenthe"
w1.Selection.TypeParagraph
w1.Selection.Font.Bold = False
w1.Selection.TypeText "Westerbrink 1"
w1.Selection.TypeParagraph
w1.Selection.TypeText "Postbus 122"
w1.Selection.TypeParagraph
w1.Selection.TypeText "9400 AC Assen"
w1.Selection.TypeParagraph
w1.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
w1.Selection.TypeParagraph
w1.Selection.Font.Bold = True
w1.Selection.Font.Size = 12
w1.Selection.TypeText "GBI Metagegevens: " + AlgemeenTextbox1.Text + " per " + Str(Date)
w1.Selection.Font.Bold = False
w1.Selection.Font.Size = 10

teller = w1.Selection.StoryLength
teller = teller - 1

Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(teller, teller), 7, 2)
otable.Borders.Enable = 0
w1.ActiveDocument.Tables(1).Columns(1).Width = 150
w1.ActiveDocument.Tables(1).Columns(2).Width = 300

w1.ActiveDocument.Tables(1).Cell(1, 1).Merge w1.ActiveDocument.Tables(1).Cell(1, 2)
w1.ActiveDocument.Tables(1).Cell(1, 1).Range.Font.Bold = True
w1.ActiveDocument.Tables(1).Cell(1, 1).Range.Font.Size = 10
w1.ActiveDocument.Tables(1).Cell(1, 1).Range.Font.Italic = True
w1.ActiveDocument.Tables(1).Cell(1, 1).Range.InsertAfter "Overzicht:"

w1.ActiveDocument.Tables(1).Cell(2, 1).Range.InsertAfter "Bestandstitel:"
w1.ActiveDocument.Tables(1).Cell(2, 2).Range.InsertAfter AlgemeenTextbox1.Text
w1.ActiveDocument.Tables(1).Cell(3, 1).Range.InsertAfter "Omschrijving:"
w1.ActiveDocument.Tables(1).Cell(3, 2).Range.InsertAfter AlgemeenTextbox2.Text
w1.ActiveDocument.Tables(1).Cell(4, 1).Range.InsertAfter "Bestandsnaam:"
w1.ActiveDocument.Tables(1).Cell(4, 2).Range.InsertAfter AlgemeenTextbox3.Text
w1.ActiveDocument.Tables(1).Cell(5, 1).Range.InsertAfter "Opbouwdatum:"
w1.ActiveDocument.Tables(1).Cell(5, 2).Range.InsertAfter AlgemeenTextbox4.Text
w1.ActiveDocument.Tables(1).Cell(6, 1).Range.InsertAfter "Bronvermelding:"
w1.ActiveDocument.Tables(1).Cell(6, 2).Range.InsertAfter AlgemeenTextbox5.Text
w1.ActiveDocument.Tables(1).Cell(7, 1).Range.InsertAfter "Brondatum:"
w1.ActiveDocument.Tables(1).Cell(7, 2).Range.InsertAfter AlgemeenTextbox6.Text

w1.Selection.GoTo wdGoToLine, wdGoToRelative, teller
w1.Selection.TypeParagraph

teller = w1.Selection.StoryLength
teller = teller - 1

Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(teller, teller), 6, 2)
otable.Borders.Enable = 0
w1.ActiveDocument.Tables(2).Columns(1).Width = 150
w1.ActiveDocument.Tables(2).Columns(2).Width = 300

w1.ActiveDocument.Tables(2).Cell(1, 1).Merge w1.ActiveDocument.Tables(2).Cell(1, 2)
w1.ActiveDocument.Tables(2).Cell(1, 1).Range.Font.Bold = True
w1.ActiveDocument.Tables(2).Cell(1, 1).Range.Font.Size = 10
w1.ActiveDocument.Tables(2).Cell(1, 1).Range.Font.Italic = True
w1.ActiveDocument.Tables(2).Cell(1, 1).Range.InsertAfter "Algemeen:"

w1.ActiveDocument.Tables(2).Cell(2, 1).Range.InsertAfter "Titel:"
w1.ActiveDocument.Tables(2).Cell(2, 2).Range.InsertAfter AlgemeenTextbox1.Text
w1.ActiveDocument.Tables(2).Cell(3, 1).Range.InsertAfter "Omschrijving:"
w1.ActiveDocument.Tables(2).Cell(3, 2).Range.InsertAfter AlgemeenTextbox2.Text
w1.ActiveDocument.Tables(2).Cell(4, 1).Range.InsertAfter "Covernaam:"
w1.ActiveDocument.Tables(2).Cell(4, 2).Range.InsertAfter AlgemeenTextbox3.Text
w1.ActiveDocument.Tables(2).Cell(5, 1).Range.InsertAfter "Gebruiksbeperking:"
w1.ActiveDocument.Tables(2).Cell(5, 2).Range.InsertAfter AlgemeenTextbox4.Text
w1.ActiveDocument.Tables(2).Cell(6, 1).Range.InsertAfter "Status bestand:"
w1.ActiveDocument.Tables(2).Cell(6, 2).Range.InsertAfter AlgemeenTextbox5.Text

w1.Selection.GoTo wdGoToLine, wdGoToRelative, teller
w1.Selection.TypeParagraph

teller = w1.Selection.StoryLength
teller = teller - 1

Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(teller, teller), 6, 2)
otable.Borders.Enable = 0
w1.ActiveDocument.Tables(3).Columns(1).Width = 150
w1.ActiveDocument.Tables(3).Columns(2).Width = 300

w1.ActiveDocument.Tables(3).Cell(1, 1).Merge w1.ActiveDocument.Tables(3).Cell(1, 2)
w1.ActiveDocument.Tables(3).Cell(1, 1).Range.Font.Bold = True
w1.ActiveDocument.Tables(3).Cell(1, 1).Range.Font.Size = 10
w1.ActiveDocument.Tables(3).Cell(1, 1).Range.Font.Italic = True
w1.ActiveDocument.Tables(3).Cell(1, 1).Range.InsertAfter "Toegang:"

w1.ActiveDocument.Tables(3).Cell(2, 1).Range.InsertAfter "Bestandsnaam:"
'w1.ActiveDocument.Tables(3).Cell(2, 2).Range.InsertAfter ToegangTextbox1.Text
'w1.ActiveDocument.Tables(3).Cell(3, 1).Range.InsertAfter "Padnaam bestand:"
'w1.ActiveDocument.Tables(3).Cell(3, 2).Range.InsertAfter ToegangTextbox2.Text
'w1.ActiveDocument.Tables(3).Cell(4, 1).Range.InsertAfter "Contactpersoon inhoud:"
'w1.ActiveDocument.Tables(3).Cell(4, 2).Range.InsertAfter ToegangTextbox3.Text
'w1.ActiveDocument.Tables(3).Cell(5, 1).Range.InsertAfter "Copyright:"
'w1.ActiveDocument.Tables(3).Cell(5, 2).Range.InsertAfter ToegangTextbox4.Text
'w1.ActiveDocument.Tables(3).Cell(6, 1).Range.InsertAfter "Nummer dataset:"
'w1.ActiveDocument.Tables(3).Cell(6, 2).Range.InsertAfter ToegangTextbox5.Text

w1.Selection.GoTo wdGoToLine, wdGoToRelative, teller
w1.Selection.TypeParagraph

teller = w1.Selection.StoryLength
teller = teller - 1

Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(teller, teller), 7, 2)
otable.Borders.Enable = 0
w1.ActiveDocument.Tables(4).Columns(1).Width = 150
w1.ActiveDocument.Tables(4).Columns(2).Width = 300

w1.ActiveDocument.Tables(4).Cell(1, 1).Merge w1.ActiveDocument.Tables(4).Cell(1, 2)
w1.ActiveDocument.Tables(4).Cell(1, 1).Range.Font.Bold = True
w1.ActiveDocument.Tables(4).Cell(1, 1).Range.Font.Size = 10
w1.ActiveDocument.Tables(4).Cell(1, 1).Range.Font.Italic = True
w1.ActiveDocument.Tables(4).Cell(1, 1).Range.InsertAfter "Inhoud:"

w1.ActiveDocument.Tables(4).Cell(2, 1).Range.InsertAfter "Brondatum:"
w1.ActiveDocument.Tables(4).Cell(2, 2).Range.InsertAfter InhoudTextbox1.Text
w1.ActiveDocument.Tables(4).Cell(3, 1).Range.InsertAfter "Opbouwdatum:"
w1.ActiveDocument.Tables(4).Cell(3, 2).Range.InsertAfter InhoudTextbox2.Text
w1.ActiveDocument.Tables(4).Cell(4, 1).Range.InsertAfter "Frequentie bijhouding:"
w1.ActiveDocument.Tables(4).Cell(4, 2).Range.InsertAfter InhoudTextbox3.Text
w1.ActiveDocument.Tables(4).Cell(5, 1).Range.InsertAfter "Contactpersoon inhoud:"
w1.ActiveDocument.Tables(4).Cell(5, 2).Range.InsertAfter InhoudTextbox4.Text
w1.ActiveDocument.Tables(4).Cell(6, 1).Range.InsertAfter "Leverancier:"
w1.ActiveDocument.Tables(4).Cell(6, 2).Range.InsertAfter InhoudTextbox5.Text
w1.ActiveDocument.Tables(4).Cell(7, 1).Range.InsertAfter "Contactpersoon leverancier:"
w1.ActiveDocument.Tables(4).Cell(7, 2).Range.InsertAfter InhoudTextbox6.Text

w1.Selection.GoTo wdGoToLine, wdGoToRelative, teller
w1.Selection.TypeParagraph

teller = w1.Selection.StoryLength
teller = teller - 1

Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(teller, teller), 9, 2)
otable.Borders.Enable = 0
w1.ActiveDocument.Tables(5).Columns(1).Width = 150
w1.ActiveDocument.Tables(5).Columns(2).Width = 300

w1.ActiveDocument.Tables(5).Cell(1, 1).Merge w1.ActiveDocument.Tables(5).Cell(1, 2)
w1.ActiveDocument.Tables(5).Cell(1, 1).Range.Font.Bold = True
w1.ActiveDocument.Tables(5).Cell(1, 1).Range.Font.Size = 10
w1.ActiveDocument.Tables(5).Cell(1, 1).Range.Font.Italic = True
w1.ActiveDocument.Tables(5).Cell(1, 1).Range.InsertAfter "Specifiek:"

w1.ActiveDocument.Tables(5).Cell(2, 1).Range.InsertAfter "Opbouwmethode:"
w1.ActiveDocument.Tables(5).Cell(2, 2).Range.InsertAfter SpecifiekTextbox1.Text
w1.ActiveDocument.Tables(5).Cell(3, 1).Range.InsertAfter "Gebied:"
w1.ActiveDocument.Tables(5).Cell(3, 2).Range.InsertAfter SpecifiekTextbox2.Text
w1.ActiveDocument.Tables(5).Cell(4, 1).Range.InsertAfter "Geometrie:"
w1.ActiveDocument.Tables(5).Cell(4, 2).Range.InsertAfter SpecifiekTextbox3.Text
w1.ActiveDocument.Tables(5).Cell(5, 1).Range.InsertAfter "Legenda-file:"
w1.ActiveDocument.Tables(5).Cell(5, 2).Range.InsertAfter SpecifiekTextbox4.Text
w1.ActiveDocument.Tables(5).Cell(6, 1).Range.InsertAfter "Positionele nauwkeurigheid:"
w1.ActiveDocument.Tables(5).Cell(6, 2).Range.InsertAfter SpecifiekTextbox5.Text
w1.ActiveDocument.Tables(5).Cell(7, 1).Range.InsertAfter "Schaal:"
w1.ActiveDocument.Tables(5).Cell(7, 2).Range.InsertAfter SpecifiekTextbox6.Text
w1.ActiveDocument.Tables(5).Cell(8, 1).Range.InsertAfter "Tekenschaal:"
w1.ActiveDocument.Tables(5).Cell(8, 2).Range.InsertAfter SpecifiekTextbox7.Text
w1.ActiveDocument.Tables(5).Cell(9, 1).Range.InsertAfter "Standaarditem:"
w1.ActiveDocument.Tables(5).Cell(9, 2).Range.InsertAfter SpecifiekTextbox8.Text

w1.Selection.GoTo wdGoToLine, wdGoToRelative, teller
w1.Selection.TypeParagraph

teller = w1.Selection.StoryLength
teller = teller - 1

aantalrij = aantalrij - 1

Set otable = w1.ActiveDocument.Tables.Add(w1.ActiveDocument.Range(teller, teller), (2 + aantalrij), 4)
otable.Borders.Enable = 0
w1.ActiveDocument.Tables(6).Columns(1).Width = 70
w1.ActiveDocument.Tables(6).Columns(2).Width = 125
w1.ActiveDocument.Tables(6).Columns(3).Width = 50
w1.ActiveDocument.Tables(6).Columns(4).Width = 210

w1.ActiveDocument.Tables(6).Cell(1, 1).Merge w1.ActiveDocument.Tables(6).Cell(1, 4)
w1.ActiveDocument.Tables(6).Cell(1, 1).Range.Font.Bold = True
w1.ActiveDocument.Tables(6).Cell(1, 1).Range.Font.Size = 10
w1.ActiveDocument.Tables(6).Cell(1, 1).Range.Font.Italic = True
w1.ActiveDocument.Tables(6).Cell(1, 1).Range.InsertAfter "Items:"

w1.ActiveDocument.Tables(6).Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
w1.ActiveDocument.Tables(6).Cell(2, 1).Range.InsertAfter "Itemnaam"
w1.ActiveDocument.Tables(6).Cell(2, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
w1.ActiveDocument.Tables(6).Cell(2, 2).Range.InsertAfter "Itemdefinitie"
w1.ActiveDocument.Tables(6).Cell(2, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
w1.ActiveDocument.Tables(6).Cell(2, 3).Range.InsertAfter "Eenheid"
w1.ActiveDocument.Tables(6).Cell(2, 4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
w1.ActiveDocument.Tables(6).Cell(2, 4).Range.InsertAfter "Mogelijke waarden"



w1.Selection.GoTo wdGoToLine, wdGoToRelative, 0

Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim SQL As String
Dim lijst As String
Dim totaal As Integer
Dim totaalset As Integer

On Error GoTo ErrorHandler

totaalset = 0

If KeyAscii = 13 Then
    If Text1.Text <> "" Then
    
    SQL = "SELECT DISTINCT DATASET.DATASET_TITEL " + _
    "FROM ((DATASET INNER JOIN MEMOTABEL ON DATASET.OMSCHRIJVING_CODE = MEMOTABEL.CODE) INNER JOIN TREFCODE ON DATASET.DATACODE = " + _
    "TREFCODE.DATACODE) INNER JOIN TREFTEXT ON TREFCODE.TREFCODE = TREFTEXT.TREFCODE " + _
    "WHERE instr(DATASET.DATASET_TITEL, '" + Text1.Text + "') OR instr(MEMOTABEL.TEKST, '" + Text1.Text + "') OR " + _
    "instr(TREFTEXT.TREFWOORD, '" + Text1.Text + "') AND DATASET.TYPE = 1 AND DATASET.VEILIGHEID = 'vrij toegankelijk' ORDER BY DATASET.DATASET_TITEL;"
        
    'SQL = "SELECT DISTINCT DATASET.DATASET_TITEL FROM " + _
    '"((TREFCODE INNER JOIN DATASET ON TREFCODE.DATACODE = DATASET.DATACODE) INNER JOIN TREFTEXT ON " + _
    '"TREFCODE.TREFCODE = TREFTEXT.TREFCODE) INNER JOIN MEMOTABEL ON DATASET.DATACODE = MEMOTABEL.CODE " + _
    '"WHERE (((DATASET.DATASET_TITEL) Like '%" + Text1.Text + "%')) " + _
    '"OR (((MEMOTABEL.TEKST) Like '%" + Text1.Text + "%')) OR (((TREFTEXT.TREFWOORD) Like '%" + Text1.Text + "%')) AND DATASET.TYPE=1 ORDER BY DATASET.DATASET_TITEL;"
                    
    Form1.List1.Clear
    
    Set rsselect = New ADODB.Recordset
    rsselect.Open SQL, db, adOpenStatic, adLockOptimistic
      
    With rsselect
      If Not .EOF Then
            .MoveFirst
            Do
            If Not IsNull(rsselect.Fields(0)) Then
              lijst = rsselect.Fields(0)
              Form1.List1.AddItem (lijst)
            End If
            rsselect.MoveNext
            totaalset = totaalset + 1
            Loop Until .EOF
        End If
        .Close
    End With
    End If
End If

LabelNamen.Caption = "Datasets (" + Str(totaalset) + " van " + Str(totaalmeta) + ")"
    
Exit Sub
ErrorHandler:
  MsgBox "Er is een fout opgetreden in de ArcGIS - GBI tool." & vbCr & vbCr & _
         "Fout Detail : " & Err.Description, vbExclamation + vbOKOnly, "Foutmelding"
End Sub
