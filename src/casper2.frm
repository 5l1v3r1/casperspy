VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.2#0"; "CODEJO~4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.2#0"; "CODEJO~3.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.2#0"; "CO85CC~1.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   Caption         =   " CasperLauncher | Full Version [Shadow Striker]"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "casper2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton9 
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
      _Version        =   983042
      _ExtentX        =   661
      _ExtentY        =   556
      _StockProps     =   79
      BackColor       =   -2147483633
      FlatStyle       =   -1  'True
      Appearance      =   2
      IconWidth       =   16
      Icon            =   "casper2.frx":1CCA
   End
   Begin VB.TextBox m 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Text            =   "Find attacker here.."
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   1095
      Left            =   7440
      TabIndex        =   0
      ToolTipText     =   "Launch Attacker"
      Top             =   2400
      Width           =   1575
      _Version        =   983042
      _ExtentX        =   2778
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Launch"
      ForeColor       =   4210752
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   2
      DrawFocusRect   =   0   'False
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "casper2.frx":2134
   End
   Begin XtremeSuiteControls.PushButton PushButton4 
      Height          =   1095
      Left            =   9120
      TabIndex        =   3
      Top             =   6645
      Width           =   1575
      _Version        =   983042
      _ExtentX        =   2778
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "About"
      ForeColor       =   4210752
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   2
      DrawFocusRect   =   0   'False
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "casper2.frx":459E
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   1095
      Left            =   9120
      TabIndex        =   2
      ToolTipText     =   "Find tutorial how to use the Attacker which you select"
      Top             =   2400
      Width           =   1575
      _Version        =   983042
      _ExtentX        =   2778
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Guide"
      ForeColor       =   4210752
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   2
      DrawFocusRect   =   0   'False
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "casper2.frx":6A08
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   1095
      Left            =   7440
      TabIndex        =   1
      ToolTipText     =   "Find specific attacker"
      Top             =   6645
      Width           =   1575
      _Version        =   983042
      _ExtentX        =   2778
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Find"
      ForeColor       =   4210752
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   2
      DrawFocusRect   =   0   'False
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "casper2.frx":8E72
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   8
      TabsPerRow      =   9
      TabHeight       =   520
      TabMaxWidth     =   882
      BackColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Attacker"
      TabPicture(0)   =   "casper2.frx":B2DC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SkinFramework1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CommonDialog1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "List1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Botnet"
      TabPicture(1)   =   "casper2.frx":B2F8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List2"
      Tab(1).Control(1)=   "Text2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Cracking"
      TabPicture(2)   =   "casper2.frx":B314
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "List3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Keylogger"
      TabPicture(3)   =   "casper2.frx":B330
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "List4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Text4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Network"
      TabPicture(4)   =   "casper2.frx":B34C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "List5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Trojan"
      TabPicture(5)   =   "casper2.frx":B368
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "List6"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Exploit"
      TabPicture(6)   =   "casper2.frx":B384
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Text7"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "List7"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Social Engginering"
      TabPicture(7)   =   "casper2.frx":B3A0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Text8"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "List8"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Web Hacking"
      TabPicture(8)   =   "casper2.frx":B3BC
      Tab(8).ControlEnabled=   -1  'True
      Tab(8).Control(0)=   "Text9"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "List9"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).ControlCount=   2
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":B3D8
         Left            =   120
         List            =   "casper2.frx":B48A
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Text            =   "casper2.frx":B6CA
         Top             =   960
         Width           =   4095
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":B74C
         Left            =   -74880
         List            =   "casper2.frx":B7FE
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -71640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "casper2.frx":BA3E
         Top             =   960
         Width           =   4095
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":BAC0
         Left            =   -74880
         List            =   "casper2.frx":BB72
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -71640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Text            =   "casper2.frx":BDB2
         Top             =   960
         Width           =   4095
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":BE34
         Left            =   -74880
         List            =   "casper2.frx":BEE6
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -71640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "casper2.frx":C126
         Top             =   960
         Width           =   4095
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":C1A8
         Left            =   -74880
         List            =   "casper2.frx":C25A
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -71640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "casper2.frx":C49A
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -71640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "casper2.frx":C51C
         Top             =   960
         Width           =   4095
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":C59E
         Left            =   -74880
         List            =   "casper2.frx":C650
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -71640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "casper2.frx":C890
         Top             =   960
         Width           =   4095
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":C912
         Left            =   -74880
         List            =   "casper2.frx":C9C4
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -71640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "casper2.frx":CC04
         Top             =   960
         Width           =   4095
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":CC86
         Left            =   -74880
         List            =   "casper2.frx":CD38
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   3135
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "casper2.frx":CF78
         Left            =   -74880
         List            =   "casper2.frx":D02A
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -71640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "casper2.frx":D26A
         Top             =   960
         Width           =   4095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -65760
         Top             =   4830
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin XtremeSkinFramework.SkinFramework SkinFramework1 
         Left            =   -66240
         Top             =   4830
         _Version        =   983042
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H80000007&
      Caption         =   "CREATE THE REPORT FIRST!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaskColor       =   &H000000FF&
      TabIndex        =   4
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "CasperLauncher | Most Complete Attacker"
      BeginProperty Font 
         Name            =   "HelveticaNeueLT Std Blk Ext"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   6015
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   10215
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   1320
      Top             =   120
      _Version        =   983042
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "casper2.frx":D2EC
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   240
      Top             =   120
      _Version        =   983042
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   7
      DesignerControls=   -1  'True
      DesignerControlsData=   "casper2.frx":36FBA
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CasperLauncher is a Hacking Framework Interface built for Windows.
' It is written in Visual Basic 6
' Copyright (C) www.casperspy.com by sillhouette
' CasperLauncher is distributed in the hope that it will be useful,

Option Explicit

Dim intResponse As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Public UseDisabledIcons As Boolean

' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Dim file1
Dim about
Public execution1 As String
Public execution2 As String
Public execution3 As String
Public execution4 As String

Dim reportlocation As String

Dim WshShell, oExec
Dim counter As Integer
Dim search, se As String

Public Function MapPath( _
    path As String _
) As String
End Function

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Try To Get System Info Program Path\Name From Registry...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validate Existance Of Known 32 Bit File Version
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error - File Can Not Be Found...
                Else
                        GoTo SysInfoErr
                End If
        ' Error - Registry Entry Can Not Be Found...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, Id As Long, Optional BeginGroup As Boolean = False) As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, Id, "")
     
    Control.BeginGroup = BeginGroup
    Set AddButton = Control
End Function

Private Sub DockRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    CommandBars.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub
Public Function GetDll() As String
    If CommandBarsGlobalSettings.ResourceImages.DllFileName = "" Then
        'if nothing has been set, then use Office 2007
        GetDll = App.path & "\..\..\..\Styles\OfficeXp.Dll"
    Else
        GetDll = CommandBarsGlobalSettings.ResourceImages.DllFileName
    End If
End Function

Public Function GetINI() As String
    If CommandBarsGlobalSettings.ResourceImages.IniFileName = "" Then
        'if nothing has been set, then use Office 2007
        GetINI = App.path & ""
    Else
        GetINI = CommandBarsGlobalSettings.ResourceImages.IniFileName
    End If
End Function




Private Sub Command10_Click()

End Sub

Private Sub Command11_Click()

If execution4 = "autorunsc" Or execution4 = "handle" Or execution4 = "Listdlls" Or execution4 = "logonsessions" Or execution4 = "psgetsid" Or execution4 = "Psinfo" Or execution4 = "pslist" Or execution4 = "psloggedon" Or execution4 = "psloglist" Or execution4 = "psservice" Then
    Set WshShell = CreateObject("WScript.Shell")
    Set oExec = WshShell.Exec(App.path & "\Programs\wft\tools\sysinternals\" & execution4 & ".exe")
    Text7.Text = oExec.StdOut.ReadAll
    Set oExec = Nothing
    Set WshShell = Nothing

    If Not Command14.Visible And Option1 Then
    Open reportlocation For Append As #1
    Print #1, Text7.Text
    Print #1, ""
    Close #1
    End If
    
ElseIf execution4 = "efsdump" Or execution4 = "ntfsinfo" Or execution4 = "psexec" Or execution4 = "psfile" Or execution4 = "pskill" Or execution4 = "pspasswd" Or execution4 = "psshutdown" Or execution4 = "pssuspend" Or execution4 = "streams" Or execution4 = "strings" Then
    Shell ("cmd.exe /k cd " & App.path & "\Programs\wft\tools\sysinternals && " & App.path & "\Programs\wft\tools\sysinternals\" & execution4 & ".exe"), vbNormalFocus
End If
End Sub

Private Sub Command12_Click()
Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)
Clipboard.Clear
Clipboard.SetText Text7.SelText
End Sub

Private Sub Command13_Click()
about = MsgBox("Copyright (C) 2009 - Giancarlo Giustini (gianchi83@gmail.com) - Nanni Bassetti (digitfor@gmail.com)", vbInformation + vbOKOnly, "WinTaylor for CAINE")
End Sub

Private Sub Command14_Click()
    
    CommonDialog1.Filter = "Text Files|*.txt"
    CommonDialog1.ShowSave
    reportlocation = CommonDialog1.FileName
   
    If Not reportlocation = "" Then
    Open reportlocation For Append As #1
    Print #1, "##################################################################################"
    Print #1, "####                                                                          ####"
    Print #1, "####                            CASPERLAUNCHER 1.0                            ####"
    Print #1, "####                              ATTACK REPORT                               ####"
    Print #1, "####                                                                          ####"
    Print #1, "####                                                                          ####"
    Print #1, "####                                                                          ####"
    Print #1, "##################################################################################"
    Print #1, ""
    Print #1, "Report creation time: "; Now
    Print #1, ""
    Print #1, ""
    Close #1
    Command14.Visible = False
    Else: Command14.Visible = True
    End If
        
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Shell (App.path & "\Programs\HashMyFiles.exe"), vbNormalFocus
    
    If Not Command14.Visible Then
    Open reportlocation For Append As #1
    Print #1, "__________________________________________________________________________________"
    Print #1, "-> "; Time; "opened program: HashMyFiles"
    Print #1, ""
    Close #1
    End If
    
End Sub

Private Sub Command5_Click()
    Shell (App.path & "\Programs\MWSnap.exe"), vbNormalFocus
    
    If Not Command14.Visible Then
    Open reportlocation For Append As #1
    Print #1, "__________________________________________________________________________________"
    Print #1, "-> "; Time; "opened program: MWSnap"
    Print #1, ""
    Close #1
    End If
    
End Sub

Private Sub Command6_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Clipboard.Clear
Clipboard.SetText Text1.SelText
End Sub

Private Sub Command7_Click()
    Shell (App.path & "\Programs\Imager\FTKImager.exe"), vbNormalFocus
    
    If Not Command14.Visible Then
    Open reportlocation For Append As #1
    Print #1, "__________________________________________________________________________________"
    Print #1, "-> "; Time; "opened program: FTKImager"
    Print #1, ""
    Close #1
    End If

End Sub

Private Sub Command8_Click()
    Shell ("cmd.exe /k cd " & App.path & "\Programs\wft && " & App.path & "\Programs\wft\wft.exe"), vbNormalFocus
        
    If Not Command14.Visible Then
    Open reportlocation For Append As #1
    Print #1, "__________________________________________________________________________________"
    Print #1, "-> "; Time; "opened program: Windows Forensic Toolchest"
    Print #1, ""
    Close #1
    End If
    
End Sub

Private Sub Command9_Click()
    Shell (App.path & "\Programs\AgileRM\Nigilant32.exe"), vbNormalFocus
    
    If Not Command14.Visible Then
    Open reportlocation For Append As #1
    Print #1, "__________________________________________________________________________________"
    Print #1, "-> "; Time; "opened program: Nigilant32"
    Print #1, ""
    Close #1
    End If
    
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   On Error Resume Next
      
    Select Case Control.Id
        Case ID_HELP_ABOUT: Intro.Show
        Case ID_ATTACKER_MORE: Dim intResponse As Integer
                               intResponse = MsgBox("Only Available on Full Version!" _
                               & "Do you want to download Full Version?", _
                               vbYesNo + vbQuestion + vbDefaultButton2, _
                               "Get Full Version")
                               If intResponse = vbYes Then
                               ShellExecute 0, vbNullString, "http://form.jotform.me/form/31453837808461", vbNullString, vbNullString, vbNormalFocus
                               End If
        Case ID_FILE_UPDATE:   intResponse = MsgBox("Only Available on Full Version!" _
                               & "Do you want to download Full Version?", _
                               vbYesNo + vbQuestion + vbDefaultButton2, _
                               "Get Full Version")
                               If intResponse = vbYes Then
                               ShellExecute 0, vbNullString, "http://form.jotform.me/form/31453837808461", vbNullString, vbNullString, vbNormalFocus
                               End If
        Case ID_ATTACKER_AGOBOT: List1.ListIndex = 0
        SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_CASPER: List1.ListIndex = 7
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_CITADEL: List1.ListIndex = 8
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_CYTHOSIA: List1.ListIndex = 10
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_SPITMO: List1.ListIndex = 40
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_SPYEYE: List1.ListIndex = 41
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_VERTEXNET: List1.ListIndex = 48
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_ZEUS: List1.ListIndex = 54
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        
        Case ID_ATTACKER_EXPLOIT1: List1.ListIndex = 9
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_EXPLOIT2: List1.ListIndex = 21
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_EXPLOIT3: List1.ListIndex = 30
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_EXPLOIT4: List1.ListIndex = 32
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        
        Case ID_ATTACKER_FORMGRABBER1: List1.ListIndex = 31
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_FORMGRABBER2: List1.ListIndex = 47
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_FORMGRABBER3: List1.ListIndex = 51
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        
        Case ID_ATTACKER_KEYLOGGER1: List1.ListIndex = 2
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_KEYLOGGER2: List1.ListIndex = 28
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_KEYLOGGER3: List1.ListIndex = 36
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_KEYLOGGER4: List1.ListIndex = 49
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        
        Case ID_ATTACKER_NETWORK1: List1.ListIndex = 1
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK2: List1.ListIndex = 13
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK3: List1.ListIndex = 14
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK4: List1.ListIndex = 17
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK5: List1.ListIndex = 20
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK6: List1.ListIndex = 23
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK7: List1.ListIndex = 25
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK8: List1.ListIndex = 29
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK9: List1.ListIndex = 33
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK10: List1.ListIndex = 35
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK11: List1.ListIndex = 37
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK12: List1.ListIndex = 39
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK13: List1.ListIndex = 50
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_NETWORK14: List1.ListIndex = 52
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        
        Case ID_ATTACKER_TROJAN1: List1.ListIndex = 4
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_TROJAN2: List1.ListIndex = 5
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_TROJAN3: List1.ListIndex = 11
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_TROJAN4: List1.ListIndex = 34
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_TROJAN5: List1.ListIndex = 42
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_TROJAN6: List1.ListIndex = 46
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_TROJAN7: List1.ListIndex = 45
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        
        Case ID_ATTACKER_WEBHACKING1: List1.ListIndex = 6
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_WEBHACKING2: List1.ListIndex = 12
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_WEBHACKING3: List1.ListIndex = 18
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_WEBHACKING4: List1.ListIndex = 24
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_WEBHACKING5: List1.ListIndex = 27
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_WEBHACKING6: List1.ListIndex = 43
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_WEBHACKING7: List1.ListIndex = 53
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_SOCIAL1: List1.ListIndex = 3
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_SOCIAL2: List1.ListIndex = 19
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_SOCIAL3: List1.ListIndex = 26
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_SOCIAL4: List1.ListIndex = 44
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_SOCIAL5: List1.ListIndex = 38
         SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case ID_ATTACKER_EXIT: Unload Me
    
        Case ID_FILE_UPDATE: MsgBox Control.Caption & " clicked", vbOKOnly, "Button Clicked"
        
        Case ID_FILE_NEWS: ShellExecute 0, vbNullString, "http://feeds.feedburner.com/casperspy/yttQ/", vbNullString, vbNullString, vbNormalFocus
        
        Case ID_ACTIONS_FULL: ShellExecute 0, vbNullString, "http://form.jotform.me/form/31453837808461", vbNullString, vbNullString, vbNormalFocus
        
        Case ID_ACTIONS_CHAT: ShellExecute 0, vbNullString, "http://casperspy.com/lounge/", vbNullString, vbNullString, vbNormalFocus

        Case ID_ACTIONS_CONTACT: ShellExecute 0, vbNullString, "http://form.jotform.me/form/31453837808461", vbNullString, vbNullString, vbNormalFocus
        
        Case ID_FILE_FORUM: ShellExecute 0, vbNullString, "http://casperspy.com/forum/", vbNullString, vbNullString, vbNormalFocus
        
        Case ID_FILE_VIDEO: List4.ListIndex = 58
        frmMain.SSTab2.Tab = 0
        SSTab2.Visible = True
        List4.Visible = True
        Text8.Visible = True
        PushButton5.Visible = True
        PushButton7.Visible = True
        PushButton6.Visible = True
        Frame1.Caption = "CasperLauncher | Guide Menu"
        If Frame1.Caption = "CasperLauncher | Guide Menu" Then
        Text10.Visible = False
        m.Visible = False
        PushButton9.Visible = False
        End If
        If SSTab2.Caption = "Ebook" Then
        List5.Visible = True
        Text11.Visible = True
        Else
        List5.Visible = False
        Text11.Visible = False
        End If
        If SSTab2.Caption = "Tutorial" Then
        List4.Visible = True
        Text8.Visible = True
        Else
        List4.Visible = False
        Text8.Visible = False
        End If
        
        Case ID_FILE_TUTORIAL: SSTab2.Visible = True
        frmMain.SSTab2.Tab = 0
        List4.Visible = True
        Text8.Visible = True
        PushButton5.Visible = True
        PushButton7.Visible = True
        PushButton6.Visible = True
        Frame1.Caption = "CasperLauncher | Guide Menu"
        If Frame1.Caption = "CasperLauncher | Guide Menu" Then
        Text10.Visible = False
        m.Visible = False
        PushButton9.Visible = False
        End If
        If SSTab2.Caption = "Ebook" Then
        List5.Visible = True
        Text11.Visible = True
        Else
        List5.Visible = False
        Text11.Visible = False
        End If
        If SSTab2.Caption = "Tutorial" Then
        List4.Visible = True
        Text8.Visible = True
        Else
        List4.Visible = False
        Text8.Visible = False
        End If
        
        Case ID_FILE_GUIDE: SSTab2.Visible = True
        List4.Visible = True
        Text8.Visible = True
        PushButton5.Visible = True
        PushButton7.Visible = True
        PushButton6.Visible = True
        Frame1.Caption = "CasperLauncher | Guide Menu"
        If Frame1.Caption = "CasperLauncher | Guide Menu" Then
        Text10.Visible = False
        m.Visible = False
        PushButton9.Visible = False
        End If
        If SSTab2.Caption = "Ebook" Then
        List5.Visible = True
        Text11.Visible = True
        Else
        List5.Visible = False
        Text11.Visible = False
        End If
        If SSTab2.Caption = "Tutorial" Then
        List4.Visible = True
        Text8.Visible = True
        Else
        List4.Visible = False
        Text8.Visible = False
        End If
      
        
        Case ID_FILE_EBOOK: SSTab2.Visible = True
        frmMain.SSTab2.Tab = 1
        List4.Visible = True
        Text8.Visible = True
        PushButton5.Visible = True
        PushButton7.Visible = True
        PushButton6.Visible = True
        Frame1.Caption = "CasperLauncher | Guide Menu"
        If Frame1.Caption = "CasperLauncher | Guide Menu" Then
        Text10.Visible = False
        m.Visible = False
        PushButton9.Visible = False
        End If
        If SSTab2.Caption = "Ebook" Then
        List5.Visible = True
        Text11.Visible = True
        Else
        List5.Visible = False
        Text11.Visible = False
        End If
        If SSTab2.Caption = "Tutorial" Then
        List4.Visible = True
        Text8.Visible = True
        Else
        List4.Visible = False
        Text8.Visible = False
        End If
        
        Case ID_FILE_GUIDE: SSTab2.Visible = True
        List4.Visible = True
        Text8.Visible = True
        PushButton5.Visible = True
        PushButton7.Visible = True
        PushButton6.Visible = True
        Frame1.Caption = "CasperLauncher | Guide Menu"
        If Frame1.Caption = "CasperLauncher | Guide Menu" Then
        Text10.Visible = False
        m.Visible = False
        PushButton9.Visible = False
        End If
        If SSTab2.Caption = "Ebook" Then
        List5.Visible = True
        Text11.Visible = True
        Else
        List5.Visible = False
        Text11.Visible = False
        End If
        If SSTab2.Caption = "Tutorial" Then
        List4.Visible = True
        Text8.Visible = True
        Else
        List4.Visible = False
        Text8.Visible = False
        End If
      
        Case ID_FILE_CASPERSPY: SSTab2.Visible = False
        List4.Visible = False
        Text8.Visible = False
        List5.Visible = False
        Text11.Visible = False
        PushButton5.Visible = False
        PushButton7.Visible = False
        PushButton6.Visible = False
        Frame1.Caption = "CasperLauncher | Most Complete Attacker"
        If Frame1.Caption = "CasperLauncher | Most Complete Attacker" Then
        Text9.Visible = False
        n.Visible = False
        PushButton8.Visible = False
        End If
        Case Else:   MsgBox Control.Caption & " clicked", vbOKOnly, "Button Clicked"
        End Select
      

End Sub


Private Sub CommandBars_Resize()
    
    On Error Resume Next
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.GetClientRect Left, Top, Right, Bottom
    Frame1.Move Left, Top, Right - Left, Bottom - Top
    End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case ID_EDIT_CUT: Control.Enabled = True
        Case ID_EDIT_PASTE: Control.Enabled = True
        Case ID_EDIT_COPY: Control.Enabled = True
        Case ID_EDIT_UNDO: Control.Enabled = True
        Case ID_PROJECT_STOP: Control.Enabled = True
        Case ID_FILE_SAVE_ALL: Control.Enabled = True
        Case ID_ACTIONS_CHAT: Control.Enabled = True
        Case ID_ACTIONS_CONTACT: Control.Enabled = True
        Case ID_ACTIONS_VIEW: Control.Enabled = True
    End Select
End Sub


Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
SkinFramework1.LoadSkin App.path & "\casper\EX2008.cjstyles", ""
SkinFramework1.ApplyWindow Me.hwnd
SSTab1.Tab = 0
If SSTab1.Caption = "Cracking" Then
PushButton1.Visible = False
PushButton2.Visible = False
PushButton3.Visible = False
PushButton4.Visible = False
End If

    CommandBarsGlobalSettings.App = App
    
        
    Dim Control As CommandBarControl
    Dim ControlFile As CommandBarPopup, ControlEdit As CommandBarPopup, ControlView As CommandBarPopup
    Dim ControlProject As CommandBarPopup, ControlHelp As CommandBarPopup
   
    Set ControlFile = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Attacker")
    With ControlFile.CommandBar
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_BOTNET, "&Botnet")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_AGOBOT, "&Agobot IRC"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_CASPER, "&CasperSpy 2.0"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_CITADEL, "&Citadel"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_CYTHOSIA, "&Cythosia"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SPYEYE, "&Spy Eye"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SPITMO, "&Spitmo"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_VERTEXNET, "&VertexNet DDOS"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_ZEUS, "&Zeus"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_EXPLOIT, "&Exploit")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_EXPLOIT1, "&Crypter"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_EXPLOIT2, "&Evil PDF"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_EXPLOIT3, "&Metasploit"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_EXPLOIT4, "&Msfpayload exec"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_FORMGRABBER, "&Form Grabber")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_FORMGRABBER1, "&MP-formgrabber"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_FORMGRABBER2, "&TW-Grabber"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_FORMGRABBER3, "&WebCrab"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_KEYLOGGER, "&Keylogger")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_KEYLOGGER1, "&Albertino"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_KEYLOGGER2, "&Keymail"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_KEYLOGGER3, "&Rapzo"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_KEYLOGGER4, "&Vulcan"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_NETWORK, "&Network")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK1, "&AirCrack-NG"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK2, "&Dmitry"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK3, "&Dnsdict"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK4, "&Driftnet"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK5, "&Ettercap"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK6, "&Gerix"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK7, "&Janidos"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK8, "&LBD"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK9, "&Nmap"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK10, "&Pyrit"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK11, "&Reaver"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK12, "&Slowloris"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK13, "&Wash"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_NETWORK14, "&Wifite"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_TROJAN, "&Trojan")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN1, "&Bifrost"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN2, "&Bionet"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN3, "&Darkmoon"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN4, "&Prorat"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN5, "&Spynet"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN6, "&Turkojan"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN7, "&Trojan Creator"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_WEBHACKING, "&Web Hacking")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING1, "&Burpsuite"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING2, "&Dirbuster"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING3, "&DVWA"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING4, "&Havij pro"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING5, "&Joomscan"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING6, "&Sqlmap"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING7, "&Wp-scan"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_SOCIAL, "&Social Engginering")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL1, "&Autoinfector"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL2, "&Edjpgcom"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL3, "&Jigsaw"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL4, "&The Harvester"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL5, "&SET"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE, "&More..."
        
        AddControl .Controls, xtpControlButton, ID_ATTACKER_EXIT, "&Exit"
        Set frmMain.CommandBars.Icons = frmMain.ImageManager.Icons
        frmMain.CommandBars.Options.SetIconSize True, 42, 35
        frmMain.CommandBars.RecalcLayout
        End With
    
      
    Set ControlView = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Guide")
    With ControlView.CommandBar
        AddControl .Controls, xtpControlButton, ID_GUIDE_EBOOK, "&Ebook"
        AddControl .Controls, xtpControlButton, ID_GUIDE_TUTORIAL, "&Tutorial"
        AddControl .Controls, xtpControlButton, ID_GUIDE_VIDEO, "&Video Tutorial"
        Set frmMain.CommandBars.Icons = frmMain.ImageManager.Icons
        frmMain.CommandBars.Options.SetIconSize True, 32, 32
        frmMain.CommandBars.RecalcLayout
    End With
        
    Set ControlHelp = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Help")
    With ControlHelp.CommandBar.Controls
        .Add xtpControlButton, ID_HELP_ABOUT, "&About"
        .Add xtpControlButton, ID_FILE_UPDATE, "&Updates Contents"
        .Add xtpControlButton, ID_ACTIONS_FULL, "&Get Full Version"
        Set frmMain.CommandBars.Icons = frmMain.ImageManager.Icons
        frmMain.CommandBars.Options.SetIconSize True, 32, 32
        frmMain.CommandBars.RecalcLayout
    End With
    
    Dim ToolBar As CommandBar
    
    Set ToolBar = CommandBars.Add("Alpha", xtpBarTop)
    
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_CASPERSPY, "&Attacker", False, "Show Attacker Menu"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_GUIDE, "&Guide", False, "Show Guide Menu"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_EBOOK, "&Ebook", False, "Download Exclusive Ebook"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_TUTORIAL, "&Tutorial", False, "See manual instrusction how to use all Attacker"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_VIDEO, "Video Tutorial", False, "Watch this and you'll know"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_NEWS, "&News", False, "News Update about CyberSecurity"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_FORUM, "Forum", False, "Visite CasperSpy Forum"
    AddControl ToolBar.Controls, xtpControlButton, ID_ACTIONS_CHAT, "Chat", False, "CasperSpy Lounge"
    AddControl ToolBar.Controls, xtpControlButton, ID_ACTIONS_CONTACT, "&Contact Botmaster", False, "Give me an Inspiration to make this app better ^_^"
    AddControl ToolBar.Controls, xtpControlButton, ID_ACTIONS_FULL, "&Get Full Version", False, "Get Full Version of CasperLauncher to activate all feature"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_UPDATE, "&Auto Update", False, "Auto Update all Contents & Feature of CasperLauncher"
    Set frmMain.CommandBars.Icons = frmMain.ImageManager.Icons
       frmMain.CommandBars.Options.LargeIcons = True
        frmMain.CommandBars.RecalcLayout
                
    Dim StatusBar As StatusBar
    Set StatusBar = CommandBars.StatusBar
    Set frmMain.CommandBars.Icons = frmMain.ImageManager.Icons
    StatusBar.Visible = True
    
    StatusBar.AddPane 0
    StatusBar.AddPane ID_INDICATOR_CAPS
    StatusBar.AddPane ID_INDICATOR_NUM
    StatusBar.AddPane ID_INDICATOR_SCRL

End Sub

Private Sub List1_Click()
If List1.Text = "Agobot" Or List1.Text = "agobot" Then
Text1.Text = "Agobot IRC Botnet" & vbCrLf & " " & vbCrLf & "Agobot (Gaobot) is a family of computer worms. Axel Gembe was responsible for writing the first version."
End If
If List1.Text = "CasperLogger" Or List1.Text = "casperlogger" Then
Text1.Text = "CasperLogger V1.0" & vbCrLf & " " & vbCrLf & "is stealth email keylogger which will send login credentials data to your email to use it you just set the email which was registered first on SMTP register."
End If
If List1.Text = "Stuxnet" Or List1.Text = "stuxnet" Then
Text1.Text = "Most insidious worm" & vbCrLf & " " & vbCrLf & "The most insidious digital weapons ever, capable of crippling water supplies, power plants, banks, and the very infrastructure that once seemed invulnerable to attack."
End If
If List1.Text = "Crypter Pack FUD" Or List1.Text = "crypter pack FUD" Then
Text1.Text = "Package of Crypter Fully Undetectable" & vbCrLf & " " & vbCrLf & "in these package are the best Cypter i ever use. Each of Crypter has its own advantages. Among the best I recommend you to use Fly Crypter"
End If
If List1.Text = "Trojan Creator" Or List1.Text = "trojan creator" Then
Text1.Text = "Advanced Steam Trojan Generator" & vbCrLf & " " & vbCrLf & "You can creating your own trojan with this malware construction kit"
file1 = "activex.exe"
End If
If List1.Text = "Turkojan" Or List1.Text = "turkojan" Then
Text1.Text = "Turkojan" & vbCrLf & " " & vbCrLf & "is a remote administration and spying tool for Microsoft Windows operating systems. You can use Turkojan for manage computers,employee monitoring and child monitoring..."
file1 = "activex.exe"
End If
If List1.Text = "Spynet" Or List1.Text = "spynet" Then
Text1.Text = "SpyNet - Remote Surveillance Trojan" & vbCrLf & " " & vbCrLf & "Spynet targets the Windows platform. This RAT identifies itself with a remote controlling server and listens for commands from an attacker. A considerable number of commands are available to an attacker, including keylog, download files, obtain system information, and list services, among many others."
file1 = "activex.exe"
End If
If List1.Text = "Prorat" Or List1.Text = "prorat" Then
Text1.Text = "ProRat is a Remote Administration Tool" & vbCrLf & " " & vbCrLf & "ProRat looks like a perfect Trojan but the real aim is to connect you to your own PC from a distance and its made only for education."
file1 = "activex.exe"
End If
If List1.Text = "Darkmoon" Or List1.Text = "darkmoon" Then
Text1.Text = "Agobot IRC Botnet" & vbCrLf & " " & vbCrLf & "Darkmoon is a Trojan horse that opens a back door on the compromised computer and has keylogging capabilities."
file1 = "activex.exe"
End If
If List1.Text = "Bionet" Or List1.Text = "bionet" Then
Text1.Text = "Bionet - Trojan RAT" & vbCrLf & " " & vbCrLf & "BioNet comes with a server editor that allows hackers and script kiddies to customize the trojan server. Notable features of the server editor is the ability to customize the trojan with scripts that are executed on a specific date and time, or even at regular intervals, once the trojan server compromises a machine. The trojan also has the ability to terminate security software, but this is nothing new as the feature was also available in the previous version of BioNet."
file1 = "activex.exe"
End If
If List1.Text = "Bifrost" Or List1.Text = "bifrost" Then
Text1.Text = "Bifrost - Trojan RAT" & vbCrLf & " " & vbCrLf & "Bifrost is a backdoor trojan horse family of more than 10 variants which can infect Windows 95 through Windows 7. Bifrost uses the typical server, server builder, and client backdoor program configuration to allow a remote attacker, who uses the client, to execute arbitrary code on the compromised machine (which runs the server whose behavior can be controlled by the server editor)."
file1 = "activex.exe"
End If
If List1.Text = "Wp-Scan" Or List1.Text = "wp-scan" Then
Text1.Text = "Wordpress Vulnerability Scanner" & vbCrLf & " " & vbCrLf & "WPScan is a black box WordPress Security Scanner written in Ruby which attempts to find known security weaknesses within WordPress installations. Its intended use it to be for security professionals or WordPress administrators to asses the security posture of their WordPress installations."
file1 = "activex.exe"
End If
If List1.Text = "Keymail" Or List1.Text = "keymail" Then
Text1.Text = "Keymail Keylogger - An E-mailing Key Logger for Windows with C Source" & vbCrLf & " " & vbCrLf & " Keymail is a stealth (somewhat) key logger that e-mails key strokes to whoever is set in the #define options at compile time.  This code is for educational uses, it should be useful for those that want to learn more about using sockets in C and Windows key loggers. Don't be an ass hat with it, I'll only answer intelligent questions about it so don't e-mail me about the code unless it's to contribute or to ask an intelligent question. Cool thing about it, if Anti-virus apps start to detect it you should be able to just change it a little and recompile it"
file1 = "activex.exe"
End If
If List1.Text = "Albertino" Or List1.Text = "albertino" Then
Text1.Text = "Albertino Keylogger Creator" & vbCrLf & " " & vbCrLf & "Albertino is keylogger generator which can make credentials\password stealer easily with user friendly interface"
file1 = "activex.exe"
End If
If List1.Text = "Rapzo" Or List1.Text = "rapzo" Then
Text1.Text = "Rapzo Keylogger Private Edition" & vbCrLf & " " & vbCrLf & "Rapzo Keylogger, Hack any account with this keylogger Rapzo keylogger Private Edition it's amazing stealer with full feature such as: Stealers [6] All Stealers Pure Code - No Drops + Runtime FUD [#] Firefox 3.5.0-3.6.X [#] DynDns [#] FileZilla [#] Pidgin [#] Imvu [#] No - Ip * Full UAC Bypass & Faster Execution * Coded in Vb.NET * Min Req Is .net 2.0 Now A days every pc Have it * Cool & user friendly GUI * Easily Undetectable * Encrypt Information * Encrypt E-mail information * 100% FUD from all AV's * 4 Extentions [ . exe | .scr | .pif | .com ] * Keylogger support - Smtp[Gmail,Hotmail,live,aol,] * Test E-mail - is it vaild or not. * Customize the to e-mail address. * Screen Logger * Cure.exe to remove server from your Computer * Usb Spread * File pumper - Built-in * Icon Changer - Preview * Logs are nice and clear * Log Letters - ABCD etc. * Log Symbols - !@#$% etc. * Log Numbers - 12345 etc."
file1 = "activex.exe"
End If
If List1.Text = "TW-Grabber" Or List1.Text = "tw-grabber" Then
Text1.Text = "TW-FormGrabber" & vbCrLf & " " & vbCrLf & "TW-Grabber Simple and Easy to use FormGrabber work perfectly on Firefox and chrome outdated"
file1 = "activex.exe"
End If
If List1.Text = "Msfpayload" Or List1.Text = "msfpayload" Then
Text1.Text = "Metasploit payload executor" & vbCrLf & " " & vbCrLf & "Exploitation your target by known vulnerability based on Metasploit you can combine this tool by Social Engginering trick"
file1 = "activex.exe"
End If
If List1.Text = "Crypter.py" Or List1.Text = "crypter.py" Then
Text1.Text = "Crypter.py is Undetectable payload for Metasploit" & vbCrLf & " " & vbCrLf & "Currently it is not fully undetectable but it is still undetectable against major antiviruses. So if i had knowledge on my victims antivirus, i can then use the respective crypter to get the job done."
file1 = "activex.exe"
End If
If List1.Text = "Gerix" Or List1.Text = "gerix" Then
Text1.Text = "Gerix Wifi Cracker NG (New Generation)" & vbCrLf & " " & vbCrLf & " a really complete GUI for Aircrack-NG which includes useful extras. Completely re-written in Python + QT, automates all the different techniques to attack Access Points and Wireless Routers. Currently it is available and supported natively by BackTrack (pre-installed on the BT4 Final version) and available on all the different Debian Based distributions (Ubuntu, etc..)."
file1 = "activex.exe"
End If
If List1.Text = "Slowloris" Or List1.Text = "slowloris" Then
Text1.Text = "Slowloris HTTP DDOS" & vbCrLf & " " & vbCrLf & "Slowloris is a piece of software written by Robert “RSnake” Hansen which allows a single machine to take down another machine’s web server with minimal bandwidth and side effects on unrelated services and ports."
file1 = "activex.exe"
End If
If List1.Text = "Joomscan" Or List1.Text = "joomscan" Then
Text1.Text = "Joomla Vulnerability Scanner" & vbCrLf & " " & vbCrLf & "Contributed by Mr Aung Khant from Ethical Hacker Group, joomla! vulnerability scanner runs a target joomla website against known vulnerabilities. It will assist web developers, penetration experts and hackers to identify possible security weaknesses on their deployed / target Joomla! sites. It will then report back the positive and negative results."
file1 = "activex.exe"
End If
If List1.Text = "Autoinfector" Or List1.Text = "autoinfector" Then
Text1.Text = "Infectious Media Generator" & vbCrLf & " " & vbCrLf & "The Infectious USB/CD/DVD module will create an autorun.inf file and a Metasploit payload. When the DVD/USB/CD is inserted, it will automatically run if autorun is enabled. So yeah as i was saying, MI shit! lol. The Infectious Media Generator can be found in the Social Engineering Toolkit and this attack vector is relatively simple in nature and relies on deploying the devices to the physical system."
file1 = "activex.exe"
End If
If List1.Text = "Pyrit" Or List1.Text = "pyrit" Then
Text1.Text = "Precomputing Hash Crack" & vbCrLf & " " & vbCrLf & "Pyrit allows to create massive databases, pre-computing part of the IEEE 802.11 WPA/WPA2-PSK authentication phase in a space-time-tradeoff. Pyrit can use your Graphic card to increase your cracking speed. Exploiting the computational power of Many-Core- and other platforms through ATI-Stream, Nvidia CUDA and OpenCL, it is currently by far the most powerful attack against one of the world’s most used security-protocols."
file1 = "activex.exe"
End If
If List1.Text = "Reaver" Or List1.Text = "reaver" Then
Text1.Text = "Reaver - Wireless Exploitation" & vbCrLf & " " & vbCrLf & "Reaver exploits a protocol design flaw in WiFi Protected Setup (WPS). This vulnerability exposes a side-channel attack against Wi-Fi Protected Access (WPA) versions 1 and 2 allowing the extraction of the Pre-Shared Key (PSK) used to secure the network. WPS allows users to enter an 8 digit PIN to connect to a secured network without having to enter a passphrase. When a user supplies the correct PIN the access point essentially gives the user the WPA/WPA2 PSK that is needed to connect to the network."
file1 = "activex.exe"
End If
If List1.Text = "SET" Or List1.Text = "set" Then
Text1.Text = "Social Engginer Toolkit" & vbCrLf & " " & vbCrLf & "The Social-Engineer Toolkit (SET) is specifically designed to perform advanced attacks against the human element. SET was designed to be released with the http://www.social-engineer.org launch and has quickly became a standard tool in a penetration testers arsenal. SET was written by David Kennedy (ReL1K) and with a lot of help from the community it has incorporated attacks never before seen in an exploitation toolset. The attacks built into the toolkit are designed to be targeted and focused attacks against a person or organization used during a penetration test."
file1 = "activex.exe"
End If
If List1.Text = "Spitmo" Or List1.Text = "spitmo" Then
Text1.Text = "Android Spyeye Botnet" & vbCrLf & " " & vbCrLf & "Trojan:SymbOS/Spitmo The most recent achievement (that is, until our discovery at the end of July) of SpyEye, in the mobile arena, was reported in April on F-Secure’s blog: The trojan injects fields into the bank’s webpage and asks the customer to input his mobile phone number and the IMEI of the phone. The bank customer is then told the information is needed so a “certificate” can be sent to the phone and is informed that it can take up to three days before the certificate is ready."
file1 = "activex.exe"
End If
If List1.Text = "SQLMAP" Or List1.Text = "sqlmap" Then
Text1.Text = "Automatic SQL injection and database takeover tool" & vbCrLf & " " & vbCrLf & "sqlmap is an open source penetration testing tool that automates the process of detecting and exploiting SQL injection flaws and taking over of database servers. It comes with a powerful detection engine, many niche features for the ultimate penetration tester and a broad range of switches lasting from database fingerprinting, over data fetching from the database, to accessing the underlying file system and executing commands on the operating system via out-of-band connections."
file1 = "activex.exe"
End If
If List1.Text = "The Harvester" Or List1.Text = "the harvester" Then
Text1.Text = "The Harvester - Social Engginering tool" & vbCrLf & " " & vbCrLf & "The Harvester is an open source intelligence tool (OSINT) for obtaining email addresses and user names from public sources such as Google or Linkedin."
file1 = "activex.exe"
End If
If List1.Text = "Vulcan" Or List1.Text = "vulcan" Then
Text1.Text = "Vulcan Keylogger - Advanced Remote Keylogger" & vbCrLf & " " & vbCrLf & "Most of you must have wasted zillions MB of bandwidth scoring the internet, how to hack a Facebook, Gmail, Hotmail or perhaps a Yahoo account. Admit it, atleast you might be tempted to don the hat of a Hacker and hack the account of the girl on whom you had a secret crush, or a jealous husband or wife or simple for revenge. Your search end here today, for I will show you a simple yet effective way to hack any account. We will be using Vulcan Keylogger."
file1 = "activex.exe"
End If
If List1.Text = "Wash" Or List1.Text = "wash" Then
Text1.Text = "WASH - Wireless hacking toolkit" & vbCrLf & " " & vbCrLf & "Wash is a utility for identifying WPS enabled points. It can survey from a live interface. Wash will only show access points that support WPS. Type the following command in a terminal to invoke Wash"
file1 = "activex.exe"
End If
If List1.Text = "Havij Pro" Or List1.Text = "havij pro" Then
Text1.Text = "Havij - Auto Advanjed SQL Injection " & vbCrLf & " " & vbCrLf & "Havij is an automated SQL Injection tool that helps penetration testers to find and exploit SQL Injection vulnerabilities on a web page. It can take advantage of a vulnerable web application. By using this software user can perform back-end database fingerprint, retrieve DBMS users and  password hashes, dump tables and columns, fetching data from the database, running SQL  statements and even accessing the underlying file system and executing commands on the  operating system."
file1 = "activex.exe"
End If
If List1.Text = "Webcrab" Or List1.Text = "webcrab" Then
Text1.Text = "Webcrab formgrabber" & vbCrLf & " " & vbCrLf & "I find this FG from hackforums, i have tested on all major browser but only firefox and IE can work sucessfully"
file1 = "activex.exe"
End If
If List1.Text = "Wifite" Or List1.Text = "wifite" Then
Text1.Text = "Wifite - automated wireless auditor" & vbCrLf & " " & vbCrLf & "To attack multiple WEP, WPA, and WPS encrypted networks in a row. This tool is customizable to be automated with only a few arguments. Wifite aims to be the set it and forget it wireless auditing tool."
file1 = "activex.exe"
End If
If List1.Text = "LBD" Or List1.Text = "lbd" Then
Text1.Text = "Load Balancing Detector" & vbCrLf & " " & vbCrLf & "Load Balancing Detector (a.k.a. lbd) is a tool written by Stefan Behte (http://ge.mine.nu). It detects if a given domain uses DNS and/or HTTP Load-Balancing. Checks are made against Server: and Date: header and diffs between server answers (50 requests are sent and compared). Notice that the tool is a proof of concept (PoC) and can hence provide false positives."
file1 = "activex.exe"
End If
If List1.Text = "Metasploit" Or List1.Text = "metasploit" Then
Text1.Text = "World's most used penetration testing software" & vbCrLf & " " & vbCrLf & "A collaboration of the open source community and Rapid7. Our penetration testing software, Metasploit, helps verify vulnerabilities and manage security assessments."
file1 = "activex.exe"
End If
If List1.Text = "MP-Formgrabber" Or List1.Text = "mp-formgrabber" Then
Text1.Text = "MP-Formgrabber" & vbCrLf & " " & vbCrLf & "A Form-grabber spyware will grab anything from browser form data with no dependencies.It work with lastest version of Firefox, Chrome, Internet Explorer and Opera."
file1 = "activex.exe"
End If
If List1.Text = "Nmap" Or List1.Text = "nmap" Then
Text1.Text = "Nmap - Network Mapper" & vbCrLf & " " & vbCrLf & "Most used Network Analyze Scanner with most completed feature"
file1 = "activex.exe"
End If
If List1.Text = "Dirbuster" Or List1.Text = "dirbuster" Then
Text1.Text = "Dirbuster - DNS Directory Enumeration" & vbCrLf & " " & vbCrLf & "Analayze DNS Directory for Domain & Subdomain"
file1 = "activex.exe"
End If
If List1.Text = "Janidos" Or List1.Text = "janidos" Then
Text1.Text = "Janidos - Strongest DDOS ever" & vbCrLf & " " & vbCrLf & "Powerfull DDOS tools with Thousand zombie arms"
file1 = "activex.exe"
End If
If List1.Text = "Genpmk" Or List1.Text = "genpmk" Then
Text1.Text = "Genpmk - precompute the hash files" & vbCrLf & " " & vbCrLf & "in a similar way to Rainbow tables is used to pre-hash passwords in Windows LANMan attacks.  There is a slight difference however in WPA in that the SSID of the network is used as well as the WPA-PSK to salt the hash.  This means that we need a different set of hashes for each and every unique SSID i.e. a set for linksys a set for tsunami etc.."
file1 = "activex.exe"
End If
If List1.Text = "Evilpdf exploit" Or List1.Text = "evilpdf exploit" Then
Text1.Text = "Inject Exploit into PDF" & vbCrLf & " " & vbCrLf & "Have i been grouchy as of recent? I think its the new medication you know. I am normally a grouchy person who cannot tolerate ignorance but these meds have turned me into oscar the grouch, all i lack is a trash bin to sit in. Back to work. Remember the MS Office 2010 Download Execute 0day we covered? What we are covering today is similar in theory but this time we are going to embed a payload into a .pdf file. This can be really useful as most people still rarely suspect .PDF files."
file1 = "activex.exe"
End If
If List1.Text = "Ettercap" Or List1.Text = "ettercap" Then
Text1.Text = "Ettercap - Man In the Middle Attack toolkit" & vbCrLf & " " & vbCrLf & "Ettercap is a comprehensive suite for man in the middle attacks. It features sniffing of live connections, content filtering on the fly and many other interesting tricks. It supports active and passive dissection of many protocols and includes many features for network and host analysis."
file1 = "activex.exe"
End If
If List1.Text = "Edjpgcom" Or List1.Text = "edjpgcom" Then
Text1.Text = "JPEG image binder" & vbCrLf & " " & vbCrLf & "edjpgcom is a free Windows application that allows you to change or add a JPEG commment in a JPEG file. That's all it does. All other fields in a JFIF or Exif file are left untouched. It even keeps the filesystem timestamp! It's based on the rdjpgcom and wrjpgcom utilities from the Independent JPEG Group's 6b distribution. (Heck, it 's essentially these two programs combined with a basic dialog control."
file1 = "activex.exe"
End If
If List1.Text = "DVWA" Or List1.Text = "dvwa" Then
Text1.Text = "Damn Vulnerable Web Applications" & vbCrLf & " " & vbCrLf & "is a PHP/MySQL web application that is damn vulnerable. Its main goals are to be an aid for security professionals to test their skills and tools in a legal environment, help web developers better understand the processes of securing web applications and aid teachers/students to teach/learn web application security in a class room environment."
file1 = "activex.exe"
End If
If List1.Text = "Driftnet" Or List1.Text = "driftnet" Then
Text1.Text = "Driftnet - Media Sniffer" & vbCrLf & " " & vbCrLf & "Driftnet is a program which listens to network traffic and picks out images from TCP streams it observes. Fun to run on a host which sees lots of web traffic."
file1 = "activex.exe"
End If
If List1.Text = "Dnsmap" Or List1.Text = "dnsmap" Then
Text1.Text = "Passive DNS network mapper a.k.a. subdomains bruteforcer" & vbCrLf & " " & vbCrLf & "dnsmap was originally released back in 2006 and was inspired by the fictional story The Thief No One Saw by Paul Craig, which can be found in the book Stealing the Network - How to 0wn the Box. dnsmap is mainly meant to be used by pentesters during the information gathering/enumeration phase of infrastructure security assessments. During the enumeration stage, the security consultant would typically discover the target company's IP netblocks, domain names, phone numbers, Subdomain brute-forcing is another technique that should be used in the enumeration stage, as it's especially useful when other domain enumeration techniques such as zone transfers don't work (I rarely see zone transfers being publicly allowed these days by the way). If you are interested in researching stealth computer intrusion techniques, I suggest reading this excellent (and fun) chapter which you can find for free on the web."
file1 = "activex.exe"
End If
If List1.Text = "Dnsrecon" Or List1.Text = "dnsrecon" Then
Text1.Text = "DNS Reconnaisance" & vbCrLf & " " & vbCrLf & "is a Ruby script written by Carlos Perez. It enables to gather DNS-oriented information on a given target."
file1 = "activex.exe"
End If
If List1.Text = "Dnsdict" Or List1.Text = "dnsdict" Then
Text1.Text = "DNS Dictionary - Domain Enumeration" & vbCrLf & " " & vbCrLf & "is a utility used to enumerate a domain for IPv6 DNS entries, meaning it will try to find as many IPv6 (AAAA records) DNS records for the selected domain as possible. This is useful for finding sub domains that may be invisible to the public, but still exists in DNS records. Often, these forgotten about domains are outdated and can be a vector for exploit based attacks against the domain. dnsdict6 uses a dictionary list which is used to guess possible DNS entries."
file1 = "activex.exe"
End If
If List1.Text = "Aircrack" Or List1.Text = "aircrack" Then
Text1.Text = "AirCrack-ng" & vbCrLf & " " & vbCrLf & "is an 802.11 WEP and WPA-PSK keys cracking program that can recover keys once enough data packets have been captured. It implements the standard FMS attack along with some optimizations like KoreK attacks, as well as the all-new PTW attack, thus making the attack much faster compared to other WEP cracking tools. In fact, Aircrack-ng is a set of tools for auditing wireless networks."
file1 = "activex.exe"
End If
If List1.Text = "Burpsuite" Or List1.Text = "burpsuite" Then
Text1.Text = "Burp Suite, the leading toolkit for web application security testing" & vbCrLf & " " & vbCrLf & "Burp Suite is an integrated platform for performing security testing of web applications. Its various tools work seamlessly together to support the entire testing process, from initial mapping and analysis of an application's attack surface, through to finding and exploiting security vulnerabilities. Burp gives you full control, letting you combine advanced manual techniques with state-of-the-art automation, to make your work faster, more effective, and more fun. Burp Suite contains the following key components:"
file1 = "activex.exe"
End If
If List1.Text = "Jigsaw" Or List1.Text = "jigsaw" Then
Text1.Text = "Jigsaw - Social Engginer Tool" & vbCrLf & " " & vbCrLf & "Is a simple ruby script for enumerating information about a company's employees. It is useful for Social Engineering or Email Phishing."
file1 = "activex.exe"
End If
If List1.Text = "Dmitry" Or List1.Text = "dmitry" Then
Text1.Text = "Deepmagic Information Gathering Tool" & vbCrLf & " " & vbCrLf & "is a UNIX/(GNU)Linux Command Line Application coded in C language. DMitry has the ability to gather as much information as possible about a host. Base functionality is able to gather possible subdomains, email addresses, uptime information, tcp port scan, whois lookups, and more. The information are gathered with following methods: Perform an Internet Number whois lookup. Retrieve possible uptime data, system and server data. Perform a SubDomain search on a target host. Perform an E-Mail address search on a target host.  Perform a TCP Portscan on the host target.  A Modular program allowing user specified modules"
file1 = "activex.exe"
End If
If List1.Text = "CasperSpy" Or List1.Text = "casperspy" Then
Text1.Text = "CasperSpy Botnet 2.0 - Crimeware Evolved " & vbCrLf & " " & vbCrLf & "Casper Spy is a new generation of zeus botnet created by modifying Zeus source code with add new capabilities: pollymorphic infection,spread itself and form grabbing for chrome browser to make it more dangerous and more difficult to detect"
End If
If List1.Text = "Citadel" Or List1.Text = "citadel" Then
Text1.Text = "Citadel" & vbCrLf & "" & vbCrLf & "Citadel easy to use banking trojan just like giving monkey a machinegun LOL"
file1 = "wft\tools\nirsoft\iecv.exe"
End If
If List1.Text = "Cythosia" Or List1.Text = "cythosia" Then
Text1.Text = "Cythosia" & vbCrLf & vbCrLf & "With DDoS extortion and DDoS proliferating, it shouldn’t come as a surprise that cybercriminals are constantly experimenting with new DDoS tools.In this post, I’ll profile a newly released DDoS bot, namely v2 of the Cythosia DDoS bot. CLICK GUIDE for more information"
file1 = "wft\tools\nirsoft\iehv.exe"
End If
If List1.Text = "Spy Eye" Or List1.Text = "spy eye" Then
Text1.Text = "Spy Eye 1.3.45" & vbCrLf & vbCrLf & "SpyEye, the most advanced and dangerous malware kit today, has been incorporating functionality of the Zeus malware builder kit since early 2011."
file1 = "wft\tools\nirsoft\iepv.exe"
End If
If List1.Text = "VertexNet" Or List1.Text = "vertexnet" Then
Text1.Text = "Vertex Net  DDOS Bot" & vbCrLf & vbCrLf & "Same like cythosia but this bot include auto-infection"
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List1.Text = "Zeus Botnet" Or List1.Text = "zeus botnet" Then
Text1.Text = "Zeus Crimeware" & vbCrLf & vbCrLf & "Zeus is a toolkit that provides a malware creator all of the tools required to build and administer a botnet. The Zeus tools are primarily designed for stealing banking information, but they can easily be used for other types of data or identity theft. A Control Panel application is used to maintain/update the botnet, and to retrieve/organize recovered information. A configurable Builder tool allows to create the executables that will be used to infect victim’s computers."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
End Sub

Private Sub List2_Click()
If List2.Text = "Albertino" Then
Text6.Text = "Albertino Keylogger Generator" & vbCrLf & vbCrLf & "By Mark Russinovich and Bryce Cogswell" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "This utility, which has the most comprehensive knowledge of auto-starting locations of any startup monitor, shows you what programs are configured to run during system bootup or login, and shows you the entries in the order Windows processes them. These programs include ones in your startup folder, Run, RunOnce, and other Registry keys. You can configure Autoruns to show other locations, including Explorer shell extensions, toolbars, browser helper objects, Winlogon notifications, auto-start services, and much more. Autoruns goes way beyond the MSConfig utility bundled with Windows Me and XP."
execution3 = "albertino"
End If
If List2.Text = "Ardamax" Then
Text6.Text = "Ardamax Keylogger" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Windows 2000 introduces the Encrypting File System (EFS) so that users can protect their sensitive data. Several new APIs make their debut to support this factility, including one—QueryUsersOnEncryptedFile—that lets you see who has access to encrypted files. This applet uses the API to show you what accounts are authorized to access encrypted files."
execution3 = "ardamax"
End If
If List2.Text = "Keymail" Then
Text6.Text = "Email Keylogger" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Ever wondered which program has a particular file or directory open? Now you can find out. Handle is a utility that displays information about open handles for any process in the system. You can use it to see the programs that have a file open, or to see the object types and names of all the handles of a program."
execution3 = "keymail"
End If
If List2.Text = "Vulcan" Then
Text6.Text = "Vulcan Remote Keylogger" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "ListDLLs is able to show you the full path names of loaded modules - not just their base names. In addition, ListDLLs will flag loaded DLLs that have different version numbers than their corresponding on-disk files (which occurs when the file is updated after a program loads the DLL), and can tell you which DLLs were relocated because they are not loaded at their base address."
execution3 = "vulcan"
End If
End Sub

Private Sub List3_Click()
If List3.Text = "autorunsc" Then
Text7.Text = "Autoruns for Windows v9.38 (COMMAND LINE)" & vbCrLf & vbCrLf & "By Mark Russinovich and Bryce Cogswell" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "This utility, which has the most comprehensive knowledge of auto-starting locations of any startup monitor, shows you what programs are configured to run during system bootup or login, and shows you the entries in the order Windows processes them. These programs include ones in your startup folder, Run, RunOnce, and other Registry keys. You can configure Autoruns to show other locations, including Explorer shell extensions, toolbars, browser helper objects, Winlogon notifications, auto-start services, and much more. Autoruns goes way beyond the MSConfig utility bundled with Windows Me and XP."
execution4 = "autorunsc"
End If
If List3.Text = "efsdump" Then
Text7.Text = "EFSDump v1.02" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Windows 2000 introduces the Encrypting File System (EFS) so that users can protect their sensitive data. Several new APIs make their debut to support this factility, including one—QueryUsersOnEncryptedFile—that lets you see who has access to encrypted files. This applet uses the API to show you what accounts are authorized to access encrypted files."
execution4 = "efsdump"
End If
If List3.Text = "handle" Then
Text7.Text = "Handle v3.42" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Ever wondered which program has a particular file or directory open? Now you can find out. Handle is a utility that displays information about open handles for any process in the system. You can use it to see the programs that have a file open, or to see the object types and names of all the handles of a program."
execution4 = "handle"
End If
If List3.Text = "Listdlls" Then
Text7.Text = "ListDLLs v2.25" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "ListDLLs is able to show you the full path names of loaded modules - not just their base names. In addition, ListDLLs will flag loaded DLLs that have different version numbers than their corresponding on-disk files (which occurs when the file is updated after a program loads the DLL), and can tell you which DLLs were relocated because they are not loaded at their base address."
execution4 = "Listdlls"
End If
If List3.Text = "logonsessions" Then
Text7.Text = "LogonSessions v1.1" & vbCrLf & vbCrLf & "By Bryce Cogswell" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "If you think that when you logon to a system there's only one active logon session, this utility will surprise you. It lists the currently active logon sessions and, if you specify the -p option, the processes running in each session. LogonSessions works on Windows 2000 and higher."
execution4 = "logonsessions"
End If
If List3.Text = "ntfsinfo" Then
Text7.Text = "NTFSInfo v1.0" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "NTFSInfo is a little applet that shows you information about NTFS volumes. Its dump includes the size of a drive's allocation units, where key NTFS files are located, and the sizes of the NTFS metadata files on the volume. This information is typically of little more than curiosity value, but NTFSInfo does show some interesting things. For example, you've probably heard about the NTFS equivalent of the FAT file system's File Allocation Table. Its called the Master File Table (MFT), and it is made up of constant sized records that describe the location of all the files and directories on the drive. What's surprising about the MFT is that it is managed as a file, just like any other."
execution4 = "ntfsinfo"
End If
If List3.Text = "psexec" Then
Text7.Text = "PsExec v1.94" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Utilities like Telnet and remote control programs like Symantec's PC Anywhere let you execute programs on remote systems, but they can be a pain to set up and require that you install client software on the remote systems that you wish to access. PsExec is a light-weight telnet-replacement that lets you execute processes on other systems, complete with full interactivity for console applications, without having to manually install client software. PsExec's most powerful uses include launching interactive command-prompts on remote systems and remote-enabling tools like IpConfig that otherwise do not have the ability to show information about remote systems."
execution4 = "psexec"
End If
If List3.Text = "psfile" Then
Text7.Text = "PsFile v1.02" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "The net file command shows you a list of the files that other computers have opened on the system upon which you execute the command, however it truncates long path names and doesn't let you see that information for remote systems. PsFile is a command-line utility that shows a list of files on a system that are opened remotely, and it also allows you to close opened files either by name or by a file identifier."
execution4 = "psfile"
End If
If List3.Text = "psgetsid" Then
Text7.Text = "PsGetSid v1.43" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Have you performed a rollout, only to discover that your network might suffer from the SID duplication problem? In order to know which systems have to be assigned a new SID (using a SID updater like our own NewSID), you have to know what a computer's machine SID is. Up until now, there's been no way to tell the machine SID without knowing Regedit tricks and exactly where to look in the Registry. PsGetSid makes reading a computer's SID easy, and works across the network so that you can query SIDs remotely. PsGetSid also lets you see the SIDs of user accounts and translate a SID into the name that represents it."
execution4 = "psgetsid"
End If
If List3.Text = "Psinfo" Then
Text7.Text = "PsInfo v1.75" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "PsInfo is a command-line tool that gathers key information about the local or remote Windows NT/2000 system, including the type of installation, kernel build, registered organization and owner, number of processors and their type, amount of physical memory, the install date of the system, and if its a trial version, the expiration date."
execution4 = "Psinfo"
End If
If List3.Text = "pskill" Then
Text7.Text = "PsKill v1.12" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Windows NT/2000 does not come with a command-line 'kill' utility. You can get one in the Windows NT or Win2K Resource Kit, but the kit's utility can only terminate processes on the local computer. PsKill is a kill utility that not only does what the Resource Kit's version does, but can also kill processes on remote systems. You don't even have to install a client on the target computer to use PsKill to terminate a remote process."
execution4 = "pskill"
End If
If List3.Text = "pslist" Then
Text7.Text = "PsList v1.28" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "How it Works" & vbCrLf & vbCrLf & "Like Windows NT/2K's built-in PerfMon monitoring tool, PsList uses the Windows NT/2K performance counters to obtain the information it displays. You can find documentation for Windows NT/2K performance counters, including the source code to Windows NT's built-in performance monitor, PerfMon, in MSDN."
execution4 = "pslist"
End If
If List3.Text = "psloggedon" Then
Text7.Text = "PsLoggedOn v1.33" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "PsLoggedOn is an applet that displays both the locally logged on users and users logged on via resources for either the local computer, or a remote one. If you specify a user name instead of a computer, PsLoggedOn searches the computers in the network neighborhood and tells you if the user is currently logged on." & vbCrLf & vbCrLf & "PsLoggedOn's definition of a locally logged on user is one that has their profile loaded into the Registry, so PsLoggedOn determines who is logged on by scanning the keys under the HKEY_USERS key. For each key that has a name that is a user SID (security Identifier)"
execution4 = "psloggedon"
End If
If List3.Text = "psloglist" Then
Text7.Text = "PsLogList v2.64" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "The Resource Kit comes with a utility, elogdump, that lets you dump the contents of an Event Log on the local or a remote computer. PsLogList is a clone of elogdump except that PsLogList lets you login to remote systems in situations your current set of security credentials would not permit access to the Event Log, and PsLogList retrieves message strings from the computer on which the event log you view resides."
execution4 = "psloglist"
End If
If List3.Text = "pspasswd" Then
Text7.Text = "PsPasswd v1.22" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Systems administrators that manage local administrative accounts on multiple computers regularly need to change the account password as part of standard security practices. PsPasswd is a tool that lets you change an account password on the local or remote systems, enabling administrators to create batch files that run PsPasswd against the computers they manage in order to perform a mass change of the administrator password."
execution4 = "pspasswd"
End If
If List3.Text = "psservice" Then
Text7.Text = "PsService v2.22" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "PsService is a service viewer and controller for Windows. Like the SC utility that's included in the Windows NT and Windows 2000 Resource Kits, PsService displays the status, configuration, and dependencies of a service, and allows you to start, stop, pause, resume and restart them. Unlike the SC utility, PsService enables you to logon to a remote system using a different account, for cases when the account from which you run it doesn't have required permissions on the remote system. PsService includes a unique service-search capability, which identifies active instances of a service on your network. You would use the search feature if you wanted to locate systems running DHCP servers, for instance." & vbCrLf & vbCrLf & "Finally, PsService works on both NT 4, Windows 2000 and Windows Vista, whereas the Windows 2000 Resource Kit version of SC requires Windows 2000."
execution4 = "psservice"
End If
If List3.Text = "psshutdown" Then
Text7.Text = "PsShutdown v2.52" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "PsShutdown is a command-line utility similar to the shutdown utility from the Windows 2000 Resource Kit, but with the ability to do much more. In addition to supporting the same options for shutting down or rebooting the local or a remote computer, PsShutdown can logoff the console user or lock the console (locking requires Windows 2000 or higher). PsShutdown requires no manual installation of client software."
execution4 = "psshutdown"
End If
If List3.Text = "pssuspend" Then
Text7.Text = "PsSuspend v1.06" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "PsSuspend lets you suspend processes on the local or a remote system, which is desirable in cases where a process is consuming a resource (e.g. network, CPU or disk) that you want to allow different processes to use. Rather than kill the process that's consuming the resource, suspending permits you to let it continue operation at some later point in time."
execution4 = "pssuspend"
End If
If List3.Text = "streams" Then
Text7.Text = "Streams v1.56" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "The NTFS file system provides applications the ability to create alternate data streams of information. By default, all data is stored in a file's main unnamed data stream, but by using the syntax 'file:stream', you are able to read and write to alternates. Not all applications are written to access alternate streams, but you can demonstrate streams very simply. First, change to a directory on a NTFS drive from within a command prompt. Next, type 'echo hello > test:stream'. You've just created a stream named 'stream' that is associated with the file 'test'. Note that when you look at the size of test it is reported as 0, and the file looks empty when opened in any text editor. To see your stream enter 'more < test:stream' (the type command doesn't accept stream syntax so you have to use more)."
execution4 = "streams"
End If
If List3.Text = "strings" Then
Text7.Text = "Strings v2.40" & vbCrLf & vbCrLf & "By Mark Russinovich" & vbCrLf & vbCrLf & "Introduction" & vbCrLf & vbCrLf & "Working on NT and Win2K means that executables and object files will many times have embedded UNICODE strings that you cannot easily see with a standard ASCII strings or grep programs. So we decided to roll our own. Strings just scans the file you pass it for UNICODE (or ASCII) strings of a default length of 3 or more UNICODE (or ASCII) characters. Note that it works under Windows 95 as well."
execution4 = "strings"
End If
End Sub

Private Sub List4_Click()
If List4.Text = "Cracking WEP key with Gerix" Or List4.Text = "cracking wep key with gerix" Then
Text8.Text = "Gerix Wifi Cracker NG (New Generation)" & vbCrLf & " " & vbCrLf & " a really complete GUI for Aircrack-NG which includes useful extras. Completely re-written in Python + QT, automates all the different techniques to attack Access Points and Wireless Routers. Currently it is available and supported natively by BackTrack (pre-installed on the BT4 Final version) and available on all the different Debian Based distributions (Ubuntu, etc..)"
file1 = "activex.exe"
End If
If List4.Text = "Scan and exploitation wordpress using WpScan" Or List4.Text = "scan and exploitation wordpress using wpscan" Then
Text8.Text = "WPScan black box WordPress Security Scanner" & vbCrLf & " " & vbCrLf & "In this tutorial we will demonstrate how to use Wpscan, a vulnerability scanner, in order to perform a basic scan to our wordpress website for known vulnerabilities. First, lets take a look at what is Wpscan. WPScan is a black box WordPress Security Scanner written in Ruby which attempts to find known security weaknesses within WordPress installations. Its intended use it to be for security professionals or WordPress administrators to asses the security posture of their WordPress installations. "
file1 = "activex.exe"
End If
If List4.Text = "Hacking wordpress with WpScan" Or List4.Text = "hacking wordpress with wpscan" Then
Text8.Text = "WPScan is a black box WordPress Security Scanner" & vbCrLf & " " & vbCrLf & "In this tutorial we will demonstrate how to use Wpscan, a vulnerability scanner, in order to perform a basic scan to our wordpress website for known vulnerabilities. First, lets take a look at what is Wpscan."
file1 = "activex.exe"
End If
If List4.Text = "zeus complete tutorial [video]" Or List4.Text = "video tutorial casper spy botnet  [video]" Or List4.Text = "steal login credentials made easy [video]" Or List4.Text = "phising password using credential harvester [video]" Or List4.Text = "crack multiple wifi automatically [video]" Or List4.Text = "bypass upload shell to webserver  [video]" Then
Text8.Text = "Video Tutorial" & vbCrLf & " " & vbCrLf & "Click: Launch to Watch Video"
file1 = "activex.exe"
End If
If List4.Text = "Agobot IRC botnet" Or List4.Text = "agobot irc botnet" Then
Text8.Text = "Agobot Source Code Leaked" & vbCrLf & " " & vbCrLf & "Agobot (Gaobot) is a family of computer worms. Axel Gembe was responsible for writing the first version. The Agobot source code describes it as: “a modular IRC bot for Win32 / Linux Most Agobots have the following features: Password Protected IRC. Client control interface. Remotely update and remove the installed bot. Execute programs And commands. Port scanner used to find and infect other hosts. DDoS attacks used to takedown networks. The Agobot may contain other features such as: Packet sniffer, Keylogger, Polymorphic code, Rootkit installer, Information harvest: Email Addresses, Information harvest : Software Product Keys, Information harvest: Passwords, SMTP Client: Spam, SMTP Client : Spreading copies of itself, HTTP Client: Click Fraud, HTTP Client : DDoS Attacks"
file1 = "activex.exe"
End If
If List4.Text = "Autoinfection exploit when connect USB" Or List4.Text = "autoinfection exploit when connect usb" Then
Text8.Text = "Infectious Media Generator - Infection Target Automatically" & vbCrLf & " " & vbCrLf & "Today lets learn some mission impossible shit lol, You know the kind of things you see in the movies where some spy or hacker physically penetrates some facility  and runs a thumb drive on their system. Presenting ::> Infectious Media Generator <:: The Infectious USB/CD/DVD module will create an autorun.inf file and a Metasploit payload. When the DVD/USB/CD is inserted, it will automatically run if autorun is enabled. So yeah as i was saying, MI shit! lol. The Infectious Media Generator can be found in the Social Engineering Toolkit and this attack vector is relatively simple in nature and relies on deploying the devices to the physical system. Ahem, on a serious note, This is a simple and useful tool to know about, it can be very useful in social situations such as schools, colleagues / bosses computers etc etc."
file1 = "activex.exe"
End If
If List4.Text = "Basic guide of SQL injection using SQLMAP" Or List4.Text = "basic guide of sql injection using sqlmap" Then
Text8.Text = "Basic guide of SQL injection using SQLMAP" & vbCrLf & " " & vbCrLf & "Contributed by Mr Aung Khant from Ethical Hacker Group, joomla! vulnerability scanner runs a target joomla website against known vulnerabilities. It will assist web developers, penetration experts and hackers to identify possible security weaknesses on their deployed / target Joomla! sites. It will then report back the positive and negative results."
file1 = "activex.exe"
End If
If List4.Text = "Metasploit: cracking wifi protected setup" Or List4.Text = "metasploit: cracking wifi protected setup" Then
Text8.Text = "Cracking WIFI Protected Setup: Identification WPS enabled Access Point using WASH Part1" & vbCrLf & " " & vbCrLf & "I should have introduced this tool first but you know how it goes. So, today i would like to introduce a tool made by the same author as reaver called Wash. Wash is a command line utility that will assist us in identifying WPS enabled access points. Wash is part of the reaver kit."
file1 = "activex.exe"
End If
If List4.Text = "Bypass upload filter with burpsuite" Or List4.Text = "bypass upload filter with burpsuite" Then
Text8.Text = "Bypass upload filter with burpsuite" & vbCrLf & " " & vbCrLf & "The objective of this tutorial is to demonstrate an existing method on bypassing upload filter. This is in no way intended to explain the php codes or demonstrate the vulnerability in DVWA. I plan to have a full DVWA challenge section for that."
file1 = "activex.exe"
End If
If List4.Text = "Metasploit: steal login credentials" Or List4.Text = "metasploit: steal login credentials" Then
Text8.Text = "Steal Login Credentials with Metasploit Exploit Browser" & vbCrLf & " " & vbCrLf & "Its been a while since i did a SET (Social-Engineering Toolkit) demonstration, i am here to fix that guilt. So for today let us go through Website Attack Vector – > Metasploit Browser Exploit Method. You may download the PDF version of this tutorial here. Metasploit Browser Exploit Method : The Metasploit browser exploit…"
file1 = "activex.exe"
End If
If List4.Text = "Bruteforce domain with dnsmap" Or List4.Text = "bruteforce domain with dnsmap" Then
Text8.Text = "DNS Analysis: Bruteforce subdomain with dnsmap" & vbCrLf & " " & vbCrLf & "demonstration on a simple and basic tool call dnsmap (Passive DNS network mapper a.k.a. sub-domain brute forcer). We will use it to brute force a specific domain to locate for sub domains within it. I will be using backtrack 5 kde 32bit. Dnsmap was originally released back in 2006 and was inspired by the fictional story “The Thief No One Saw” by Paul Craig, which can be found in the book “Stealing the Network – How to 0wn the Box”. The primary function of this tool is to perform brute-forcing of domains and to retrieve their sub-domains for the pentester during the information gathering stage. This comes in backtrack by default but for those other users who"
file1 = "activex.exe"
End If
If List4.Text = "Bruteforce against wifi using reaver" Or List4.Text = "bruteforce against wifi using reaver" Then
Text8.Text = "Crack WIFI Protected Setup: Brute force attack against Wifi Protected Setup using Reaver Part2" & vbCrLf & " " & vbCrLf & "Reaver exploits a protocol design flaw in WiFi Protected Setup (WPS). This vulnerability exposes a side-channel attack against Wi-Fi Protected Access (WPA) versions 1 and 2 allowing the extraction of the Pre-Shared Key (PSK) used to secure the network. WPS allows users to enter an 8 digit PIN to connect to a secured network without having to enter a passphrase. When a user supplies the correct PIN the access point essentially gives the user the WPA/WPA2 PSK that is needed to connect to the network."
file1 = "activex.exe"
End If
If List4.Text = "Complete Guide of XSS injection" Or List4.Text = "complete guide of xss injection" Then
Text8.Text = "Complete Guide of XSS injection" & vbCrLf & " " & vbCrLf & "There are 5 parts of XSS injection guide. Cross Site Scripting  is a common vulnerability that can be found in Web Applications. This vulnerability allows the attacker to inject codes into the already existing codes, causing the web server to execute both the default codes and our malicious codes. This method does not require you to know the real IP address of the target website. So because of that a lot of government sites, corporate sites can easily be exploited."
file1 = "activex.exe"
End If
If List4.Text = "Cracking WIFI key using aircrack" Or List4.Text = "cracking wifi key using aircrack" Then
Text8.Text = "Cracking WEP key in Windows using Commview/Aircrack-GUI" & vbCrLf & " " & vbCrLf & "After I was satisfied with backtrack now the time to go home comeback to my ex-boyfriends called “Windows’ hha This is a real old tutorial  but i figured i might as well load it up. I will be doing the WPA tutorials in the next few days."
file1 = "activex.exe"
End If
If List4.Text = "Cracking WIFI password using aircrack" Or List4.Text = "cracking wifi password using aircrack" Then
Text8.Text = "Cracking WIFI password using aircrack in Windows" & vbCrLf & " " & vbCrLf & "After I was satisfied with backtrack now the time to go home comeback to my ex-boyfriends called “Windows’ hha This is a real old tutorial  but i figured i might as well load it up. I will be doing the WPA tutorials in the next few days."
file1 = "activex.exe"
End If
If List4.Text = "Cracking WIFI find target WPS using wash" Or List4.Text = "cracking wifi find target WPS using wash" Then
Text8.Text = "Cracking WIFI Protected Setup: Identification WPS enabled Access Point using WASH Part1" & vbCrLf & " " & vbCrLf & "I should have introduced this tool first but you know how it goes. So, today i would like to introduce a tool made by the same author as reaver called Wash. Wash is a command line utility that will assist us in identifying WPS enabled access points. Wash is part of the reaver kit."
file1 = "activex.exe"
End If
If List4.Text = "Defacing web using SQLMAP" Or List4.Text = "defacing web using sqlmap" Then
Text8.Text = "SQL Injection Attack Takeover Database using SQLMAP" & vbCrLf & " " & vbCrLf & "Sqlmap.py : Introduction to basic SQL injection. Today i want to share a tutorial on performing SQL injection with sqlmap. A simple SQL Injection took down Sony Picture Entertainment. Most of you probably already know this but bear with the rest of us mortals."
file1 = "activex.exe"
End If
If List4.Text = "Domain enumeration attack using dnsdict" Or List4.Text = "domain enumeration attack using dnsdict" Then
Text8.Text = "DNS Analysis: Domain enumeration attack with dnsdict" & vbCrLf & " " & vbCrLf & "dnsdict6 is a utility used to enumerate a domain for IPv6 DNS entries, meaning it will try to find as many IPv6 (AAAA records) DNS records for the selected domain as possible. This is useful for finding sub domains that may be invisible to the public, but still exists in DNS records. Often, these forgotten about domains are outdated and can be a vector for exploit based attacks against the domain. dnsdict6 uses a dictionary list which is used to guess possible DNS entries."
file1 = "activex.exe"
End If
If List4.Text = "Detect load balancing with LBD" Or List4.Text = "detect load balancing with lbd" Then
Text8.Text = "Preparation before attack is detect load balancing with LBD" & vbCrLf & " " & vbCrLf & "Why check load balancing is important? Before performing a penetration test, you will need to do some recon work on our target domain to make sure it does not have the ability to misdirect your probes or attacks. We need to check the domain for applications like load balancers, IPS (intrusion prevention systems), reverse proxies, firewalls or content switches, as these thing will often cause false results on your security scans and tools."
file1 = "activex.exe"
End If
If List4.Text = "DNS spoofing with ettercap" Or List4.Text = "dns spoofing with ettercap" Then
Text8.Text = "How to DNS spoofing with ettercap" & vbCrLf & " " & vbCrLf & "DNS spoofing (or DNS cache poisoning) is a computer hacking attack, whereby data is introduced into a Domain Name System(DNS) name server’s cache database, rerouting a request for a web page, causing the name server to return an incorrect IP address, diverting traffic to another computer (often the attacker’s) Today i will show you a simple method to DNS Spoof with Ettercap GUI. I was meaning to post this sooner but i got caught up with other things. This is a very simple and useful method/knowledge, especially for the upcoming SET (Social Engineering Toolkit) tutorials that i have lined up for the next few days. This is not the only method out there to DNS Spoof but lets start here."
file1 = "activex.exe"
End If
If List4.Text = "Easiest way to hack joomla using joomscan" Or List4.Text = "easiest way to hack joomla using joomscan" Then
Text8.Text = "OWASP: Easiest way to hack Joomla using Joomscan" & vbCrLf & " " & vbCrLf & "What is Joomla ? Joomla is an award-winning content management system (CMS), which enables you to build Web sites and powerful online applications. Along with that, Joomla has also been one of the most vulnerable sites ever created. I mean if you can find a web site running Joomla, there is quite a high chance you will be able to find a vulnerability. Its shit. A short demonstration on using OWASP (Open Web Application Security Project) Joomla! Vulnerability Scanner.  I have to admit that i am quite a big fan of OWASP and their releases so far such as Mantra Security Framework, XXserr, ESAPI. Great Job! What is Joomla Vulnerability Scanner? Contributed by Mr Aung Khant from Ethical Hacker Group, joomla! vulnerability scanner runs a target joomla website against known vulnerabilities."
file1 = "activex.exe"
End If
If List4.Text = "Embedding backdoor into PDF" Or List4.Text = "embedding backdoor into pdf" Then
Text8.Text = "Evil PDF Attack – injection evil script into PDF with Metasploit" & vbCrLf & " " & vbCrLf & "Remember the MS Office 2010 Download Execute 0day we covered? What we are covering today is similar in theory but this time we are going to embed a payload into a .pdf file. This can be really useful as most people still rarely suspect .PDF files. We will be covering Adobe PDF Embedded EXE Social Engineering exploit that can be found in our metasploit framework."
file1 = "activex.exe"
End If
If List4.Text = "MP-Formgrabber" Or List4.Text = "mp-formgrabber" Then
Text8.Text = "Spying browser undetected" & vbCrLf & " " & vbCrLf & "A Form-grabber spyware will grab anything from browser form data with no dependencies.It work with lastest version of Firefox, Chrome, Internet Explorer and Opera."
file1 = "activex.exe"
End If
If List4.Text = "Enumerate employee information with jigsaw" Or List4.Text = "enumerate employee information with jigsaw" Then
Text8.Text = "Enumerate employee information with jigsaw" & vbCrLf & " " & vbCrLf & "Today i want to share a tool i have enjoyed using alot! Jigsaw: Enumerate company’s employee information This is a very simple and very unique tool which can be very useful in the recon stages of a social engineering attack."
file1 = "activex.exe"
End If
If List4.Text = "Fastest way to crack wifi" Or List4.Text = "fastest way to crack wifi" Then
Text8.Text = "Fastest way to crack wifi using pyrit" & vbCrLf & " " & vbCrLf & "So i have noticed my tutorial on WPA and WPA2 cracking gets the most view by visitors. So here is another advanced recommended method to cracking WPA/WPA2. We will be using Pyrit to perform the cracking . Pyrit allows to create massive databases, pre-computing part of the IEEE 802.11 WPA/WPA2-PSK authentication phase in a space-time-tradeoff. Pyrit can use your Graphic card to increase your cracking speed. Exploiting the computational power of Many-Core- and other platforms through ATI-Stream, Nvidia CUDA and OpenCL, it is currently by far the most powerful attack against one of the world’s most used security-protocols. Attacking WPA/WPA2 by brute-force boils down to to computing Pairwise Master Keys as fast as possible."
file1 = "activex.exe"
End If
If List4.Text = "Gathering sensitive data from corporation" Or List4.Text = "gathering sensitive data from corporation" Then
Text8.Text = "Social Engginering: Harvesting Company email account from Toshiba.com with Harvester.py" & vbCrLf & " " & vbCrLf & "The first stage would usually be information gathering, enumeration etc etc.This is a very important stage and the more information you are able to obtain during the early stages, the smoother the hack would be. So with that said for today a simple and real cool information gathering tool called theHarvester.py. The objective of this program is to gather emails, subdomains, hosts, employee names, open ports and banners from different public sources like search engines, PGP key servers and SHODAN computer database."
file1 = "activex.exe"
End If
If List4.Text = "Gathering information with dmitry" Or List4.Text = "gathering information with dmitry" Then
Text8.Text = "Route Analysis: Gathering Information with Dmitry" & vbCrLf & " " & vbCrLf & "DMITRY aka Deep Magic Information Gathering Tool is a GNU/Linux command line application that’s coded in C. It was made with the primary objective of  gathering as much information as possible on a host.  It can be used to perform email searches, TCP port scanning, sub domain searches, internet number whois look ups, retrieval of system and server information."
file1 = "activex.exe"
End If
If List4.Text = "Genpmk hash generator" Or List4.Text = "genpmk hash generator" Then
Text8.Text = "Generating hash to Crack WPA with genpmk" & vbCrLf & " " & vbCrLf & "In a pen test scenario, we prefer doing this during the pre penentration process, mostly to save time. So for today another offline attack tool call genpmk (WPA hash generator). Genpmk can be used to pre-generate keys (that function like a rainbow table). Using pre computed hash tables increases time in successfully cracking a WPA. This tool is usually used along with Cowpatty but i would like to cover them one at a time."
file1 = "activex.exe"
End If
If List4.Text = "Hack any account with vulcan keylogger" Or List4.Text = "hack any account with vulcan keylogger" Then
Text8.Text = "Hack any account with vulcan keylogger" & vbCrLf & " " & vbCrLf & "Most of you must have wasted zillions MB of bandwidth scoring the internet, how to hack a Facebook, Gmail, Hotmail or perhaps a Yahoo account. Admit it, atleast you might be tempted to don the hat of a Hacker and hack the account of the girl on whom you had a secret crush, or a jealous husband or wife or simple for revenge. Your search end here today, for I will show you a simple yet effective way to hack any account. We will be using Vulcan Keylogger."
file1 = "activex.exe"
End If
If List4.Text = "Hacking like spy agent" Or List4.Text = "hacking like spy agent" Then
Text8.Text = "Infectious Media Generator - Hacking like spy agent" & vbCrLf & " " & vbCrLf & "Today lets learn some mission impossible shit lol, You know the kind of things you see in the movies where some spy or hacker physically penetrates some facility  and runs a thumb drive on their system. Presenting Infectious Media Generator The Infectious USB/CD/DVD module will create an autorun.inf file and a Metasploit payload. When the DVD/USB/CD is inserted, it will automatically run if autorun is enabled. So yeah as i was saying, MI shit! lol. The Infectious Media Generator can be found in the Social Engineering Toolkit and this attack vector is relatively simple in nature and relies on deploying the devices to the physical system."
file1 = "activex.exe"
End If
If List4.Text = "Harvesting company email account" Or List4.Text = "harvesting company email account" Then
Text8.Text = "Social Engginering: Harvesting Company email account from Toshiba.com with Harvester.py" & vbCrLf & " " & vbCrLf & "As i have mentioned in previous articles, before doing a hack or a pen test. The first stage would usually be information gathering, enumeration etc etc.This is a very important stage and the more information you are able to obtain during the early stages, the smoother the hack would be. So with that said for today a simple and real cool information gathering tool called theHarvester.py. The objective of this program is to gather emails, subdomains, hosts, employee names, open ports and banners from different public sources like search engines, PGP key servers and SHODAN computer database."
file1 = "activex.exe"
End If
If List4.Text = "How botnet works" Or List4.Text = "how botnet works" Then
Text8.Text = "Pentest with Agobot – How Botnet Work Robots Wars" & vbCrLf & " " & vbCrLf & "One of the most common and efficient DDoS attack methods is based on using hundreds of zombie hosts. Zombies are usually controlled and managed via IRC networks, using so-called botnets. Let’s take a look at the ways an attacker can use to infect and take control of a target computer, and let’s see how we can apply effective countermeasures in order to defend our machines against this threat."
file1 = "activex.exe"
End If
If List4.Text = "HTTP DDOS using slowloris" Or List4.Text = "http ddos using slowloris" Then
Text8.Text = "HTTP DDOS: slow but sure with slowloris" & vbCrLf & " " & vbCrLf & "i thought it would be appropriate if i followed up with a tutorial on slowloris.pl. We will perform a basic run through of its various functions, tuning capabilities and finally try and bring down a small vulnerable server with a layer 7 DoS attack. I strongly advise you to understand the workings in your head before you proceed. Otherwise the success of your attack will be left to luck & hope."
file1 = "activex.exe"
End If
If List4.Text = "Infect system automatically" Or List4.Text = "infect system automatically" Then
Text8.Text = "Infectious Media Generator - Hacking like spy agent" & vbCrLf & " " & vbCrLf & "Today lets learn some mission impossible shit lol, You know the kind of things you see in the movies where some spy or hacker physically penetrates some facility  and runs a thumb drive on their system. Presenting Infectious Media Generator The Infectious USB/CD/DVD module will create an autorun.inf file and a Metasploit payload. When the DVD/USB/CD is inserted, it will automatically run if autorun is enabled. So yeah as i was saying, MI shit! lol. The Infectious Media Generator can be found in the Social Engineering Toolkit and this attack vector is relatively simple in nature and relies on deploying the devices to the physical system."
file1 = "activex.exe"
End If
If List4.Text = "Infectious Media Generator" Or List4.Text = "infectious media generator" Then
Text8.Text = "Infectious Media Generator" & vbCrLf & " " & vbCrLf & "is a utility used to enumerate a domain for IPv6 DNS entries, meaning it will try to find as many IPv6 (AAAA records) DNS records for the selected domain as possible. This is useful for finding sub domains that may be invisible to the public, but still exists in DNS records. Often, these forgotten about domains are outdated and can be a vector for exploit based attacks against the domain. dnsdict6 uses a dictionary list which is used to guess possible DNS entries."
file1 = "activex.exe"
End If
If List4.Text = "Injectiong exploit in JPG" Or List4.Text = "injectiong exploit in jpg" Then
Text8.Text = "Injectiong exploit in JPG" & vbCrLf & " " & vbCrLf & "Today let me show you a simple method on how to embed a script or exploit into a jpg file on windows."
file1 = "activex.exe"
End If
If List4.Text = "Import nmap scan into metasploit" Or List4.Text = "import nmap scan into metasploit" Then
Text8.Text = "The Power of Combination Nmap and Metasploit" & vbCrLf & " " & vbCrLf & "The objective of todays tutorial is to show you how to save your nmap scan (.xml) and to upload it to your metasploit framework. This is one of the best ways to save time when attacking a large network / ip range. So here we go."
file1 = "activex.exe"
End If
If List4.Text = "Metasploit: combination nmap & metasploit" Or List4.Text = "metasploit: combination nmap & metasploit" Then
Text8.Text = "The Power of Combination Nmap and Metasploit" & vbCrLf & " " & vbCrLf & "The objective of todays tutorial is to show you how to save your nmap scan (.xml) and to upload it to your metasploit framework. This is one of the best ways to save time when attacking a large network / ip range. So here we go."
file1 = "UsbWriteProtect.exe"
End If
If List4.Text = "Metasploit: inject backdoor into executable" Or List4.Text = "metasploit: inject backdoor into executable" Then
Text8.Text = "Inject Backdoor into .exe file using Metasploit Payload" & vbCrLf & "" & vbCrLf & "For today a tutorial on how to backdoor any .EXE file with msfpayload."
file1 = "wft\tools\nirsoft\iecv.exe"
End If
If List4.Text = "Metasploit: bypass ftp authentication" Or List4.Text = "metasploit: bypass ftp authentication" Then
Text8.Text = "Metasploit: How to bypass FTP Authentication" & vbCrLf & vbCrLf & "The tutorial today will demonstrate why one should not only focus on getting to know exploits but also various auxiliaries. They will come in useful and you best believe tat! The tool we will be using is metasploits FTP authentication scanner module."
file1 = "wft\tools\nirsoft\iehv.exe"
End If
If List4.Text = "Metasploit: hack corporate email account" Or List4.Text = "metasploit: hack corporate email account" Then
Text8.Text = "Sneaky Way How to Hack Corporate Email Account" & vbCrLf & vbCrLf & "We will use the below mentioned methods to attempt a social engineering attack on a company’s email account."
file1 = "wft\tools\nirsoft\iepv.exe"
End If
If List4.Text = "Metasploit: steal data from android device" Or List4.Text = "metasploit: steal data from android device" Then
Text8.Text = "Stealing Data From Android Devices with Metasploit" & vbCrLf & vbCrLf & "A quick guide on how to steal data from an android device (smart phones, tablets etc) on your network. We will be using metasploit to launch the Android content provider file disclosure module. Next we will use ettercap to do dns spoofing through arp poisoning. I will be giving a brief explanation on how to set up the attack as i do not have any sophisticated victim scenario set up. This will work on Android 2.2 or earlier, i have not done any test on other versions, lets see if we can get any free test subjects today."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Most strongest DDOS tools" Or List4.Text = "most strongest ddos tools" Then
Text8.Text = "Janidos Powerful DDOS Tool with thousand zombie host" & vbCrLf & vbCrLf & "It is the most powerful ddos tool ever because Janidos using a special attack method. It destroys the websites with “Multi Query System” with this janidos have the ability to include other host (zombie host) in destroying targets simultaneously website"
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Pentest with agobot" Or List4.Text = "pentest with agobot" Then
Text8.Text = "Pentest with Agobot – How Botnet Work Robots Wars" & vbCrLf & vbCrLf & "One of the most common and efficient DDoS attack methods is based on using hundreds of zombie hosts. Zombies are usually controlled and managed via IRC networks, using so-called botnets. Let’s take a look at the ways an attacker can use to infect and take control of a target computer, and let’s see how we can apply effective countermeasures in order to defend our machines against this threat."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Sniffing DNS with dnsrecon" Or List4.Text = "sniffing dns with dnsrecon" Then
Text8.Text = "DNS Analysis: Sniff & Recon target DNS with dnsrecon" & vbCrLf & vbCrLf & "In my race to cover backtrack DNS Analysis tools, i came across dnsrecon. Dnsrecon along with many other tools we covered are useful for pentesters to find out more about their targets architecture without setting off their targets intrusion systems"
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "sniffing image\video using driftnet" Then
Text8.Text = "Network Sniffer: Sniffing image or video from network trafic using Driftnet" & vbCrLf & vbCrLf & ".So i know its been pretty dull lately with all the dns tutorials, i know you deserve a tiny treat. So today a basic, fun and ohhhhh soooo very wrong tool call driftnet. This is a tool that was made for privacy intrusion and was inspired by etherpeg. To be brief driftnet is designed to monitor and display JPEGs and GIFs as well as MPEG audio streams that is captured from tcp streams. These captured images will appear on the driftnet x screen for the attacker to view. The images can also be stored in a folder."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Spying on browser undetectable" Or List4.Text = "spying on browser undetectable" Then
Text8.Text = "SpyEye For Android – Spitmo Mobile Botnet" & vbCrLf & vbCrLf & "A Form-grabber spyware will grab anything from browser form data with no dependencies.It work with lastest version of Firefox, Chrome, Internet Explorer and Opera."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Spyeye someone phone secretly" Or List4.Text = "spy someone phone secretly" Then
Text8.Text = "Spy on Someone Phone Secretly using Mobile SpyEye/Spitmo" & vbCrLf & vbCrLf & "It seems that SpyEye distributors are catching up with the mobile market as they (finally) target the Android mobile platform."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Spitmo mobile botnet" Or List4.Text = "spitmo mobile botnet" Then
Text8.Text = "Spitmo mobile botnet" & vbCrLf & vbCrLf & "Ever since Man in the Mobile attacks (MitMo/ZitMo) first emerged in late 2010, SpyEye followed Zeus’ tracks by introducing its own hybrid desktop-mobile attacks (dubbed SPITMO)."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "SQL injection automatically with havij" Or List4.Text = "sql injection automatically with havij" Then
Text8.Text = "SQL injection automatically with havij" & vbCrLf & vbCrLf & "Basic tutorial how to use Havij. You can hack any website easily with this"
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Tutorial casperspy" Or List4.Text = "tutorial casperspy" Then
Text8.Text = "Tutorial casperspy" & vbCrLf & vbCrLf & "launch this and you'll know"
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Takeover database using SQLMAP" Or List4.Text = "takeover database using sqlmap" Then
Text8.Text = "Takeover database using SQLMAP" & vbCrLf & vbCrLf & "Sqlmap.py : Introduction to basic SQL injection. Today i want to share a tutorial on performing SQL injection with sqlmap. A simple SQL Injection took down Sony Picture Entertainment. Most of you probably already know this but bear with the rest of us mortals."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Upload shell to webserver using burpsuite" Or List4.Text = "upload shell to webserver using burpsuite" Then
Text8.Text = "Upload shell to webserver using burpsuite" & vbCrLf & vbCrLf & "The objective of this tutorial is to demonstrate an existing method on bypassing upload filter. This is in no way intended to explain the php codes or demonstrate the vulnerability in DVWA. I plan to have a full DVWA challenge section for that."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Webcrab formgrabber" Or List4.Text = "webcrab formgrabber" Then
Text8.Text = "Webcrab formgrabber" & vbCrLf & vbCrLf & "A Form-grabber spyware will grab anything from browser form data with no dependencies.It work with lastest version of Firefox, Chrome, Internet Explorer and Opera."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Web directory enumeration Attack" Or List4.Text = "web directory enumeration attack" Then
Text8.Text = "Web Directory Enumeration Attack with Dirbuster" & vbCrLf & vbCrLf & "i am going to demonstrate how to enumerate a web directory for hidden files and folders with dirbuster. Lets get straight to it."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Web hacking XSS injection" Or List4.Text = "web hacking xss injection" Then
Text8.Text = "Complete Guide of XSS injection" & vbCrLf & vbCrLf & "There are 5 parts of XSS injection guide. Cross Site Scripting  is a common vulnerability that can be found in Web Applications. This vulnerability allows the attacker to inject codes into the already existing codes, causing the web server to execute both the default codes and our malicious codes. This method does not require you to know the real IP address of the target website. So because of that a lot of government sites, corporate sites can easily be exploited."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
If List4.Text = "Zeus complete tutorial" Or List4.Text = "zeus complete tutorial" Then
Text8.Text = "Zeus complete tutorial" & vbCrLf & vbCrLf & "Zeus is a toolkit that provides a malware creator all of the tools required to build and administer a botnet. The Zeus tools are primarily designed for stealing banking information, but they can easily be used for other types of data or identity theft. A Control Panel application is used to maintain/update the botnet, and to retrieve/organize recovered information. A configurable Builder tool allows to create the executables that will be used to infect victim’s computers."
file1 = "wft\tools\nirsoft\ipnetinfo.exe"
End If
End Sub

Private Sub List5_Click()
If List5.Text = "hackers handbook [ebook]" Then
Text11.Text = "Hackers Underground Handbook - Step by Step Tutorial" & vbCrLf & " " & vbCrLf & "Learn the secret to Hack! even in most secure system"
End If
If List5.Text = "hackers blackbook [ebook]" Then
Text11.Text = "Hackers Blackhat Book - from deepest darkness" & vbCrLf & " " & vbCrLf & "This unique report gives a perfect overview, what hackers are able to, how they act and how to protect a system. It's easy to read (not heavy to read like all other books around hacking) and fascinating. More than 200,000 peoples in all over the world (USA and Germany) already read this report. Now they act more carefully in the internet. PLUS: The book contains a login and password to an online readers area with tools, articles and a special hacking and security search engine"
End If
If List5.Text = "buffer overflows [ebook]" Then
Text11.Text = "Buffer Overflow Attacks: Detect, Exploit, Prevent" & vbCrLf & " " & vbCrLf & "Buffer overflows make up one of the largest collections of vulnerabilities in existence; And a large percentage of possible remote exploits are of the overflow variety. Almost all of the most devastating computer attacks to hit the Internet in recent years including SQL Slammer, Blaster, and I Love You attacks. If executed properly, an overflow vulnerability will allow an attacker to run arbitrary code on the victim's machine with the equivalent rights of whichever process was overflowed. This is often used to provide a remote shell onto the victim machine, which can be used for further exploitation."
End If
If List5.Text = "using multicore to suport security-related applications [ebook]" Then
Text11.Text = "Using Multi-core to Support Security-related Applications" & vbCrLf & " " & vbCrLf & "Part 1: Introduction to multi-core.   Part 2: Background of multiprocessing   Part 3: Security-related applications   Part 4: Using multi-core to support  security-related applications"
End If
If List5.Text = "stealing network – how to own the box (hacking novel) [ebook]" Then
Text11.Text = "Stealing the Network: How to Own the Box" & vbCrLf & " " & vbCrLf & "Stealing the Network: How to Own the Box is NOT intended to be a install, configure, update, troubleshoot, and defend book. It is also NOT another one of the countless Hacker books out there. So, what IS it? It is an edgy, provocative, attack-oriented series of chapters written in a first hand, conversational style. World-renowned network security personalities present a series of 25 to 30 page chapters written from the point of an attacker who is gaining access to a particular system. This book portrays the street fighting tactics used to attack networks and systems. "
Else
Text11.Text = List5.Text & vbCrLf & " " & vbCrLf & "[LAUNCH] to see ebook.                 [FIND] to search specific ebook.       [DOWNLOAD] to download a content."
End If
End Sub
Private Sub PushButton1_Click()
If ((SSTab1.Caption = "Attacker") And (List1.Text = "CasperLogger" Or List1.Text = "casperlogger")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/cowpatty-genpmk-4.0-win32.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Genpmk" Or List1.Text = "genpmk")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/cowpatty-genpmk-4.0-win32.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Stuxnet" Or List1.Text = "stuxnet")) Then
ShellExecute 0, vbNullString, "http://www.megapanzer.com/wp-content/uploads/Stuxnet.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "VertexNet" Or List1.Text = "vertexnet")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/VertexNet-sillhouette.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Crypter Pack FUD" Or List1.Text = "crypter pack fud")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/crypters-pack.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "TW-Grabber" Or List1.Text = "tw-grabber")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/formgrabber/TW-Grabber.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Vulcan" Or List1.Text = "vulcan")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/keylogger/vulcan.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Agobot" Or List1.Text = "agobot")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/agobot3.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Citadel" Or List1.Text = "citadel")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/citadel-sillhouette.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Cythosia" Or List1.Text = "cythosia")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/cythosia-sillhouette.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Crypter.py" Or List1.Text = "crypter.py")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/exploit/crypter.py.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Nmap" Or List1.Text = "nmap")) Then
ShellExecute 0, vbNullString, "http://nmap.org/download.html", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "LBD" Or List1.Text = "lbd")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/lbd-ge.mine.nu-master.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Evilpdf exploit" Or List1.Text = "evilpdf exploit")) Then
ShellExecute 0, vbNullString, "http://www.rapid7.com/products/metasploit/download.jsp", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dmitry" Or List1.Text = "dmitry")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/DMitry-1.3a.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Metasploit" Or List1.Text = "metasploit")) Then
ShellExecute 0, vbNullString, "http://www.rapid7.com/products/metasploit/download.jsp", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Msfpayload" Or List1.Text = "msfpayload")) Then
ShellExecute 0, vbNullString, "http://www.rapid7.com/products/metasploit/download.jsp", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "The Harvester" Or List1.Text = "the harvester")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/theHarvester-2.1_BH2011_Arsenal.tar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Trojan Creator" Or List1.Text = "trojan creator")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/astgen11.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Keymail" Or List1.Text = "keymail")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/keylogger/keymail.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dirbuster" Or List1.Text = "dirbuster")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/DirBuster-0.12-Setup.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Aircrack" Or List1.Text = "aircrack")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/aircrack-airmon-ng-1.2-win32.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Burpsuite" Or List1.Text = "burpsuite")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/burpsuite_free_v1.5.jar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dnsmap" Or List1.Text = "dnsmap")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/dnsmap-0.30.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Driftnet" Or List1.Text = "driftnet")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/driftnet-0.1.6.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Ettercap" Or List1.Text = "ettercap")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/ettercap-A%20suite%20for%20man%20in%20the%20middle%20attacks.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Rapzo" Or List1.Text = "rapzo")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/keylogger/rapzo.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Autoinfector" Or List1.Text = "autoinfector")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/setoolkit-3.5.1.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Albertino" Or List1.Text = "albertino")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/keylogger/albertino.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Gerix" Or List1.Text = "gerix")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/gerix-wifi-cracker-master.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Reaver" Or List1.Text = "reaver")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/reaver-1.4.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Wash" Or List1.Text = "wash")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/reaver-1.4.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Wifite" Or List1.Text = "wifite")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/wifite-2.0r85.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Pyrit" Or List1.Text = "pyrit")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/pyrit-0.4.0-WPAWPA2-PSK-key-cracker.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Joomscan" Or List1.Text = "joomscan")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/joomscan-OWASP%20Joomla%20Vulnerability%20Scanner.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Slowloris" Or List1.Text = "slowloris")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/Slowloris_Gui-release1.0_SRC-httpDDOS.7z", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Bifrost" Or List1.Text = "bifrost")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/Bifrost%201.2.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Bionet" Or List1.Text = "bionet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/bionet%20rootkit%20by%20fabio%20grunge.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Turkojan" Or List1.Text = "turkojan")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/turkojan%204.0.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Prorat" Or List1.Text = "prorat")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/TROJAN%20PRORAT.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Spynet" Or List1.Text = "spynet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/Spy_Net.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Darkmoon" Or List1.Text = "darkmoon")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/Dark%20Moon.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "SET" Or List1.Text = "set")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/setoolkit-3.5.1.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Jigsaw" Or List1.Text = "jigsaw")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/jigsaw-master.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Edjpgcom" Or List1.Text = "edjpgcom")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/edjpgcom-binder.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "DVWA" Or List1.Text = "dvwa")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Web%20Hacking/DVWA-1.0.8.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Havij Pro" Or List1.Text = "havij pro")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Web%20Hacking/Havij-PRO.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "SQLMAP" Or List1.Text = "sqlmap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Web%20Hacking/sqlmapproject-sqlmap-0.9-3248-gdbb0d7f.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Wp-Scan" Or List1.Text = "wp-scan")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Web%20Hacking/wpscanteam-wpscan-WordPress%20Security%20Scanner.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dnsdict" Or List1.Text = "dnsdict")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/dns-analyze-tool.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dnsrecon" Or List1.Text = "dnsrecon")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/dnsrecon-master.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dnsmap" Or List1.Text = "dnsmap")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/dnsmap-0.30.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "CasperSpy" Or List1.Text = "casperspy")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/casperspy-stable-version-ver2.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Spy Eye" Or List1.Text = "spy eye")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/spyeye-botpack.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Spitmo" Or List1.Text = "spitmo")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/spyeye_android.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Zeus Botnet" Or List1.Text = "zeus botnet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/zeuss_botpack.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Janidos" Or List1.Text = "janidos")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/janidos.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "MP-Formgrabber" Or List1.Text = "Webcrab")) Then
intResponse = MsgBox("Only Available on Full Version! Are you sure you want to Get Full Version?", _
vbYesNo + vbQuestion, _
"Get Full Version")
If intResponse = vbYes Then
ShellExecute 0, vbNullString, "http://form.jotform.me/form/31453837808461", vbNullString, vbNullString, vbNormalFocus
End If

If ((SSTab1.Caption = "Exploit") And (execution2 = "mdd")) Then
MsgBox "System Information Is Unavailable At This Time", vbOKOnly

ElseIf execution2 = "win32dd" Then
    Shell ("cmd.exe /k cd " & App.path & "\Programs\RAM\win32dd && " & App.path & "\Programs\RAM\win32dd\win32dd"), vbNormalFocus

ElseIf execution2 = "winen" Then
    Shell ("cmd.exe /k cd " & App.path & "\Programs\RAM\winen && " & App.path & "\Programs\RAM\winen\winen -h"), vbNormalFocus

ElseIf execution2 = "Fport" Then
    Set WshShell = CreateObject("WScript.Shell")
    Set oExec = WshShell.Exec(App.path + "\Programs\fport.exe")
    Text2.Text = oExec.StdOut.ReadAll
    Set oExec = Nothing
    Set WshShell = Nothing

ElseIf execution2 = "TCPView" Then
    Shell (App.path & "\Programs\wft\tools\sysinternals\TCPView.exe"), vbNormalFocus
    
ElseIf execution2 = "ALS" Then
    Shell (App.path & "\Programs\ALS\ALS.exe"), vbNormalFocus
End If

If SSTab1.Caption = "Keylogger" Then
Shell (App.path & "\Programs\wft\tools\sysinternals\" & execution3), vbNormalFocus
End If

If SSTab1.Caption = "Network" Or execution4 = "autorunsc" Or execution4 = "handle" Or execution4 = "Listdlls" Or execution4 = "logonsessions" Or execution4 = "psgetsid" Or execution4 = "Psinfo" Or execution4 = "pslist" Or execution4 = "psloggedon" Or execution4 = "psloglist" Or execution4 = "psservice" Then
    Set WshShell = CreateObject("WScript.Shell")
    Set oExec = WshShell.Exec(App.path & "\Programs\wft\tools\sysinternals\" & execution4 & ".exe")
    Text7.Text = oExec.StdOut.ReadAll
    Set oExec = Nothing
    Set WshShell = Nothing
    ElseIf execution4 = "efsdump" Or execution4 = "ntfsinfo" Or execution4 = "psexec" Or execution4 = "psfile" Or execution4 = "pskill" Or execution4 = "pspasswd" Or execution4 = "psshutdown" Or execution4 = "pssuspend" Or execution4 = "streams" Or execution4 = "strings" Then
    Shell ("cmd.exe /k cd " & App.path & "\Programs\wft\tools\sysinternals && " & App.path & "\Programs\wft\tools\sysinternals\" & execution4 & ".exe"), vbNormalFocus
    End If
    End If
    End Sub

Private Sub PushButton2_Click()
Text10.Visible = True
m.Visible = True
PushButton9.Visible = True
End Sub

Private Sub PushButton3_Click()
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Gerix" Or List1.Text = "gerix")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/hack-wifi-crack-wep-with-gerix/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Stuxnet" Or List1.Text = "stuxnet")) Then
ShellExecute 0, vbNullString, " http://casperspy.com/stuxnet-timeline-most-insidious-cyber-weapon/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Albertino" Or List1.Text = "albertino")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tutorial-how-to-use-albertino-keylogger", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Wp-Scan" Or List1.Text = "wp-scan")) Then
ShellExecute 0, vbNullString, "http://antmanaras.wordpress.com/2012/12/30/tutorial-scan-a-wordpress-website-with-wpscan-part-1-basic-scan/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "TW-Grabber" Or List1.Text = "tw-grabber")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/category/form-grabber/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Crypter.py" Or List1.Text = "crypter.py")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/sneaky-way-how-to-hack-corporate-email-account/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Msfpayload" Or List1.Text = "msfpayload")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/inject-backdoor-using-metasploit/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Agobot" Or List1.Text = "agobot")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/pentest-with-agobot-how-botnet-work-robots-wars/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Slowloris" Or List1.Text = "slowloris")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/http-ddos-with-slowloris/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Joomscan" Or List1.Text = "joomscan")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/easiest-way-to-hack-joomla-using-joomscan/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Autoinfector" Or List1.Text = "autoinfector")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/infectious-media-generator-infecting-target-automatically/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Aircrack" Or List1.Text = "aircrack")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wep-key-in-windows-using-commviewaircrack-gui/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Burpsuite" Or List1.Text = "burpsuite")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/bypassing-upload-filter-with-burpsuite/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "CasperSpy" Or List1.Text = "casperspy")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/casperspy-overview/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Citadel" Or List1.Text = "citadel")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tutorial-citadel-botnet-user-friendly-botnet/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Cythosia" Or List1.Text = "cythosia")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/destroying-webserver-easily-using-cythosia-ddos-bot/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dirbuster" Or List1.Text = "dirbuster")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/webapps-fuzzer-webdir-enumeration-attack-dirbuster/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dnsdict" Or List1.Text = "dnsdict")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/dns-analysis-domain-enumeration-with-dnsdict/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dmitry" Or List1.Text = "dmitry")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/gathering-information-with-dmitry-route-analysis/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dnsmap" Or List1.Text = "dnsmap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/bruteforce-subdomain-with-dnsmap/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Dnsrecon" Or List1.Text = "dnsrecon")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/reconnaissance-dns-with-dnsrecon/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Driftnet" Or List1.Text = "driftnet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/network-sniffer-sniffing-media-files-image-video-with-driftnet/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "DVWA" Or List1.Text = "dvwa")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/hacking-website-complete-guide-of-xss-injection-part-5/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Edjpgcom" Or List1.Text = "edjpgcom")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/injecting-exploit-in-jpg/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Ettercap" Or List1.Text = "ettercap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/network-sniffing-how-to-dns-spoofing-with-ettercap/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Evilpdf exploit" Or List1.Text = "evilpdf exploit")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/inject-exploit-into-pdf-with-metasploit-evilpdf/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Genpmk" Or List1.Text = "genpmk")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/generating-hash-to-crack-wpa-with-genpmk/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Havij Pro" Or List1.Text = "havij pro")) Then
ShellExecute 0, vbNullString, "http://www.itsecteam.com/files/havij/havij_help-english.pdf", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Janidos" Or List1.Text = "janidos")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/janidos-powerfull-ddos-with-zombie-army/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Jigsaw" Or List1.Text = "jigsaw")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/enumeration-employee-info-with-jigsaw/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "LBD" Or List1.Text = "lbd")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/detect-load-balancing-with-lbd/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Metasploit" Or List1.Text = "metasploit")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/category/metasploit-2/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "MP-Formgrabber" Or List1.Text = "mp-formgrabber")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/mp-formgrabber-stable-on-all-major-browser/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Nmap" Or List1.Text = "nmap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/import-nmap-scan-xml-into-metasploit/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Pyrit" Or List1.Text = "pyrit")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/fastest-way-to-crack-wpa-wpa2-pyrit/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Reaver" Or List1.Text = "reaver")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wifi-protected-setup-part2-brute-force-attack-against-wifi-protected-setup-using-reaver/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "SET" Or List1.Text = "set")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/tag/set/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Spitmo" Or List1.Text = "spitmo")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/spyeye-for-android-spitmo/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Spy Eye" Or List1.Text = "spy eye")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/category/spyeye/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "SQLMAP" Or List1.Text = "sqlmap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/complete-guide-how-to-hack-website/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "The Harvester" Or List1.Text = "the harvester")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/gathering-sensitive-information-from-corporation-with-theharvester/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "VertexNet" Or List1.Text = "vertexnet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/mp-formgrabber-stable-on-all-major-browser/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Vulcan" Or List1.Text = "vulcan")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/hack-any-account-using-vulcan-remote-keylogger/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Wash" Or List1.Text = "wash")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wifi-protected-setup-part-1/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Webcrab" Or List1.Text = "webcrab")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/webcrab-underground-formgrabber/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Wifite" Or List1.Text = "wifite")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-multiple-wifi-automatically-using-wifite-v2/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab1.Caption = "Attacker") And (List1.Text = "Zeus Botnet" Or List1.Text = "zeus botnet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/zeus-complete-tutorial/", vbNullString, vbNullString, vbNormalFocus
End If
End Sub

Private Sub PushButton4_Click()
Intro.Show
End Sub

Private Sub PushButton5_Click()
If ((SSTab2.Caption = "Ebook") And (List5.Text = "analysis of rxbot [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/ANALYSIS%20OF%20RXBOT%20.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "an introduction the world of botnets [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/An%20Introduction%20Into%20the%20World.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "an overview of rfi bot [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/An%20Overview%20Of%20The%20Casper%20RFI%20Bot.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "anatomy of a botnet [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Anatomy%20of%20a%20Botnet.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "andbot towards advanced mobile botnets [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Andbot%20Towards%20Advanced%20Mobile%20Botnets.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "byob build your own botnet [ebook]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/BYOB%20Build%20Your%20Own%20Botnet.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "botnet communication topologies [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Botnet%20Communication%20Topologies.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "botnets detection, measurement, disinfection & defence [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Botnets%20Detection,%20Measurement,%20Disinfection%20&%20Defence.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "fighting against botnet [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Fighting_against_Botnet.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "inside carberp botnet [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Inside%20Carberp%20Botnet.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "modern malware for dummies")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Modern%20Malware%20for%20Dummies.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "on botnets that use dns for command and control [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/On%20Botnets%20that%20use%20DNS%20for%20Command%20and%20Control.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "on the analysis of the zeus botnet crimeware [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/On%20the%20Analysis%20of%20the%20Zeus%20Botnet%20Crimeware.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "riorey_taxonomy_ddos_attacks [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/RioRey_Taxonomy_DDoS_Attacks_2.2_2011.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "the anatomy of clickbot.a [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/The%20Anatomy%20of%20Clickbot.A.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "the zombie roundup-understanding, detecting, and disrupting botnets [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/The%20Zombie%20Roundup-Understanding,%20Detecting,%20and%20Disrupting%20Botnets.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "tracking multiple c&c botnets [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Tracking%20Multiple%20C&C%20Botnets.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "working the revitalising zombie army [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/Working%20the%20revitalising%20ombie%20army.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "russian-underground-101 [paper]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/ebook/russian-underground-101.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "hackers handbook [ebook]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/hacker-handbook.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "hackers blackbook [ebook]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/hackers-blackbook-eng.rar", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "buffer overflows [ebook]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/BufferOverflow.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "using multicore to suport security-related applications [ebook]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/multi-core_tutorial_3.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "stealing network – how to own the box (hacking novel) [ebook]")) Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/hacking-novel-Stealing_the_Network_-_How_to_Own_the_Box.zip", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Ebook") And (List5.Text = "the art of intrussion ~kevin mitnick [ebook]" Or List5.Text = "the art of exploitation 2nd edition ~kevin mitnick [ebook]" Or List5.Text = "ghost in the wired ~kevin mitnick [ebook]" Or List5.Text = "social engineering: the art of human hacking [ebook]" Or List5.Text = "metasploit: the penetration tester guide [ebook]" Or List5.Text = "hacking exposed wireless 2nd edition [ebook]" Or List5.Text = "the art of deception ~kevin mitnick [ebook]")) Then
intResponse = MsgBox("Only Available on Full Version! Are you sure you want to Get Full Version?", _
vbYesNo + vbQuestion, _
"Get Full Version")
If intResponse = vbYes Then
ShellExecute 0, vbNullString, "http://form.jotform.me/form/31453837808461", vbNullString, vbNullString, vbNormalFocus
End If
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Agobot IRC botnet" Or List4.Text = "agobot irc botnet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/pentest-with-agobot-how-botnet-work-robots-wars/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Autoinfection exploit when connect USB" Or List4.Text = "autoinfection exploit when connect usb")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/infectious-media-generator-infecting-target-automatically/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Basic guide of SQL injection using SQLMAP" Or List4.Text = "basic guide of sql injection using sqlmap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/complete-guide-how-to-hack-website/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Bypass upload filter with burpsuite" Or List4.Text = "bypass upload filter with burpsuite")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/bypassing-upload-filter-with-burpsuite/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Bruteforce domain with dnsmap" Or List4.Text = "bruteforce domain with dnsmap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/bruteforce-subdomain-with-dnsmap/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Bruteforce against wifi using reaver" Or List4.Text = "bruteforce against wifi using reaver")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wifi-protected-setup-part2-brute-force-attack-against-wifi-protected-setup-using-reaver/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Complete Guide of XSS injection" Or List4.Text = "complete guide of xss injection")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/hacking-website-complete-guide-of-xss-injection-part-5/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Cracking WIFI key using aircrack" Or List4.Text = "cracking wifi key using aircrack")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wep-key-in-windows-using-commviewaircrack-gui/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Cracking WIFI password using aircrack" Or List4.Text = "cracking wifi password using aircrack")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wep-key-in-windows-using-commviewaircrack-gui/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Cracking WIFI find target WPS using wash" Or List4.Text = "cracking wifi find target wps using wash")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wifi-protected-setup-part-1/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Defacing web using SQLMAP" Or List4.Text = "defacing web using sqlmap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/complete-guide-how-to-hack-website/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Domain enumeration attack using dnsdict")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/dns-analysis-domain-enumeration-with-dnsdict/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Detect load balancing with LBD")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/detect-load-balancing-with-lbd/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "DNS spoofing with ettercap")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/network-sniffing-how-to-dns-spoofing-with-ettercap/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Easiest way to hack joomla using joomscan")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/easiest-way-to-hack-joomla-using-joomscan/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Embedding backdoor into PDF")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/inject-exploit-into-pdf-with-metasploit-evilpdf/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Enumerate employee information with jigsaw")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/enumeration-employee-info-with-jigsaw/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Fastest way to crack wifi")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/fastest-way-to-crack-wpa-wpa2-pyrit/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Gathering sensitive data from corporation")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/gathering-sensitive-information-from-corporation-with-theharvester/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Gathering information with dmitry")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/gathering-information-with-dmitry-route-analysis/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Genpmk hash generator")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/generating-hash-to-crack-wpa-with-genpmk/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Hack any account with vulcan keylogger")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/hack-any-account-using-vulcan-remote-keylogger/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Hacking like spy agent")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/infectious-media-generator-infecting-target-automatically/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Harvesting company email account")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/gathering-sensitive-information-from-corporation-with-theharvester/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "How botnet works")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/pentest-with-agobot-how-botnet-work-robots-wars/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "HTTP DDOS using slowloris")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/http-ddos-with-slowloris/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Infect system automatically")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/infectious-media-generator-infecting-target-automatically/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Infectious Media Generator")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/infectious-media-generator-infecting-target-automatically/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Injectiong exploit in JPG")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/injecting-exploit-in-jpg/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Import nmap scan into metasploit")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/import-nmap-scan-xml-into-metasploit/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Metasploit: steal login credentials")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/metasploit-can-steal-login-credentials/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Metasploit: cracking wifi protected setup")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wifi-protected-setup-part-1/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Metasploit: combination nmap & metasploit")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/cracking-wifi-protected-setup-part-1/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Metasploit: inject backdoor into executable")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/inject-backdoor-using-metasploit/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Metasploit: bypass ftp authentication")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/metasploit-how-to-bypass-ftp-authentication/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Metasploit: hack corporate email account")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/sneaky-way-how-to-hack-corporate-email-account/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Metasploit: steal data from android device")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/stealing-data-from-android-devices-with-metasploit/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Most strongest DDOS tools")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/janidos-powerfull-ddos-with-zombie-army/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Pentest with agobot")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/pentest-with-agobot-how-botnet-work-robots-wars/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Sniffing DNS with dnsrecon")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/zeus-complete-tutorial/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "sniffing image\video using driftnet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/network-sniffer-sniffing-media-files-image-video-with-driftnet/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Spying on browser undetectable")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/mp-formgrabber-stable-on-all-major-browser/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Spy someone phone secretly")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/spy-android-phone-using-spitmo/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Spitmo mobile botnet")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/spyeye-for-android-spitmo/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "SQL injection automatically with havij")) Then
ShellExecute 0, vbNullString, "www.itsecteam.com/files/havij/havij_help-english.pdf", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Takeover database using SQLMAP")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/complete-guide-how-to-hack-website/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Upload shell to webserver using burpsuite")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/bypassing-upload-filter-with-burpsuite/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Webcrab formgrabber")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/webcrab-underground-formgrabber/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Web directory enumeration Attack")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/webapps-fuzzer-webdir-enumeration-attack-dirbuster/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Web hacking XSS injection ")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/hacking-website-complete-guide-of-xss-injection-part-5/", vbNullString, vbNullString, vbNormalFocus
End If
If ((SSTab2.Caption = "Tutorial") And (List4.Text = "Zeus complete tutorial")) Then
ShellExecute 0, vbNullString, "http://casperspy.com/zeus-complete-tutorial/", vbNullString, vbNullString, vbNormalFocus
End If
End Sub

Private Sub PushButton7_Click()
Text9.Visible = True
PushButton8.Visible = True
n.Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "Botnet" Or SSTab1.Caption = "Keylogger" Or SSTab1.Caption = "Network" Or SSTab1.Caption = "Trojan" Or SSTab1.Caption = "Exploit" Or SSTab1.Caption = "Social Engginering" Or SSTab1.Caption = "Web Hacking" Or SSTab1.Caption = "Cracking" Then
SSTab1.Tab = 0
intResponse = MsgBox("Only Available on Full Version! Are you sure you want to Get Full Version?", _
vbYesNo + vbQuestion, _
"Get Full Version")
If intResponse = vbYes Then
ShellExecute 0, vbNullString, "http://form.jotform.me/form/31453837808461", vbNullString, vbNullString, vbNormalFocus
End If
End If
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
If SSTab2.Caption = "Ebook" Then
List5.Visible = True
Text11.Visible = True
Else
List5.Visible = False
Text11.Visible = False
End If
If SSTab2.Caption = "Tutorial" Then
List4.Visible = True
Text8.Visible = True
Else
List4.Visible = False
Text8.Visible = False
End If
End Sub

Private Sub Text10_Change()
List1.Clear
m = 0
If Len(Text10) <> 0 Then
se = Text10.Text
m = 0
Open App.path & "\programs\attacker.txt" For Input As #1
While Not EOF(1)
m = m + 1
Input #1, search
For counter = 1 To Len(search)
If Mid(se, 1, Len(Text10)) = Mid(search, counter, Len(Text10)) Then
List1.AddItem (search)
End If
Next
Wend
Close #1
End If
m.Text = List1.ListCount & "/" & n
End Sub

Private Sub Text10_Click()
Text10.Text = ""
End Sub

Private Sub Text9_Change()
List4.Clear
List5.Clear
n = 0
If Len(Text9) <> 0 Then
se = Text9.Text
n = 0
Open App.path & "\programs\guidemenu.txt" For Input As #1
While Not EOF(1)
n = n + 1
Input #1, search
For counter = 1 To Len(search)
If Mid(se, 1, Len(Text9)) = Mid(search, counter, Len(Text9)) Then
List4.AddItem (search)
List5.AddItem (search)
End If
Next
Wend
Close #1
End If
n.Text = List4.ListCount & "/" & n
End Sub

Private Sub Text9_Click()
Text9.Text = ""
End Sub
