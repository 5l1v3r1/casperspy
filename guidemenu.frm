VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.2#0"; "CODEJO~4.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.2#0"; "CODEJO~2.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.2#0"; "CO85CC~1.OCX"
Begin VB.Form guide 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   Caption         =   "CasperLauncher | Full Version [Shadow Striker]"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "guidemenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton4 
      Height          =   1095
      Left            =   8760
      TabIndex        =   12
      Top             =   7080
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
      Appearance      =   6
      DrawFocusRect   =   0   'False
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "guidemenu.frx":1CCA
   End
   Begin XtremeSuiteControls.PushButton PushButton5 
      Height          =   1095
      Left            =   4080
      TabIndex        =   13
      ToolTipText     =   "Launch selected content "
      Top             =   7080
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "guidemenu.frx":4134
   End
   Begin XtremeSuiteControls.PushButton PushButton7 
      Height          =   1095
      Left            =   5640
      TabIndex        =   14
      ToolTipText     =   "Find specific Guide content"
      Top             =   7080
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "guidemenu.frx":659E
   End
   Begin XtremeSuiteControls.PushButton PushButton6 
      Height          =   1095
      Left            =   7200
      TabIndex        =   15
      ToolTipText     =   "Download content on guide menu."
      Top             =   7080
      Width           =   1590
      _Version        =   983042
      _ExtentX        =   2813
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Download"
      ForeColor       =   4210752
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "guidemenu.frx":8A08
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   10575
      _Version        =   983042
      _ExtentX        =   18653
      _ExtentY        =   12303
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   10
      Color           =   32
      PaintManager.Layout=   5
      PaintManager.Position=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   6
      SelectedItem    =   3
      Item(0).Caption =   "System Attack"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "Text6"
      Item(0).Control(1)=   "List8"
      Item(0).Control(2)=   "Text7"
      Item(0).Control(3)=   "List9"
      Item(0).Control(4)=   "Text12"
      Item(0).Control(5)=   "List10"
      Item(0).Control(6)=   "Text13"
      Item(0).Control(7)=   "List11"
      Item(0).Control(8)=   "TabControl2"
      Item(1).Caption =   "Network Attack"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControl3"
      Item(2).Caption =   "Web Attack"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControl4"
      Item(3).Caption =   "Tutorial"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "Text25"
      Item(3).Control(1)=   "List23"
      Item(4).Caption =   "Video"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "Text26"
      Item(4).Control(1)=   "List24"
      Item(4).Control(2)=   "PushButton9"
      Item(5).Caption =   "Ebook"
      Item(5).ControlCount=   3
      Item(5).Control(0)=   "Text27"
      Item(5).Control(1)=   "List25"
      Item(5).Control(2)=   "PushButton8"
      Begin XtremeSuiteControls.PushButton PushButton9 
         Height          =   615
         Left            =   -67600
         TabIndex        =   62
         ToolTipText     =   "Auto Update all Contents & Feature of CasperLauncher"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
         _Version        =   983042
         _ExtentX        =   1720
         _ExtentY        =   1085
         _StockProps     =   79
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
         TextAlignment   =   10
         Appearance      =   2
         ImageAlignment  =   6
         TextImageRelation=   0
         IconWidth       =   32
         Icon            =   "guidemenu.frx":AE72
      End
      Begin XtremeSuiteControls.PushButton PushButton8 
         Height          =   615
         Left            =   -62320
         TabIndex        =   61
         ToolTipText     =   "Auto Update all Contents & Feature of CasperLauncher"
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
         _Version        =   983042
         _ExtentX        =   1720
         _ExtentY        =   1085
         _StockProps     =   79
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
         TextAlignment   =   10
         Appearance      =   2
         ImageAlignment  =   6
         TextImageRelation=   0
         IconWidth       =   32
         Icon            =   "guidemenu.frx":BEDC
      End
      Begin VB.ListBox List25 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         ItemData        =   "guidemenu.frx":CF46
         Left            =   -67840
         List            =   "guidemenu.frx":CFAD
         TabIndex        =   60
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         Left            =   -63880
         MultiLine       =   -1  'True
         TabIndex        =   59
         Text            =   "guidemenu.frx":D2FF
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ListBox List23 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5475
         ItemData        =   "guidemenu.frx":D36C
         Left            =   2160
         List            =   "guidemenu.frx":D424
         Sorted          =   -1  'True
         TabIndex        =   58
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         Left            =   6600
         MultiLine       =   -1  'True
         TabIndex        =   57
         Text            =   "guidemenu.frx":DA6E
         Top             =   240
         Width           =   3855
      End
      Begin VB.ListBox List24 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         ItemData        =   "guidemenu.frx":DA93
         Left            =   -67840
         List            =   "guidemenu.frx":DAA0
         TabIndex        =   56
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         Left            =   -63880
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
      End
      Begin XtremeSuiteControls.TabControl TabControl4 
         Height          =   6975
         Left            =   -67960
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   8535
         _Version        =   983042
         _ExtentX        =   15055
         _ExtentY        =   12303
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         ItemCount       =   5
         SelectedItem    =   4
         Item(0).Caption =   "Introduction"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "List18"
         Item(0).Control(1)=   "Text20"
         Item(1).Caption =   "Information Gathering"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "List19"
         Item(1).Control(1)=   "Text21"
         Item(2).Caption =   "Web Exploitation"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "List20"
         Item(2).Control(1)=   "Text22"
         Item(3).Caption =   "Cross Site Scripting (XSS)"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "List21"
         Item(3).Control(1)=   "Text23"
         Item(4).Caption =   "SQL Injection"
         Item(4).ControlCount=   2
         Item(4).Control(0)=   "List22"
         Item(4).Control(1)=   "Text24"
         Begin VB.TextBox Text24 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4905
            Left            =   4080
            MultiLine       =   -1  'True
            TabIndex        =   54
            Text            =   "guidemenu.frx":DAF9
            Top             =   720
            Width           =   4215
         End
         Begin VB.ListBox List22 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":DD67
            Left            =   120
            List            =   "guidemenu.frx":DDA4
            TabIndex        =   53
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox Text20 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   52
            Text            =   "guidemenu.frx":DF71
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List18 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4905
            ItemData        =   "guidemenu.frx":E0E7
            Left            =   -69880
            List            =   "guidemenu.frx":E103
            TabIndex        =   51
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text23 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   50
            Text            =   "guidemenu.frx":E17F
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List21 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":E3C4
            Left            =   -69880
            List            =   "guidemenu.frx":E3F5
            TabIndex        =   49
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text22 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   48
            Text            =   "guidemenu.frx":E535
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List20 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":E75E
            Left            =   -69880
            List            =   "guidemenu.frx":E771
            TabIndex        =   47
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text21 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4860
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   46
            Text            =   "guidemenu.frx":E7CF
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List19 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4785
            ItemData        =   "guidemenu.frx":E988
            Left            =   -69880
            List            =   "guidemenu.frx":E9D1
            TabIndex        =   45
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl3 
         Height          =   7335
         Left            =   -67720
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   8295
         _Version        =   983042
         _ExtentX        =   14631
         _ExtentY        =   12938
         _StockProps     =   68
         Appearance      =   10
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         ItemCount       =   6
         SelectedItem    =   5
         Item(0).Caption =   "Information Gathering"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "Text11"
         Item(0).Control(1)=   "List7"
         Item(1).Caption =   "Scanning"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "Text14"
         Item(1).Control(1)=   "List12"
         Item(2).Caption =   "Enumeration"
         Item(2).ControlCount=   2
         Item(2).Control(0)=   "Text15"
         Item(2).Control(1)=   "List13"
         Item(3).Caption =   "Sniffing and MITM"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "Text16"
         Item(3).Control(1)=   "List14"
         Item(4).Caption =   "Vulnerability Exploitation"
         Item(4).ControlCount=   2
         Item(4).Control(0)=   "Text17"
         Item(4).Control(1)=   "List15"
         Item(5).Caption =   "Social Engginering"
         Item(5).ControlCount=   2
         Item(5).Control(0)=   "Text18"
         Item(5).Control(1)=   "List16"
         Begin VB.ListBox List16 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":EC5A
            Left            =   120
            List            =   "guidemenu.frx":EC88
            TabIndex        =   42
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox Text18 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   3960
            MultiLine       =   -1  'True
            TabIndex        =   41
            Text            =   "guidemenu.frx":ED8A
            Top             =   720
            Width           =   4215
         End
         Begin VB.ListBox List15 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":EF19
            Left            =   -69880
            List            =   "guidemenu.frx":EF80
            TabIndex        =   40
            Top             =   720
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox Text17 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -66040
            MultiLine       =   -1  'True
            TabIndex        =   39
            Text            =   "guidemenu.frx":F27F
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List14 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":F456
            Left            =   -69880
            List            =   "guidemenu.frx":F49C
            TabIndex        =   38
            Top             =   720
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -66040
            MultiLine       =   -1  'True
            TabIndex        =   37
            Text            =   "guidemenu.frx":F61F
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List13 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":F85F
            Left            =   -69880
            List            =   "guidemenu.frx":F89F
            TabIndex        =   36
            Top             =   720
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox Text15 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -66040
            MultiLine       =   -1  'True
            TabIndex        =   35
            Text            =   "guidemenu.frx":F9CD
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List12 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":FCBE
            Left            =   -69880
            List            =   "guidemenu.frx":FD0A
            TabIndex        =   34
            Top             =   720
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox Text14 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -66040
            MultiLine       =   -1  'True
            TabIndex        =   33
            Text            =   "guidemenu.frx":FE69
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List7 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":100F3
            Left            =   -69880
            List            =   "guidemenu.frx":1014B
            TabIndex        =   32
            Top             =   720
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -66040
            MultiLine       =   -1  'True
            TabIndex        =   31
            Text            =   "guidemenu.frx":10344
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -1.36640e5
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Text            =   "guidemenu.frx":105AE
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
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
         ItemData        =   "guidemenu.frx":10630
         Left            =   -1.39880e5
         List            =   "guidemenu.frx":10649
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -1.36640e5
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "guidemenu.frx":10690
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
      End
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
         ItemData        =   "guidemenu.frx":10712
         Left            =   -1.39880e5
         List            =   "guidemenu.frx":107C4
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -1.36640e5
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "guidemenu.frx":10A04
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.ListBox List10 
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
         ItemData        =   "guidemenu.frx":10A86
         Left            =   -1.39880e5
         List            =   "guidemenu.frx":10B38
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -1.36640e5
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "guidemenu.frx":10D78
         Top             =   1200
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.ListBox List11 
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
         ItemData        =   "guidemenu.frx":10DFA
         Left            =   -1.39880e5
         List            =   "guidemenu.frx":10E13
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
      End
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   7335
         Left            =   -67960
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   8535
         _Version        =   983042
         _ExtentX        =   15055
         _ExtentY        =   12938
         _StockProps     =   68
         AllowReorder    =   -1  'True
         Appearance      =   10
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         ItemCount       =   7
         SelectedItem    =   6
         Item(0).Caption =   "Introduction"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "List1"
         Item(0).Control(1)=   "Text1"
         Item(1).Caption =   "Anonymity"
         Item(1).ControlCount=   2
         Item(1).Control(0)=   "Text19"
         Item(1).Control(1)=   "List17"
         Item(2).Caption =   "Buffer Overflow"
         Item(2).ControlCount=   3
         Item(2).Control(0)=   "List2"
         Item(2).Control(1)=   "Text2"
         Item(2).Control(2)=   "Text8"
         Item(3).Caption =   "Cracking"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "Text3"
         Item(3).Control(1)=   "List3"
         Item(4).Caption =   "Shell Coding"
         Item(4).ControlCount=   2
         Item(4).Control(0)=   "Text4"
         Item(4).Control(1)=   "List4"
         Item(5).Caption =   "Malware"
         Item(5).ControlCount=   2
         Item(5).Control(0)=   "Text5"
         Item(5).Control(1)=   "List5"
         Item(6).Caption =   "Rootkit Coding"
         Item(6).ControlCount=   3
         Item(6).Control(0)=   "List6"
         Item(6).Control(1)=   "Text9"
         Item(6).Control(2)=   "Text10"
         Begin VB.ListBox List17 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":10E5A
            Left            =   -69880
            List            =   "guidemenu.frx":10E73
            TabIndex        =   44
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text19 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   43
            Text            =   "guidemenu.frx":10EF3
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   4080
            MultiLine       =   -1  'True
            TabIndex        =   30
            Text            =   "guidemenu.frx":1105B
            Top             =   720
            Width           =   4215
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   29
            Text            =   "guidemenu.frx":11218
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4875
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   28
            Text            =   "guidemenu.frx":1142D
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4860
            Left            =   -66640
            MultiLine       =   -1  'True
            TabIndex        =   27
            Text            =   "guidemenu.frx":11595
            Top             =   720
            Width           =   4095
         End
         Begin VB.ListBox List1 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":11617
            Left            =   -69880
            List            =   "guidemenu.frx":1163F
            TabIndex        =   26
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.ListBox List2 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":116FD
            Left            =   -69880
            List            =   "guidemenu.frx":11734
            Sorted          =   -1  'True
            TabIndex        =   25
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4965
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   24
            Text            =   "guidemenu.frx":118A0
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List3 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":11A50
            Left            =   -69880
            List            =   "guidemenu.frx":11A8D
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   22
            Text            =   "guidemenu.frx":11C16
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List4 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":11DDB
            Left            =   -69880
            List            =   "guidemenu.frx":11E15
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4890
            Left            =   -65920
            MultiLine       =   -1  'True
            TabIndex        =   20
            Text            =   "guidemenu.frx":11FF2
            Top             =   720
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.ListBox List5 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":12130
            Left            =   -69880
            List            =   "guidemenu.frx":12188
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4860
            Left            =   -1.36760e5
            MultiLine       =   -1  'True
            TabIndex        =   18
            Text            =   "guidemenu.frx":12371
            Top             =   720
            Visible         =   0   'False
            Width           =   4095
         End
         Begin VB.ListBox List6 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4920
            ItemData        =   "guidemenu.frx":123F3
            Left            =   120
            List            =   "guidemenu.frx":12406
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   720
            Width           =   3855
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   7215
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   10575
      _Version        =   983042
      _ExtentX        =   18653
      _ExtentY        =   12726
      _StockProps     =   79
      Caption         =   " CasperLauncher | Guide Menu"
      ForeColor       =   4210752
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HelveticaNeueLT Std Blk Ext"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
      BorderStyle     =   1
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   1320
      Top             =   120
      _Version        =   983042
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "guidemenu.frx":12487
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4560
      Top             =   3840
      _Version        =   983042
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   240
      Top             =   120
      _Version        =   983042
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   6
      DesignerControls=   -1  'True
      DesignerControlsData=   "guidemenu.frx":40B91
   End
End
Attribute VB_Name = "guide"
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




Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next
Select Case Control.Id
    
 'Attacker menubar
        Case ID_ATTACKER_ATTACK1: Shell App.path & "\programs\builder\casperspy.exe", vbNormalFocus
        
        Case ID_ATTACKER_ATTACK2: frmmain2.Show
        frmmain2.wb1.Navigate "http://casperspy.com/dashboard/"
        
        Case ID_ATTACKER_ATTACK3: Dim intResponse As Integer
        intResponse = MsgBox("I have problem when connecting CasperSpy! did you like to help me?", _
vbYesNo + vbQuestion, _
"Help me")

If intResponse = vbYes Then

ShellExecute 0, vbNullString, "https://groups.google.com/forum/#!forum/casperspy", vbNullString, vbNullString, vbNormalFocus

End If
        
        Case ID_ATTACKER_AGOBOT: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 0
                        
        Case ID_ATTACKER_CASPER: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 8
      
        Case ID_ATTACKER_CITADEL: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 9
        
        Case ID_ATTACKER_CYTHOSIA: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 12
        
        Case ID_ATTACKER_VERTEXNET: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 52
        
        Case ID_ATTACKER_ZEUS: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 58
        
        Case ID_ATTACKER_EXPLOIT1: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 10
       
        Case ID_ATTACKER_EXPLOIT2: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 23
        
        Case ID_ATTACKER_EXPLOIT3: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 32
        
        Case ID_ATTACKER_EXPLOIT4:  Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 34
        
        Case ID_ATTACKER_FORMGRABBER1: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 33
        
        Case ID_ATTACKER_FORMGRABBER2: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 50
       
        Case ID_ATTACKER_FORMGRABBER3: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 54
        
        Case ID_ATTACKER_KEYLOGGER1: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 2
        
        Case ID_ATTACKER_KEYLOGGER2: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 30
     
        Case ID_ATTACKER_KEYLOGGER3: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 38
        
        Case ID_ATTACKER_KEYLOGGER4: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 52
        
        Case ID_ATTACKER_NETWORK1: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 1
        
        Case ID_ATTACKER_NETWORK2: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 15
        
        Case ID_ATTACKER_NETWORK3: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 16
        
        Case ID_ATTACKER_NETWORK4:  Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 19
        
        Case ID_ATTACKER_NETWORK5: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 22
        
        Case ID_ATTACKER_NETWORK6: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 25
        
        Case ID_ATTACKER_NETWORK7: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 27
        
        Case ID_ATTACKER_NETWORK8: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 31
        
        Case ID_ATTACKER_NETWORK9: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 35
        
        
        Case ID_ATTACKER_NETWORK10: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 37
      
        Case ID_ATTACKER_NETWORK11: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 39
        
        Case ID_ATTACKER_NETWORK12: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 41
        
        Case ID_ATTACKER_NETWORK13: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 53
        
        Case ID_ATTACKER_NETWORK14: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 55
        
        Case ID_ATTACKER_TROJAN1:  Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 4
        
        Case ID_ATTACKER_TROJAN2: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 5
        
        Case ID_ATTACKER_TROJAN3: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 13
        
        Case ID_ATTACKER_TROJAN4: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 36
        
        Case ID_ATTACKER_TROJAN5: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 44
        
        Case ID_ATTACKER_TROJAN6: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 49
        
        Case ID_ATTACKER_TROJAN7: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 48
        
        Case ID_ATTACKER_WEBHACKING1: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 6
         
        Case ID_ATTACKER_WEBHACKING2: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 14
        
        Case ID_ATTACKER_WEBHACKING3:  Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 20
        
        Case ID_ATTACKER_WEBHACKING4: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 26
      
        Case ID_ATTACKER_WEBHACKING5: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 29
        
        Case ID_ATTACKER_WEBHACKING6: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 45
        
        Case ID_ATTACKER_WEBHACKING7: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 56
    
        Case ID_ATTACKER_SOCIAL1: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 3
        
        Case ID_ATTACKER_SOCIAL2: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 21
        
        Case ID_ATTACKER_SOCIAL3: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 28
        
        Case ID_ATTACKER_SOCIAL4: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 47
        
        Case ID_ATTACKER_SOCIAL5: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
        frmMain.List1.ListIndex = 40
        
        Case ID_ATTACKER_EXIT: End
        
        Case ID_ATTACKER_MORE1: TabControl1.SelectedItem = 1
        Case ID_ATTACKER_MORE2: TabControl1.SelectedItem = 3
        Case ID_ATTACKER_MORE3: TabControl1.SelectedItem = 0
                                List1.ListIndex = 33
        Case ID_ATTACKER_MORE4: TabControl1.SelectedItem = 4
        Case ID_ATTACKER_MORE5: TabControl1.SelectedItem = 5
        Case ID_ATTACKER_MORE6: TabControl1.SelectedItem = 7
        Case ID_ATTACKER_MORE7: TabControl1.SelectedItem = 8
        Case ID_ATTACKER_MORE8: TabControl1.SelectedItem = 6
        
        'Guide Menubar
        Case ID_GUIDE_SYSTEM: Me.Hide
                              guide.Show
                              guide.TabControl1.SelectedItem = 0
                              
        Case ID_GUIDE_NETWORK:  Me.Hide
                                guide.Show
                                guide.TabControl1.SelectedItem = 1
                                
        Case ID_GUIDE_WEB:  Me.Hide
                            guide.Show
                            guide.TabControl1.SelectedItem = 2
                            
        Case ID_GUIDE_TUTORIAL: Me.Hide
                                guide.Show
                                guide.TabControl1.SelectedItem = 3
                                
        Case ID_FILE_VIDEO: Me.Hide
                            guide.Show
                            guide.TabControl1.SelectedItem = 4
                            
        'Help Menubar
        Case ID_HELP_ABOUT: Intro.Show
        Case ID_FILE_UPDATE: MsgBox Control.Caption & " UPDATED", vbOKOnly, "Updated"
        
        'Toolbar
        Case ID_FILE_UPDATE: MsgBox Control.Caption & " clicked", vbOKOnly, "Button Clicked"
        
        Case ID_FILE_NEWS: frmmain2.Show
        frmmain2.wb1.Navigate "http://feeds.feedburner.com/casperspybotnet"
                   
        Case ID_ACTIONS_CHAT: frmmain2.Show
        frmmain2.wb1.Navigate "http://form.jotform.me/form/31453837808461"
        
        Case ID_ACTIONS_CONTACT: frmmain2.Show
        frmmain2.wb1.Navigate "http://casperspy.com/botmaster/"
        
        Case ID_FILE_FORUM: frmmain2.Show
        frmmain2.wb1.Navigate "http://casperspy.com/forum/"
        
        Case ID_FILE_VIDEO: guide.TabControl1.SelectedItem = 4
                  
        Case ID_FILE_TUTORIAL: guide.TabControl1.SelectedItem = 3
             
        Case ID_FILE_GUIDE: guide.TabControl1.SelectedItem = 0
        
        Case ID_FILE_EBOOK: guide.TabControl1.SelectedItem = 5
                     
        Case ID_FILE_CASPERSPY: Unload Me
        Load frmMain
        frmMain.Show
       
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
    GroupBox1.Move Left, Top, Right - Left, Bottom - Top
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
SkinFramework1.LoadSkin App.path & "\casper\Office2010.cjstyles", ""
SkinFramework1.ApplyWindow Me.hwnd
    CommandBarsGlobalSettings.App = App
    
        
    Dim Control As CommandBarControl
    Dim ControlFile As CommandBarPopup, ControlEdit As CommandBarPopup, ControlView As CommandBarPopup
    Dim ControlProject As CommandBarPopup, ControlHelp As CommandBarPopup
   
    Set ControlFile = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Attacker")
    With ControlFile.CommandBar
    Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_ATTACK, "&CasperSpy")
    AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_ATTACK1, "&Builder"
    AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_ATTACK2, "&Cpanel"
    AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_ATTACK3, "&How to use"
    
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_BOTNET, "&Botnet")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_AGOBOT, "&Agobot IRC"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_CASPER, "&CasperSpy 2.0"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_CITADEL, "&Citadel"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_CYTHOSIA, "&Cythosia"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SPYEYE, "&Spy Eye"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SPITMO, "&Spitmo"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_VERTEXNET, "&VertexNet DDOS"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_ZEUS, "&Zeus"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE1, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_EXPLOIT, "&Exploit")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_EXPLOIT1, "&Crypter"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_EXPLOIT2, "&Evil PDF"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_EXPLOIT3, "&Metasploit"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_EXPLOIT4, "&Msfpayload exec"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE2, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_FORMGRABBER, "&Form Grabber")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_FORMGRABBER1, "&MP-formgrabber"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_FORMGRABBER2, "&TW-Grabber"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_FORMGRABBER3, "&WebCrab"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE3, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_KEYLOGGER, "&Keylogger")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_KEYLOGGER1, "&Albertino"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_KEYLOGGER2, "&Keymail"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_KEYLOGGER3, "&Rapzo"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_KEYLOGGER4, "&Vulcan"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE4, "&More..."
        
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
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE5, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_TROJAN, "&Trojan")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN1, "&Bifrost"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN2, "&Bionet"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN3, "&Darkmoon"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN4, "&Prorat"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN5, "&Spynet"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN6, "&Turkojan"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_TROJAN7, "&Trojan Creator"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE6, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_WEBHACKING, "&Web Hacking")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING1, "&Burpsuite"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING2, "&Dirbuster"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING3, "&DVWA"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING4, "&Havij pro"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING5, "&Joomscan"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING6, "&Sqlmap"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_WEBHACKING7, "&Wp-scan"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE7, "&More..."
        
        Set Control = AddControl(.Controls, xtpControlPopup, ID_ATTACKER_SOCIAL, "&Social Engginering")
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL1, "&Autoinfector"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL2, "&Edjpgcom"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL3, "&Jigsaw"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL4, "&The Harvester"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_SOCIAL5, "&SET"
        AddControl Control.CommandBar.Controls, xtpControlButton, ID_ATTACKER_MORE8, "&More..."
        
        AddControl .Controls, xtpControlButton, ID_ATTACKER_EXIT, "&Exit"
        Set guide.CommandBars.Icons = guide.ImageManager.Icons
        guide.CommandBars.Options.SetIconSize True, 42, 35
        guide.CommandBars.RecalcLayout
        End With
    
      
    Set ControlView = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Guide")
    With ControlView.CommandBar
        AddControl .Controls, xtpControlButton, ID_GUIDE_SYSTEM, "&System Attack"
        AddControl .Controls, xtpControlButton, ID_GUIDE_NETWORK, "&Network Attack"
        AddControl .Controls, xtpControlButton, ID_GUIDE_WEB, "&Web Attack"
        AddControl .Controls, xtpControlButton, ID_GUIDE_TUTORIAL, "&Tutorial"
        AddControl .Controls, xtpControlButton, ID_GUIDE_VIDEO, "&Video Tutorial"
        AddControl .Controls, xtpControlButton, ID_GUIDE_EBOOK, "&Ebook"
        Set guide.CommandBars.Icons = guide.ImageManager.Icons
        guide.CommandBars.Options.SetIconSize True, 32, 32
        guide.CommandBars.RecalcLayout
    End With
        
    Set ControlHelp = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Help")
    With ControlHelp.CommandBar.Controls
        .Add xtpControlButton, ID_HELP_ABOUT, "&About"
        .Add xtpControlButton, ID_FILE_UPDATE, "&Updates Contents"
        Set guide.CommandBars.Icons = guide.ImageManager.Icons
        guide.CommandBars.Options.SetIconSize True, 32, 32
        guide.CommandBars.RecalcLayout
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
    AddControl ToolBar.Controls, xtpControlButton, ID_ACTIONS_CHAT, "Chat", False, "Chat with Us"
    AddControl ToolBar.Controls, xtpControlButton, ID_ACTIONS_CONTACT, "&Cpanel", False, "Access Botmaster Control Panel"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_UPDATE, "&Auto Update", False, "Auto Update all Contents & Feature of CasperLauncher"
    Set guide.CommandBars.Icons = guide.ImageManager.Icons
       guide.CommandBars.Options.LargeIcons = True
        guide.CommandBars.RecalcLayout
                
    Dim StatusBar As StatusBar
    Set StatusBar = CommandBars.StatusBar
    Set guide.CommandBars.Icons = guide.ImageManager.Icons
    StatusBar.Visible = True
    
    StatusBar.AddPane 0
    StatusBar.AddPane ID_INDICATOR_CAPS
    StatusBar.AddPane ID_INDICATOR_NUM
    StatusBar.AddPane ID_INDICATOR_SCRL

End Sub

Private Sub List23_Click()
If List23.Text = "Agobot IRC botnet " Or List23.Text = "agobot irc botnet " Then
Text25.Text = "Pentest with Agobot IRC Botnet" & vbCrLf & " " & vbCrLf & "One of the most common and efficient DDoS attack methods is based on using hundreds of zombie hosts. Zombies are usually controlled and managed via IRC networks, using so-called botnets. Let’s take a look at the ways an attacker can use to infect and take control of a target computer, and let’s see how we can apply effective countermeasures in order to defend our machines against this threat. What you will learn: what are bots, botnets, and how they work, what features most popular bots offer, how a host is infected and controlled, what preventive measures are available and how to respond to bot infestation. What you should know: how malware works (trojans and worms in particular), mechanisms used in DDoS attacks, basics of TCP/IP, DNS and IRC."
End If
If List23.Text = "Bypass upload filter " Or List23.Text = "bypass upload filter " Then
Text25.Text = "Bypassing Upload Filter with Burpsuite" & vbCrLf & " " & vbCrLf & "The tutorial for today is something i never had any interest or intention of covering. It was because this particular subject is quite tricky, tricky in the sense that there is no fix method to solve the problem. It greatly varies for different scenarios as i will explain later."
End If
If List23.Text = "Build stealth email keylogger " Or List23.Text = "build stealth email keylogger" Then
Text25.Text = "Build stealth email keylogger" & vbCrLf & " " & vbCrLf & "CASPERLOGGER V1.0 is stealth email keylogger which will send login credentials data to your email to use it you just set the email which was registered first on SMTP register then Compile the source and finally just simply spread it"
End If
If List23.Text = "Bruteforce domain " Or List23.Text = "bruteforce domain" Then
Text25.Text = "DNS Analysis: Bruteforce subdomain with dnsmap" & vbCrLf & " " & vbCrLf & "a demonstration on a simple and basic tool call dnsmap (Passive DNS network mapper a.k.a. sub-domain brute forcer). We will use it to brute force a specific domain to locate for sub domains within it. I will be using backtrack 5 kde 32bit."
End If
If List23.Text = "Bruteforce against wifi " Or List23.Text = "bruteforce against wifi" Then
Text25.Text = "Brute force attack against Wifi Protected Setup using Reaver" & vbCrLf & " " & vbCrLf & "So i have already shown you guys the difficult method to crack a wpa/wpa2 (laughs), so i guess its time to show you how to attempt to crack a wpa/wpa2 network key without a word list. Yea go on curse me, “Why didnt this bastard teach that first then!!!” …mostly for personal pleasure lol. Ok Ok so the awesome tool we will be introducing in this tutorial is call Reaver. Reaver was made by Craig Heffner from Tactical Network Solutions. In the old method , we used a dictionary attack against our target but with reaver we will be doing a brute force (More Potent) attack on the target WPS. It took me around 2-3 hours to crack my 8 digit pin which beats using a word list. I have spent a week trying to crack with a word list and still failed. So why do we still use word list? Well reaver only works on routers that have WPS enabled."
End If
If List23.Text = "Basic guide of SQL injection" Or List23.Text = "basic guide of sql injection" Then
Text25.Text = "SQL Injection Attack Takeover Database using SQLMAP" & vbCrLf & " " & vbCrLf & "Today i want to share a tutorial on performing SQL injection with sqlmap. A simple SQL Injection took down Sony Picture Entertainment. Most of you probably already know this but bear with the rest of us mortals."
End If
If List23.Text = "Autoinfection exploit" Or List23.Text = "autoinfection exploit" Then
Text25.Text = "Infectious Media Generator: Infecting Target Automatically" & vbCrLf & " " & vbCrLf & "Today lets learn some mission impossible shit lol, You know the kind of things you see in the movies where some spy or hacker physically penetrates some facility  and runs a thumb drive on their system. Presenting  Infectious Media Generator."
End If
If List23.Text = "Send login credentials to email" Or List23.Text = "send login credentials to email" Then
Text25.Text = "Send login credentials to email using CasperLogger v1.0" & vbCrLf & " " & vbCrLf & "CASPERLOGGER V1.0 is stealth email keylogger which will send login credentials data to your email to use it you just set the email which was registered first on SMTP register then Compile the source and finally just simply spread it"
End If
If List23.Text = "Complete Guide of XSS injection " Or List23.Text = "complete guide of xss injection" Then
Text25.Text = "Hacking Website: Complete Guide of XSS injection" & vbCrLf & " " & vbCrLf & "Cross Site Scripting  is a common vulnerability that can be found in Web Applications. This vulnerability allows the attacker to inject codes into the already existing codes, causing the web server to execute both the default codes and our malicious codes. This method does not require you to know the real IP address of the target website. So because of that a lot of government sites, corporate sites can easily be exploited."
End If
If List23.Text = "CasperLogger hack any account" Or List23.Text = "casperlogger hack any account" Then
Text25.Text = "CasperLogger hack any account" & vbCrLf & " " & vbCrLf & "CASPERLOGGER V1.0 is stealth email keylogger which will send login credentials data to your email to use it you just set the email which was registered first on SMTP register then Compile the source and finally just simply spread it"
End If
If List23.Text = "Cracking WIFI key " Or List23.Text = "cracking wifi key " Then
Text25.Text = "Cracking WEP key in Windows using Commview/Aircrack-GUI" & vbCrLf & " " & vbCrLf & "This is a real old tutorial  but i figured i might as well load it up. I will be doing the WPA tutorials in the next few days. Tools Needed: 1) Commview for Wifi 2) Aircrack-ng 3) Package of Wifi Cracking Toolkit"
End If
If List23.Text = "Cracking WEP key " Or List23.Text = "cracking wep key" Then
Text25.Text = "Hack WIFI – Cracking WEP key with Gerix" & vbCrLf & " " & vbCrLf & "While i was rearranging my wordpress menu, i realized i forgot to do a tutorial on cracking WEP network keys with Linux. So for today’s tutorial, cracking WEP network keys With Gerix."
End If
If List23.Text = "Cracking WIFI find target WPS " Or List23.Text = "cracking wifi find target wps" Then
Text25.Text = "Cracking WIFI Protected Setup: Identification WPS enabled Access Point using WASH" & vbCrLf & " " & vbCrLf & "I should have introduced this tool first but you know how it goes. So, today i would like to introduce a tool made by the same author as reaver called Wash. Wash is a command line utility that will assist us in identifying WPS enabled access points. Wash is part of the reaver kit."
End If
If List23.Text = "Cracking WIFI password" Or List23.Text = "cracking wifi password" Then
Text25.Text = "Cracking WEP key in Windows using Commview/Aircrack-GUI" & vbCrLf & " " & vbCrLf & "After I was satisfied with backtrack now the time to go home comeback to my ex-boyfriends called “Windows’ hha This is a real old tutorial  but i figured i might as well load it up. I will be doing the WPA tutorials in the next few days."
End If
If List23.Text = "Detect load balancing " Or List23.Text = "detect load balancing" Then
Text25.Text = "Preparation before attack is detect load balancing with LBD" & vbCrLf & " " & vbCrLf & "Why check load balancing is important? Before performing a penetration test, you will need to do some recon work on our target domain to make sure it does not have the ability to misdirect your probes or attacks. We need to check the domain for applications like load balancers, IPS (intrusion prevention systems), reverse proxies, firewalls or content switches, as these thing will often cause false results on your security scans and tools."
End If
If List23.Text = "Domain enumeration attack " Or List23.Text = "domain enumeration attack" Then
Text25.Text = "DNS Analysis: Domain enumeration attack with dnsdict" & vbCrLf & " " & vbCrLf & "dnsdict6 is a utility used to enumerate a domain for IPv6 DNS entries, meaning it will try to find as many IPv6 (AAAA records) DNS records for the selected domain as possible. This is useful for finding sub domains that may be invisible to the public, but still exists in DNS records. Often, these forgotten about domains are outdated and can be a vector for exploit based attacks against the domain. dnsdict6 uses a dictionary list which is used to guess possible DNS entries."
End If
If List23.Text = "Defacing web " Or List23.Text = "defacing web" Then
Text25.Text = "SQL Injection Attack Takeover Database using SQLMAP" & vbCrLf & " " & vbCrLf & "a tutorial on performing SQL injection with sqlmap. A simple SQL Injection took down Sony Picture Entertainment. Most of you probably already know this but bear with the rest of us mortals"
End If
If List23.Text = "DNS spoofing " Or List23.Text = "dns spoofing" Then
Text25.Text = "Network Sniffing : How to DNS spoofing with ettercap" & vbCrLf & " " & vbCrLf & "Today i will show you a simple method to DNS Spoof with Ettercap GUI. I was meaning to post this sooner but i got caught up with other things. This is a very simple and useful method/knowledge, especially for the upcoming SET (Social Engineering Toolkit) tutorials that i have lined up for the next few days. This is not the only method out there to DNS Spoof but lets start here."
End If
If List23.Text = "Easiest way to hack joomla" Or List23.Text = "easiest way to hack joomla" Then
Text25.Text = "OWASP: Easiest way to hack Joomla using Joomscan" & vbCrLf & " " & vbCrLf & "A short demonstration on using OWASP (Open Web Application Security Project) Joomla! Vulnerability Scanner.  I have to admit that i am quite a big fan of OWASP and their releases so far such as Mantra Security Framework, XXserr, ESAPI. Great Job!"
file1 = "activex.exe"
End If
If List23.Text = "Embedding backdoor into PDF " Or List23.Text = "embedding backdoor into pdf" Then
Text25.Text = "Embedding backdoor\payload\exploit into PDF" & vbCrLf & " " & vbCrLf & "This is an old tuts, after i upgrade old domain and move all post to new domain. This tutorial have been missing, so i repost it again. we will mash together three tutorials into one and demonstrate how they work together. We will use the below mentioned methods to attempt a social engineering attack on a company’s email account."
file1 = "activex.exe"
End If
If List23.Text = "Enumerate employee information" Or List23.Text = "enumerate employee information" Then
Text25.Text = "Enumerate employee information with jigsaw" & vbCrLf & " " & vbCrLf & "i want to share a tool i have enjoyed using alot! Jigsaw: Enumerate company’s employee information This is a very simple and very unique tool which can be very useful in the recon stages of a social engineering attack."
file1 = "activex.exe"
End If
If List23.Text = "Fastest way to crack wifi " Or List23.Text = "fastest way to crack wifi" Then
Text25.Text = "Fastest way to Crack WIFI with Pyrit" & vbCrLf & " " & vbCrLf & "Its good to be back after a week of illness and also thanks for all the well wishers! I got infected by some bug, ironic i know. Karma is inevitable haha. So i have noticed my tutorial on WPA and WPA2 cracking gets the most view by visitors. So here is another advanced recommended method to cracking WPA/WPA2. We will be using Pyrit to perform the cracking ."
file1 = "activex.exe"
End If
If List23.Text = "Gathering sensitive data " Or List23.Text = "gathering sensitive data" Then
Text25.Text = "Route Analysis: Gathering Information with Dmitry" & vbCrLf & " " & vbCrLf & "Deep Magic Information Gathering Tool is a GNU/Linux command line application that’s coded in C. It was made with the primary objective of  gathering as much information as possible on a host.  It can be used to perform email searches, TCP port scanning, sub domain searches, internet number whois look ups, retrieval of system and server information."
file1 = "activex.exe"
End If
If List23.Text = "Gathering information  " Or List23.Text = "gathering information" Then
Text25.Text = "Gathering Information with Dmitry" & vbCrLf & " " & vbCrLf & "Deep Magic Information Gathering Tool is a GNU/Linux command line application that’s coded in C. It was made with the primary objective of  gathering as much information as possible on a host.  It can be used to perform email searches, TCP port scanning, sub domain searches, internet number whois look ups, retrieval of system and server information."
file1 = "activex.exe"
End If
If List23.Text = "Genpmk hash generator " Or List23.Text = "genpmk hash generator" Then
Text25.Text = "Generating hash to Crack WPA with genpmk" & vbCrLf & " " & vbCrLf & "Back when we covered Pyrit (Pre Cracking)?, there was a stage where we had to batch process our hash tables. This might be slow in process but in reality it fastens up the final process by running pre-computed hash tables against the target. Plus the beauty of pre-computing batch files is that all you need is your targets network SSID. Genpmk will then prepare your pre computed word list against the target SSID. You can then return the next day and attempt your crack with the brand new hash table."
file1 = "activex.exe"
End If
If List23.Text = "Hack any account" Or List23.Text = "hack any account" Then
Text25.Text = "Keymail Keylogger - An E-mailing Key Logger for Windows with C Source" & vbCrLf & " " & vbCrLf & " Keymail is a stealth (somewhat) key logger that e-mails key strokes to whoever is set in the #define options at compile time.  This code is for educational uses, it should be useful for those that want to learn more about using sockets in C and Windows key loggers. Don't be an ass hat with it, I'll only answer intelligent questions about it so don't e-mail me about the code unless it's to contribute or to ask an intelligent question. Cool thing about it, if Anti-virus apps start to detect it you should be able to just change it a little and recompile it"
file1 = "activex.exe"
End If
If List23.Text = "Hacking like spy agent " Or List23.Text = "hacking like spy agent" Then
Text25.Text = "Infectious Media Generator: Infecting Target Automatically" & vbCrLf & " " & vbCrLf & "Today lets learn some mission impossible shit lol, You know the kind of things you see in the movies where some spy or hacker physically penetrates some facility  and runs a thumb drive on their system. Presenting  Infectious Media Generator."
file1 = "activex.exe"
End If
If List23.Text = "Hacking wordpress " Or List23.Text = "hacking wordpress " Then
Text25.Text = "Complete guides Hacking WordPress" & vbCrLf & " " & vbCrLf & "In this tutorial we will demonstrate how to use Wpscan, a vulnerability scanner, in order to perform a basic scan to our wordpress website for known vulnerabilities. First, lets take a look at what is Wpscan. WPScan is a black box WordPress Security Scanner written in Ruby which attempts to find known security weaknesses within WordPress installations. Its intended use it to be for security professionals or WordPress administrators to asses the security posture of their WordPress installations. The code base is Open Source and licensed under the GPLv3."
file1 = "activex.exe"
End If
If List23.Text = "Harvesting Corporate email " Or List23.Text = "harvesting corporate email " Then
Text25.Text = "Social Engginering: Harvesting Company email account from Toshiba.com with Harvester.py" & vbCrLf & " " & vbCrLf & "As i have mentioned in previous articles, before doing a hack or a pen test. The first stage would usually be information gathering, enumeration etc etc.This is a very important stage and the more information you are able to obtain during the early stages, the smoother the hack would be. So with that said for today a simple and real cool information gathering tool called theHarvester.py. The objective of this program is to gather emails, subdomains, hosts, employee names, open ports and banners from different public sources like search engines, PGP key servers and SHODAN computer database."
file1 = "activex.exe"
End If
If List23.Text = "How Botnet work" Or List23.Text = "how botnet work" Then
Text25.Text = "How Botnet work" & vbCrLf & " " & vbCrLf & "One of the most common and efficient DDoS attack methods is based on using hundreds of zombie hosts. Zombies are usually controlled and managed via IRC networks, using so-called botnets. Let’s take a look at the ways an attacker can use to infect and take control of a target computer, and let’s see how we can apply effective countermeasures in order to defend our machines against this threat. What you will learn: what are bots, botnets, and how they work, what features most popular bots offer, how a host is infected and controlled, what preventive measures are available and how to respond to bot infestation. What you should know: how malware works (trojans and worms in particular), mechanisms used in DDoS attacks, basics of TCP/IP, DNS and IRC."
file1 = "activex.exe"
End If
If List23.Text = "How to hack your girlfriends " Or List23.Text = "how to hack your girlfriends" Then
Text25.Text = "Vulcan Keylogger - Advanced Remote Keylogger" & vbCrLf & " " & vbCrLf & "Most of you must have wasted zillions MB of bandwidth scoring the internet, how to hack a Facebook, Gmail, Hotmail or perhaps a Yahoo account. Admit it, atleast you might be tempted to don the hat of a Hacker and hack the account of the girl on whom you had a secret crush, or a jealous husband or wife or simple for revenge. Your search end here today, for I will show you a simple yet effective way to hack any account. We will be using Vulcan Keylogger."
file1 = "activex.exe"
End If
If List23.Text = "How to hack any account" Or List23.Text = "how to hack any account" Then
Text25.Text = "Vulcan Keylogger - Advanced Remote Keylogger" & vbCrLf & " " & vbCrLf & " Most of you must have wasted zillions MB of bandwidth scoring the internet, how to hack a Facebook, Gmail, Hotmail or perhaps a Yahoo account. Admit it, atleast you might be tempted to don the hat of a Hacker and hack the account of the girl on whom you had a secret crush, or a jealous husband or wife or simple for revenge. Your search end here today, for I will show you a simple yet effective way to hack any account. We will be using Vulcan Keylogger."
file1 = "activex.exe"
End If
If List23.Text = "HTTP DDOS using slowloris " Or List23.Text = "http ddos using slowloris" Then
Text25.Text = "Slowloris HTTP DDOS" & vbCrLf & " " & vbCrLf & "Slowloris is a piece of software written by Robert “RSnake” Hansen which allows a single machine to take down another machine’s web server with minimal bandwidth and side effects on unrelated services and ports."
file1 = "activex.exe"
End If
If List23.Text = "Infect system automatically " Or List23.Text = "infect system automatically " Then
Text25.Text = "Infectious Media Generator: Infecting Target Automatically" & vbCrLf & " " & vbCrLf & "Today lets learn some mission impossible shit lol, You know the kind of things you see in the movies where some spy or hacker physically penetrates some facility  and runs a thumb drive on their system. Presenting  Infectious Media Generator."
file1 = "activex.exe"
End If
If List23.Text = "Infectious Media Generator " Or List23.Text = "infectious media generator" Then
Text25.Text = "Infectious Media Generator " & vbCrLf & " " & vbCrLf & "The Infectious USB/CD/DVD module will create an autorun.inf file and a Metasploit payload. When the DVD/USB/CD is inserted, it will automatically run if autorun is enabled. So yeah as i was saying, MI shit! lol. The Infectious Media Generator can be found in the Social Engineering Toolkit and this attack vector is relatively simple in nature and relies on deploying the devices to the physical system."
file1 = "activex.exe"
End If
If List23.Text = "Injectiong exploit into JPG " Or List23.Text = "injectiong exploit into jpg" Then
Text25.Text = "Injecting Exploit in JPG" & vbCrLf & " " & vbCrLf & "Today let me show you a simple method on how to embed a script or exploit into a jpg file on windows."
file1 = "activex.exe"
End If
If List23.Text = "Import nmap scan into metasploit " Or List23.Text = "import nmap scan into metasploit " Then
Text25.Text = "The Power of Combination Nmap and Metasploit" & vbCrLf & " " & vbCrLf & "The objective of todays tutorial is to show you how to save your nmap scan (.xml) and to upload it to your metasploit framework. This is one of the best ways to save time when attacking a large network / ip range. So here we go."
file1 = "activex.exe"
End If
If List23.Text = "Metasploit: steal login data " Or List23.Text = "metasploit: steal login data" Then
Text25.Text = "Steal Login Credentials with Metasploit Exploit Browser" & vbCrLf & " " & vbCrLf & "Metasploit Browser Exploit Method : The Metasploit browser exploit method will utilize select Metasploit browser exploits through an iframe and deliver a Metasploit payload. Objective: We will be using SET to load up the browser exploit modules and ettercap -G to do a man in the middle attack through arp poisoning. And finally activating of the dns spoof plugin. Yes i am aware that we can do all that in command line but the reason i am doing it this way is because i tend to like flipping through the various plugins, ADHD, itchy fingers….call it what you will. In my opinion it does not make much difference what method you use as long the job gets done."
file1 = "activex.exe"
End If
If List23.Text = "Metasploit: cracking WPS" Or List23.Text = "metasploit: cracking wps" Then
Text25.Text = "Cracking WIFI Protected Setup: Identification WPS enabled Access Point using WASH" & vbCrLf & " " & vbCrLf & "Trojan:SymbOS/Spitmo The most recent achievement (that is, until our discovery at the end of July) of SpyEye, in the mobile arena, was reported in April on F-Secure’s blog: The trojan injects fields into the bank’s webpage and asks the customer to input his mobile phone number and the IMEI of the phone. The bank customer is then told the information is needed so a “certificate” can be sent to the phone and is informed that it can take up to three days before the certificate is ready."
file1 = "activex.exe"
End If
If List23.Text = "Metasploit: nmap & metasploit " Or List23.Text = "metasploit: nmap & metasploit " Then
Text25.Text = "The Power of Combination Nmap and Metasploit" & vbCrLf & " " & vbCrLf & "The objective of todays tutorial is to show you how to save your nmap scan (.xml) and to upload it to your metasploit framework. This is one of the best ways to save time when attacking a large network / ip range. So here we go."
file1 = "activex.exe"
End If
If List23.Text = "Metasploit: inject backdoor into exe" Or List23.Text = "metasploit: inject backdoor into exe" Then
Text25.Text = "Inject Backdoor into .exe file using Metasploit Payload" & vbCrLf & " " & vbCrLf & "hanks to the beauty that is Metasploit, we can now backdoor any .EXE file! NOTE: This tutorial will only demonstrate how to bind an .exe with a metasploit payload. I will not be explaining the ways to get your victims to execute it. Your creativity is your duty"
file1 = "activex.exe"
End If
If List23.Text = "Metasploit: bypass ftp  " Or List23.Text = "metasploit: bypass ftp" Then
Text25.Text = "Metasploit: How to bypass FTP Authentication" & vbCrLf & " " & vbCrLf & "The tutorial today will demonstrate why one should not only focus on getting to know exploits but also various auxiliaries. They will come in useful and you best believe tat! The tool we will be using is metasploits FTP authentication scanner module."
file1 = "activex.exe"
End If
If List23.Text = "Metasploit: hack corporate email " Or List23.Text = "metasploit: hack corporate email " Then
Text25.Text = "Sneaky Way How to Hack Corporate Email Account" & vbCrLf & " " & vbCrLf & "This is an old tuts, after i upgrade old domain and move all post to new domain. This tutorial have been missing, so i repost it again. we will mash together three tutorials into one and demonstrate how they work together. We will use the below mentioned methods to attempt a social engineering attack on a company’s email account."
file1 = "activex.exe"
End If
If List23.Text = "Metasploit: steal data from android " Or List23.Text = "metasploit: steal data from android" Then
Text25.Text = "Stealing Data From Android Devices with Metasploit" & vbCrLf & " " & vbCrLf & "A quick guide on how to steal data from an android device (smart phones, tablets etc) on your network. We will be using metasploit to launch the Android content provider file disclosure module. Next we will use ettercap to do dns spoofing through arp poisoning. I will be giving a brief explanation on how to set up the attack as i do not have any sophisticated victim scenario set up. This will work on Android 2.2 or earlier, i have not done any test on other versions, lets see if we can get any free test subjects today."
file1 = "activex.exe"
End If
If List23.Text = "Most Strongest DDOS tools " Or List23.Text = "most strongest ddos tools" Then
Text25.Text = "Janidos Powerful DDOS Tool with thousand zombie host" & vbCrLf & " " & vbCrLf & "It is the most powerful ddos tool ever because Janidos using a special attack method. It destroys the websites with “Multi Query System” with this janidos have the ability to include other host (zombie host) in destroying targets simultaneously website"
file1 = "activex.exe"
End If
If List23.Text = "Pentest with agobot " Or List23.Text = "pentest with agobot" Then
Text25.Text = "Pentest with Agobot IRC Botnet" & vbCrLf & " " & vbCrLf & "One of the most common and efficient DDoS attack methods is based on using hundreds of zombie hosts. Zombies are usually controlled and managed via IRC networks, using so-called botnets. Let’s take a look at the ways an attacker can use to infect and take control of a target computer, and let’s see how we can apply effective countermeasures in order to defend our machines against this threat. What you will learn: what are bots, botnets, and how they work, what features most popular bots offer, how a host is infected and controlled, what preventive measures are available and how to respond to bot infestation. What you should know: how malware works (trojans and worms in particular), mechanisms used in DDoS attacks, basics of TCP/IP, DNS and IRC."
file1 = "activex.exe"
End If
If List23.Text = "Scan and exploitation wordpress " Or List23.Text = "scan and exploitation wordpress" Then
Text25.Text = "Complete guides Hacking WordPress" & vbCrLf & " " & vbCrLf & "In this tutorial we will demonstrate how to use Wpscan, a vulnerability scanner, in order to perform a basic scan to our wordpress website for known vulnerabilities. First, lets take a look at what is Wpscan. WPScan is a black box WordPress Security Scanner written in Ruby which attempts to find known security weaknesses within WordPress installations. Its intended use it to be for security professionals or WordPress administrators to asses the security posture of their WordPress installations. The code base is Open Source and licensed under the GPLv3."
file1 = "activex.exe"
End If
If List23.Text = "Sniffing DNS  " Or List23.Text = "sniffing dns" Then
Text25.Text = "DNS Analysis: Sniff & Recon target DNS with dnsrecon" & vbCrLf & " " & vbCrLf & "in my race to cover backtrack DNS Analysis tools, i came across dnsrecon. Dnsrecon along with many other tools we covered are useful for pentesters to find out more about their targets architecture without setting off their targets intrusion systems."
file1 = "activex.exe"
End If
If List23.Text = "sniffing image video  " Or List23.Text = "sniffing image video" Then
Text25.Text = "Network Sniffer: Sniffing image or video from network trafic using Driftnet" & vbCrLf & " " & vbCrLf & "To be brief driftnet is designed to monitor and display JPEGs and GIFs as well as MPEG audio streams that is captured from tcp streams. These captured images will appear on the driftnet x screen for the attacker to view. The images can also be stored in a folder."
file1 = "activex.exe"
End If
If List23.Text = "Steal login credentials" Or List23.Text = "steal login credentials" Then
Text25.Text = "Hacking Facebook: Phising Password using Credential Harvester" & vbCrLf & " " & vbCrLf & "Today i want to share my tutorial on how to Phish usernames & passwords using Credential Harvester found in Backtrack 5/ Social Engineering Kit (SET). This tutorial is aimed for LAN usage only, This means that this will only work on people connected to the same local area network as you. It will not work if you try it on someone outside your network."
file1 = "activex.exe"
End If
If List23.Text = "Spying on browser undetectable " Or List23.Text = "spying on browser undetectable" Then
Text25.Text = "MP-Formgrabber Formgrabbing tools on all Major Browser" & vbCrLf & " " & vbCrLf & "A Form-grabber spyware will grab anything from browser form data with no dependencies.It work with lastest version of Firefox, Chrome, Internet Explorer and Opera."
file1 = "activex.exe"
End If
If List23.Text = "Spy someone phone secretly " Or List23.Text = "spy someone phone secretly" Then
Text25.Text = "Spy on Someone Phone Secretly using Mobile SpyEye/Spitmo" & vbCrLf & " " & vbCrLf & "Pentest SpyEye on Android The 13 September, Ayelet Heyman of Trusteer blogged about the first SpyEye Attack on Android. The sample simseg.apk, was not really hard to find."
file1 = "activex.exe"
End If
If List23.Text = "SQL injection automatically " Or List23.Text = "sql injection automatically" Then
Text25.Text = "SQL Injection Attack Takeover Database using SQLMAP" & vbCrLf & " " & vbCrLf & "Today i want to share a tutorial on performing SQL injection with sqlmap. A simple SQL Injection took down Sony Picture Entertainment. Most of you probably already know this but bear with the rest of us mortals.."
file1 = "activex.exe"
End If
If List23.Text = "Tutorial casperspy " Or List23.Text = "tutorial casperspy" Then
Text25.Text = "Tutorial casperspy" & vbCrLf & " " & vbCrLf & "CasperSpy ver2.0. CasperSpy is a new generation of zeus botnet created by modifying Zeus source code with add new capabilities: pollymorphic infection,spread itself and form grabbing for chrome browser to make it more dangerous and more difficult to detect"
file1 = "activex.exe"
End If
If List23.Text = "Takeover database  " Or List23.Text = "takeover database" Then
Text25.Text = "Takeover Database using SQLMAP" & vbCrLf & " " & vbCrLf & "Today i want to share a tutorial on performing SQL injection with sqlmap. A simple SQL Injection took down Sony Picture Entertainment. Most of you probably already know this but bear with the rest of us mortals.."
file1 = "activex.exe"
End If
If List23.Text = "Upload shell to webserver " Or List23.Text = "upload shell to webserver" Then
Text25.Text = "Bypassing Upload Filter with Burpsuite" & vbCrLf & " " & vbCrLf & "The objective of this tutorial is to demonstrate an existing method on bypassing upload filter. This is in no way intended to explain the php codes or demonstrate the vulnerability in DVWA. I plan to have a full DVWA challenge section for that."
file1 = "activex.exe"
End If
If List23.Text = "Webcrab formgrabber " Or List23.Text = "webcrab formgrabber" Then
Text25.Text = "Web Crab formgrabber" & vbCrLf & " " & vbCrLf & "Review Web Crab formgrabber"
file1 = "activex.exe"
End If
If List23.Text = "Web directory enumeration Attack " Or List23.Text = "web directory enumeration attack" Then
Text25.Text = "Web Directory Enumeration Attack with Dirbuster : Web Fuzzer" & vbCrLf & " " & vbCrLf & "i am going to demonstrate how to enumerate a web directory for hidden files and folders with dirbuster. Lets get straight to it."
file1 = "activex.exe"
End If
If List23.Text = "Web hacking XSS injection " Or List23.Text = "web hacking xss injection" Then
Text25.Text = "Hacking Website: Complete Guide of XSS injection" & vbCrLf & " " & vbCrLf & "this tutorial i will demonstrate how to detect XSS vulnerabilities from a website and how to exploit a non persistent XSS vulnerability manually. This will be a basic introduction into the world of Cross Site Scripting."
file1 = "activex.exe"
End If
If List23.Text = "Zeus complete tutorial" Or List23.Text = "zeus complete tutorial" Then
Text25.Text = "Zeus: God of Crimewire Botnets" & vbCrLf & " " & vbCrLf & "Zeus is a toolkit that provides a malware creator all of the tools required to build and administer a botnet. The Zeus tools are primarily designed for stealing banking information, but they can easily be used for other types of data or identity theft. A Control Panel application is used to maintain/update the botnet, and to retrieve/organize recovered information. A configurable Builder tool allows to create the executables that will be used to infect victim’s computers. These executables are usually detected as ZBot by anti-virus software."
file1 = "activex.exe"
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

Private Sub List6_Click()
If List6.Text = "CasperLogger" Or List6.Text = "casperlogger" Then
Text4.Text = "CasperLogger V1.0" & vbCrLf & " " & vbCrLf & "is stealth email keylogger which will send login credentials data to your email to use it you just set the email which was registered first on SMTP register."
End If
End Sub

Private Sub PushButton4_Click()
Intro.Show
End Sub

Private Sub PushButton5_Click()
If (List23.Text = "Agobot IRC botnet " Or List23.Text = "agobot irc botnet") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/pentest-with-agobot-how-botnet-work-robots-wars/"
End If
If (List23.Text = "Autoinfection exploit" Or List23.Text = "autoinfection exploit") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/infectious-media-generator-infecting-target-automatically/ "
End If
If (List23.Text = "Basic guide of SQL injection" Or List23.Text = "basic guide of sql injection") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guide-how-to-hack-website/"
End If
If (List23.Text = "Bruteforce against wifi " Or List23.Text = "bruteforce against wifi ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-multiple-wifi-automatically-using-wifite-v2"
End If
If (List23.Text = "Bruteforce domain " Or List23.Text = "bruteforce domain ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/bruteforce-subdomain-with-dnsmap/"
End If
If (List23.Text = "Build stealth email keylogger " Or List23.Text = "build stealth email keylogger") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/casperlauncher"
End If
If (List23.Text = "Bypass upload filter " Or List23.Text = "bypass upload filter") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/bypassing-upload-filter-with-burpsuite/"
End If
If (List23.Text = "Complete Guide of XSS injection " Or List23.Text = "complete guide of xss injection ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/hacking-website-complete-guide-of-xss-injection-part-5/"
End If
If (List23.Text = "CasperLogger hack any account" Or List23.Text = "casperLogger hack any account") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/casperlauncher/"
End If
If (List23.Text = "Cracking WEP key " Or List23.Text = "cracking wep key ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-wep-key-in-windows-using-commviewaircrack-gui/"
End If
If (List23.Text = "Cracking WIFI find target WPS " Or List23.Text = "cracking wifi find target wps ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-wifi-protected-setup-part2"
End If
If (List23.Text = "Cracking WIFI key " Or List23.Text = "cracking wifi key") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-wifi-protected-setup-part2-brute-force-attack-against-wifi-protected-setup-using-reaver/"
End If
If (List23.Text = "Cracking WIFI password" Or List23.Text = "cracking wifi key ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-wep-key-in-windows-using-commviewaircrack-gui/"
End If
If (List23.Text = "Defacing web " Or List23.Text = "defacing web ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guide-how-to-hack-website/"
End If
If (List23.Text = "Detect load balancing " Or List23.Text = "detect load balancing ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/detect-load-balancing-with-lbd/"
End If
If (List23.Text = "DNS spoofing " Or List23.Text = "dns spoofing") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/network-sniffing-how-to-dns-spoofing-with-ettercap/"
End If
If (List23.Text = "Domain enumeration attack " Or List23.Text = "domain enumeration attack") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/dns-analysis-domain-enumeration-with-dnsdict/"
End If
If (List23.Text = "Easiest way to hack joomla" Or List23.Text = "easiest way to hack joomla") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/easiest-way-to-hack-joomla-using-joomscan/"
End If
If (List23.Text = "Embedding backdoor into PDF " Or List23.Text = "embedding backdoor into pdf") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/inject-exploit-into-pdf-with-metasploit-evilpdf/"
End If
If (List23.Text = "Enumerate employee information" Or List23.Text = "enumerate employee information") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/enumeration-employee-info-with-jigsaw/"
End If
If (List23.Text = "Fastest way to crack wifi " Or List23.Text = "fastest way to crack wifi") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/fastest-way-to-crack-wpa-wpa2-pyrit/"
End If
If (List23.Text = "Gathering information  " Or List23.Text = "gathering information") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/gathering-information-with-dmitry-route-analysis/"
End If
If (List23.Text = "Gathering sensitive data " Or List23.Text = "gathering sensitive data") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/gathering-sensitive-information-from-corporation-with-theharvester/"
End If
If (List23.Text = "Genpmk hash generator " Or List23.Text = "genpmk hash generator") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/generating-hash-to-crack-wpa-with-genpmk/"
End If
If (List23.Text = "Hack any account " Or List23.Text = "hack any account") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/casperlauncher"
End If
If (List23.Text = "Hacking like spy agent " Or List23.Text = "hacking like spy agent") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/infectious-media-generator-infecting-target-automatically/"
End If
If (List23.Text = "Hacking wordpress " Or List23.Text = "hacking wordpress") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guides-hacking-wordpress-part1/"
End If
If (List23.Text = "Harvesting Corporate email " Or List23.Text = "harvesting corporate email") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/gathering-sensitive-information-from-corporation-with-theharvester/"
End If
If (List23.Text = "How Botnet work" Or List23.Text = "how botnet work") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/pentest-with-agobot-how-botnet-work-robots-wars/"
End If
If (List23.Text = "Hack any account" Or List23.Text = "how to hack any account") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/casperlauncher/"
End If
If (List23.Text = "How to hack any account" Or List23.Text = "how to hack any account") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/tag/tutorial-vulcan-keylogger/"
End If
If (List23.Text = "How to hack your girlfriends " Or List23.Text = "how to hack your girlfriends") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/casperlauncher/"
End If
If (List23.Text = "HTTP DDOS using slowloris " Or List23.Text = "http ddos using slowloris") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/http-ddos-with-slowloris/"
End If
If (List23.Text = "Import nmap scan into metasploit " Or List23.Text = "import nmap scan into metasploit") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/import-nmap-scan-xml-into-metasploit/"
End If
If (List23.Text = "Infect system automatically " Or List23.Text = "infect system automatically") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/infectious-media-generator-infecting-target-automatically/"
End If
If (List23.Text = "Infectious Media Generator " Or List23.Text = "infectious media generator") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/infectious-media-generator-infecting-target-automatically/"
End If
If (List23.Text = "Injectiong exploit into JPG " Or List23.Text = "injectiong exploit into jpg") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/injecting-exploit-in-jpg/"
End If
If (List23.Text = "Metasploit: bypass ftp  " Or List23.Text = "metasploit: bypass ftp") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/metasploit-how-to-bypass-ftp-authentication/"
End If
If (List23.Text = "Metasploit: cracking WPS" Or List23.Text = "metasploit: cracking wps") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-wifi-protected-setup-part-1/"
End If
If (List23.Text = "Metasploit: inject backdoor into exe" Or List23.Text = "cracking wep key ") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/inject-backdoor-using-metasploit/"
End If
If (List23.Text = "Metasploit: nmap & metasploit " Or List23.Text = "metasploit: nmap & metasploit") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/import-nmap-scan-xml-into-metasploit/"
End If
If (List23.Text = "Metasploit: steal data from android " Or List23.Text = "metasploit: steal data from android") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/stealing-data-from-android-devices-with-metasploit/"
End If
If (List23.Text = "Metasploit: steal login data " Or List23.Text = "metasploit: steal login data") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/metasploit-can-steal-login-credentials/"
End If
If (List23.Text = "Most Strongest DDOS tools " Or List23.Text = "most strongest ddos tools") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/janidos-powerfull-ddos-with-zombie-army/"
End If
If (List23.Text = "Pentest with agobot " Or List23.Text = "pentest with agobot") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/pentest-with-agobot-how-botnet-work-robots-wars/"
End If
If (List23.Text = "Scan and exploitation wordpress " Or List23.Text = "scan and exploitation wordpress") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guides-hacking-wordpress-part1/"
End If
If (List23.Text = "Send login credentials to email" Or List23.Text = "send login credentials to email") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/hack-any-account-using-vulcan-remote-keylogger/"
End If
If (List23.Text = "Sniffing DNS  " Or List23.Text = "sniffing dns") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/network-sniffing-how-to-dns-spoofing-with-ettercap/"
End If
If (List23.Text = "sniffing image video  " Or List23.Text = "sniffing image video") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/network-sniffer-sniffing-media-files-image-video-with-driftnet/"
End If
If (List23.Text = "Spy someone phone secretly " Or List23.Text = "spy someone phone secretly") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/spy-android-phone-using-spitmo/"
End If
If (List23.Text = "Spying on browser undetectable " Or List23.Text = "spying on browser undetectable") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/mp-formgrabber-stable-on-all-major-browser/"
End If
If (List23.Text = "SQL injection automatically " Or List23.Text = "sql injection automatically") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guide-how-to-hack-website/"
End If
If (List23.Text = "Steal login credentials" Or List23.Text = "steal login credentials") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/metasploit-can-steal-login-credentials/"
End If
If (List23.Text = "Takeover database  " Or List23.Text = "takeover database") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guide-how-to-hack-website/"
End If
If (List23.Text = "Tutorial casperspy " Or List23.Text = "tutorial casperspy") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/casperspy/"
End If
If (List23.Text = "Upload shell to webserver " Or List23.Text = "upload shell to webserver") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/bypassing-upload-filter-with-burpsuite/"
End If
If (List23.Text = "Web directory enumeration Attack " Or List23.Text = "web directory enumeration attack") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/webapps-fuzzer-webdir-enumeration-attack-dirbuster/"
End If
If (List23.Text = "Web hacking XSS injection " Or List23.Text = "web hacking xss injection") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guide-of-xss-injection-part1/"
End If
If (List23.Text = "Webcrab formgrabber " Or List23.Text = "webcrab formgrabber") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/webcrab-underground-formgrabber/"
End If
If (List23.Text = "Zeus complete tutorial" Or List23.Text = "zeus complete tutorial") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/zeus-complete-tutorial/"
End If
End Sub

Private Sub PushButton8_Click()
Shell App.path & "\programs\autoupdate.exe", vbNormalFocus
End Sub

Private Sub PushButton9_Click()
Shell App.path & "\programs\autoupdate.exe", vbNormalFocus
End Sub

Private Sub Text9_Click()
Text9.Text = ""
End Sub

