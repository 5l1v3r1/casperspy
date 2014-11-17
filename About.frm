VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.2#0"; "CODEJO~4.OCX"
Begin VB.Form Intro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "About"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9240
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
      _Version        =   983042
      _ExtentX        =   2566
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "Like"
      ForeColor       =   -2147483634
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   2
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "About.frx":058A
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
      _Version        =   983042
      _ExtentX        =   2566
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "Follow "
      ForeColor       =   -2147483634
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   2
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "About.frx":29F4
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1455
      _Version        =   983042
      _ExtentX        =   2566
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "Fanpage"
      ForeColor       =   -2147483637
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   2
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "About.frx":4E5E
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      Picture         =   "About.frx":72C8
      ScaleHeight     =   5745
      ScaleWidth      =   9225
      TabIndex        =   0
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub PushButton1_Click()
frmmain2.Show
frmmain2.wb1.Navigate "http://m.facebook.com/casperspyofficial"
End Sub

Private Sub PushButton2_Click()
frmmain2.Show
frmmain2.wb1.Navigate "http://mobile.twitter.com/casperspy"
End Sub

Private Sub PushButton3_Click()
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/support.html"
End Sub
