VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.2#0"; "CODEJO~4.OCX"
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleMode       =   0  'User
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4200
      Width           =   9330
      _Version        =   983042
      _ExtentX        =   16457
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   0
      Appearance      =   6
      UseVisualStyle  =   0   'False
      FlatStyle       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   0
      Picture         =   "Splash.frx":0000
      ScaleHeight     =   5865
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   720
         Top             =   0
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1200
         Top             =   0
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1680
         Top             =   0
      End
      Begin VB.Timer Timer5 
         Interval        =   100
         Left            =   2160
         Top             =   0
      End
      Begin VB.Timer Timer1 
         Interval        =   1500
         Left            =   240
         Top             =   0
      End
      Begin VB.Label LblStatus 
         BackColor       =   &H80000007&
         Caption         =   "Initializing..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   720
         TabIndex        =   3
         Top             =   4555
         Width           =   4335
      End
      Begin VB.Label LblAnim 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Caption         =   "[-]"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   4560
         Width           =   495
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
    ProgressBar1.Value = ProgressBar1.Value + 10
    ProgressBar1.Value = ProgressBar1.Value + 10
    LblStatus.Caption = "Initializing... Checking Attacker.."
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
    ProgressBar1.Value = ProgressBar1.Value + 30
    LblStatus.Caption = "Initializing... Checking Guide..."
    Timer2.Enabled = False
    Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
ProgressBar1.Value = ProgressBar1.Value + 50
    Me.MousePointer = 11
    On Error GoTo Koneksidatabase
    'Call Buka
    
    If ProgressBar1.Value = ProgressBar1.Max Then
    LblStatus.Caption = "Access Granted..."
    Timer3.Enabled = False
    Timer4.Enabled = True
    End If
    Exit Sub
Koneksidatabase:
        LblStatus.Caption = "Connection Failed; Check your internet connection..."
        MsgBox "Checking Internet Connection..." & vbCrLf & "" _
             & "Visite official blog www.casperspy.com ", vbCritical, "Connection Database Error..."
        Set Splash = Nothing
        End
End Sub

Private Sub Timer4_Timer()
    Timer4.Interval = 10
    Timer4.Enabled = True
    LblStatus.Caption = "Exit..."
    Unload Me
    frmMain.Show
End Sub

Private Sub AnimateText()
On Error Resume Next
If LblAnim.Caption = "[-]" Then
   LblAnim.Caption = "[\]"
ElseIf LblAnim.Caption = "[\]" Then
   LblAnim.Caption = "[|]"
ElseIf LblAnim.Caption = "[|]" Then
   LblAnim.Caption = "[/]"
ElseIf LblAnim.Caption = "[/]" Then
   LblAnim.Caption = "[-]"
End If
End Sub

Private Sub Timer5_Timer()
    AnimateText
End Sub


