VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.2#0"; "CODEJO~4.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.2#0"; "CO85CC~1.OCX"
Begin VB.Form frmDownLoad 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000015&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDownLoad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   1140
      Picture         =   "frmDownLoad.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   6045
      TabIndex        =   10
      Top             =   840
      Width           =   6045
   End
   Begin VB.Timer Timer2 
      Interval        =   700
      Left            =   8160
      Top             =   3300
   End
   Begin VB.PictureBox picCheckingForUpdates 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   870
      ScaleHeight     =   645
      ScaleWidth      =   4335
      TabIndex        =   8
      Top             =   1290
      Width           =   4335
      Begin VB.Label lblCheckingForUpdates 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checking for updates please wait...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   9
         Top             =   180
         Width           =   3405
      End
   End
   Begin VB.Timer Timer1 
      Left            =   8160
      Top             =   3690
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdCheckForUpdates 
      Height          =   1095
      Left            =   1800
      TabIndex        =   11
      ToolTipText     =   "Check for updates"
      Top             =   3765
      Width           =   1575
      _Version        =   983042
      _ExtentX        =   2778
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Updates"
      ForeColor       =   -2147483635
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
      Icon            =   "frmDownLoad.frx":869E
   End
   Begin XtremeSuiteControls.PushButton cmdCancel 
      Height          =   1095
      Left            =   3480
      TabIndex        =   12
      ToolTipText     =   "Cancel Update progress"
      Top             =   3765
      Width           =   1575
      _Version        =   983042
      _ExtentX        =   2778
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Cancel"
      ForeColor       =   -2147483635
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
      Icon            =   "frmDownLoad.frx":AB08
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Visit CasperLauncher Overview to find more updates"
      Top             =   3765
      Width           =   1575
      _Version        =   983042
      _ExtentX        =   2778
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Help"
      ForeColor       =   -2147483635
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
      Icon            =   "frmDownLoad.frx":CF72
   End
   Begin XtremeSuiteControls.PushButton lblDownLoadMinor 
      Height          =   1095
      Left            =   5160
      TabIndex        =   14
      ToolTipText     =   "When Update found Download it!"
      Top             =   3765
      Width           =   1590
      _Version        =   983042
      _ExtentX        =   2813
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Download"
      ForeColor       =   -2147483635
      BackColor       =   -2147483632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      Appearance      =   2
      TextImageRelation=   1
      IconWidth       =   48
      Icon            =   "frmDownLoad.frx":F3DC
   End
   Begin XtremeSuiteControls.PushButton cmdExit 
      Height          =   1095
      Left            =   6840
      TabIndex        =   15
      Top             =   3765
      Width           =   1575
      _Version        =   983042
      _ExtentX        =   2778
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Exit"
      ForeColor       =   -2147483635
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
      Icon            =   "frmDownLoad.frx":11846
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   3600
      Top             =   2160
      _Version        =   983042
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblDownLoadMajor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Download Major Update (Not Free)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3300
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2040
      Width           =   3285
   End
   Begin VB.Label lblReadMajorFeatures 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Read Deatails Online"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2040
      Width           =   2010
   End
   Begin VB.Label lblMajorVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   720
      TabIndex        =   5
      Top             =   1740
      Width           =   180
   End
   Begin VB.Label lblMinorVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xx"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1110
      TabIndex        =   4
      Top             =   3030
      Width           =   330
   End
   Begin VB.Label lblDownLoadProgress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   810
      TabIndex        =   3
      Top             =   3090
      Width           =   180
   End
   Begin VB.Image Image2 
      Height          =   990
      Left            =   7530
      Picture         =   "frmDownLoad.frx":13CB0
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading..."
      Height          =   195
      Left            =   6000
      TabIndex        =   2
      Top             =   540
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   7410
      Stretch         =   -1  'True
      Top             =   780
      Width           =   630
   End
   Begin VB.Label lblMessaage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "xx"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   300
   End
End
Attribute VB_Name = "frmDownLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmDownLoad
' DateTime  : 7/26/2007 14:34
' Author    : Robert Moore
' Purpose   :To update your existing exe with a new one from the internet without any external files
'This program contains the following
' 4 progress indicators eg: 1 progress bar colored fore and back colors,1 percent complete on form caption,
'1 expanding image control, 1 kb completed out total kb's
'1.ini class for handling ini files (think i got this here at planet)
'1. resource file that uses bmp's and custom jpg (bas module handles the jpg conversion)
'1 google search not visible on form load, use it if you want or not
'1 gradient class, unused in this project, set your colors in basGradient
'1 bas error handler
'the check for major updates is unfinished, finish you want or not

'---------------------------------------------------------------------------------------

'Important notes !!!!!!!!!!!!!!:
'Be sure program being updated is not dirty(work unsaved, etc) any unsaved work will be gone, including this file
'and be replaced with a new exe.
'Check list before using
'Files needed
'1.remote text file holding only version numbers, using this format xx.xx
'2. remote exe that is a newer version than this one
'3. path-name to tmp folder to hold downloaded exe
'4. this exe name with path
'5. optional backup exe name
'optional remote web page for minor upgrade information
'optional remote web page for major upgrade information



Option Explicit
Private Declare Function SendMessage Lib _
  "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long


Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)



Private clsINIFile As cInifile


Private lngConnectionID As Long

Private Const INTERNET_CONNECTION_CONFIGURED = &H40
Private Const INTERNET_CONNECTION_LAN = &H2
Private Const INTERNET_CONNECTION_MODEM = &H1
Private Const INTERNET_CONNECTION_OFFLINE = &H20
Private Const INTERNET_CONNECTION_PROXY = &H4
Private Const INTERNET_RAS_INSTALLED = &H10
Private Declare Function InternetGetConnectedState Lib "Wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

'these declare for overwriting file in use

Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As VbAppWinStyle) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3


   Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long


   Private Declare Function ReplaceFileW Lib "Kernel32.dll" (ByVal lpReplacedFileName As Long, ByVal lpReplacementFileName As Long, _
    ByVal lpBackupFileName As Long, ByVal dwReplaceFlags As Long, ByVal lpExclude As Long, ByVal lpReserved As Long) As Boolean
''''''''''end declare for overwriting file in use

   Dim IsCancelRequested As Boolean

   Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
   Private hFile As Integer


    Private WithEvents OpenPage As WinHttpRequest
Attribute OpenPage.VB_VarHelpID = -1
    Private WithEvents http As WinHttpRequest
Attribute http.VB_VarHelpID = -1
    Private WithEvents httpul As WinHttpRequest
Attribute httpul.VB_VarHelpID = -1
    Private WithEvents httCheckVersion As WinHttpRequest
Attribute httCheckVersion.VB_VarHelpID = -1
    
    Dim mFirstLoad As Boolean
    Dim mMinorUpGradeAvailable As Boolean
    Dim mMajorUpGradeAvailable As Boolean
    
    Dim mDateCheckedForUpdates As String
    Dim mDateMajorUpdated As String
    Dim mDateMinorUpdated As String
    Dim mDaysCheckedForUpdates As Integer
    Dim mDaysSinceLastMinorUpdate As Integer
    Dim mImage1Height As Long
    Dim mFormFirstOpen As Boolean 'only true on first load
    Dim mCanceledByUser As Boolean
    Dim mDownloadStarted As Boolean
    Dim mLocalExeToReplace As String
    Dim mRemoteTextFile As String
    Dim mRemoteFile As String
    Dim mTmpLocalFile As String
    Dim mLocalBackUpFileName As String
    Private mContentLength As Long
    Private mProgress As Long



Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
SkinFramework1.LoadSkin App.Path & "\casper\EX2008.cjstyles", ""
SkinFramework1.ApplyWindow Me.hWnd
'put our files names here, be sure remote filenames exist***********************************
'********************************************************************************************
       mLocalExeToReplace = App.Path & "\VBDownloader.exe"
       mTmpLocalFile = App.Path & "\tmp\VBDownloader.exe"
       mLocalBackUpFileName = App.Path & "\VBDownloader.backup"
       mRemoteFile = "http://casperspy.com/VBDownloader.exe"
       mRemoteTextFile = "http://casperspy.com/ver.txt"
'**********************************************************************************************
'**********************************************************************************************

    Call ControlState("FormLoad")
    End Sub


Sub ControlState(Status As String)
    Select Case Status
    
   Case "FormLoad"
       Set clsINIFile = New cInifile
       
      
       Call FormStartSettings 'get our .ini values
       
     mDaysSinceLastMinorUpdate = DateDiff("d", mDateMinorUpdated, Date)
    If mDaysSinceLastMinorUpdate > 30 Then '30 days
       MsgBox "You last updated your software , " & mDaysSinceLastMinorUpdate & " days ago"
     End If
       
       mDaysCheckedForUpdates = DateDiff("d", mDateCheckedForUpdates, Date)
    If mDaysCheckedForUpdates > 30 Then '30 days ' need to check this against mDaysSinceLastMinorUpdate, can't be more
       MsgBox "You last checked for updates, " & mDaysCheckedForUpdates & " days ago"
     End If
             
     If mFirstLoad = True Then 'get from ini file
        MsgBox "Very important!!Very important!! " & vbCrLf _
      & "If you have not already done so replace the file names on form load", vbCritical, "For Planet users"
      Unload Me
     End If
     
   Call UpDateFirstLoad 'only for planet users delete
      Me.picCheckingForUpdates.Visible = False
      Me.Timer2.Enabled = False
      mImage1Height = Image1.Height
      lblMessaage.Caption = "You are using " & App.Title & " Version" & App.Major & "." & App.Minor
      Timer1.Enabled = False
      Me.ProgressBar1.Visible = False
      Me.Image1.Visible = False
      Label1.Visible = False
      Me.lblDownLoadMinor.Enabled = False
      Me.lblMinorVersion.Visible = False
      Me.lblMajorVersion.Visible = False
      Me.lblReadMajorFeatures.Visible = False
      Me.lblDownLoadProgress.Visible = False
      Me.lblDownLoadMajor.Visible = False
      Me.Image1.Picture = LoadResPicture(110, vbResBitmap) 'progress picture
      
      'Set the ProgressBar Barcolor with green color
      SendMessage ProgressBar1.hWnd, PBM_SETBARCOLOR, 0, ByVal RGB(0, 252, 64)
    ' Set the ProgressBar Backcolor with yellow color
      SendMessage ProgressBar1.hWnd, PBM_SETBKCOLOR, 0, ByVal RGB(252, 234, 56)
       
       
       Me.lblReadMajorFeatures.MouseIcon = LoadResPicture(101, vbResCursor)
       Me.lblDownLoadMajor.MouseIcon = LoadResPicture(101, vbResCursor)

          '
  Case "NewMinorVersionAvailable"
       Me.lblMinorVersion.Enabled = True
      
       Me.cmdCheckForUpdates.Enabled = False
       Me.lblDownLoadMinor.Visible = True
       Me.picCheckingForUpdates.Visible = False
       Me.Timer2.Enabled = False
       
  Case "NewMajorVersionAvailable"
       Me.lblMajorVersion.Visible = True
       Me.lblReadMajorFeatures.Visible = True
       Me.cmdCheckForUpdates.Enabled = False
       Me.lblDownLoadMajor.Visible = True
    
  Case "DownLoadOnStart"
       Me.ProgressBar1.Visible = True
       Me.Image1.Visible = True
       Timer1.Enabled = True
       Me.lblDownLoadProgress.Visible = True
  
  Case "DownLoadInterrupted"
       Timer1.Enabled = False
       Me.ProgressBar1.Value = 0
       Me.ProgressBar1.Visible = False
       Me.Caption = "Program update incomplete"
    If mMinorUpGradeAvailable = True Then
       Me.lblMinorVersion.Visible = True
       Me.lblReadMajorFeatures.Visible = True
       Me.lblDownLoadMinor.Visible = True
   End If
   
  If mMajorUpGradeAvailable = True Then
     Me.lblReadMajorFeatures.Visible = True
     Me.lblDownLoadMajor.Visible = True
     Me.lblDownLoadProgress.Visible = False
     Me.Label1.Visible = False
     Me.Image1.Visible = False 'reset the size here
   End If
           
 End Select
        End Sub

Private Sub Form_Initialize()
    mFormFirstOpen = True 'check for internet connection on load
    Me.Caption = App.Title
    'GradObj Me.Pic 'make pic gradient
    End Sub

'connection has been lost during download, inform the user
Private Sub http_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
      Dim msg As String
        msg = "Your internet connection has been disconnected" & vbCrLf & _
         "Connect to the internet and try again"
         
        Call ControlState("DownLoadInterrupted")
      
      End Sub


Private Sub httpul_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String) 'check for connection failed
'If ErrorNumber = -2147012889 then "Check your internet connection, cables,etc (With dialup could be a lot of things)
'If ErrorNumber = -214701867 then "Your firewall is blocking your connection
'put connection error information in help file, firewalls, timeouts etc
     Dim msg As String
      If ErrorNumber = -214701867 Then
          msg = "Error number " & ErrorNumber & vbCrLf _
          & "Your firewall may be blocking this connection, make adjustments and try again"
     Else
            msg = "An internet connection could not be established, connect to the internet befoe proceeding"
     End If
     
         MsgBox msg, vbExclamation + vbOK, "Internet connection unavailable" 'add help button on final release
            
         httpul.Abort
        Set httpul = Nothing
       
       
    End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim response
    Dim msg As String
    msg = "Stop program update ?" & vbCrLf & vbCrLf & _
    "To resume update click Cancel" & vbCrLf & _
    "To continue without the program update click OK"
    
 If UnloadMode <> 1 Then 'skip if user clicked the cancel button
   If mDownloadStarted = True Then
     response = MsgBox(msg, vbOKCancel + vbQuestion, App.Title) 'add help button on final release
     On Error Resume Next
       Select Case response
        Case vbOK
          http.Abort
          Cancel = 0 ' 0 allows closing
          
        Case vbNo
          Exit Sub
        
        End Select
        
    End If
 End If
        End Sub

Private Sub http_OnResponseDataAvailable(Data() As Byte)
'  On Error GoTo e
  Call ControlState("DownLoadOnStart")
  Timer1.Interval = 600
  Call DownloadProgress(mProgress, mContentLength)
        mProgress = mProgress + UBound(Data) + 1
        ProgressBar1.Value = mProgress
'
      Me.Caption = EstimatedEndTime(mProgress, mContentLength) & "% Percent complete"

     
        Put #hFile, , Data
        Exit Sub
e:
    ErrHandler "Sub http_OnResponseDataAvailable ", Err.Number, Err.Description
    http.Abort
'    End
    End Sub

Private Sub http_OnResponseFinished()
'      On Error GoTo e
        Close #hFile
        mDownloadStarted = False
        MsgBox "Download Complete", vbInformation, App.Title
        http.Abort
        Set http = Nothing
        
        Me.ProgressBar1.Visible = False
        Timer1.Enabled = False
        Me.lblDownLoadProgress.Visible = False
      mDateMinorUpdated = Date
      Call DateMinorUpdated 'update ini file with date
        
       Call ReplaceExeInUse
bail:
        Exit Sub
'e:
'        ErrHandler "Sub http_OnResponseFinished ", Err.Number, Err.Description
        
    End Sub

Private Sub http_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
       On Error GoTo e
        hFile = FreeFile  'assign a file handle
    If http.Status <> 200 Then
        ' If it's not 200 we throw an error... we'll
        ' check for it and others later.
        Err.Raise 1, "HttpRequester", "Invalid HTTP Response Code"
     End If
                
        mProgress = 0
        mContentLength = CLng(http.GetResponseHeader("Content-Length"))
       
        ProgressBar1.Max = mContentLength
      Call ControlState("DownLoadOnStart")
         mDownloadStarted = True
          Open mTmpLocalFile For Binary As #hFile  'Open file for binary method
         Exit Sub
          
e:
        http.Abort
         ErrHandler "Sub http_OnResponseStart ", Err.Number, Err.Description
         
        End Sub

Function EstimatedEndTime(ProcessedCount As Long, TotalCount As Long)
        On Error GoTo e
        EstimatedEndTime = Round((ProcessedCount / TotalCount) * 100, 0)
        Exit Function
e:
        ErrHandler "Function EstimatedEndTime ", Err.Number, Err.Description
        End Function
Private Sub cmdCancel_Click() 'we need to check how far we have progreeded,if updates available and download started
   Dim response As Integer
   If mDownloadStarted = True Then 'we know check for updates is complete, if we get this far
            response = MsgBox("Stop program update ?", vbInformation + vbOKCancel) 'add help button on final release ?
          ControlState ("DownLoadInterrupted")
       If response = vbOK Then
                http.Abort
                Set http = Nothing
                mCanceledByUser = True
              Call ControlState("DownLoadInterrupted")
               mDownloadStarted = False
              GoTo bail
           Else
           Call ControlState("DownLoadInterrupted")
         End If
       GoTo bail
    End If
 'if we get this far then check if update has been done
    
    If mMinorUpGradeAvailable = False Then 'start fresh
         Call ControlState("FormLoad") 'reset everything
       Else
         Call ControlState("DownLoadInterrupted")
     End If
         
        Exit Sub
bail:
    
        End Sub
Private Sub cmdCheckForUpdates_Click() 'open txt file on the internet and compare to current version exe
' create an indicator here for slow pc's and slow connections to inform user of progress
    
    Call CheckConnection
      Set httCheckVersion = New WinHttpRequest
        Dim sDownloadURL As String
        httCheckVersion.SetTimeouts 1000, 1000, 1000, 1000
        sDownloadURL = mRemoteTextFile '"http://mysite.com/TextFiles/aa.txt"
        httCheckVersion.Open "GET", sDownloadURL, True  'True means asynch.
        httCheckVersion.Send
        'DoEvents
     
       mFirstLoad = False
        End Sub
Private Sub httCheckVersion_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
MsgBox "The online site could reached, Check the URL and spelling and try again", vbOKOnly + vbInformation

End Sub

Sub CheckConnection() 'check for a connection
    MousePointer = vbHourglass
    Set httpul = New WinHttpRequest
    httpul.SetTimeouts 10000, 10000, 10000, 10000
    httpul.Open "GET", "http://google.com", False

    On Error Resume Next
    httpul.Send
    
    MousePointer = vbDefault
    Exit Sub
    

End Sub


Private Sub httCheckVersion_OnResponseStart(ByVal Status As Long, ByVal ContentType As String)
     If Status <> 200 Then
     MsgBox "Remote file not found"
     On Error Resume Next
     httCheckVersion.Abort
     End If
    
    End Sub

Private Sub httCheckVersion_OnResponseDataAvailable(Data() As Byte)
'we are checking the app Major & Minor number
    Dim response As Integer
    Dim RemoteTextFileVersion As String
    Dim RemoteMajorVersion As String ' not free
    Dim RemoteMinorVersion As String ' free
    Dim AppMinorVersion As String
    Dim AppMajorVersion As String
   
       
        AppMajorVersion = App.Major 'get the current Major version
        AppMinorVersion = App.Minor 'get the current Minor version
        
        RemoteMajorVersion = Left(httCheckVersion.ResponseText, 2) 'this holds your remote files data
        RemoteMinorVersion = Right(httCheckVersion.ResponseText, 2) 'revision numbers are always xx.xx
        
      MousePointer = vbDefault
            'Remote textfile format xx.xx
  If Val(AppMinorVersion) < Val(RemoteMinorVersion) Then 'if minor file remote version higher
          Me.lblMinorVersion.Caption = App.Title & " Version " & AppMajorVersion & "." & RemoteMinorVersion & " is available !"
          mMinorUpGradeAvailable = True
         Call ControlState("NewMinorVersionAvailable")
     Else
           Me.lblMinorVersion.Caption = "You have the latest version, no minor updates currently available" 'add help button on final release ?
   End If
   
      'lets ckeck for major version (not Free)exe must have the same name, but be in a different remote folder and need new reg numbers
'  If Val(AppMajorVersion) < Val(RemoteMajorVersion) Then 'if major file remote version higher
'      Call ControlState("NewMajorVersionAvailable")
'      mMajorUpGradeAvailable = True
'      lblMajorVersion.Caption = App.Title & " Version " & RemoteMajorVersion & "." & RemoteMinorVersion & " is available !"
'
'   End If
'       mDateCheckedForUpdates = Date
'       Call DateCheckedForUpdates 'update our ini file
bail:
         Set httCheckVersion = Nothing
            
        End Sub


'Activate this up if we want to show bytes recieved and total bytes
Private Sub DownloadProgress(ByVal ReceivedBytes As Long, ByVal TotalBytes As Long)
    On Error GoTo e
    Me.lblDownLoadProgress.Caption = "Recieved " & Format$(ReceivedBytes / 1024, "0.00") & " Kb of " _
     & Format$(TotalBytes / 1024, "0.00") & " Kb"
   
bail:
    Exit Sub

e:
    ErrHandler "Sub DownloadProgress ", Err.Number, Err.Description
End Sub



Public Function ReplaceFile(ByVal lpReplacedFileName As String, ByVal lpReplacementFileName As String, _
    ByVal lpBackupFileName As String) As Boolean

  'If lpBackupFileName parameter is NULL,'no backup file is created.

  Const REPLACEFILE_WRITE_THROUGH& = &H1

  ReplaceFile = ReplaceFileW(StrPtr(lpReplacedFileName), StrPtr(lpReplacementFileName), StrPtr(lpBackupFileName), _
      REPLACEFILE_WRITE_THROUGH, ByVal 0&, ByVal 0&)

End Function

Private Sub ReplaceExeInUse()
   Dim result As Long
    
    mDateMinorUpdated = Date
   Call DateMinorUpdated 'update our .ini file
   
   Call ReplaceFile(mLocalExeToReplace, mTmpLocalFile, mLocalBackUpFileName)
    'Call ReplaceFile(App.Path & "\" & App.EXEName & ".exe", App.Path & "\tmp\" & App.EXEName & ".exe", App.Path & "\" & App.EXEName & ".backup")
    
    result = ShellExecute(Me.hWnd, "open", mLocalExeToReplace, "/updaterestart " & GetCurrentProcessId(), "", 1)
    
    End
End Sub


Private Sub lblDownLoadMinor_Click()
          On Error GoTo e
            ' Create the WinHTTPRequest ActiveX Object.
            Set http = New WinHttpRequest
            ' Open an HTTP connection.
            'Set the receive timeout to 3 seconds.
           http.SetTimeouts 60000, 60000, 60000, 60000
           http.Open "GET", mRemoteFile, True  'True means asynch.'
          
           On Error Resume Next 'force overwrite if \tmp already existing
           MkDir App.Path & "\tmp"
           On Error GoTo e
           http.Send
            Me.ProgressBar1.Visible = True
        Exit Sub
e:
      http.Abort
        ErrHandler "cmdDownLoad_Click ", Err.Number, Err.Description
         
End Sub

Private Sub lblReadMinorFeatures_Click() 'open webpage to view new features
    ShellExecute ByVal 0&, "open", "http://yourwebsite.com", vbNullString, vbNullString, SW_SHOWMAXIMIZED
    End Sub
Private Sub lblReadMajorFeatures_Click()
ShellExecute ByVal 0&, "open", "http://google.com", vbNullString, vbNullString, SW_SHOWMAXIMIZED
End Sub


Private Sub lblSearch_Click()
Call SearchGoogle(txtGoogleSearch)
End Sub

Private Sub Timer1_Timer()
    Dim maxHeight As Long
      maxHeight = 2000
    
 If Me.Image1.Height > maxHeight Then
      Me.Image1.Height = mImage1Height
      Me.Label1.Visible = False
  Else
    Me.Image1.Height = Me.Image1.Height + 100
    Me.Label1.Visible = True
 End If
End Sub


Private Sub SearchGoogle(strSearchString As String)
    If InStr(1, strSearchString, " ") Then
        strSearchString = Replace(strSearchString, " ", "+")
    End If

    Call ShellExecute(Me.hWnd, "Open", "www.google.com/search?hl=en&q=" & strSearchString & "&meta=", _
                      vbNullString, vbNullString, vbNormal)
                     

End Sub

Private Sub Timer2_Timer() 'run when checking for connection and updates on a fast pc with dsl will hardly be seen
  If Me.lblCheckingForUpdates.Visible = True Then
    Me.lblCheckingForUpdates.Visible = False
   Else
    Me.lblCheckingForUpdates.Visible = True
  End If
    End Sub
    
  Sub FormStartSettings()
    Dim MyPath As String
    MyPath = App.Path

         MyPath = App.Path & "\VbDownloader.ini"
         
          With clsINIFile
           .Path = MyPath

          .Section = "Dates"
          .Key = "LastChecked"
          .Default = Date
           mDateCheckedForUpdates = .Value

          .Section = "Dates"
          .Key = "Minor"
          .Default = Date     '
           mDateMinorUpdated = .Value

          .Section = "Dates"
          .Key = "Major"
          .Default = Date
           mDateMajorUpdated = .Value
           
           .Section = "StartSettings"
           .Key = "FirstLoad"
           .Default = True
           mFirstLoad = .Value
           
      End With
    End Sub
    Sub DateCheckedForUpdates()
    Dim MyPath As String
    MyPath = App.Path
    MyPath = MyPath & "\VbDownloader.ini"
       With clsINIFile
        .Path = MyPath
        .Section = "Dates"
        .Key = "LastChecked"
        .Default = "0"       '
        .Value = mDateCheckedForUpdates
        End With
        End Sub
   Sub DateMinorUpdated()
    Dim MyPath As String
    MyPath = App.Path
    MyPath = MyPath & "\VbDownloader.ini"
       With clsINIFile
        .Path = MyPath
        .Section = "Dates"
        .Key = "Minor"
        .Default = "0"       '
         .Value = mDateMinorUpdated
    End With
    End Sub
Sub DateMajorUpdated()
    Dim MyPath As String
    MyPath = App.Path
    MyPath = MyPath & "\VbDownloader.ini"
       With clsINIFile
        .Section = "Dates"
        .Key = "Major"
        .Default = "0"
        .Value = mDateMajorUpdated
       End With
    End Sub
    
Sub UpDateFirstLoad() 'delete this
Dim MyPath As String
    MyPath = App.Path
    MyPath = MyPath & "\VbDownloader.ini"
       With clsINIFile
        .Section = "StartSettings"
        .Key = "FirstLoad"
        .Default = True
        .Value = False
       End With


End Sub


Private Sub cmdExit_Click()
 Dim response As Integer
   If mDownloadStarted = True Then 'we know check for updates is complete, if we get this far
            response = MsgBox("Stop program update ?", vbInformation + vbOKCancel) 'add help button on final release ?
          ControlState ("DownLoadInterrupted")
          On Error Resume Next
       If response = vbOK Then
                http.Abort
                Set http = Nothing
                mCanceledByUser = True
              GoTo bail
           Else
           Call ControlState("DownLoadInterrupted")
           Exit Sub
         End If
    End If
 
    
    If mMinorUpGradeAvailable = True Then 'start fresh
         mDateCheckedForUpdates = Date
        Call DateCheckedForUpdates 'update the ini file
      
     End If
         
        
bail:
    Unload Me
End Sub










