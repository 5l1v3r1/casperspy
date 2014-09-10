VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.2#0"; "CODEJO~4.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.2#0"; "CODEJO~2.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.2#0"; "CO85CC~1.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   Caption         =   " CasperLauncher | Full Version [Shadow Striker]"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   Icon            =   "attacker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "CasperLauncher | Most Complete Attacker"
      BeginProperty Font 
         Name            =   "HelveticaNeueLT Std Blk Ext"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10935
      Begin VB.ListBox List1 
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
         Height          =   5190
         ItemData        =   "attacker.frx":1CCA
         Left            =   240
         List            =   "attacker.frx":1D7F
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "attacker.frx":1FC8
         Top             =   1680
         Width           =   3855
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   1095
         Left            =   7440
         TabIndex        =   3
         ToolTipText     =   "Launch Attacker"
         Top             =   1680
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
         Icon            =   "attacker.frx":204E
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   1095
         Left            =   9120
         TabIndex        =   4
         Top             =   5925
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
         Icon            =   "attacker.frx":44B8
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   1095
         Left            =   9000
         TabIndex        =   5
         ToolTipText     =   "Find tutorial how to use the Attacker which you select"
         Top             =   1680
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
         Icon            =   "attacker.frx":6922
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   1095
         Left            =   7440
         TabIndex        =   6
         ToolTipText     =   "Find specific attacker"
         Top             =   5925
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
         Icon            =   "attacker.frx":8D8C
      End
      Begin XtremeSuiteControls.PushButton PushButton9 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
         _Version        =   983042
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   79
         BackColor       =   -2147483633
         Enabled         =   0   'False
         Transparent     =   -1  'True
         FlatStyle       =   -1  'True
         Appearance      =   2
         BorderGap       =   0
         IconWidth       =   16
         Icon            =   "attacker.frx":B1F6
      End
      Begin VB.TextBox m 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   2520
         TabIndex        =   1
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Text            =   "Find attacker here.."
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   8295
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   10935
         _Version        =   983042
         _ExtentX        =   19288
         _ExtentY        =   14631
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   8
         PaintManager.BoldSelected=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         ItemCount       =   9
         SelectedItem    =   4
         Item(0).Caption =   "Attacker"
         Item(0).ControlCount=   0
         Item(1).Caption =   "Botnet"
         Item(1).ControlCount=   0
         Item(2).Caption =   "Cracking"
         Item(2).ControlCount=   0
         Item(3).Caption =   "Exploit"
         Item(3).ControlCount=   0
         Item(4).Caption =   "Keylogger"
         Item(4).ControlCount=   0
         Item(5).Caption =   "Network"
         Item(5).ControlCount=   0
         Item(6).Caption =   "Social Engginering"
         Item(6).ControlCount=   0
         Item(7).Caption =   "Trojan"
         Item(7).ControlCount=   0
         Item(8).Caption =   "Web Hacking"
         Item(8).ControlCount=   0
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4560
      Top             =   3840
      _Version        =   983042
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   1320
      Top             =   120
      _Version        =   983042
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "attacker.frx":B660
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
      DesignerControlsData=   "attacker.frx":39D6A
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




Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   On Error Resume Next
      
    Select Case Control.Id
    
      'Attacker menubar
        Case ID_ATTACKER_ATTACK1:
        If Dir("C:\Program Files\CasperLauncher\programs\builder\CSBuilder.exe") <> "" Then
 Shell App.path & "\programs\builder\CSBuilder.exe", vbNormalFocus
 
Else
 
    ShellExecute 0, vbNullString, "http://casperspy.com/download/casperspy-builder/", vbNullString, vbNullString, vbNormalFocus

 
End If
        
        Case ID_ATTACKER_ATTACK2: frmmain2.Show
        frmmain2.wb1.Navigate "http://casperspy.com/dashboard/"
        
        Case ID_ATTACKER_ATTACK3: Dim intResponse As Integer
       
ShellExecute 0, vbNullString, "http://blog.casperspy.com/how-to-use/", vbNullString, vbNullString, vbNormalFocus

        
        Case ID_ATTACKER_AGOBOT: TabControl1.SelectedItem = 0
        List1.ListIndex = 0
                
        Case ID_ATTACKER_CASPER: TabControl1.SelectedItem = 0
        List1.ListIndex = 8
      
        Case ID_ATTACKER_CITADEL: TabControl1.SelectedItem = 0
        List1.ListIndex = 9
        
        Case ID_ATTACKER_CYTHOSIA: TabControl1.SelectedItem = 0
        List1.ListIndex = 12
        
        Case ID_ATTACKER_VERTEXNET: TabControl1.SelectedItem = 0
        List1.ListIndex = 52
        
        Case ID_ATTACKER_ZEUS: TabControl1.SelectedItem = 0
        List1.ListIndex = 58
        
        Case ID_ATTACKER_EXPLOIT1: TabControl1.SelectedItem = 0
        List1.ListIndex = 10
       
        Case ID_ATTACKER_EXPLOIT2: TabControl1.SelectedItem = 0
        List1.ListIndex = 23
        
        Case ID_ATTACKER_EXPLOIT3: TabControl1.SelectedItem = 0
        List1.ListIndex = 32
        
        Case ID_ATTACKER_EXPLOIT4: TabControl1.SelectedItem = 0
        List1.ListIndex = 34
        
        Case ID_ATTACKER_FORMGRABBER1: TabControl1.SelectedItem = 0
        List1.ListIndex = 33
        
        Case ID_ATTACKER_FORMGRABBER2: TabControl1.SelectedItem = 0
        List1.ListIndex = 50
       
        Case ID_ATTACKER_FORMGRABBER3: TabControl1.SelectedItem = 0
        List1.ListIndex = 54
        
        Case ID_ATTACKER_KEYLOGGER1: TabControl1.SelectedItem = 0
        List1.ListIndex = 2
        
        Case ID_ATTACKER_KEYLOGGER2: TabControl1.SelectedItem = 0
        List1.ListIndex = 30
     
        Case ID_ATTACKER_KEYLOGGER3: TabControl1.SelectedItem = 0
        List1.ListIndex = 38
        
        Case ID_ATTACKER_KEYLOGGER4: TabControl1.SelectedItem = 0
        List1.ListIndex = 52
        
        Case ID_ATTACKER_NETWORK1: TabControl1.SelectedItem = 0
        List1.ListIndex = 1
        
        Case ID_ATTACKER_NETWORK2: TabControl1.SelectedItem = 0
        List1.ListIndex = 15
        
        Case ID_ATTACKER_NETWORK3: TabControl1.SelectedItem = 0
        List1.ListIndex = 16
        
        Case ID_ATTACKER_NETWORK4: TabControl1.SelectedItem = 0
        List1.ListIndex = 19
        
        Case ID_ATTACKER_NETWORK5: TabControl1.SelectedItem = 0
        List1.ListIndex = 22
        
        Case ID_ATTACKER_NETWORK6: TabControl1.SelectedItem = 0
        List1.ListIndex = 25
        
        Case ID_ATTACKER_NETWORK7: TabControl1.SelectedItem = 0
        List1.ListIndex = 27
        
        Case ID_ATTACKER_NETWORK8: TabControl1.SelectedItem = 0
        List1.ListIndex = 31
        
        Case ID_ATTACKER_NETWORK9: TabControl1.SelectedItem = 0
        List1.ListIndex = 35
        
        
        Case ID_ATTACKER_NETWORK10: TabControl1.SelectedItem = 0
        List1.ListIndex = 37
      
        Case ID_ATTACKER_NETWORK11: TabControl1.SelectedItem = 0
        List1.ListIndex = 39
        
        Case ID_ATTACKER_NETWORK12: TabControl1.SelectedItem = 0
        List1.ListIndex = 41
        
        Case ID_ATTACKER_NETWORK13: TabControl1.SelectedItem = 0
        List1.ListIndex = 53
        
        Case ID_ATTACKER_NETWORK14: TabControl1.SelectedItem = 0
        List1.ListIndex = 55
        
        Case ID_ATTACKER_TROJAN1: TabControl1.SelectedItem = 0
        List1.ListIndex = 4
        
        Case ID_ATTACKER_TROJAN2: TabControl1.SelectedItem = 0
        List1.ListIndex = 5
        
        Case ID_ATTACKER_TROJAN3: TabControl1.SelectedItem = 0
        List1.ListIndex = 13
        
        Case ID_ATTACKER_TROJAN4: TabControl1.SelectedItem = 0
        List1.ListIndex = 36
        
        Case ID_ATTACKER_TROJAN5: TabControl1.SelectedItem = 0
        List1.ListIndex = 44
        
        Case ID_ATTACKER_TROJAN6: TabControl1.SelectedItem = 0
        List1.ListIndex = 49
        
        Case ID_ATTACKER_TROJAN7: TabControl1.SelectedItem = 0
        List1.ListIndex = 48
        
        Case ID_ATTACKER_WEBHACKING1: TabControl1.SelectedItem = 0
        List1.ListIndex = 6
         
        Case ID_ATTACKER_WEBHACKING2: TabControl1.SelectedItem = 0
        List1.ListIndex = 14
        
        Case ID_ATTACKER_WEBHACKING3: TabControl1.SelectedItem = 0
        List1.ListIndex = 20
        
        Case ID_ATTACKER_WEBHACKING4: TabControl1.SelectedItem = 0
        List1.ListIndex = 26
      
        Case ID_ATTACKER_WEBHACKING5: TabControl1.SelectedItem = 0
        List1.ListIndex = 29
        
        Case ID_ATTACKER_WEBHACKING6: TabControl1.SelectedItem = 0
        List1.ListIndex = 45
        
        Case ID_ATTACKER_WEBHACKING7: TabControl1.SelectedItem = 0
        List1.ListIndex = 56
    
        Case ID_ATTACKER_SOCIAL1: TabControl1.SelectedItem = 0
        List1.ListIndex = 3
        
        Case ID_ATTACKER_SOCIAL2: TabControl1.SelectedItem = 0
        List1.ListIndex = 21
        
        Case ID_ATTACKER_SOCIAL3: TabControl1.SelectedItem = 0
        List1.ListIndex = 28
        
        Case ID_ATTACKER_SOCIAL4: TabControl1.SelectedItem = 0
        List1.ListIndex = 47
        
        Case ID_ATTACKER_SOCIAL5: TabControl1.SelectedItem = 0
        List1.ListIndex = 40
        
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
        Case ID_FILE_UPDATE: Shell App.path & "\programs\autoupdate.exe", vbNormalFocus
        
        'ToolBar
        Case ID_FILE_NEWS: frmmain2.Show
        frmmain2.wb1.Navigate "http://feeds.feedburner.com/casperspybotnet"
                   
        Case ID_ACTIONS_CHAT: frmmain2.Show
        frmmain2.wb1.Navigate "http://form.jotform.me/form/31453837808461"
        
        Case ID_ACTIONS_CONTACT: frmmain2.Show
        frmmain2.wb1.Navigate "http://casperspy.com/botmaster/"
        
        Case ID_FILE_FORUM: frmmain2.Show
        frmmain2.wb1.Navigate "http://casperspy.com/forum/"
        
        Case ID_FILE_TUTORIAL: Me.Hide
        guide.Show
        guide.TabControl1.SelectedItem = 3
        
        Case ID_FILE_GUIDE: Unload Me
        Load guide
        guide.Show
        Case ID_FILE_EBOOK: Me.Hide
        guide.Show
        guide.TabControl1.SelectedItem = 5
             
        Case ID_FILE_CASPERSPY: Me.Hide
        frmMain.Show
        frmMain.TabControl1.SelectedItem = 0
               
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
TabControl1.SelectedItem = 0
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
        Set frmMain.CommandBars.Icons = frmMain.ImageManager.Icons
        frmMain.CommandBars.Options.SetIconSize True, 42, 35
        frmMain.CommandBars.RecalcLayout
        End With
    
      
    Set ControlView = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Guide")
    With ControlView.CommandBar
        AddControl .Controls, xtpControlButton, ID_GUIDE_SYSTEM, "&System Attack"
        AddControl .Controls, xtpControlButton, ID_GUIDE_NETWORK, "&Network Attack"
        AddControl .Controls, xtpControlButton, ID_GUIDE_WEB, "&Web Attack"
        AddControl .Controls, xtpControlButton, ID_GUIDE_TUTORIAL, "&Tutorial"
        AddControl .Controls, xtpControlButton, ID_GUIDE_VIDEO, "&Video Tutorial"
        AddControl .Controls, xtpControlButton, ID_GUIDE_EBOOK, "&Ebook"

        Set frmMain.CommandBars.Icons = frmMain.ImageManager.Icons
        frmMain.CommandBars.Options.SetIconSize True, 32, 32
        frmMain.CommandBars.RecalcLayout
    End With
        
    Set ControlHelp = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Help")
    With ControlHelp.CommandBar.Controls
        .Add xtpControlButton, ID_HELP_ABOUT, "&About"
        .Add xtpControlButton, ID_FILE_UPDATE, "&Updates Contents"
        Set frmMain.CommandBars.Icons = frmMain.ImageManager.Icons
        frmMain.CommandBars.Options.SetIconSize True, 32, 32
        frmMain.CommandBars.RecalcLayout
    End With
    
    Dim ToolBar As CommandBar
    
    Set ToolBar = CommandBars.Add("Alpha", xtpBarTop)
    
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_CASPERSPY, "&Attacker", False, "Show Attacker Menu"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_GUIDE, "&Guide", False, "Show Guide Menu"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_EBOOK, "&Ebook", False, "Download Exclusive Ebook"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_TUTORIAL, "&Tutorial", False, "See manual instruction how to use all Attacker"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_VIDEO, "Video Tutorial", False, "Watch this and you'll know"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_NEWS, "&News", False, "News Update about CyberSecurity"
    AddControl ToolBar.Controls, xtpControlButton, ID_FILE_FORUM, "Forum", False, "Visite CasperSpy Forum"
    AddControl ToolBar.Controls, xtpControlButton, ID_ACTIONS_CHAT, "Chat", False, "Chat with Us"
    AddControl ToolBar.Controls, xtpControlButton, ID_ACTIONS_CONTACT, "&Cpanel", False, "Access Botmaster Control Panel"
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
If List1.Text = "Cybergate" Or List1.Text = "cybergate" Then
Text1.Text = "CyberGate - Rat" & vbCrLf & " " & vbCrLf & "CyberGate was built to be a tool for various possible applications, ranging from assisting Users with routine maintenance tasks, to remotely monitoring your Children, captures regular user activities and maintain a backup of your typed data automatically. It can also be used as a monitoring device for detecting unauthorized access."
End If
If List1.Text = "Subseven" Or List1.Text = "subseven" Then
Text1.Text = "SubSeven - Remote Administration Trojan" & vbCrLf & " " & vbCrLf & "SubSeven 2.3 is a simple, easy to use remote administration tool (RAT) designed to work on all current Windows platforms, both 32bit and 64bit. This tool is aimed at people who want that little  bit more power and control over remote computer management. Please use this tool responsibly and read and accept the disclaimer prior to use. If you do not agree with the disclaimer, please do not use the tool. You accept full liability and responsibility for your actions when using SubSeven. Do not use this tool on computers you are not authorized to control. U can download it from the clicking the download  now button."
End If
If List1.Text = "SoftIce" Or List1.Text = "softice" Then
Text1.Text = "SoftIce | W32Dasm - Windows Disassembler | Hacker's View (HIEW) | Cracking toolkit" & vbCrLf & " " & vbCrLf & "SoftICE is a kernel mode debugger for Microsoft Windows. Crucially, it is designed to run underneath Windows such that the operating system is unaware of its presence. Unlike an application debugger, SoftICE is capable of suspending all operations in Windows when instructed. For driver debugging this is critical due to how hardware is accessed and the kernel of the operating system functions. Because of its low-level capabilities, SoftICE is also popular as a software cracking tool."
End If
If List1.Text = "ResHack" Or List1.Text = "reshack" Then
Text1.Text = "Resource Hacker" & vbCrLf & " " & vbCrLf & "Resource HackerTM is a freeware utility to view, modify, rename, add, delete and extract resources in 32bit & 64bit Windows executables and resource files (*.res). It incorporates an internal resource script compiler and decompiler and works on all (Win95 - Win7) Windows operating systems."
End If
If List1.Text = "PeExplorer" Or List1.Text = "peexplorer" Then
Text1.Text = "View, Edit, and Reverse Engineer EXE and DLL Files." & vbCrLf & " " & vbCrLf & "PE Explorer is the most feature-packed program for inspecting the inner workings of your own software, and more importantly, third party Windows applications and libraries for which you do not have source code."
End If
If List1.Text = "OllyDbg" Or List1.Text = "ollydbg" Then
Text1.Text = "Olly Debugger" & vbCrLf & " " & vbCrLf & "OllyDbg is a 32-bit assembler level analysing debugger for Microsoft® Windows®. Emphasis on binary code analysis makes it particularly useful in cases where source is unavailable."
End If
If List1.Text = "IdaPro" Or List1.Text = "idapro" Then
Text1.Text = "Interactive DissassAsembler" & vbCrLf & " " & vbCrLf & "IDA is a Windows, Linux or Mac OS X hosted multi-processor disassembler and debugger that offers so many features it is hard to describe them all. Just grab an evaluation version if you want a test drive. An executive summary is provided for the non-technical user."
End If
If List1.Text = "Hiew32" Or List1.Text = "hiew32" Then
Text1.Text = "Hiew32 - Old School Cracking Tools" & vbCrLf & " " & vbCrLf & "view and edit files of any length in text, hex, and decode modes, x86-64 disassembler & assembler (AVX instructions include), physical & logical drive view & edit, support for NE, LE, LX, PE/PE32+ and little-endian ELF/ELF64 executable formats, support for Netware Loadable Modules like NLM, DSK, LAN, etc"
End If
If List1.Text = "Hexedit" Or List1.Text = "hexedit" Then
Text1.Text = "Hexadecimal Editor" & vbCrLf & " " & vbCrLf & "With a hex editor, a user can see or edit the raw and exact contents of a file, as opposed to the interpretation of the same content that other, higher level application software may associate with the file format. For example, this could be raw image data, in contrast to the way image editing software would interpret and show the same file."
End If
If List1.Text = "RxBot" Or List1.Text = "rxbot" Then
Text1.Text = "RxBot IRC Botnet" & vbCrLf & " " & vbCrLf & "Rxbot” (also known as “rBot”) is a win32 computer IRC worm (written in C++) that spreads to computers running Windows XP"
End If
If List1.Text = "Crime Pack" Or List1.Text = "crimepack" Then
Text1.Text = "CrimePack - Exploit Kit" & vbCrLf & " " & vbCrLf & "is a malicious code present on fraudulent websites or illegally injected on legitimate but hacked websites without the knowledge of the administrator."
End If
If List1.Text = "Picebot" Or List1.Text = "picebot" Then
Text1.Text = "Picebot Pharming Botnet" & vbCrLf & " " & vbCrLf & "Pharming is a cyber attack intended to redirect a website's traffic to another, bogus site. Pharming can be conducted either by changing the hosts file on a victim's computer or by exploitation of a vulnerability in DNS server software."
End If
If List1.Text = "Andromeda" Or List1.Text = "andromeda" Then
Text1.Text = "Andromeda Botnet" & vbCrLf & " " & vbCrLf & "Botnet with Flexible modules bot. Based on this product, you can build a botnet with extremely diverse opportunities. Bot extended functions with the help of the plug-in can be loaded right quantity and at any time."
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
Text1.Text = "Darkmoon - RAT" & vbCrLf & " " & vbCrLf & "Darkmoon is a Trojan horse that opens a back door on the compromised computer and has keylogging capabilities."
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
Text1.Text = "Rapzo Keylogger Private Edition" & vbCrLf & " " & vbCrLf & "Rapzo Keylogger, Hack any account with this keylogger Rapzo keylogger Private Edition it's amazing stealer with full feature such as: Stealers All Stealers Pure Code - No Drops + Runtime FUD [#] Firefox 3.5.0-3.6.X [#] DynDns [#] FileZilla [#] Pidgin [#] Imvu [#] No - Ip * Full UAC Bypass & Faster Execution Extentions [ . exe | .scr | .pif | .com ] * Keylogger support - Smtp[Gmail,Hotmail,live,aol,] * Screen Logger * Cure.exe to remove server from your Computer * Usb Spread * File pumper - Built-in * Icon Changer."
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
Text1.Text = "Automatic SQL injection and database takeover tool" & vbCrLf & " " & vbCrLf & "sqlmap is an open source penetration testing tool that automates the process of detecting and exploiting SQL injection flaws and taking over of database servers. Many niche features for the ultimate penetration tester and a broad range of switches lasting from database fingerprinting, over data fetching from the database, to accessing the underlying file system and executing commands on the operating system via out-of-band connections."
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
Text1.Text = "Passive DNS network mapper a.k.a. subdomains bruteforcer" & vbCrLf & " " & vbCrLf & "Dnsmap is mainly meant to be used by pentesters during the information gathering/enumeration phase of infrastructure security assessments. Pentester would typically discover the target company's IP netblocks, domain names, phone numbers, Subdomain brute-forcing is another technique that should be used in the enumeration stage."
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
Text1.Text = "CasperSpy Botnet 2.1 - Botnet for CyberSecurity Defense " & vbCrLf & " " & vbCrLf & "Casper Spy is a new generation of zeus botnet created by modifying Zeus source code with add new capabilities: pollymorphic infection,spread itself and form grabbing for chrome browser to make it more dangerous and more difficult to detect"
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

Private Sub PushButton1_Click()
If (List1.Text = "CasperLogger" Or List1.Text = "casperlogger") Then
Shell App.path & "\programs\Keylogger\casperlogger.exe", vbNormalFocus
End If
If (List1.Text = "Crime Pack" Or List1.Text = "crimepack") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/crimepack-exploit-kit.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "OllyDbg" Or List1.Text = "ollydbg") Then
ShellExecute 0, vbNullString, "https://www.dropbox.com/s/f2qnf82wur8gvnw/ollydbg2.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "SoftIce" Or List1.Text = "softice") Then
ShellExecute 0, vbNullString, "https://www.dropbox.com/s/5l5n1o0fbcz4rfd/hiew_softice_w32dasm_crack_toolkit.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Hiew32" Or List1.Text = "hiew32") Then
ShellExecute 0, vbNullString, "https://www.dropbox.com/s/5l5n1o0fbcz4rfd/hiew_softice_w32dasm_crack_toolkit.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Subseven" Or List1.Text = "subseven") Then
ShellExecute 0, vbNullString, "https://www.dropbox.com/s/ilvbgyymuxrfeu1/subseven.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Cybergate" Or List1.Text = "cybergate") Then
ShellExecute 0, vbNullString, "https://www.dropbox.com/s/l7ki5p79vu1kdn8/cybergate.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "IdaPro" Or List1.Text = "idapro") Then
ShellExecute 0, vbNullString, "https://www.hex-rays.com/products/ida/support/download_freeware.shtml", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "ResHack" Or List1.Text = "reshack") Then
ShellExecute 0, vbNullString, "http://www.angusj.com/resourcehacker/", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "PeExplorer" Or List1.Text = "peexplorer") Then
ShellExecute 0, vbNullString, "http://www.heaventools.com/overview.htm", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "RxBot" Or List1.Text = "rxbot") Then
ShellExecute 0, vbNullString, "ftp://caspersp@casperspy.com/public_html/tools/rxbot_digital_spreader.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Picebot" Or List1.Text = "picebot") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/picebot.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Andromeda" Or List1.Text = "andromeda") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/andromeda-botnet.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Genpmk" Or List1.Text = "genpmk") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/cowpatty-genpmk-4.0-win32.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Stuxnet" Or List1.Text = "stuxnet") Then
ShellExecute 0, vbNullString, "http://www.megapanzer.com/wp-content/uploads/Stuxnet.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "VertexNet" Or List1.Text = "vertexnet") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/VertexNet-sillhouette.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Crypter Pack FUD" Or List1.Text = "crypter pack fud") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/crypters-pack.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "TW-Grabber" Or List1.Text = "tw-grabber") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/formgrabber/TW-Grabber.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Vulcan" Or List1.Text = "vulcan") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/keylogger/vulcan.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Agobot" Or List1.Text = "agobot") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/agobot3.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Citadel" Or List1.Text = "citadel") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/citadel-sillhouette.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Cythosia" Or List1.Text = "cythosia") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/cythosia-sillhouette.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Crypter.py" Or List1.Text = "crypter.py") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/exploit/crypter.py.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Nmap" Or List1.Text = "nmap") Then
ShellExecute 0, vbNullString, "http://nmap.org/download.html", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "LBD" Or List1.Text = "lbd") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/lbd-ge.mine.nu-master.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Evilpdf exploit" Or List1.Text = "evilpdf exploit") Then
ShellExecute 0, vbNullString, "http://www.rapid7.com/products/metasploit/download.jsp", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Dmitry" Or List1.Text = "dmitry") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/DMitry-1.3a.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Metasploit" Or List1.Text = "metasploit") Then
ShellExecute 0, vbNullString, "http://www.rapid7.com/products/metasploit/download.jsp", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Msfpayload" Or List1.Text = "msfpayload") Then
ShellExecute 0, vbNullString, "http://www.rapid7.com/products/metasploit/download.jsp", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "The Harvester" Or List1.Text = "the harvester") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/theHarvester-2.1_BH2011_Arsenal.tar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Trojan Creator" Or List1.Text = "trojan creator") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/astgen11.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Keymail" Or List1.Text = "keymail") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/keylogger/keymail.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Dirbuster" Or List1.Text = "dirbuster") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/DirBuster-0.12-Setup.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Aircrack" Or List1.Text = "aircrack") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/aircrack-airmon-ng-1.2-win32.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Burpsuite" Or List1.Text = "burpsuite") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/burpsuite_free_v1.5.jar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Dnsmap" Or List1.Text = "dnsmap") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/dnsmap-0.30.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Driftnet" Or List1.Text = "driftnet") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/driftnet-0.1.6.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Ettercap" Or List1.Text = "ettercap") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/ettercap-A%20suite%20for%20man%20in%20the%20middle%20attacks.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Rapzo" Or List1.Text = "rapzo") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/keylogger/rapzo.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Autoinfector" Or List1.Text = "autoinfector") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/setoolkit-3.5.1.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Albertino" Or List1.Text = "albertino") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/keylogger/albertino.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Gerix" Or List1.Text = "gerix") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/gerix-wifi-cracker-master.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Reaver" Or List1.Text = "reaver") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/reaver-1.4.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Wash" Or List1.Text = "wash") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/reaver-1.4.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Wifite" Or List1.Text = "wifite") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/wifite-2.0r85.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Pyrit" Or List1.Text = "pyrit") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/pyrit-0.4.0-WPAWPA2-PSK-key-cracker.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Joomscan" Or List1.Text = "joomscan") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/joomscan-OWASP%20Joomla%20Vulnerability%20Scanner.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Slowloris" Or List1.Text = "slowloris") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/Slowloris_Gui-release1.0_SRC-httpDDOS.7z", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Bifrost" Or List1.Text = "bifrost") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/Bifrost%201.2.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Bionet" Or List1.Text = "bionet") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/bionet%20rootkit%20by%20fabio%20grunge.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Turkojan" Or List1.Text = "turkojan") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/turkojan%204.0.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Prorat" Or List1.Text = "prorat") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/TROJAN%20PRORAT.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Spynet" Or List1.Text = "spynet") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/Spy_Net.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Darkmoon" Or List1.Text = "darkmoon") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Trojan/Dark%20Moon.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "SET" Or List1.Text = "set") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/setoolkit-3.5.1.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Jigsaw" Or List1.Text = "jigsaw") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/jigsaw-master.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Edjpgcom" Or List1.Text = "edjpgcom") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Social%20Engginering/edjpgcom-binder.rar", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "DVWA" Or List1.Text = "dvwa") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Web%20Hacking/DVWA-1.0.8.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Havij Pro" Or List1.Text = "havij pro") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Web%20Hacking/Havij-PRO.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "SQLMAP" Or List1.Text = "sqlmap") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Web%20Hacking/sqlmapproject-sqlmap-0.9-3248-gdbb0d7f.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Wp-Scan" Or List1.Text = "wp-scan") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/Web%20Hacking/wpscanteam-wpscan-WordPress%20Security%20Scanner.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Dnsdict" Or List1.Text = "dnsdict") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/dns-analyze-tool.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Dnsrecon" Or List1.Text = "dnsrecon") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/dnsrecon-master.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Dnsmap" Or List1.Text = "dnsmap") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/network/dnsmap-0.30.tar.gz", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "CasperSpy" Or List1.Text = "casperspy") Then
ShellExecute 0, vbNullString, "http://casperspy.com/download/casperspy-builder/", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Spy Eye" Or List1.Text = "spy eye") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/spyeye-botpack.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Spitmo" Or List1.Text = "spitmo") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/spyeye_roid.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Zeus Botnet" Or List1.Text = "zeus botnet") Then
ShellExecute 0, vbNullString, "http://casperspy.com/tools/zeuss_botpack.zip", vbNullString, vbNullString, vbNormalFocus
End If
If (List1.Text = "Janidos" Or List1.Text = "janidos") Then
ShellExecute 0, vbNullString, "http://anggeraw.com/attacker/botnet/janidos.zip", vbNullString, vbNullString, vbNormalFocus
End If
End Sub

Private Sub PushButton2_Click()
Text10.Visible = True
m.Visible = True
PushButton9.Visible = True
End Sub
Private Sub PushButton3_Click()
If (List1.Text = "Gerix" Or List1.Text = "gerix") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/hack-wifi-crack-wep-with-gerix/"
End If
If (List1.Text = "Stuxnet" Or List1.Text = "stuxnet") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/stuxnet-timeline-most-insidious-cyber-weapon/"
End If
If (List1.Text = "Albertino" Or List1.Text = "albertino") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/tutorial-how-to-use-albertino-keylogger"
End If
If (List1.Text = "Wp-Scan" Or List1.Text = "wp-scan") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guide-hacking-wordpress-part1"
End If
If (List1.Text = "TW-Grabber" Or List1.Text = "tw-grabber") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/category/form-grabber/"
End If
If (List1.Text = "Crypter.py" Or List1.Text = "crypter.py") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/sneaky-way-how-to-hack-corporate-email-account/"
End If
If (List1.Text = "Msfpayload" Or List1.Text = "msfpayload") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/inject-backdoor-using-metasploit/"
End If
If (List1.Text = "Agobot" Or List1.Text = "agobot") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/pentest-with-agobot-how-botnet-work-robots-wars/"
End If
If (List1.Text = "Slowloris" Or List1.Text = "slowloris") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/http-ddos-with-slowloris/"
End If
If (List1.Text = "Joomscan" Or List1.Text = "joomscan") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/easiest-way-to-hack-joomla-using-joomscan/"
End If
If (List1.Text = "Autoinfector" Or List1.Text = "autoinfector") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/infectious-media-generator-infecting-target-automatically/"
End If
If (List1.Text = "Aircrack" Or List1.Text = "aircrack") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-wep-key-in-windows-using-commviewaircrack-gui/"
End If
If (List1.Text = "Burpsuite" Or List1.Text = "burpsuite") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/bypassing-upload-filter-with-burpsuite/"
End If
If (List1.Text = "CasperSpy" Or List1.Text = "casperspy") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/casperspy-overview/"
End If
If (List1.Text = "Citadel" Or List1.Text = "citadel") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/tutorial-citadel-botnet-user-friendly-botnet/"
End If
If (List1.Text = "Cythosia" Or List1.Text = "cythosia") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/destroying-webserver-easily-using-cythosia-ddos-bot/"
End If
If (List1.Text = "Dirbuster" Or List1.Text = "dirbuster") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/webapps-fuzzer-webdir-enumeration-attack-dirbuster/"
End If
If (List1.Text = "Dnsdict" Or List1.Text = "dnsdict") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/dns-analysis-domain-enumeration-with-dnsdict/"
End If
If (List1.Text = "Dmitry" Or List1.Text = "dmitry") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/gathering-information-with-dmitry-route-analysis/"
End If
If (List1.Text = "Dnsmap" Or List1.Text = "dnsmap") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/bruteforce-subdomain-with-dnsmap/"
End If
If (List1.Text = "Dnsrecon" Or List1.Text = "dnsrecon") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/reconnaissance-dns-with-dnsrecon/"
End If
If (List1.Text = "Driftnet" Or List1.Text = "driftnet") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/network-sniffer-sniffing-media-files-image-video-with-driftnet/"
End If
If (List1.Text = "DVWA" Or List1.Text = "dvwa") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/hacking-website-complete-guide-of-xss-injection-part-5/"
End If
If (List1.Text = "Edjpgcom" Or List1.Text = "edjpgcom") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/injecting-exploit-in-jpg/"
End If
If (List1.Text = "Ettercap" Or List1.Text = "ettercap") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/network-sniffing-how-to-dns-spoofing-with-ettercap/"
End If
If (List1.Text = "Evilpdf exploit" Or List1.Text = "evilpdf exploit") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/inject-exploit-into-pdf-with-metasploit-evilpdf/"
End If
If (List1.Text = "Genpmk" Or List1.Text = "genpmk") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/generating-hash-to-crack-wpa-with-genpmk/"
End If
If (List1.Text = "Havij Pro" Or List1.Text = "havij pro") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://www.itsecteam.com/files/havij/havij_help-english.pdf"
End If
If (List1.Text = "Janidos" Or List1.Text = "janidos") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/janidos-powerfull-ddos-with-zombie-army/"
End If
If (List1.Text = "Jigsaw" Or List1.Text = "jigsaw") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/enumeration-employee-info-with-jigsaw/"
End If
If (List1.Text = "LBD" Or List1.Text = "lbd") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/detect-load-balancing-with-lbd/"
End If
If (List1.Text = "Metasploit" Or List1.Text = "metasploit") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/category/metasploit-2/"
End If
If (List1.Text = "MP-Formgrabber" Or List1.Text = "mp-formgrabber") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/mp-formgrabber-stable-on-all-major-browser/"
End If
If (List1.Text = "Nmap" Or List1.Text = "nmap") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/import-nmap-scan-xml-into-metasploit/"
End If
If (List1.Text = "Pyrit" Or List1.Text = "pyrit") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/fastest-way-to-crack-wpa-wpa2-pyrit/"
End If
If (List1.Text = "Reaver" Or List1.Text = "reaver") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-wifi-protected-setup-part2-brute-force-attack-against-wifi-protected-setup-using-reaver/"
End If
If (List1.Text = "SET" Or List1.Text = "set") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/tag/set/"
End If
If (List1.Text = "Spitmo" Or List1.Text = "spitmo") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/spyeye-for-android-spitmo/"
End If
If (List1.Text = "Spy Eye" Or List1.Text = "spy eye") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/category/spyeye/"
End If
If (List1.Text = "SQLMAP" Or List1.Text = "sqlmap") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/complete-guide-how-to-hack-website/"
End If
If (List1.Text = "The Harvester" Or List1.Text = "the harvester") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/gathering-sensitive-information-from-corporation-with-theharvester/"
End If
If (List1.Text = "VertexNet" Or List1.Text = "vertexnet") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/mp-formgrabber-stable-on-all-major-browser/"
End If
If (List1.Text = "Vulcan" Or List1.Text = "vulcan") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/hack-any-account-using-vulcan-remote-keylogger/"
End If
If (List1.Text = "Wash" Or List1.Text = "wash") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-wifi-protected-setup-part-1/"
End If
If (List1.Text = "Webcrab" Or List1.Text = "webcrab") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/webcrab-underground-formgrabber/"
End If
If (List1.Text = "Wifite" Or List1.Text = "wifite") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/cracking-multiple-wifi-automatically-using-wifite-v2/"
End If
If (List1.Text = "Zeus Botnet" Or List1.Text = "zeus botnet") Then
frmmain2.Show
frmmain2.wb1.Navigate "http://casperspy.com/zeus-complete-tutorial/"
End If

End Sub

Private Sub PushButton4_Click()
Intro.Show
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If TabControl1.Selected.Caption = "Attacker" Then
List1.Clear
Open App.path & "\programs\attackerlist.txt" For Input As #1
While Not EOF(1)
Input #1, search
List1.AddItem (search)
Wend
Close #1
End If
If TabControl1.Selected.Caption = "Botnet" Then
List1.Clear
List1.AddItem "Agobot"
List1.AddItem "Andromeda"
List1.AddItem "CasperSpy"
List1.AddItem "Citadel"
List1.AddItem "Crime Pack"
List1.AddItem "Cythosia"
List1.AddItem "Janidos"
List1.AddItem "Picebot"
List1.AddItem "RxBot"
List1.AddItem "Subseven"
List1.AddItem "Spitmo"
List1.AddItem "Spy Eye"
List1.AddItem "VertexNet"
List1.AddItem "Zeus Botnet"
End If
If TabControl1.Selected.Caption = "Exploit" Then
List1.Clear
List1.AddItem "Crypter.py"
List1.AddItem "Evilpdf exploit"
List1.AddItem "Metasploit"
List1.AddItem "Msfpayload"
End If
If TabControl1.Selected.Caption = "Cracking" Then
List1.Clear
List1.AddItem "Aircrack"
List1.AddItem "Hexedit"
List1.AddItem "Hiew32"
List1.AddItem "IdaPro"
List1.AddItem "OllyDbg"
List1.AddItem "PeExplorer"
List1.AddItem "ResHack"
List1.AddItem "SoftIce"
End If
If TabControl1.Selected.Caption = "Keylogger" Then
List1.Clear
List1.AddItem "Albertino"
List1.AddItem "CasperLogger"
List1.AddItem "Keymail"
List1.AddItem "Rapzo"
List1.AddItem "Vulcan"
End If
If TabControl1.Selected.Caption = "Network" Then
List1.Clear
List1.AddItem "Aircrack"
List1.AddItem "Dmitry"
List1.AddItem "Dnsdict"
List1.AddItem "Driftnet"
List1.AddItem "Ettercap"
List1.AddItem "Gerix"
List1.AddItem "Janidos"
List1.AddItem "LBD"
List1.AddItem "Nmap"
List1.AddItem "Pyrit"
List1.AddItem "Reaver"
List1.AddItem "Slowloris"
List1.AddItem "Wash"
List1.AddItem "Wifite"
End If
If TabControl1.Selected.Caption = "Social Engginering" Then
List1.Clear
List1.AddItem "Autoinfector"
List1.AddItem "Edjpgcom"
List1.AddItem "Jigsaw"
List1.AddItem "SET"
List1.AddItem "The Harvester"
End If
If TabControl1.Selected.Caption = "Trojan" Then
List1.Clear
List1.AddItem "Bifrost"
List1.AddItem "Bionet"
List1.AddItem "Cybergate"
List1.AddItem "Spynet"
List1.AddItem "Subseven"
List1.AddItem "Trojan Creator"
List1.AddItem "Turkojan"
End If
If TabControl1.Selected.Caption = "Web Hacking" Then
List1.Clear
List1.AddItem "Burpsuite"
List1.AddItem "Dirbuster"
List1.AddItem "DVWA"
List1.AddItem "Havij Pro"
List1.AddItem "Joomscan"
List1.AddItem "SQLMAP"
List1.AddItem "Wp-Scan"
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
m.Text = List1.ListCount & "/" & m
End Sub

Private Sub Text10_Click()
Text10.Text = ""
End Sub

