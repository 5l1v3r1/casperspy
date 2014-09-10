VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.2#0"; "CODEJO~4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain2 
   BackColor       =   &H002A2A2A&
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10785
   Icon            =   "casperbrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar SB2 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   7845
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8235
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "CasperBrowser @casperspy"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.WebBrowser wb1 
      Height          =   7575
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   13695
      _Version        =   983042
      _ExtentX        =   24156
      _ExtentY        =   13361
      _StockProps     =   173
      ForeColor       =   2763306
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScrollBarStyle  =   1
      ScriptErrorsSuppressed=   -1  'True
   End
   Begin VB.Image Botao_Navegador 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "casperbrowser.frx":058A
      ToolTipText     =   "Retroceder"
      Top             =   240
      Width           =   240
   End
   Begin VB.Image Botao_Navegador 
      Height          =   240
      Index           =   2
      Left            =   1080
      Picture         =   "casperbrowser.frx":08CC
      ToolTipText     =   "Actualizar"
      Top             =   240
      Width           =   240
   End
   Begin VB.Image Botao_Navegador 
      Height          =   240
      Index           =   1
      Left            =   600
      Picture         =   "casperbrowser.frx":0C0E
      ToolTipText     =   "Avançar"
      Top             =   240
      Width           =   240
   End
   Begin VB.Image Botao_Navegador 
      Height          =   240
      Index           =   3
      Left            =   1560
      Picture         =   "casperbrowser.frx":0F50
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frmmain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Botao_Navegador_Click(Index As Integer)
    Select Case Index
        Case 0:
            goback
        Case 1:
            goforward
        Case 2:
            wb1.Refresh
            
    End Select
End Sub

Private Sub Form_Load()
    wb1.Navigate "www.casperspy.com"
End Sub

Private Sub WB1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Rem Notifies when page is done loading
SB2.Panels(1).Text = "Document done"
End Sub
Private Sub WB1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
Rem Notifies you what percent of the page is done, would look nice with a progress bar
On Error Resume Next
Dim Stuff As Long, Crap As Long, LastCrap As Long
Stuff = ProgressMax / 100
Crap = Progress / Stuff
If Crap < 1 Then Exit Sub
If Crap > 100 Then Exit Sub
If Crap < LastCrap Then Exit Sub
SB2.Panels(1).Text = Crap & "% Complete"
LastCrap = Crap
End Sub
Private Sub WB1_StatusTextChange(ByVal Text As String)
Rem Description of page
If wb1.LocationURL = "about:<b><body%20bgcolor%20=%20""gray""><font%20face%20=%20""arial""%20size%20=%20""2""%20color%20=%20""white""><center><br><br><br><br><br><br><br><br><br><br><br><br>Example%20web%20browser%20in%20vb!<br>by:%20keith_escalade<br><hr%20color%20=%20""white""><a%20href%20=%20""http://www.yahpro.org"">Yah-Pro<br><img%20src%20=%20""http://www.yahpro.org/themes/XP-Silver/images/logo.gif"">" Then: Exit Sub
Me.Caption = wb1.LocationName
SB1.SimpleText = Text
End Sub
Private Sub WB1_TitleChange(ByVal Text As String)
Rem Changes text1 when title is changed
If wb1.LocationURL = "about:<b><body%20bgcolor%20=%20""gray""><font%20face%20=%20""arial""%20size%20=%20""2""%20color%20=%20""white""><center><br><br><br><br><br><br><br><br><br><br><br><br>Example%20web%20browser%20in%20vb!<br>by:%20keith_escalade<br><hr%20color%20=%20""white""><a%20href%20=%20""http://www.yahpro.org"">Yah-Pro<br><img%20src%20=%20""http://www.yahpro.org/themes/XP-Silver/images/logo.gif"">" Then: Exit Sub

End Sub
Private Sub Form_Resize()
Rem used to get the web browser control just right to fit the form
On Error Resume Next
If Me.Height > 11130 Then
    Me.Height = 11130
    wb1.Width = Me.Width - 200
    wb1.Height = Me.Height - 1900
    PB1.Width = wb1.Width
    PB1.Top = SB1.Top - 120
    
End If
End Sub
