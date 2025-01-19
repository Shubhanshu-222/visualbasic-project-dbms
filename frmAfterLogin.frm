VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAfterLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome screen"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3855
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAfterLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsername 
      Height          =   1335
      Left            =   3240
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtPrimary 
      Height          =   1335
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdOpenExistingRecords 
      Appearance      =   0  'Flat
      Caption         =   "Open Existing"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      ToolTipText     =   "Open existing record"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnterNewRecord 
      Appearance      =   0  'Flat
      Caption         =   "Enter New Record"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Enter new record"
      Top             =   1680
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   11
      Scrolling       =   1
   End
   Begin VB.Label lblUserPass 
      BackStyle       =   0  'Transparent
      Caption         =   "What would you like to do? Please choose."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label lblMaster 
      BackStyle       =   0  'Transparent
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "Rockwell Extra Bold"
         Size            =   36
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblVer 
      BackStyle       =   0  'Transparent
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   3480
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   3480
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "frmAfterLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnterNewRecord_Click()

    'This will open the main form in "Enter new record" mode
    frmGeneral.txtPrimary.Text = txtPrimary.Text
    mdiMain.Show
    frmGeneral.txtAfterLogin.Text = "New"
    frmGeneral.lblGeneralUsername.Caption = txtUsername.Text
    
    Unload Me
    
End Sub

Private Sub cmdOpenExistingRecords_Click()

    'This will open the main form and display the first record
    frmGeneral.txtPrimary.Text = txtPrimary.Text
    mdiMain.Show
    frmGeneral.txtAfterLogin.Text = "Existing"
    frmGeneral.lblGeneralUsername.Caption = txtUsername.Text
   
    Unload Me

End Sub

Private Sub Form_Load()
    
    'Loading Progress Bar in frmLogin
    frmLogin.ProgressBar.Visible = True
    frmLogin.ProgressBar.Value = 0
    frmLogin.ProgressBar.Value = 10
    frmLogin.ProgressBar.Value = 20
    frmLogin.ProgressBar.Value = 30
    frmLogin.ProgressBar.Value = 40
    frmLogin.ProgressBar.Value = 50
    frmLogin.ProgressBar.Value = 60
    frmLogin.ProgressBar.Value = 70
    frmLogin.ProgressBar.Value = 80
    frmLogin.ProgressBar.Value = 90
    frmLogin.ProgressBar.Value = 100
    frmLogin.ProgressBar.Visible = False

End Sub
