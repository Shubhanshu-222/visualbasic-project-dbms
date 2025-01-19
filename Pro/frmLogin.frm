VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdForgotPass 
      Appearance      =   0  'Flat
      Caption         =   "Forgot Password"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdNewUser 
      Appearance      =   0  'Flat
      Caption         =   "New User"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      Caption         =   "Login"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   3
      Text            =   "Password"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "Username"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lblUserPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your username and password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblSlogan 
      BackStyle       =   0  'Transparent
      Caption         =   "One for all"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3225
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   4080
      Width           =   255
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
      Left            =   5040
      TabIndex        =   1
      Top             =   480
      Width           =   255
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForgotPass_Click()

    frmForgotPassword.Show
    frmLogin.Hide
    
End Sub

Private Sub cmdLogin_Click()

    frmLogin.Hide
    mdiMain.Show
    
End Sub

Private Sub cmdNewUser_Click()

    frmNewUser.Show
    frmLogin.Hide

End Sub

Private Sub Form_Load()

    txtUsername.Text = "Username"       'Entering text in login textboxes
    txtPassword.PasswordChar = ""
    txtPassword.Text = "Password"
    
End Sub

Private Sub txtPassword_Click()

    txtPassword.Text = ""               'Making sure password can be entered
    txtPassword.PasswordChar = "*"
    If txtUsername.Text = "" Then
        txtUsername.Text = "Username"
    End If
    
End Sub

Private Sub txtUsername_Click()

    txtUsername.Text = ""               'Making sure username can be entered
    If txtPassword.Text = "" Then
        txtPassword.PasswordChar = ""
        txtPassword.Text = "Password"
    End If
    
End Sub
