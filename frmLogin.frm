VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   4440
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer5 
      Interval        =   1050
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer Timer4 
      Interval        =   900
      Left            =   120
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Interval        =   450
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   400
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   350
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   4575
      Left            =   0
      MousePointer    =   11  'Hourglass
      Picture         =   "frmLogin.frx":7FC7
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   10
      Top             =   120
      Width           =   7335
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   4575
      Left            =   0
      Picture         =   "frmLogin.frx":10C94
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   11
      Top             =   120
      Width           =   7335
   End
   Begin VB.PictureBox pic3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   4575
      Left            =   0
      Picture         =   "frmLogin.frx":197F0
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   12
      Top             =   120
      Width           =   7335
   End
   Begin VB.PictureBox pic4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   4575
      Left            =   0
      Picture         =   "frmLogin.frx":21FFF
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   13
      Top             =   120
      Width           =   7335
   End
   Begin VB.PictureBox pic5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   4575
      Left            =   120
      Picture         =   "frmLogin.frx":26371
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   14
      Top             =   120
      Width           =   7335
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Username"
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   4
      Text            =   "Password"
      Top             =   2280
      Width           =   3615
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
      TabIndex        =   0
      Top             =   3120
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
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
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
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtPrimary 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc ado 
      Height          =   375
      Left            =   4680
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Hp\Desktop\VB6new\Security.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Hp\Desktop\VB6new\Security.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Security"
      Caption         =   "Login"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   6840
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   440
      Y1              =   0
      Y2              =   0
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
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   24
      X2              =   464
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   24
      X2              =   464
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   24
      X2              =   464
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "General Database Management System"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label lblUserPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your username and password"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   2895
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
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
    
    If Trim(txtUsername.Text) = "" And Trim(txtPassword.Text) = "" Then
    
        'In case of Incorrect Username and Password
        MsgBox "Enter you Username and Password", vbCritical, "Blank field(s)"
        
    Else
    
        'Verifying Username and Password
        ado.RecordSource = "select * from Security where Username = '" + txtUsername.Text + "'" _
        & "and Password = '" + txtPassword.Text + "'"
        ado.Refresh
        
        If ado.Recordset.EOF Then
            MsgBox "Please try again", vbExclamation, "Incorrect Username or Password"
            txtUsername = "Enter Username"
            txtPassword = "Enter Password"
            
        Else
        
            txtPrimary.Text = ado.Recordset.Fields("Primary")
            frmAfterLogin.txtPrimary.Text = txtPrimary.Text
            frmAfterLogin.txtUsername.Text = txtUsername.Text
            
            frmAfterLogin.Show
            
            ado.Recordset.Close
            
            Unload Me
            
        End If
    End If
    
    Set frmNewUser = Nothing
    
End Sub

Private Sub cmdForgotPass_Click()
    
    frmForgotPassword.Show
  
    Unload Me
    
End Sub

Private Sub cmdNewUser_Click()
    
    frmNewUser.Refresh
    frmNewUser.Show
   
    Unload Me

End Sub

Private Sub Timer1_Timer()

    pic1.Visible = False
    Timer1.Enabled = False
    Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()

    pic2.Visible = False
    Timer2.Enabled = False
    Timer3.Enabled = True
    
End Sub

Private Sub Timer3_Timer()

    pic3.Visible = False
    Timer3.Enabled = False
    Timer4.Enabled = True
    
End Sub

Private Sub Timer4_Timer()

    pic4.Visible = False
    Timer4.Enabled = False
    Timer5.Enabled = True
    
End Sub

Private Sub Timer5_Timer()

    pic5.Visible = False
    Timer5.Enabled = False

End Sub

Private Sub txtPassword_Click()

    If Trim(txtPassword.Text) = "" Or txtPassword.Text = "Enter Password" Then
        
        'Making sure password can be entered
        txtPassword.Text = ""
        txtPassword.PasswordChar = "*"
        
        If Trim(txtUsername.Text) = "" Then
        
            txtUsername.Text = "Enter Username"
            
        End If
    End If
    
End Sub

Private Sub txtPassword_GotFocus()

     If Trim(txtPassword.Text) = "" Or txtPassword.Text = "Enter Password" Then
        
        'Making sure password can be entered
        txtPassword.Text = ""
        txtPassword.PasswordChar = "*"
        
        If Trim(txtUsername.Text) = "" Then
        
            txtUsername.Text = "Enter Username"
            
        End If
    End If

End Sub

Private Sub txtUsername_Click()

    If Trim(txtUsername.Text) = "" Or txtUsername.Text = "Enter Username" Then
        
        'Making sure username can be entered
        txtUsername.Text = ""
        
        If Trim(txtPassword.Text) = "" Then
        
            txtPassword.PasswordChar = ""
            txtPassword.Text = "Enter Password"
            
        End If
    End If
    
End Sub

Private Sub txtUsername_GotFocus()

    If Trim(txtUsername.Text) = "" Or txtUsername.Text = "Enter Username" Then
        
        'Making sure username can be entered
        txtUsername.Text = ""
        
        If Trim(txtPassword.Text) = "" Then
        
            txtPassword.PasswordChar = ""
            txtPassword.Text = "Enter Password"
            
        End If
    End If

End Sub

Private Sub Form_Load()

    'Playing the song when the form loads
    wmp1.URL = "C:\Users\Hp\Desktop\VB6new\Wav\Master.wav"

    'Disabling timers for animation
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False

    'Displaying text on load
    txtUsername.Text = "Enter Username"       'Entering text in login textboxes
    txtPassword.PasswordChar = ""
    txtPassword.Text = "Enter Password"

End Sub
