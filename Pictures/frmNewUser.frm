VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewUser 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboGender 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox cboProvider2 
      Height          =   315
      Left            =   5640
      TabIndex        =   12
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   360
      TabIndex        =   17
      Top             =   -240
      Width           =   7455
      Begin VB.TextBox txtPrimary 
         Height          =   375
         Left            =   3720
         TabIndex        =   48
         Top             =   2760
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ComboBox cboProvider1 
         Height          =   315
         Left            =   4320
         TabIndex        =   10
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txtCCode 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
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
         Left            =   3000
         TabIndex        =   16
         Top             =   9600
         Width           =   1455
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   5880
         Width           =   2535
      End
      Begin VB.TextBox txtCID 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   6480
         Width           =   2535
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   7080
         Width           =   2535
      End
      Begin VB.TextBox txtCPass 
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   7680
         Width           =   2535
      End
      Begin VB.ComboBox cboCity 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox txtContact 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   5280
         Width           =   2175
      End
      Begin VB.PictureBox Picture2 
         Height          =   255
         Left            =   2400
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   29
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkDeclare 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   8760
         Width           =   255
      End
      Begin VB.TextBox txtLName 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtFName 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   1080
         Width           =   3735
      End
      Begin VB.PictureBox Picture4 
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   28
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   3960
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   27
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture6 
         Height          =   255
         Left            =   3960
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   26
         Top             =   4080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture7 
         Height          =   255
         Left            =   5280
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture8 
         Height          =   255
         Left            =   5280
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture10 
         Height          =   255
         Left            =   3960
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   23
         Top             =   4680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture11 
         Height          =   255
         Left            =   4200
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   22
         Top             =   5280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture9 
         Height          =   255
         Left            =   5520
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   21
         Top             =   5880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture12 
         Height          =   255
         Left            =   6480
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   20
         Top             =   6480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture13 
         Height          =   255
         Left            =   3960
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   19
         Top             =   7080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture15 
         Height          =   255
         Left            =   4800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   18
         Top             =   7680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cboCountry 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   4680
         Width           =   2655
      End
      Begin VB.ComboBox cboState 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   4080
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtDOB 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   143523841
         CurrentDate     =   42724
      End
      Begin MSComctlLib.ImageList iml 
         Left            =   6720
         Top             =   9480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   47
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "@"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   46
         Top             =   5880
         Width           =   255
      End
      Begin VB.Label lblFName 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   1080
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLName 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   1680
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblGender 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   43
         Top             =   2280
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDOB 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   2880
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCity 
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   3480
         Width           =   495
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   4080
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCountry 
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   4680
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Email ID:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   5880
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm E-mail ID:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   6480
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   36
         Top             =   7080
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   7680
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPStrength 
         BackStyle       =   0  'Transparent
         Caption         =   "Password Strength:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   34
         Top             =   8280
         Width           =   1935
      End
      Begin VB.Label lblContact 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label lblWelcome 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to registration form."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   32
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblKindly 
         BackStyle       =   0  'Transparent
         Caption         =   "Kindly fill in your details and get started with Master."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   31
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label lblDeclare 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmNewUser.frx":0000
         Height          =   615
         Left            =   720
         TabIndex        =   30
         Top             =   8880
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdSubmit_Click()

    rs.MoveLast
    rs.AddNew
    
    rs.Fields("Primary").Value = txtFName.Text & "." & txtLName.Text
    rs.Fields("Username").Value = txtID.Text & "@" & cboProvider1.Text
    rs.Fields("Email Id").Value = txtID.Text
    rs.Fields("Email Provider").Value = cboProvider1.Text
    rs.Fields("Password").Value = txtPass.Text
    rs.Fields("First Name").Value = txtFName.Text
    rs.Fields("Last Name").Value = txtLName.Text
    rs.Fields("DOB").Value = dtDOB.Value
    rs.Fields("Gender").Value = cboGender.Text
    rs.Fields("City").Value = cboCity.Text
    rs.Fields("State").Value = cboState.Text
    rs.Fields("Country").Value = cboCountry.Text
    rs.Fields("Contact").Value = txtContact.Text
    
    txtPrimary.Text = rs.Fields("Primary").Value
    
    rs.Update
    frmSecurity.Show
    frmSecurity.txtPrimary.Text = frmNewUser.txtPrimary.Text
    
    Set frmNewUser = Nothing
    Unload Me
    
End Sub

Private Sub Form_Load()

    con.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source= C:\Users\Hp\Desktop\VB6new\Security.mdb; Persist Security Info= False"
    rs.Open "Select * from Security", con, adOpenDynamic, adLockOptimistic
    
End Sub

