VERSION 5.00
Begin VB.Form frmSecurity 
   BackColor       =   &H00000000&
   Caption         =   "Security"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
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
      Height          =   2715
      Left            =   360
      TabIndex        =   2
      Top             =   -120
      Width           =   7455
      Begin VB.ComboBox cboSecurity 
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtPrimary 
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtAnswer 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Just one last step. It will secure your account and act as another password."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label lblKindly 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Kindly fill the fields."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblSecurityQues 
         BackStyle       =   0  'Transparent
         Caption         =   "Security Question:"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblEnterAns 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter your answer:"
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
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdSubmit_Click()

    rs.Fields("Security").Value = cboSecurity.Text
    rs.Fields("Answer").Value = txtAnswer.Text
    
    rs.Update
    
    frmLogin.Show
    
    Set frmSecurity = Nothing
    Unload Me
    
End Sub

Private Sub Form_Load()

    con.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source= C:\Users\Hp\Desktop\VB6new\Security.mdb; Persist Security Info= False"
    rs.Open "Select * from Security where Primary = '" + txtPrimary.Text + "'", con, adOpenDynamic, adLockOptimistic
    
End Sub
