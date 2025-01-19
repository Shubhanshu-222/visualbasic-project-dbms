VERSION 5.00
Begin VB.Form frmCriminal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17190
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   17190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
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
      Left            =   4680
      TabIndex        =   22
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
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
      Left            =   4680
      TabIndex        =   21
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddNew 
      Appearance      =   0  'Flat
      Caption         =   "Add New"
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
      Left            =   13680
      TabIndex        =   20
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox txtDu 
      Height          =   495
      Left            =   14280
      TabIndex        =   19
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtPr 
      Height          =   495
      Left            =   8040
      TabIndex        =   18
      Top             =   2160
      Width           =   6015
   End
   Begin VB.TextBox txtAg 
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtCr 
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtSr 
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdLast 
      Appearance      =   0  'Flat
      Caption         =   ">>"
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
      Left            =   11520
      TabIndex        =   14
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdFirst 
      Appearance      =   0  'Flat
      Caption         =   "<<"
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
      Left            =   6840
      TabIndex        =   13
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      Caption         =   ">"
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
      Left            =   9960
      TabIndex        =   12
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrevious 
      Appearance      =   0  'Flat
      Caption         =   "<"
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
      Left            =   8400
      TabIndex        =   11
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdTravelling 
      Appearance      =   0  'Flat
      Caption         =   "Travelling"
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
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdEducational 
      Appearance      =   0  'Flat
      Caption         =   "Educational"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdMedical 
      Appearance      =   0  'Flat
      Caption         =   "Medical"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdGeneral 
      Appearance      =   0  'Flat
      Caption         =   "General"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Line Line9 
      X1              =   17040
      X2              =   17040
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line8 
      X1              =   1800
      X2              =   17040
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lbl 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   960
      Width           =   735
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   120
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
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   23
      Top             =   0
      Width           =   3135
   End
   Begin VB.Line Line7 
      X1              =   1800
      X2              =   1800
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line6 
      X1              =   1800
      X2              =   17040
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line5 
      X1              =   14160
      X2              =   14160
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line4 
      X1              =   7920
      X2              =   7920
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line3 
      X1              =   6960
      X2              =   6960
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line2 
      X1              =   2880
      X2              =   2880
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   17040
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblSrNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sr. No."
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
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblDuration 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration"
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
      Left            =   15240
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblPrisonRecord 
      BackStyle       =   0  'Transparent
      Caption         =   "Prison Record"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblAge 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Left            =   7320
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblCrime 
      BackStyle       =   0  'Transparent
      Caption         =   "Crime"
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
      Left            =   4560
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Criminal Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   7575
   End
End
Attribute VB_Name = "frmCriminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEdit_Click()

    cmdEdit.Visible = False
    cmdSave.Visible = True
    
End Sub

Private Sub cmdEducational_Click()

    frmCriminal.Hide
    frmEducational.Show
    
End Sub

Private Sub cmdGeneral_Click()

    frmCriminal.Hide
    frmGeneral.Show
    
End Sub

Private Sub cmdMedical_Click()

    frmCriminal.Hide
    frmMedical.Show
    
End Sub

Private Sub cmdSave_Click()

    cmdSave.Visible = False
    cmdEdit.Visible = True
    
End Sub

Private Sub cmdTravelling_Click()

    frmCriminal.Hide
    frmTravelling.Show
    
End Sub

Private Sub Form_Load()

    cmdSave.Visible = False
    
End Sub
