VERSION 5.00
Begin VB.Form frmMedical 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WWWWW"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   17190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAd 
      Height          =   495
      Left            =   10800
      TabIndex        =   36
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtIl 
      Height          =   495
      Left            =   3000
      TabIndex        =   33
      Top             =   2640
      Width           =   3255
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
      TabIndex        =   15
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdCriminal 
      Appearance      =   0  'Flat
      Caption         =   "Criminal"
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
      TabIndex        =   14
      Top             =   2040
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
      TabIndex        =   13
      Top             =   2640
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
      TabIndex        =   12
      Top             =   3240
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   7920
      Width           =   1455
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
      TabIndex        =   8
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox txtSr 
      Height          =   495
      Left            =   1920
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtSIN 
      Height          =   495
      Left            =   13320
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtAg 
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox txtDo 
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtSt 
      Height          =   495
      Left            =   14280
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   7920
      Width           =   1455
   End
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
      TabIndex        =   0
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Line Line10 
      X1              =   17040
      X2              =   17040
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line9 
      X1              =   1800
      X2              =   17040
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblAd 
      Height          =   375
      Left            =   10920
      TabIndex        =   35
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
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
      Left            =   12000
      TabIndex        =   34
      Top             =   1560
      Width           =   735
   End
   Begin VB.Line Line8 
      X1              =   10680
      X2              =   10680
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Label lblIl 
      Height          =   375
      Left            =   3120
      TabIndex        =   32
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblIllness 
      Caption         =   "Illness"
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
      Left            =   4320
      TabIndex        =   31
      Top             =   1560
      Width           =   615
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
      TabIndex        =   30
      Top             =   960
      Width           =   7695
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Medical Data"
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
      TabIndex        =   29
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblSr 
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSINno 
      Height          =   375
      Left            =   13440
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblAg 
      Height          =   375
      Left            =   6600
      TabIndex        =   26
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblDo 
      Height          =   375
      Left            =   7560
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblSt 
      Height          =   375
      Left            =   14400
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblSIN 
      BackStyle       =   0  'Transparent
      Caption         =   "Social Insurance Number:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   23
      Top             =   960
      Width           =   2775
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
      Left            =   6720
      TabIndex        =   22
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblDoctor 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor (Hospital)"
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
      Left            =   8280
      TabIndex        =   21
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      TabIndex        =   20
      Top             =   1560
      Width           =   735
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
      TabIndex        =   19
      Top             =   1560
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   17040
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   2880
      X2              =   2880
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line3 
      X1              =   6360
      X2              =   6360
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line4 
      X1              =   7320
      X2              =   7320
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line5 
      X1              =   14160
      X2              =   14160
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line6 
      X1              =   1800
      X2              =   17040
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line7 
      X1              =   1800
      X2              =   1800
      Y1              =   1440
      Y2              =   7680
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
      TabIndex        =   18
      Top             =   0
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   255
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
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmMedical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCriminal_Click()

    frmMedical.Hide
    frmCriminal.Show
    
End Sub

Private Sub cmdEducational_Click()

    frmMedical.Hide
    frmEducational.Show
    
End Sub

Private Sub cmdGeneral_Click()

    frmMedical.Hide
    frmGeneral.Show
    
End Sub

Private Sub cmdTravelling_Click()

    frmMedical.Hide
    frmTravelling.Show
    
End Sub
