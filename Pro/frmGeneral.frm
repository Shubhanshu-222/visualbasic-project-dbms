VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGeneral 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "General Data"
   ClientHeight    =   8640
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   20400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog C 
      Left            =   9000
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   7080
      TabIndex        =   90
      Top             =   7440
      Width           =   1335
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
      Left            =   8640
      TabIndex        =   89
      Top             =   7440
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
      Left            =   4920
      TabIndex        =   88
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox txtFavPlaces 
      Height          =   405
      Left            =   12960
      TabIndex        =   87
      Top             =   6840
      Width           =   6615
   End
   Begin VB.TextBox txtHobbies 
      Height          =   405
      Left            =   12120
      TabIndex        =   86
      Top             =   6120
      Width           =   7455
   End
   Begin VB.ComboBox cboMarital 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   12720
      TabIndex        =   85
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ComboBox Combo14 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   14280
      TabIndex        =   84
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ComboBox Combo13 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   17040
      TabIndex        =   83
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ComboBox cboCity3 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   11640
      TabIndex        =   82
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtEmpAdd 
      Height          =   405
      Left            =   13320
      TabIndex        =   78
      Top             =   4200
      Width           =   6255
   End
   Begin VB.TextBox txtEmployer 
      Height          =   285
      Left            =   15240
      TabIndex        =   76
      Top             =   3600
      Width           =   4335
   End
   Begin VB.ComboBox cboJob 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   12240
      TabIndex        =   74
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ComboBox cboEducation 
      Height          =   315
      Left            =   13200
      TabIndex        =   73
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txtSisters 
      Height          =   375
      Left            =   16080
      TabIndex        =   71
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txtBrothers 
      Height          =   375
      Left            =   14400
      TabIndex        =   70
      Text            =   " "
      Top             =   2160
      Width           =   495
   End
   Begin VB.ComboBox cboSiblings 
      Height          =   315
      Left            =   12120
      TabIndex        =   69
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txtMothers 
      Height          =   285
      Left            =   17160
      TabIndex        =   68
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtFathers 
      Height          =   285
      Left            =   12720
      TabIndex        =   67
      Top             =   1560
      Width           =   2415
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
      Left            =   4920
      TabIndex        =   66
      Top             =   8040
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
      Left            =   14160
      TabIndex        =   65
      Top             =   8040
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
      Left            =   10320
      TabIndex        =   64
      Top             =   8040
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
      Left            =   12000
      TabIndex        =   63
      Top             =   8040
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
      Left            =   8640
      TabIndex        =   62
      Top             =   8040
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
      Left            =   6960
      TabIndex        =   61
      Top             =   8040
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
      Left            =   12000
      TabIndex        =   60
      Top             =   7440
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
      Left            =   10320
      TabIndex        =   59
      Top             =   7440
      Width           =   1455
   End
   Begin VB.ComboBox cboOccupation 
      Height          =   315
      Left            =   1800
      TabIndex        =   58
      Top             =   7200
      Width           =   3495
   End
   Begin VB.ComboBox cboProvider 
      Height          =   315
      Left            =   15120
      TabIndex        =   56
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   12120
      TabIndex        =   55
      Top             =   960
      Width           =   2535
   End
   Begin VB.ComboBox cboReligion 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   6600
      TabIndex        =   53
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtTelCode 
      Height          =   285
      Left            =   2040
      TabIndex        =   52
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtHome 
      Height          =   285
      Left            =   3120
      TabIndex        =   51
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox txtCCode 
      Height          =   285
      Left            =   1320
      TabIndex        =   50
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox txtPCode 
      Height          =   315
      Left            =   1920
      TabIndex        =   49
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox txtPCode1 
      Height          =   315
      Left            =   1800
      TabIndex        =   48
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ComboBox cboCountry2 
      DataField       =   "City"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   6480
      TabIndex        =   43
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ComboBox cboState2 
      DataField       =   "City"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   3600
      TabIndex        =   42
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ComboBox cboCity2 
      DataField       =   "City"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   960
      TabIndex        =   41
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtResAdd 
      Height          =   405
      Left            =   2520
      TabIndex        =   37
      Top             =   4080
      Width           =   5415
   End
   Begin VB.TextBox txtPerAdd 
      Height          =   405
      Left            =   2520
      TabIndex        =   36
      Top             =   2160
      Width           =   5535
   End
   Begin VB.ComboBox cboHealthy 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   9240
      TabIndex        =   35
      Top             =   4200
      Width           =   735
   End
   Begin VB.ComboBox cboCriminal 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   9240
      TabIndex        =   34
      Top             =   3600
      Width           =   735
   End
   Begin VB.ComboBox cboGender 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   1320
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload Image"
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
      TabIndex        =   18
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox cboState1 
      DataField       =   "State"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   3600
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox cboCountry1 
      DataField       =   "Country"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   6480
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtFName 
      DataField       =   "FName"
      DataSource      =   "Adodc"
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtLName 
      DataField       =   "LName"
      DataSource      =   "Adodc"
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtContact 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   6000
      Width           =   2175
   End
   Begin VB.ComboBox cboCity1 
      DataField       =   "City"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox picSelf 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   8400
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtdob 
      DataField       =   "DOB"
      DataSource      =   "Adodc"
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   117243905
      CurrentDate     =   42724
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
      Left            =   360
      TabIndex        =   94
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
      Left            =   3480
      TabIndex        =   93
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblHealthySay 
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      Height          =   315
      Left            =   9240
      TabIndex        =   92
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblCriminalSay 
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      Height          =   315
      Left            =   9240
      TabIndex        =   91
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label cboCountry3 
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
      Height          =   375
      Left            =   16080
      TabIndex        =   81
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label cboState3 
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
      Height          =   375
      Left            =   13560
      TabIndex        =   80
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label lblCity3 
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
      Height          =   375
      Left            =   11040
      TabIndex        =   79
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label lblEmpAdd 
      BackStyle       =   0  'Transparent
      Caption         =   "Employer's Addresss:"
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
      Left            =   11040
      TabIndex        =   77
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label lblEmployer 
      BackStyle       =   0  'Transparent
      Caption         =   "Employer:"
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
      Left            =   14160
      TabIndex        =   75
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter number of brothers and sisters you have."
      Height          =   495
      Left            =   13320
      TabIndex        =   72
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblOccupation 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation:"
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
      TabIndex        =   57
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblAtRate 
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
      Left            =   14760
      TabIndex        =   54
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblPCode2 
      BackStyle       =   0  'Transparent
      Caption         =   "Postal Code:"
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
      TabIndex        =   47
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblPCode1 
      BackStyle       =   0  'Transparent
      Caption         =   "Postal Code:"
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
      TabIndex        =   46
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblEducation 
      BackStyle       =   0  'Transparent
      Caption         =   "Education (Highest):"
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
      Left            =   11040
      TabIndex        =   45
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label lblSiblings 
      BackStyle       =   0  'Transparent
      Caption         =   " Siblings:"
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
      Left            =   11040
      TabIndex        =   44
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblCountry2 
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
      Height          =   375
      Left            =   5520
      TabIndex        =   40
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblState2 
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
      Left            =   2880
      TabIndex        =   39
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label lblCity2 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lblHContact 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Contact:"
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
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label lblFavPlaces 
      BackStyle       =   0  'Transparent
      Caption         =   "Favourite Places:"
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
      Left            =   11040
      TabIndex        =   32
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblHobbies 
      BackStyle       =   0  'Transparent
      Caption         =   "Hobbies:"
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
      Left            =   11040
      TabIndex        =   31
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label lblJob 
      BackStyle       =   0  'Transparent
      Caption         =   "Job Status:"
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
      Left            =   11040
      TabIndex        =   30
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblReligion 
      BackStyle       =   0  'Transparent
      Caption         =   "Religion:"
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
      Left            =   5520
      TabIndex        =   29
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblResAdd 
      BackStyle       =   0  'Transparent
      Caption         =   "Residential Address:"
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
      TabIndex        =   28
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label lblPerAdd 
      BackStyle       =   0  'Transparent
      Caption         =   "Permanent Address:"
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
      TabIndex        =   27
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblMarital 
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status:"
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
      Left            =   11040
      TabIndex        =   26
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblSisters 
      BackStyle       =   0  'Transparent
      Caption         =   "Sisters:"
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
      Left            =   15120
      TabIndex        =   25
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblBrothers 
      BackStyle       =   0  'Transparent
      Caption         =   "Brothers:"
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
      Left            =   13320
      TabIndex        =   24
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblMothers 
      BackStyle       =   0  'Transparent
      Caption         =   "Mother's Name:"
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
      Left            =   15360
      TabIndex        =   23
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblFathers 
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name:"
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
      Left            =   11040
      TabIndex        =   22
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblHealthy 
      BackStyle       =   0  'Transparent
      Caption         =   "Healthy:"
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
      Left            =   8220
      TabIndex        =   21
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblCriminal 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Criminal:"
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
      Left            =   8160
      TabIndex        =   20
      Top             =   3600
      Width           =   975
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
      TabIndex        =   17
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lblID 
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
      Left            =   11040
      TabIndex        =   16
      Top             =   960
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCountry1 
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
      Left            =   5520
      TabIndex        =   15
      Top             =   2880
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblState1 
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
      Left            =   2880
      TabIndex        =   14
      Top             =   2880
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCity1 
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
      TabIndex        =   13
      Top             =   2880
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDOB 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB:"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   1560
      Width           =   615
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
      TabIndex        =   11
      Top             =   1560
      Width           =   855
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
      Left            =   4320
      TabIndex        =   10
      Top             =   960
      Width           =   1215
      WordWrap        =   -1  'True
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
      TabIndex        =   9
      Top             =   960
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   10200
      X2              =   10200
      Y1              =   840
      Y2              =   7440
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "General Data"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCriminal_Click()

    frmCriminal.Show
    
End Sub

Private Sub cmdEdit_Click()

    cmdEdit.Visible = False
    cmdSave.Visible = True
    
End Sub

Private Sub cmdEducational_Click()

    frmEducational.Show
    
End Sub

Private Sub cmdMedical_Click()

    frmMedical.Show
    
End Sub

Private Sub cmdSave_Click()

    cmdSave.Visible = False
    cmdEdit.Visible = True
    
End Sub

Private Sub cmdTravelling_Click()

    frmTravelling.Show

End Sub

Private Sub cmdVerification_Click()

    frmVerification.Show
    
End Sub

Private Sub Form_Load()

    cmdSave.Visible = False
    
End Sub
