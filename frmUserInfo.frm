VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAboutMe 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmUserInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   625
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   5520
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   90
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picSelf 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   5520
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   89
      Top             =   1680
      Width           =   1575
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
      Left            =   5640
      TabIndex        =   9
      ToolTipText     =   "Upload image"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture40 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   88
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic14 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   87
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic13 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   86
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cboSecurity 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2760
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtAnswer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtok11 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   81
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok12 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   80
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
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
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Go back"
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   5520
      TabIndex        =   2
      ToolTipText     =   "Save changes"
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtPrimary 
      Height          =   285
      Left            =   5160
      TabIndex        =   44
      Top             =   570
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cboCity 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtContact 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox pic3 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   43
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtLName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtFName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox pic4 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   42
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic5 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   41
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic6 
      Height          =   255
      Left            =   5280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   40
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   39
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   38
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic7 
      Height          =   255
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   37
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic8 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   36
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cboCountry 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cboState 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3720
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cboGender 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboProvider1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
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
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "Edit account"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtID 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtCID 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   14
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtCPass 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      TabIndex        =   17
      Top             =   7200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox pic9 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   35
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic10 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic11 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   33
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic12 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   32
      Top             =   7200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cboProvider2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   16
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picStrength1 
      Height          =   165
      Left            =   1800
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   31
      Top             =   5880
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtok1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   29
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok5 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   26
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok6 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6840
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok7 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok8 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok9 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6840
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtok10 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picStrength2 
      Height          =   165
      Left            =   1800
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   20
      Top             =   7560
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Width           =   1215
         WordWrap        =   -1  'True
      End
   End
   Begin MSComCtl2.DTPicker dtDOB 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   124846081
      CurrentDate     =   42724
   End
   Begin MSComDlg.CommonDialog PicDialog 
      Left            =   6240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   -360
      TabIndex        =   95
      Top             =   9120
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   11
      Scrolling       =   1
   End
   Begin VB.Label lblMyCriminal 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   94
      Top             =   2520
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyHealthy 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   93
      Top             =   3000
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFirst 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   5520
      TabIndex        =   92
      Top             =   3480
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   5520
      TabIndex        =   91
      Top             =   3840
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyAnswer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   85
      Top             =   7200
      Visible         =   0   'False
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMySecurity 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   84
      Top             =   6720
      Visible         =   0   'False
      Width           =   3255
      WordWrap        =   -1  'True
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   83
      Top             =   6720
      Visible         =   0   'False
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   82
      Top             =   7200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   24
      X2              =   472
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Label lblMyPassword 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   79
      Top             =   5520
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyID 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   78
      Top             =   5040
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyContact 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2160
      TabIndex        =   77
      Top             =   4560
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyCCode 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   76
      Top             =   4560
      Width           =   375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyCountry 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   75
      Top             =   4080
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyState 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   74
      Top             =   3600
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyCity 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   73
      Top             =   3600
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyDOB 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   72
      Top             =   3120
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyGender 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   70
      Top             =   2640
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyLastName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   69
      Top             =   2160
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyFirstName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   68
      Top             =   1680
      Width           =   3975
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   67
      Top             =   1680
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   66
      Top             =   2160
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   65
      Top             =   2640
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   64
      Top             =   3120
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   63
      Top             =   3600
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   62
      Top             =   3600
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   61
      Top             =   4080
      Width           =   855
      WordWrap        =   -1  'True
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   60
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lblAboutMe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Account"
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
      Left            =   1920
      TabIndex        =   59
      Tag             =   " "
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblHere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Here, you can edit your account information"
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
      Height          =   375
      Left            =   1800
      TabIndex        =   58
      Top             =   1080
      Width           =   3735
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   57
      Top             =   6600
      Visible         =   0   'False
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   56
      Top             =   5040
      Width           =   255
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   55
      Top             =   5040
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Email ID:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   54
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   53
      Top             =   5520
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
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   720
      TabIndex        =   52
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
      WordWrap        =   -1  'True
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
      TabIndex        =   51
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   50
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   24
      X2              =   472
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   24
      X2              =   472
      Y1              =   608
      Y2              =   608
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   24
      X2              =   472
      Y1              =   528
      Y2              =   528
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   49
      Top             =   4560
      Width           =   135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCom1 
      BackStyle       =   0  'Transparent
      Caption         =   ".com"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   48
      Top             =   5040
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCom2 
      BackStyle       =   0  'Transparent
      Caption         =   ".com"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   47
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPassStat1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   46
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPassStat2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   45
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset
Private lb As New ADODB.Recordset
Private str As String

'Following two Lost Focus events will tell the strength of the password
Private Sub txtCPass_LostFocus()

    'Checking strength of Password
    'First criteria : Length
    If Len(Trim(txtPass.Text)) > 0 And Len(Trim(txtPass.Text)) < 6 Then
            
        'Second criteria : Is numeric or alphanumeric?
        If IsNumeric(txtPass.Text) Then
        
            lblPassStat2.Visible = True
            lblPassStat2.Caption = "Weak"
            picStrength2.Visible = True
            picStrength2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Weak.jpg")
                    
        Else
                    
            lblPassStat2.Visible = True
            lblPassStat2.Caption = "Normal"
            picStrength2.Visible = True
            picStrength2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Average.jpg")
                    
        End If
    End If
            
    'First
    If Len(Trim(txtPass.Text)) > 5 And Len(Trim(txtPass.Text)) < 11 Then
                
        'Second
        If IsNumeric(txtPass.Text) Then
        
            lblPassStat2.Visible = True
            lblPassStat2.Caption = "Normal"
            picStrength2.Visible = True
            picStrength2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Average.jpg")
                    
        Else
        
            lblPassStat2.Visible = True
            lblPassStat2.Caption = "Good"
            picStrength2.Visible = True
            picStrength2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Good.jpg")
                    
        End If
    End If
            
    'First
    If Len(Trim(txtPass.Text)) > 10 Then
            
        'Second
        If IsNumeric(txtPass.Text) Then
        
            lblPassStat2.Visible = True
            lblPassStat2.Caption = "Good"
            picStrength2.Visible = True
            picStrength2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Good.jpg")
                    
        Else
        
            lblPassStat2.Visible = True
            lblPassStat2.Caption = "Excellent"
            picStrength2.Visible = True
            picStrength2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Good.jpg")
                    
        End If
    End If

End Sub

Private Sub txtPass_LostFocus()

    'Checking strength of Password
    'First criteria : Length
    If Len(Trim(txtPass.Text)) > 0 And Len(Trim(txtPass.Text)) < 6 Then
            
        'Second criteria : Is numeric or alphanumeric?
        If IsNumeric(txtPass.Text) Then
        
            lblPassStat1.Visible = True
            lblPassStat1.Caption = "Weak"
            picStrength1.Visible = True
            picStrength1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Weak.jpg")
                    
        Else
                    
            lblPassStat1.Visible = True
            lblPassStat1.Caption = "Normal"
            picStrength1.Visible = True
            picStrength1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Average.jpg")
                    
        End If
    End If
            
    'First
    If Len(Trim(txtPass.Text)) > 5 And Len(Trim(txtPass.Text)) < 11 Then
                
        'Second
        If IsNumeric(txtPass.Text) Then
        
            lblPassStat1.Visible = True
            lblPassStat1.Caption = "Normal"
            picStrength1.Visible = True
            picStrength1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Average.jpg")
                    
        Else
                    
            lblPassStat1.Visible = True
            lblPassStat1.Caption = "Good"
            picStrength1.Visible = True
            picStrength1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Good.jpg")
                    
        End If
    End If
    
    'First
    If Len(Trim(txtPass.Text)) > 10 Then
            
        'Second
        If IsNumeric(txtPass.Text) Then
        
            lblPassStat1.Visible = True
            lblPassStat1.Caption = "Good"
            picStrength1.Visible = True
            picStrength1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Good.jpg")
                    
        Else
                    
            lblPassStat1.Visible = True
            lblPassStat1.Caption = "Excellent"
            picStrength1.Visible = True
            picStrength1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Password_Good.jpg")
                    
        End If
    End If
            
End Sub

Private Sub cmdSave_Click()
    
    Set rs = New Recordset
    
    rs.Open "Select * from Security where Primary = '" + frmGeneral.txtPrimary.Text + "'", _
    con, adOpenDynamic, adLockOptimistic
    
    ProgressBar.Visible = True
    ProgressBar.Value = 0
    
    'Save the data in the database
    pic1.Visible = False
    pic2.Visible = False
    pic3.Visible = False
    pic4.Visible = False
    pic5.Visible = False
    pic6.Visible = False
    pic7.Visible = False
    pic8.Visible = False
    pic9.Visible = False
    pic10.Visible = False
    pic11.Visible = False
    pic12.Visible = False
    pic13.Visible = False
    pic14.Visible = False
    
    'Preparing Validation
    txtok1.Text = ""
    txtok2.Text = ""
    txtok3.Text = ""
    txtok4.Text = ""
    txtok5.Text = ""
    txtok6.Text = ""
    txtok7.Text = ""
    txtok8.Text = ""
    txtok9.Text = ""
    txtok10.Text = ""
    txtok11.Text = ""
    txtok12.Text = ""

    
    'Validating First Name
    'Checking for non-numerical first name or simply alphabetical first name
    If Not IsAlphabetical(Trim(txtFName.Text)) Or Trim(txtFName.Text) = "" Then
    
        'MsgBox "Please re-enter your first name", vbCritical, "Invalid field"
        pic1.Visible = True
            
    Else
            
        'First Name is ok
        txtok1.Text = "ok"
            
    End If
    
    ProgressBar.Value = 10
    
    'Validating Last Name
    'Checking for the same in last name
    If Not IsAlphabetical(Trim(txtLName.Text)) Or Trim(txtLName.Text) = "" Then
    
        'MsgBox "Please re-enter your Last name", vbCritical, "Invalid field"
        pic2.Visible = True
            
    Else
            
        'Last Name is ok
        txtok2.Text = "ok"
    
    End If
    
    'Validating Gender
    'MALE & FEMALE
    If Trim(cboGender.Text) <> "" Then

        If Trim(cboGender.Text) = "Male" Or Trim(cboGender.Text) = "Female" Then
        
            'Gender is ok
            txtok3.Text = "ok"
            
        Else
            
            pic3.Visible = True
                
        End If
    Else
            
        pic3.Visible = True
        
    End If
       
    ProgressBar.Value = 20
        
    'Validating DOB
    If dtDOB.Value = "12/20/2016" Then
        
        If MsgBox("Is your DOB, 12/20/2016?", vbInformation + vbYesNo, "Note") = vbNo Then
        
            MsgBox "Please cboose your DOB", vbCritical, "Invalid field"
            pic4.Visible = True
                
        Else
            
            txtok4.Text = "ok"
                
        End If
            
    Else
           
        'DOB is ok
        txtok4.Text = "ok"
            
    End If
        
        
    'Validating City
    If Trim(cboCity.Text) <> "" Then
    
        If cboCity.Text = "Mumbai" Or cboCity.Text = "Navi Mumbai" Or _
        cboCity.Text = "Pune" Or cboCity.Text = "Nagpur" Or _
        cboCity.Text = "Thane" Or cboCity.Text = "Pimpri" Or _
        cboCity.Text = "Chinchwad" Or cboCity.Text = "Nashik" Or _
        cboCity.Text = "Kalyan" Or cboCity.Text = "Dombivali" Or _
        cboCity.Text = "Pune" Or cboCity.Text = "Vasai" Or _
        cboCity.Text = "Virar" Or cboCity.Text = "Aurangabad" Or _
        cboCity.Text = "Solapur" Or cboCity.Text = "Mira" Or _
        cboCity.Text = "Bhayandar" Or cboCity.Text = "Bhiwandi" Or _
        cboCity.Text = "Nizampur" Or cboCity.Text = "Amravati" Or _
        cboCity.Text = "Nanded" Or cboCity.Text = "Waghala" Or _
        cboCity.Text = "Panvel" Or cboCity.Text = "Sangli" Or _
        cboCity.Text = "Akola" Or cboCity.Text = "Ahmednagar" Or _
        cboCity.Text = "Parbhani" Or cboCity.Text = "Chandrapur" Or _
        cboCity.Text = "Dhule" Or cboCity.Text = "Malegaon" Or _
        cboCity.Text = "Jalgaon" Or cboCity.Text = "Kolhapur" Or _
        cboCity.Text = "Nashik" Or cboCity.Text = "Latur" Then
                    
            txtok5.Text = "ok"
                
        Else
        
            pic5.Visible = True
                
        End If
            
    Else
            
        'MsgBox "Please select from given cities", vbCritical, "Note"
        pic5.Visible = True
            
    End If
    
    ProgressBar.Value = 30
    
    'Validating State
    If Trim(cboState.Text) <> "" Then
        
        If Trim(cboState.Text) = "Maharashtra" Then
        
            'State is ok
            txtok6.Text = "ok"
            
        Else
    
            'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
            pic6.Visible = True
                
        End If
    
    Else
        
        'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
        pic6.Visible = True
                
    End If
        
        
    'Validating Country
    If Trim(cboCountry.Text) <> "" Then
            
        If Trim(cboCountry.Text) = "India" Then
        
            'Country is ok
            txtok7.Text = "ok"
        
        Else
            
            'MsgBox "Please select from given state(s)", vbCritical, "Note"
            pic7.Visible = True
            
        End If
            
    Else
        
        pic7.Visible = True
            
    End If
    
    ProgressBar.Value = 40

    'Validating Contact
    If Not IsNumeric(Trim(txtContact.Text)) Or Len(Trim(txtContact.Text)) <> 10 Then
    
        pic8.Visible = True
        'MsgBox "Please enter valid contact", vbCritical, "Invalid contact"
            
    Else
        
        'Contact is ok
        txtok8.Text = "ok"
            
    End If
    
    ProgressBar.Value = 50
    
    'Validating ID
    If Trim(txtID.Text) <> Trim(txtCID.Text) Or Trim(cboProvider1.Text) <> Trim(cboProvider2.Text) Then
    
        pic9.Visible = True
        pic10.Visible = True
        'MsgBox "Please re-enter your email id", vbCritical, "Unequal email ids"
            
    Else
            
        If Trim(cboProvider1.Text) = "Gmail" Or Trim(cboProvider1.Text) = "iCloud" Or _
        Trim(cboProvider1.Text) = "GMX" Or Trim(cboProvider1.Text) = "Outlook" Or _
        Trim(cboProvider1.Text) = "Yahoo" Or Trim(cboProvider1.Text) = "Aol" Or _
        Trim(cboProvider1.Text) = "Zoho" Or Trim(cboProvider1.Text) = "Mail" Or _
        Trim(cboProvider1.Text) = "Yandex" Or Trim(cboProvider1.Text) = "ProtonMail" Then
            
            If Trim(cboProvider2.Text) = "Gmail" Or Trim(cboProvider2.Text) = "iCloud" Or _
            Trim(cboProvider2.Text) = "GMX" Or Trim(cboProvider2.Text) = "Outlook" Or _
            Trim(cboProvider2.Text) = "Yahoo" Or Trim(cboProvider2.Text) = "Aol" Or _
            Trim(cboProvider2.Text) = "Zoho" Or Trim(cboProvider2.Text) = "Mail" Or _
            Trim(cboProvider2.Text) = "Yandex" Or Trim(cboProvider2.Text) = "ProtonMail" Then
            
                'Email ids are ok
                txtok9.Text = "ok"
            
            End If
            
        Else
        
            pic9.Visible = True
            pic10.Visible = True
            
        End If
    End If
    
    ProgressBar.Value = 60
            
    'Validating Password
    If Trim(txtPass.Text) <> Trim(txtCPass.Text) Then
        
        pic11.Visible = True
        pic12.Visible = True
        'MsgBox "Please re-enter password", vbCritical, "Unequal passwords"
            
    Else
    
        'Password is ok
        txtok10.Text = "ok"
            
    End If
    
    ProgressBar.Value = 70
    
    'Validating security question
    If Trim(cboSecurity.Text) <> "" Then
    
        If cboSecurity.Text = "What is my name?" Or cboSecurity.Text = "Where do i live?" Or _
        cboSecurity.Text = "What do i like?" Or cboSecurity.Text = "What do i want to become?" Or _
        cboSecurity.Text = "What will i do tomorrow?" Or cboSecurity.Text = "When will i die?" Or _
        cboSecurity.Text = "What does clouds look like?" Or cboSecurity.Text = "When do i wake up?" Then
            
            txtok11.Text = "ok"

        Else
            
            'MsgBox "Please select any given question", vbCritical, "Invalid field"
            pic13.Visible = True
            
        End If
            
    Else
        
        pic13.Visible = True
        'MsgBox "Please select any given question", vbCritical, "Empty field"
        
    End If
    
    ProgressBar.Value = 80
    
    'Validating Answer
    If Trim(txtAnswer.Text) <> "" Then
    
        txtok12.Text = "ok"
        
    Else
        
        pic14.Visible = True
        'MsgBox "Please enter your answer", vbCritical, "Empty field"
        
    End If
    
    ProgressBar.Value = 100
    ProgressBar.Visible = False
            
    
    'Saving the DATA
    If txtok1.Text = "ok" And txtok2.Text = "ok" And txtok3.Text = "ok" And _
    txtok4.Text = "ok" And txtok5.Text = "ok" And txtok6.Text = "ok" And _
    txtok7.Text = "ok" And txtok8.Text = "ok" And txtok9.Text = "ok" And _
    txtok10.Text = "ok" And txtok11.Text = "ok" And txtok12.Text = "ok" Then
    
        ProgressBar.Visible = True
        ProgressBar.Value = 0
            
        rs.Fields("Primary").Value = Trim(txtFName.Text) & "." & Trim(txtLName.Text)
        rs.Fields("Username").Value = Trim(txtID.Text) & "@" & cboProvider1.Text & ".com"
        rs.Fields("Email Id").Value = Trim(txtID.Text)
        rs.Fields("Email Provider").Value = cboProvider1.Text
        rs.Fields("Password").Value = Trim(txtPass.Text)
        rs.Fields("First Name").Value = Trim(txtFName.Text)
        rs.Fields("Last Name").Value = Trim(txtLName.Text)
        rs.Fields("DOB").Value = dtDOB.Value
        rs.Fields("Gender").Value = cboGender.Text
        rs.Fields("City").Value = cboCity.Text
        rs.Fields("State").Value = cboState.Text
        rs.Fields("Country").Value = cboCountry.Text
        rs.Fields("Contact").Value = Trim(txtContact.Text)
            
        If str <> "" Then
            
            rs.Fields("Picture").Value = str
                
        End If
        
        ProgressBar.Value = 10
            
        txtPrimary.Text = rs.Fields("Primary").Value

        rs.Fields("Security").Value = cboSecurity.Text
        rs.Fields("Answer").Value = txtAnswer.Text
            
        rs.Update
        
        ProgressBar.Value = 100
        ProgressBar.Visible = False
        
        MsgBox "Success", vbInformation, "Date updated"
            
        rs.Close
            
        'Off all objects
        'Off Main objects
        cmdUpload.Visible = False
            
        lblCom1.Visible = False
        Label3.Visible = False
        txtFName.Visible = False
        txtLName.Visible = False
        cboGender.Visible = False
        dtDOB.Visible = False
        cboCity.Visible = False
        cboState.Visible = False
        cboCountry.Visible = False
        txtContact.Visible = False
        txtID.Visible = False
        cboProvider1.Visible = False
        txtPass.Visible = False
            
        'Security Bundle Hidden
        lblMySecurity.Visible = False
        lblMyAnswer.Visible = False
        cboSecurity.Visible = False
        txtAnswer.Visible = False
        lblSecurityQues.Visible = False
        lblEnterAns.Visible = False
            
        'Security Bundle Movement
        lblSecurityQues.Left = "56"
        lblSecurityQues.Top = "448"
        lblEnterAns.Left = "56"
        lblEnterAns.Top = "480"
        cboSecurity.Left = "184"
        cboSecurity.Top = "448"
        txtAnswer.Left = "184"
        txtAnswer.Top = "480"
        lblMySecurity.Left = "184"
        lblMySecurity.Top = "448"
        lblMyAnswer.Left = "184"
        lblMyAnswer.Top = "480"
            
        lblEnterAns.Caption = "Your answer:"
            
        'Security Bundle Show
        lblMySecurity.Visible = True
        lblMyAnswer.Visible = True
        lblSecurityQues.Visible = True
        lblEnterAns.Visible = True
            
        'Off extra objects
        lblCEmail.Visible = False
        txtCID.Visible = False
        Label13.Visible = False
        cboProvider2.Visible = False
        lblCom2.Visible = False
        lblCPassword.Visible = False
        txtCPass.Visible = False
        Label1.Visible = False
        lblPassStat2.Visible = False
            
        'off Password Strength
        picStrength1.Visible = False
        picStrength2.Visible = False
        lblPassStat1.Visible = False
        lblPassStat2.Visible = False
            
        'Off cmd
        cmdBack.Visible = False
        cmdSave.Visible = False
            
        'on Edit
        cmdEdit.Visible = True
            
        'On Labels
        lblMyFirstName.Visible = True
        lblMyLastName.Visible = True
        lblMyGender.Visible = True
        lblMyDOB.Visible = True
        lblMyCity.Visible = True
        lblMyState.Visible = True
        lblMyCountry.Visible = True
        lblMyContact.Visible = True
        lblMyID.Visible = True
        lblMyPassword.Visible = True
            
        'Open lb recordset
        Set lb = New Recordset
        lb.Open "Select * from Security where Primary = '" + txtPrimary.Text + "'", _
        con, adOpenDynamic, adLockOptimistic
            
        frmGeneral.txtPrimary.Text = txtPrimary.Text
            
        'Display user data
        lblMyFirstName.Caption = lb.Fields("First Name").Value
        lblMyLastName.Caption = lb.Fields("Last Name").Value
        lblMyGender.Caption = lb.Fields("Gender").Value
        lblMyDOB.Caption = lb.Fields("DOB").Value
        lblMyCity.Caption = lb.Fields("City").Value
        lblMyState.Caption = lb.Fields("State").Value
        lblMyCountry.Caption = lb.Fields("Country").Value
        lblMyCCode.Caption = "+91"
        lblMyContact.Caption = lb.Fields("Contact").Value
        lblMyID.Caption = lb.Fields("Email ID").Value & "@" & lb.Fields("Email Provider").Value & ".com"
        lblMyPassword.Caption = lb.Fields("Password").Value
        
        If lb.Fields("Picture").Value <> "" Then
            
            picResize.Picture = LoadPicture(lb!Picture)
        
            picSelf.PaintPicture picResize.Picture, 0, 0, picSelf.ScaleWidth, picSelf.ScaleHeight, _
            0, 0, picResize.ScaleWidth, picResize.ScaleHeight, vbSrcCopy
                
        End If
            
        lblMyAnswer.Left = 145
        
    Else
        
        'MsgBox will display because fields are not filled
        MsgBox "Please fill correctly", vbCritical, "Incorrect Data"
            
    End If
    
End Sub

Private Sub cmdBack_Click()
    
    'Off cross
    pic1.Visible = False
    pic2.Visible = False
    pic3.Visible = False
    pic4.Visible = False
    pic5.Visible = False
    pic6.Visible = False
    pic7.Visible = False
    pic8.Visible = False
    pic9.Visible = False
    pic10.Visible = False
    pic11.Visible = False
    pic12.Visible = False
    pic13.Visible = False
    pic14.Visible = False
    
    picStrength1.Visible = False
    picStrength2.Visible = False
    
    
    'Off Main objects
    lblCom1.Visible = False
    Label3.Visible = False
    txtFName.Visible = False
    txtLName.Visible = False
    cboGender.Visible = False
    dtDOB.Visible = False
    cboCity.Visible = False
    cboState.Visible = False
    cboCountry.Visible = False
    txtContact.Visible = False
    txtID.Visible = False
    cboProvider1.Visible = False
    txtPass.Visible = False
    
    'Security Bundle Hidden
    lblMySecurity.Visible = False
    lblMyAnswer.Visible = False
    cboSecurity.Visible = False
    txtAnswer.Visible = False
    lblSecurityQues.Visible = False
    lblEnterAns.Visible = False
    
    'Security Bundle Movement
    lblSecurityQues.Left = "56"
    lblSecurityQues.Top = "448"
    lblEnterAns.Left = "56"
    lblEnterAns.Top = "480"
    cboSecurity.Left = "184"
    cboSecurity.Top = "448"
    txtAnswer.Left = "184"
    txtAnswer.Top = "480"
    lblMySecurity.Left = "184"
    lblMySecurity.Top = "448"
    lblMyAnswer.Left = "184"
    lblMyAnswer.Top = "480"
    
    lblEnterAns.Caption = "Your answer:"
    
    'Security Bundle Show
    lblMySecurity.Visible = True
    lblMyAnswer.Visible = True
    lblSecurityQues.Visible = True
    lblEnterAns.Visible = True
    
    'Off extra objects
    lblCEmail.Visible = False
    txtCID.Visible = False
    Label13.Visible = False
    cboProvider2.Visible = False
    lblCom2.Visible = False
    lblCPassword.Visible = False
    txtCPass.Visible = False
    Label1.Visible = False
    lblPassStat2.Visible = False
    
    'Off cmd
    cmdBack.Visible = False
    cmdSave.Visible = False
    cmdUpload.Visible = False
    
    'on Edit
    cmdEdit.Visible = True
    
    'On Labels
    lblMyFirstName.Visible = True
    lblMyLastName.Visible = True
    lblMyGender.Visible = True
    lblMyDOB.Visible = True
    lblMyCity.Visible = True
    lblMyState.Visible = True
    lblMyCountry.Visible = True
    lblMyContact.Visible = True
    lblMyID.Visible = True
    lblMyPassword.Visible = True
    
    'Open lb recordset
    lb.Open "Select * from Security where Primary = '" + frmGeneral.txtPrimary.Text + "'", _
    con, adOpenDynamic, adLockOptimistic
    
    'Display user data
    lblMyFirstName.Caption = lb.Fields("First Name").Value
    lblMyLastName.Caption = lb.Fields("Last Name").Value
    lblMyGender.Caption = lb.Fields("Gender").Value
    lblMyDOB.Caption = lb.Fields("DOB").Value
    lblMyCity.Caption = lb.Fields("City").Value
    lblMyState.Caption = lb.Fields("State").Value
    lblMyCountry.Caption = lb.Fields("Country").Value
    lblMyCCode.Caption = "+91"
    lblMyContact.Caption = lb.Fields("Contact").Value
    lblMyID.Caption = lb.Fields("Email ID").Value & "@" & lb.Fields("Email Provider").Value & ".com"
    lblMyPassword.Caption = lb.Fields("Password").Value
    
    If lb.Fields("Picture").Value <> "" Then
    
        picResize.Picture = LoadPicture(lb!Picture)
        
        picSelf.PaintPicture picResize.Picture, 0, 0, picSelf.ScaleWidth, _
        picSelf.ScaleHeight, 0, 0, picResize.ScaleWidth, picResize.ScaleHeight, vbSrcCopy

    End If
    
    lblMyAnswer.Left = 145
    
End Sub

Private Sub cmdEdit_Click()

    cmdEdit.Visible = False
    lblMyAnswer.Left = 184
    cmdUpload.Visible = True
    
    'Off labels
    lblMyFirstName.Visible = False
    lblMyLastName.Visible = False
    lblMyGender.Visible = False
    lblMyDOB.Visible = False
    lblMyCity.Visible = False
    lblMyState.Visible = False
    lblMyCountry.Visible = False
    lblMyContact.Visible = False
    lblMyID.Visible = False
    lblMyPassword.Visible = False
    
    'Security Bundle Hidden
    lblMySecurity.Visible = False
    lblMyAnswer.Visible = False
    cboSecurity.Visible = False
    txtAnswer.Visible = False
    lblSecurityQues.Visible = False
    lblEnterAns.Visible = False
    
    'Security Bundle Movement
    lblSecurityQues.Left = "56"
    lblSecurityQues.Top = "544"
    lblEnterAns.Left = "56"
    lblEnterAns.Top = "576"
    cboSecurity.Left = "184"
    cboSecurity.Top = "544"
    txtAnswer.Left = "184"
    txtAnswer.Top = "576"
    lblMySecurity.Left = "184"
    lblMySecurity.Top = "544"
    lblMyAnswer.Left = "184"
    lblMyAnswer.Top = "576"
    
    lblEnterAns.Caption = "Enter your answer:"
    
    'Security Bundle Show
    cboSecurity.Visible = True
    txtAnswer.Visible = True
    lblSecurityQues.Visible = True
    lblEnterAns.Visible = True
    
    
    'On Main objects
    lblCom1.Visible = True
    Label3.Visible = True
    txtFName.Visible = True
    txtLName.Visible = True
    cboGender.Visible = True
    dtDOB.Visible = True
    cboCity.Visible = True
    cboState.Visible = True
    cboCountry.Visible = True
    txtContact.Visible = True
    txtID.Visible = True
    cboProvider1.Visible = True
    txtPass.Visible = True
    
    
    'Adding extra objects
    lblCEmail.Visible = True
    txtCID.Visible = True
    Label13.Visible = True
    cboProvider2.Visible = True
    lblCom2.Visible = True
    lblCPassword.Visible = True
    txtCPass.Visible = True
    Label1.Visible = True
    lblPassStat2.Visible = True
    
    'on cmd
    cmdBack.Visible = True
    cmdSave.Visible = True
    
    'close lb
    lb.Close
    
    'Show Password
    txtPass.PasswordChar = ""
    txtCPass.PasswordChar = ""
    
    'Open & load rs
    rs.Open "Select * from Security where Primary = '" + frmGeneral.txtPrimary.Text + "'", _
    con, adOpenDynamic, adLockOptimistic

    txtFName.Text = rs.Fields("First Name").Value
    txtLName.Text = rs.Fields("Last Name").Value
    cboGender.Text = rs.Fields("Gender").Value
    dtDOB.Value = rs.Fields("DOB").Value
    cboCity.Text = rs.Fields("City").Value
    cboState.Text = rs.Fields("State").Value
    cboCountry.Text = rs.Fields("Country").Value
    txtContact.Text = rs.Fields("Contact").Value
    txtID.Text = rs.Fields("Email ID").Value
    txtCID.Text = rs.Fields("Email ID").Value
    cboProvider1.Text = rs.Fields("Email Provider").Value
    cboProvider2.Text = rs.Fields("Email Provider").Value
    txtPass.Text = rs.Fields("Password").Value
    txtCPass.Text = rs.Fields("Password").Value
    
    If rs!Security <> "" Then
    
        cboSecurity.Text = rs.Fields("Security").Value
        
    End If
    
    If rs!Answer <> "" Then
    
        txtAnswer.Text = rs.Fields("Answer").Value
        
    End If
    
    If rs.Fields("Picture").Value <> "" Then
    
        picResize.Picture = LoadPicture(rs!Picture)
        
        picSelf.PaintPicture picResize.Picture, 0, 0, picSelf.ScaleWidth, _
        picSelf.ScaleHeight, 0, 0, picResize.ScaleWidth, picResize.ScaleHeight, vbSrcCopy

    End If
    
    rs.Close
    
End Sub

Private Sub cmdUpload_Click()

    PicDialog.ShowOpen
    PicDialog.Filter = "Jpeg|*.jpg"
    str = PicDialog.FileName
    
    If str <> "" Then
    
        picResize.Picture = LoadPicture(str)
        picSelf.PaintPicture picResize.Picture, 0, 0, picSelf.ScaleWidth, _
        picSelf.ScaleHeight, 0, 0, picResize.ScaleWidth, picResize.ScaleHeight, vbSrcCopy
        
    End If
    
End Sub

'Following two Got Focus events are meant to support respective Lost Focus events
Private Sub txtCPass_GotFocus()

     txtCPass.PasswordChar = "*"
     
End Sub

Private Sub txtPass_GotFocus()

    txtPass.PasswordChar = "*"
    
End Sub

Private Function IsAlphabetical(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String

    'This will return true if all characters are alphabets else false
    
    sTemp = TestString
    iLen = Len(sTemp)
    
    If iLen > 0 Then
    
        For iCtr = 1 To iLen
        
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[A-Za-z' ']" Then Exit Function
        
        Next
    
        IsAlphabetical = True
    
    End If

End Function

Private Function IsNumeric(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String
    
    'This will return true if all characters are numbers else false
    
    sTemp = TestString
    iLen = Len(sTemp)
    
    If iLen > 0 Then
    
        For iCtr = 1 To iLen
        
            sChar = Mid(sTemp, iCtr, 1)
            If Not sChar Like "[0-9]" Then Exit Function
            
        Next
    
        IsNumeric = True
        
    End If

End Function

Private Function Clear()

    'This will clear the respective record
    rs.Fields("Primary").Value = ""
    rs.Fields("Username").Value = ""
    rs.Fields("Email Id").Value = ""
    rs.Fields("Email Provider").Value = ""
    rs.Fields("Password").Value = ""
    rs.Fields("First Name").Value = ""
    rs.Fields("Last Name").Value = ""
    rs.Fields("DOB").Value = ""
    rs.Fields("Gender").Value = ""
    rs.Fields("City").Value = ""
    rs.Fields("State").Value = ""
    rs.Fields("Country").Value = ""
    rs.Fields("Contact").Value = ""
            
    rs.Fields("Security").Value = ""
    rs.Fields("Answer").Value = ""
            
End Function

Private Sub Form_Load()

    con.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source= " _
    & "C:\Users\Hp\Desktop\VB6new\Security.mdb; Persist Security Info= False"
    lb.Open "Select * from Security where Primary = " _
    & "'" + frmGeneral.txtPrimary.Text + "'", con, adOpenDynamic, adLockOptimistic
    
    lblMyAnswer.Left = 145
    cmdUpload.Visible = False
    
    'on edit
    cmdEdit.Visible = True
    
    'On Security Question
    lblSecurityQues.Visible = True
    lblEnterAns.Visible = True
    lblMySecurity.Visible = True
    lblMyAnswer.Visible = True
    
    lblEnterAns.Caption = "Your answer:"
    
    'Displaying User Information
    lblMyFirstName.Caption = lb.Fields("First Name").Value
    lblMyLastName.Caption = lb.Fields("Last Name").Value
    lblMyGender.Caption = lb.Fields("Gender").Value
    lblMyDOB.Caption = lb.Fields("DOB").Value
    lblMyCity.Caption = lb.Fields("City").Value
    lblMyState.Caption = lb.Fields("State").Value
    lblMyCountry.Caption = lb.Fields("Country").Value
    lblMyCCode.Caption = "+91"
    lblMyContact.Caption = lb.Fields("Contact").Value
    lblMyID.Caption = lb.Fields("Email ID").Value & "@" & lb.Fields("Email Provider").Value & ".com"
    lblMyPassword.Caption = lb.Fields("Password").Value
    
    If lb!Security <> "" Then
    
        lblMySecurity.Caption = lb.Fields("Security").Value
        
    End If
    
    If lb!Answer <> "" Then
    
        lblMyAnswer.Caption = lb.Fields("Answer").Value
        
    End If
    
    If lb.Fields("Picture").Value <> "" Then
            
        picResize.Picture = LoadPicture(lb!Picture)
        
        picSelf.PaintPicture picResize.Picture, 0, 0, picSelf.ScaleWidth, picSelf.ScaleHeight, _
        0, 0, picResize.ScaleWidth, picResize.ScaleHeight, vbSrcCopy
                
    End If
    
    'Adding the wrong entre image in boxes pic 1 to 12
    pic1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic3.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic4.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic5.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic6.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic7.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic8.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic9.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic10.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic11.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic12.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic13.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    pic14.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
    
    'Off extra objects
    lblCEmail.Visible = False
    txtCID.Visible = False
    Label13.Visible = False
    cboProvider2.Text = False
    lblCom2.Visible = False
    Label3.Visible = False
    lblCom1.Visible = False
    lblCPassword.Visible = False
    txtCPass.Visible = False
    Label1.Visible = False
    lblPassStat2.Visible = False
    
    'Gender
    cboGender.AddItem "Male"
    cboGender.AddItem "Female"
    
    'City
    cboCity.AddItem "Mumbai"
    cboCity.AddItem "Navi Mumbai"
    cboCity.AddItem "Pune"
    cboCity.AddItem "Nagpur"
    cboCity.AddItem "Thane"
    cboCity.AddItem "Pimpri"
    cboCity.AddItem "Chinchwad"
    cboCity.AddItem "Nashik"
    cboCity.AddItem "Kalyan"
    cboCity.AddItem "Dombivali"
    cboCity.AddItem "Pune"
    cboCity.AddItem "Vasai"
    cboCity.AddItem "Virar"
    cboCity.AddItem "Aurangabad"
    cboCity.AddItem "Solapur"
    cboCity.AddItem "Mira"
    cboCity.AddItem "Bhayandar"
    cboCity.AddItem "Bhiwandi"
    cboCity.AddItem "Nizampur"
    cboCity.AddItem "Amravati"
    cboCity.AddItem "Nanded"
    cboCity.AddItem "Waghala"
    cboCity.AddItem "Panvel"
    cboCity.AddItem "Sangli"
    cboCity.AddItem "Akola"
    cboCity.AddItem "Ahmednagar"
    cboCity.AddItem "Parbhani"
    cboCity.AddItem "Chandrapur"
    cboCity.AddItem "Dhule"
    cboCity.AddItem "Malegaon"
    cboCity.AddItem "Jalgaon"
    cboCity.AddItem "Kolhapur"
    cboCity.AddItem "Nashik"
    cboCity.AddItem "Latur"

    'State
    cboState.AddItem "Maharashtra"
    
    'Country
    cboCountry.AddItem "India"
     
    'Email Service Providers
    '1
    cboProvider1.AddItem "Gmail"
    cboProvider1.AddItem "Outlook"
    cboProvider1.AddItem "Yahoo"
    cboProvider1.AddItem "Aol"
    cboProvider1.AddItem "Zoho"
    cboProvider1.AddItem "Mail"
    cboProvider1.AddItem "Yandex"
    cboProvider1.AddItem "ProtonMail"
    cboProvider1.AddItem "GMX"
    cboProvider1.AddItem "iCloud"
    
    '2
    cboProvider2.AddItem "Gmail"
    cboProvider2.AddItem "Outlook"
    cboProvider2.AddItem "Yahoo"
    cboProvider2.AddItem "Aol"
    cboProvider2.AddItem "Zoho"
    cboProvider2.AddItem "Mail"
    cboProvider2.AddItem "Yandex"
    cboProvider2.AddItem "ProtonMail"
    cboProvider2.AddItem "GMX"
    cboProvider2.AddItem "iCloud"

    'Adding questions to security questions
    cboSecurity.AddItem "What is my name?"
    cboSecurity.AddItem "Where do i live?"
    cboSecurity.AddItem "What do i like?"
    cboSecurity.AddItem "What do i want to become?"
    cboSecurity.AddItem "What will i do tomorrow?"
    cboSecurity.AddItem "When will i die?"
    cboSecurity.AddItem "What does clouds look like?"
    cboSecurity.AddItem "When do i wake up?"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'This will close the connection when the form unloads
    con.Close

End Sub
