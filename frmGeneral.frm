VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGeneral 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9360
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   20400
   ForeColor       =   &H00000000&
   Icon            =   "frmGeneral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   20400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSearch 
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
      Left            =   20040
      TabIndex        =   235
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtRecordReportID 
      Height          =   285
      Left            =   7440
      TabIndex        =   232
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cboKeyword 
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
      Left            =   15120
      TabIndex        =   30
      ToolTipText     =   "Sub main field"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text42 
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
      Left            =   19560
      TabIndex        =   228
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
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
      Left            =   9480
      TabIndex        =   225
      ToolTipText     =   "Clear fields"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox Picture40 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   19080
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   224
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtResidentialPlotNo 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtEmployerLocality 
      Height          =   285
      Left            =   16680
      TabIndex        =   49
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txtEmployerBuildingName 
      Height          =   285
      Left            =   14040
      TabIndex        =   48
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txtEmployerPlotNo 
      Height          =   285
      Left            =   12840
      TabIndex        =   47
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox txtPermanentLocality 
      Height          =   285
      Left            =   6600
      TabIndex        =   10
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtPermanentBuildingName 
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtPermanentPlotNo 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtResidentialLocality 
      Height          =   285
      Left            =   6600
      TabIndex        =   17
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtResidentialBuildingName 
      Height          =   285
      Left            =   3840
      TabIndex        =   16
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
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
      Left            =   9480
      TabIndex        =   60
      ToolTipText     =   "Delete record"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text41 
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
      Left            =   19560
      TabIndex        =   222
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text40 
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
      Left            =   18720
      TabIndex        =   221
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text39 
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
      Left            =   18240
      TabIndex        =   220
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text38 
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
      Left            =   17760
      TabIndex        =   219
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text37 
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
      Left            =   17280
      TabIndex        =   218
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text36 
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
      Left            =   16800
      TabIndex        =   217
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text35 
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
      Left            =   15960
      TabIndex        =   216
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text34 
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
      Left            =   15480
      TabIndex        =   215
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text33 
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
      Left            =   15000
      TabIndex        =   214
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text32 
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
      Left            =   14520
      TabIndex        =   213
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text31 
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
      Left            =   14040
      TabIndex        =   212
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text30 
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
      Left            =   18720
      TabIndex        =   211
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text29 
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
      Left            =   18240
      TabIndex        =   210
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text28 
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
      Left            =   17760
      TabIndex        =   209
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text27 
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
      Left            =   17280
      TabIndex        =   208
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text26 
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
      Left            =   16800
      TabIndex        =   207
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text25 
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
      Left            =   15960
      TabIndex        =   206
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text24 
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
      Left            =   15480
      TabIndex        =   205
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text23 
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
      Left            =   15000
      TabIndex        =   204
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text22 
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
      Left            =   14520
      TabIndex        =   203
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text21 
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
      Left            =   14040
      TabIndex        =   202
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text20 
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
      Left            =   5040
      TabIndex        =   201
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text19 
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
      Left            =   4560
      TabIndex        =   200
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text18 
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
      Left            =   4080
      TabIndex        =   199
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text17 
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
      Left            =   3600
      TabIndex        =   198
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text16 
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
      Left            =   3120
      TabIndex        =   197
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text15 
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
      Left            =   2280
      TabIndex        =   196
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text14 
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
      Left            =   1800
      TabIndex        =   195
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text13 
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
      Left            =   1320
      TabIndex        =   194
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text12 
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
      Left            =   840
      TabIndex        =   193
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text11 
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
      Left            =   360
      TabIndex        =   192
      Top             =   8520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text10 
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
      Left            =   5040
      TabIndex        =   191
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text9 
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
      Left            =   4560
      TabIndex        =   190
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text8 
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
      Left            =   4080
      TabIndex        =   189
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text7 
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
      Left            =   3600
      TabIndex        =   188
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text6 
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
      Left            =   3120
      TabIndex        =   187
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text5 
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
      Left            =   2280
      TabIndex        =   186
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text4 
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
      Left            =   1800
      TabIndex        =   185
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
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
      Left            =   1320
      TabIndex        =   184
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text2 
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
      Left            =   840
      TabIndex        =   183
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
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
      Left            =   360
      TabIndex        =   182
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture39 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   19200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   181
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture38 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   19200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   180
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture37 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   19200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   179
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture36 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   14040
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   178
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture35 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   19200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   177
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture34 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15840
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   176
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture33 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   175
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture32 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   19200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   174
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture31 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   19200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   173
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture30 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13440
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   172
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture29 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   171
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture28 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   16200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   170
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture27 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   14400
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   169
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture26 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   168
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture25 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   167
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture24 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   166
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture23 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   165
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture22 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   15240
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   164
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture21 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   17760
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   163
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture20 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   17760
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   162
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture19 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   161
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture18 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   160
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture17 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   159
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   158
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   157
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   156
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   155
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   154
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   153
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   152
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   151
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   150
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   149
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   148
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   147
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   146
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   145
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   144
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   143
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtRecordID 
      Height          =   285
      Left            =   5760
      TabIndex        =   110
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtAfterLogin 
      Height          =   285
      Left            =   9120
      TabIndex        =   109
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cboMothersO 
      Height          =   315
      Left            =   12840
      TabIndex        =   39
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox cboFathersO 
      Height          =   315
      Left            =   12840
      TabIndex        =   37
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1440
      TabIndex        =   27
      Top             =   6840
      Width           =   2535
   End
   Begin VB.ComboBox cboProvider 
      Height          =   315
      Left            =   4440
      Sorted          =   -1  'True
      TabIndex        =   28
      Top             =   6840
      Width           =   1215
   End
   Begin VB.ComboBox cboField 
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
      Left            =   13440
      TabIndex        =   29
      ToolTipText     =   "Main field"
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtKeyword 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   16800
      TabIndex        =   31
      ToolTipText     =   "Keyword"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
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
      Left            =   18480
      TabIndex        =   32
      ToolTipText     =   "Search"
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdReport 
      Appearance      =   0  'Flat
      Caption         =   "Show Report"
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
      Left            =   18480
      TabIndex        =   34
      ToolTipText     =   "Show report"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cboType 
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
      Left            =   16800
      TabIndex        =   33
      ToolTipText     =   "Report type"
      Top             =   1200
      Width           =   1575
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
      Left            =   6960
      TabIndex        =   57
      ToolTipText     =   "Go to First record"
      Top             =   8040
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
      Left            =   8640
      TabIndex        =   59
      ToolTipText     =   "Go to Previous record"
      Top             =   8040
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
      Left            =   6960
      TabIndex        =   58
      ToolTipText     =   "Edit record"
      Top             =   8520
      Width           =   1455
   End
   Begin VB.TextBox txtPrimary 
      DataSource      =   " "
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
      Left            =   10920
      TabIndex        =   61
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtPCode3 
      Height          =   315
      Left            =   12120
      MaxLength       =   6
      TabIndex        =   53
      Top             =   6000
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog PicDialog 
      Left            =   18960
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFavPlaces 
      Height          =   315
      Left            =   12600
      TabIndex        =   56
      Top             =   7080
      Width           =   6615
   End
   Begin VB.TextBox txtHobbies 
      Height          =   285
      Left            =   11760
      TabIndex        =   55
      Top             =   6480
      Width           =   7455
   End
   Begin VB.ComboBox cboMarital 
      Height          =   315
      Left            =   17640
      TabIndex        =   54
      Top             =   6000
      Width           =   1575
   End
   Begin VB.ComboBox cboState3 
      Height          =   315
      Left            =   14280
      Sorted          =   -1  'True
      TabIndex        =   51
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ComboBox cboCountry3 
      Height          =   315
      Left            =   17640
      Sorted          =   -1  'True
      TabIndex        =   52
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ComboBox cboCity3 
      Height          =   315
      Left            =   11160
      Sorted          =   -1  'True
      TabIndex        =   50
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtEmployer 
      Height          =   285
      Left            =   14880
      TabIndex        =   46
      Top             =   4440
      Width           =   4335
   End
   Begin VB.ComboBox cboJob 
      Height          =   315
      Left            =   11880
      Sorted          =   -1  'True
      TabIndex        =   45
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox cboEducation 
      Height          =   315
      Left            =   6840
      TabIndex        =   23
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox txtSisters 
      Height          =   285
      Left            =   15720
      MaxLength       =   2
      TabIndex        =   43
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox txtBrothers 
      Height          =   285
      Left            =   13920
      MaxLength       =   2
      TabIndex        =   42
      Text            =   " "
      Top             =   3600
      Width           =   495
   End
   Begin VB.ComboBox cboSiblings 
      Height          =   315
      Left            =   11640
      TabIndex        =   41
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtMothers 
      Height          =   285
      Left            =   12360
      TabIndex        =   36
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtFathers 
      Height          =   285
      Left            =   12240
      TabIndex        =   35
      Top             =   1680
      Width           =   3015
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
      Left            =   12000
      TabIndex        =   2
      ToolTipText     =   "Add new record"
      Top             =   8520
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
      TabIndex        =   1
      ToolTipText     =   "Go to Last record"
      Top             =   8040
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
      TabIndex        =   0
      ToolTipText     =   "Go to Next record"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.ComboBox cboOccupation 
      Height          =   315
      Left            =   1800
      TabIndex        =   26
      Top             =   6360
      Width           =   3495
   End
   Begin VB.ComboBox cboReligion 
      DataSource      =   " "
      Height          =   315
      Left            =   8040
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtTelCode 
      Height          =   285
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   24
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox txtHome 
      Height          =   285
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   25
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox txtPCode2 
      Height          =   315
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   21
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtPCode1 
      DataSource      =   " "
      Height          =   315
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   14
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ComboBox cboCountry2 
      Height          =   315
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox cboState2 
      Height          =   315
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox cboCity2 
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox cboHealthy 
      Height          =   315
      Left            =   17040
      TabIndex        =   40
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cboCriminal 
      Height          =   315
      Left            =   17040
      TabIndex        =   38
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox cboGender 
      DataSource      =   " "
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   1800
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
      Left            =   18360
      TabIndex        =   44
      ToolTipText     =   "Upload image"
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ComboBox cboState1 
      DataSource      =   " "
      Height          =   315
      Left            =   4320
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox cboCountry1 
      DataSource      =   " "
      Height          =   315
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtLastName 
      DataSource      =   " "
      Height          =   285
      Left            =   6480
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtContact 
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   22
      Top             =   5400
      Width           =   2175
   End
   Begin VB.ComboBox cboCity1 
      DataSource      =   " "
      Height          =   315
      Left            =   840
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox picSelf 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   18360
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   226
      Top             =   1800
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtdob 
      DataSource      =   " "
      Height          =   315
      Left            =   4440
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   124780545
      CurrentDate     =   42724
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
      Left            =   6960
      TabIndex        =   63
      ToolTipText     =   "Save record"
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
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
      TabIndex        =   223
      ToolTipText     =   "Cancel"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.PictureBox picResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   18360
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   227
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   -120
      TabIndex        =   231
      Top             =   9000
      Visible         =   0   'False
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   11
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdOpenExisting 
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
      Left            =   12000
      TabIndex        =   234
      ToolTipText     =   "Edit record"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSearchDOB 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(mm/dd/yyyy)"
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
      Left            =   11640
      TabIndex        =   233
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   18240
      TabIndex        =   230
      Top             =   3960
      Width           =   1815
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
      Left            =   18240
      TabIndex        =   229
      Top             =   3600
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyFavouritePlaces 
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
      Height          =   375
      Left            =   12480
      TabIndex        =   142
      Top             =   7080
      Width           =   7215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyHobbies 
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
      Height          =   375
      Left            =   11760
      TabIndex        =   141
      Top             =   6480
      Width           =   7575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyMaritalStatus 
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
      Height          =   375
      Left            =   17640
      TabIndex        =   140
      Top             =   6000
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyEmployer 
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
      Left            =   14880
      TabIndex        =   138
      Top             =   4440
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyJobStatus 
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
      Left            =   11880
      TabIndex        =   137
      Top             =   4440
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyEducationHighest 
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
      Left            =   6840
      TabIndex        =   136
      Top             =   5400
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMySisters 
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
      Left            =   15720
      TabIndex        =   135
      Top             =   3600
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyBrothers 
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
      Left            =   14040
      TabIndex        =   134
      Top             =   3600
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMySiblings 
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
      Left            =   11640
      TabIndex        =   133
      Top             =   3600
      Width           =   735
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyMothersName 
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
      Height          =   375
      Left            =   12360
      TabIndex        =   132
      Top             =   2160
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyFathersName 
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
      Height          =   375
      Left            =   12360
      TabIndex        =   131
      Top             =   1680
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyMothersOccupation 
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
      Left            =   12840
      TabIndex        =   130
      Top             =   3120
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyFathersOccupation 
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
      Left            =   12840
      TabIndex        =   129
      Top             =   2640
      Width           =   3015
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
      Left            =   17040
      TabIndex        =   128
      Top             =   3120
      Width           =   1215
      WordWrap        =   -1  'True
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
      Left            =   17040
      TabIndex        =   127
      Top             =   2640
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblcom 
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
      Height          =   255
      Left            =   5760
      TabIndex        =   126
      Top             =   6840
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyEmailID 
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
      Left            =   1320
      TabIndex        =   125
      Top             =   6840
      Width           =   5055
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyOccupation 
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
      Left            =   1680
      TabIndex        =   124
      Top             =   6360
      Width           =   3615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyHomeContact 
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
      Left            =   2880
      TabIndex        =   123
      Top             =   5880
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyHomeCode 
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
      Left            =   1920
      TabIndex        =   122
      Top             =   5880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyCCode 
      BackStyle       =   0  'Transparent
      Caption         =   "+91"
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   121
      Top             =   5400
      Width           =   375
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
      Height          =   255
      Left            =   2040
      TabIndex        =   120
      Top             =   5400
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyReligion 
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
      Left            =   8040
      TabIndex        =   117
      Top             =   1800
      Width           =   1695
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
      Height          =   255
      Left            =   4440
      TabIndex        =   116
      Top             =   1800
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
      Height          =   255
      Left            =   1320
      TabIndex        =   115
      Top             =   1800
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
      Height          =   255
      Left            =   6480
      TabIndex        =   114
      Top             =   1320
      Width           =   3135
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
      Height          =   375
      Left            =   1560
      TabIndex        =   113
      Top             =   1320
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblGeneralUsername 
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
      Height          =   375
      Left            =   1560
      TabIndex        =   112
      Top             =   7800
      Width           =   5055
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUsernameGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Username: "
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
      Left            =   360
      TabIndex        =   111
      Top             =   7800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2760
      TabIndex        =   108
      Top             =   5880
      Width           =   135
      WordWrap        =   -1  'True
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   107
      Top             =   5400
      Width           =   135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMothersO 
      BackStyle       =   0  'Transparent
      Caption         =   "Mother's Occupation:"
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
      Left            =   10680
      TabIndex        =   106
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblFathersO 
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Occupation:"
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
      Left            =   10680
      TabIndex        =   105
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   19920
      Y1              =   7680
      Y2              =   7680
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
      Left            =   360
      TabIndex        =   104
      Top             =   6840
      Width           =   1215
      WordWrap        =   -1  'True
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
      Left            =   4080
      TabIndex        =   103
      Top             =   6840
      Width           =   255
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   19920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblPCode3 
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
      Left            =   10680
      TabIndex        =   102
      Top             =   6000
      Width           =   1575
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
      TabIndex        =   101
      Top             =   240
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
      TabIndex        =   100
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblCountry3 
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
      Left            =   16680
      TabIndex        =   99
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lblState3 
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
      TabIndex        =   98
      Top             =   5520
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
      Left            =   10680
      TabIndex        =   97
      Top             =   5520
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
      Left            =   10680
      TabIndex        =   96
      Top             =   4920
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
      Left            =   13800
      TabIndex        =   95
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter number of brothers and sisters you have."
      Height          =   495
      Left            =   12960
      TabIndex        =   94
      Top             =   3960
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
      TabIndex        =   93
      Top             =   6360
      Width           =   1215
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
      TabIndex        =   92
      Top             =   4920
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
      TabIndex        =   91
      Top             =   3360
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
      Left            =   4680
      TabIndex        =   90
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label lblSiblings 
      Alignment       =   2  'Center
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
      Left            =   10560
      TabIndex        =   89
      Top             =   3600
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
      Left            =   6960
      TabIndex        =   88
      Top             =   4440
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
      Left            =   3720
      TabIndex        =   87
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label lblCity2 
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
      Left            =   360
      TabIndex        =   86
      Top             =   4440
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
      TabIndex        =   85
      Top             =   5880
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
      Left            =   10680
      TabIndex        =   84
      Top             =   7080
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
      Left            =   10680
      TabIndex        =   83
      Top             =   6480
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
      Left            =   10680
      TabIndex        =   82
      Top             =   4440
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
      Left            =   7080
      TabIndex        =   81
      Top             =   1800
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
      TabIndex        =   80
      Top             =   3840
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
      TabIndex        =   79
      Top             =   2280
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
      Left            =   16080
      TabIndex        =   78
      Top             =   6000
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
      Left            =   14880
      TabIndex        =   77
      Top             =   3600
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
      Left            =   12960
      TabIndex        =   76
      Top             =   3600
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
      Left            =   10680
      TabIndex        =   75
      Top             =   2160
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
      Left            =   10680
      TabIndex        =   74
      Top             =   1680
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
      Left            =   16080
      TabIndex        =   73
      Top             =   3120
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
      Left            =   16080
      TabIndex        =   72
      Top             =   2640
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
      TabIndex        =   71
      Top             =   5400
      Width           =   855
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
      Left            =   6960
      TabIndex        =   70
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
      Left            =   3720
      TabIndex        =   69
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
      TabIndex        =   68
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
      Left            =   3720
      TabIndex        =   67
      Top             =   1800
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
      TabIndex        =   66
      Top             =   1800
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
      Left            =   5280
      TabIndex        =   65
      Top             =   1320
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
      TabIndex        =   64
      Top             =   1320
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   10200
      X2              =   10200
      Y1              =   1080
      Y2              =   7680
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "General Database Management System"
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
      TabIndex        =   62
      Top             =   600
      Width           =   6855
   End
   Begin VB.Label lblMyEmployersAddress 
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
      Height          =   495
      Left            =   10680
      TabIndex        =   139
      Top             =   5400
      Width           =   8535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyPermanentAddress 
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
      Height          =   495
      Left            =   360
      TabIndex        =   118
      Top             =   2760
      Width           =   9135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMyResidentialAddress 
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
      Height          =   495
      Left            =   360
      TabIndex        =   119
      Top             =   4320
      Width           =   9255
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private con As New ADODB.Connection
Private rs As New ADODB.Recordset
Private lb As New ADODB.Recordset
Private cu As New ADODB.Recordset
Private str As String

'Adding Sub Procedures for MDI Main in order to access from MDI Main
Sub AddNewGeneral()

    'This will be used in place of cmdAddNew
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
        
        If cmdAddNew.Visible = True Then
        
            'rs.Close
            lb.Close
        
            'Hide the Labels
            HideLabels
            
            'Show Objects
            ShowObjects
            
            Set rs = New Recordset
            rs.Open "Select * from Data", con, adOpenDynamic, adLockOptimistic
            
            rs.AddNew
            Clear
            
            'Entering Text in Few address boxes
            txtPermanentPlotNo.Text = "Plot No"
            txtPermanentBuildingName.Text = "Building Name"
            txtPermanentLocality.Text = "Locality"
            txtResidentialPlotNo.Text = "Plot No"
            txtResidentialBuildingName.Text = "Building Name"
            txtResidentialLocality.Text = "Locality"
            txtEmployerPlotNo.Text = "Plot No"
            txtEmployerBuildingName.Text = "Building Name"
            txtEmployerLocality.Text = "Locality"
            
            'Altering some properties
            cmdEdit.Visible = False
            cmdSave.Visible = True
            cmdSave.TabIndex = "58"
            
            cmdAddNew.Visible = False
            cmdCancel.Visible = True
            cmdCancel.TabIndex = "2"
            
            cmdDelete.Visible = False
            cmdClear.Visible = True
            cmdClear.TabIndex = "60"
            
            cmdFirst.Visible = False
            cmdPrevious.Visible = False
            cmdNext.Visible = False
            cmdLast.Visible = False
            
        Else
        
            MsgBox "Please save the file before adding new", vbExclamation, "Invalid Operation"
            
        End If
    End If
        
End Sub

Sub DeleteGeneral()

    'This will be used in place of cmdDelete
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdDelete.Visible = True Then

            If MsgBox("Are you sure?", vbCritical + vbYesNo, "Note") = vbYes Then
            
                lb.Delete adAffectCurrent
                MsgBox "Deleted", vbInformation, "Success"
                lb.Update
                lb.MoveNext
                If Not lb.EOF Then
                
                    DisplayLB
                    txtRecordID.Text = lb!ID
                    
                Else
                
                    lb.MoveFirst
                    DisplayLB
                    txtRecordID.Text = lb!ID
                    
                End If
            End If
            
        Else
        
            MsgBox "Please switch format", vbCritical, "Invalid operation"
            
        End If
    End If

End Sub

Sub SaveGeneral()

    'This will be used in place of cmdSave
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdSave.Visible = True Then
        
            ProgressBar.Visible = True
            ProgressBar.Value = 0
        
            'Erasing Textual matter from Textboxes(1-41) leaving 4th special case
            'This will prepare this form for validating input
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text10.Text = ""
            Text11.Text = ""
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text30.Text = ""
            Text31.Text = ""
            Text32.Text = ""
            Text33.Text = ""
            Text34.Text = ""
            Text35.Text = ""
            Text36.Text = ""
            Text37.Text = ""
            Text38.Text = ""
            Text39.Text = ""
            Text40.Text = ""
            Text41.Text = ""
            Text42.Text = ""
            
            ProgressBar.Value = 10
            
            'Invisible Correction Picture boxes
            Picture1.Visible = False
            Picture2.Visible = False
            Picture3.Visible = False
            Picture4.Visible = False
            Picture5.Visible = False
            Picture6.Visible = False
            Picture7.Visible = False
            Picture8.Visible = False
            Picture9.Visible = False
            Picture10.Visible = False
            Picture11.Visible = False
            Picture12.Visible = False
            Picture13.Visible = False
            Picture14.Visible = False
            Picture15.Visible = False
            Picture16.Visible = False
            Picture17.Visible = False
            Picture18.Visible = False
            Picture19.Visible = False
            Picture20.Visible = False
            Picture21.Visible = False
            Picture21.Visible = False
            Picture22.Visible = False
            Picture23.Visible = False
            Picture24.Visible = False
            Picture25.Visible = False
            Picture26.Visible = False
            Picture27.Visible = False
            Picture28.Visible = False
            Picture29.Visible = False
            Picture30.Visible = False
            Picture31.Visible = False
            Picture32.Visible = False
            Picture33.Visible = False
            Picture34.Visible = False
            Picture35.Visible = False
            Picture36.Visible = False
            Picture37.Visible = False
            Picture38.Visible = False
            Picture39.Visible = False
            Picture40.Visible = False
            
            
            'Validation START with Cross Picture
            
            'First Name
            If Not IsAlphabetical(Trim(txtFirstName.Text)) Or Trim(txtFirstName.Text) = "" Then
            
                'MsgBox "Please re-enter your first name", vbCritical, "Invalid field"
                Picture1.Visible = True
                    
            Else
                    
                'First Name is ok
                Text1.Text = "ok"
                    
            End If
            
            
            'Last Name
            If Not IsAlphabetical(Trim(txtLastName.Text)) Or Trim(txtLastName.Text) = "" Then
            
                Picture2.Visible = True
                    
            Else
                    
                Text2.Text = "ok"
                    
            End If
            
            ProgressBar.Value = 20
            
            'Gender
            If Trim(cboGender.Text) <> "" Then
        
                If Trim(cboGender.Text) = "Male" Or Trim(cboGender.Text) = "Female" Then
                
                    'Gender is ok
                    Text3.Text = "ok"
                    
                Else
                    
                    Picture3.Visible = True
                        
                End If
                
            Else
                    
                Picture3.Visible = True
                
            End If
                    
                
            'Validating DOB
            If dtDOB.Value = "12/20/2016" Then
                
                If Text4.Text <> "ok" Then
                
                    If MsgBox("Is your DOB, 12/20/2016?", vbInformation + vbYesNo, "Note") = _
                    vbNo Then
                    
                        MsgBox "Please cboose your DOB", vbExclamation, "Invalid field"
                        Picture4.Visible = True
                        
                    Else
                        
                        Text4.Text = "ok"
                            
                    End If
                End If
                    
            Else
                    
                'DOB is ok
                Text4.Text = "ok"
                    
            End If
            
            
            'Religion
            If Trim(cboReligion.Text) <> "" Then
        
                If Trim(cboReligion.Text) = "Hinduism" Or Trim(cboReligion.Text) = "Islam" Or _
                Trim(cboReligion.Text) = "Christianity" Or Trim(cboReligion.Text) = "Sikhism" Or _
                Trim(cboReligion.Text) = "Buddhism" Or Trim(cboReligion.Text) = "Jainism" Then
                
                    Text5.Text = "ok"
                    
                Else
                    
                    Picture5.Visible = True
                        
                End If
                
            Else
                    
                Picture5.Visible = True
                
            End If
            
            ProgressBar.Value = 30
            
            'Permanent Address
            If Trim(txtPermanentPlotNo.Text) <> "" And Trim(txtPermanentBuildingName.Text) <> "" And _
            Trim(txtPermanentLocality.Text) <> "" Then
            
                If Not IsNumeric(Trim(txtPermanentPlotNo.Text)) Then
                
                    Picture6.Visible = True
                    
                Else
                    
                    Text6.Text = "ok"
                    
                End If
                
                If Not IsAlphabetical(Trim(txtPermanentBuildingName.Text)) Then
                
                    Picture6.Visible = True
                    
                Else
                
                    If Text6.Text = "" Then
                    
                        Text6.Text = ""
                        
                    Else
                        
                        Text6.Text = "ok"
                        
                    End If
                End If
                
                If Not IsAlphabetical(Trim(txtPermanentLocality.Text)) Then
                
                    Picture6.Visible = True
                    
                Else
                
                    If Text6.Text = "" Then
                    
                        Text6.Text = ""
                        
                    Else
                        
                        Text6.Text = "ok"
                        
                    End If
                End If
                
            Else
            
                Picture6.Visible = True
                
            End If
            
            
            'Permanent City State Country PostalCode
            'Validating City
                If Trim(cboCity1.Text) <> "" Then
                
                    If cboCity1.Text = "Mumbai" Or cboCity1.Text = "Navi Mumbai" Or _
                    cboCity1.Text = "Pune" Or cboCity1.Text = "Nagpur" Or _
                    cboCity1.Text = "Thane" Or cboCity1.Text = "Pimpri" Or _
                    cboCity1.Text = "Chinchwad" Or cboCity1.Text = "Nashik" Or _
                    cboCity1.Text = "Kalyan" Or cboCity1.Text = "Dombivali" Or _
                    cboCity1.Text = "Pune" Or cboCity1.Text = "Vasai" Or _
                    cboCity1.Text = "Virar" Or cboCity1.Text = "Aurangabad" Or _
                    cboCity1.Text = "Solapur" Or cboCity1.Text = "Mira" Or _
                    cboCity1.Text = "Bhayandar" Or cboCity1.Text = "Bhiwandi" Or _
                    cboCity1.Text = "Nizampur" Or cboCity1.Text = "Amravati" Or _
                    cboCity1.Text = "Nanded" Or cboCity1.Text = "Waghala" Or _
                    cboCity1.Text = "Panvel" Or cboCity1.Text = "Sangli" Or _
                    cboCity1.Text = "Akola" Or cboCity1.Text = "Ahmednagar" Or _
                    cboCity1.Text = "Parbhani" Or cboCity1.Text = "Chandrapur" Or _
                    cboCity1.Text = "Dhule" Or cboCity1.Text = "Malegaon" Or _
                    cboCity1.Text = "Jalgaon" Or cboCity1.Text = "Kolhapur" Or _
                    cboCity1.Text = "Nashik" Or cboCity1.Text = "Latur" Then
                            
                        Text7.Text = "ok"
                        
                    Else
                    
                        Picture7.Visible = True
                        
                    End If
                    
                Else
                    
                    'MsgBox "Please select from given cities", vbCritical, "Note"
                    Picture7.Visible = True
                    
                End If
            
                
                'Validating State
                If Trim(cboState1.Text) <> "" Then
                
                    If Trim(cboState1.Text) = "Maharashtra" Then
                
                    'State is ok
                    Text8.Text = "ok"
                    
                    Else
                
                        'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
                        Picture8.Visible = True
                        
                    End If
                
                
                Else
                
                    'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
                    Picture8.Visible = True
                        
                End If
                
                
                'Validating Country
                If Trim(cboCountry1.Text) <> "" Then
                    
                    If Trim(cboCountry1.Text) = "India" Then
                
                        'Country is ok
                        Text9.Text = "ok"
                
                    Else
                    
                        'MsgBox "Please select from given Country", vbCritical, "Note"
                        Picture9.Visible = True
                    
                    End If
                    
                Else
                
                    Picture9.Visible = True
                    
                End If
                
            
            'Permanent Postal Code
            If Not IsNumeric(Trim(txtPCode1.Text)) Or Trim(txtPCode1.Text) = "" Then
            
                Picture10.Visible = True
                
            Else
            
                Text10.Text = "ok"
                
            End If
            
            ProgressBar.Value = 40
            
            'Residential Address
            If Trim(txtResidentialPlotNo.Text) <> "" And _
            Trim(txtResidentialBuildingName.Text) <> "" And Trim(txtResidentialLocality.Text) <> "" Then
            
                If Not IsNumeric(Trim(txtResidentialPlotNo.Text)) Then
                
                    Picture11.Visible = True
                    
                Else
                    
                    Text11.Text = "ok"
                    
                End If
                
                If Not IsAlphabetical(Trim(txtResidentialBuildingName.Text)) Then
                
                    Picture11.Visible = True
                    
                Else
                
                    If Text11.Text = "" Then
                    
                        Text11.Text = ""
                        
                    Else
                        
                        Text11.Text = "ok"
                        
                    End If
                End If
                
                If Not IsAlphabetical(Trim(txtResidentialLocality.Text)) Then
                
                    Picture11.Visible = True
                    
                Else
                
                    If Text11.Text = "" Then
                    
                        Text11.Text = ""
                        
                    Else
                        
                        Text11.Text = "ok"
                        
                    End If
                End If
                
            Else
            
                Picture11.Visible = True
                
            End If
            
            'Residential City State Country PostalCode
            'Validating City
                If Trim(cboCity2.Text) <> "" Then
                
                    If cboCity2.Text = "Mumbai" Or cboCity2.Text = "Navi Mumbai" Or _
                    cboCity2.Text = "Pune" Or cboCity2.Text = "Nagpur" Or _
                    cboCity2.Text = "Thane" Or cboCity2.Text = "Pimpri" Or _
                    cboCity2.Text = "Chinchwad" Or cboCity2.Text = "Nashik" Or _
                    cboCity2.Text = "Kalyan" Or cboCity2.Text = "Dombivali" Or _
                    cboCity2.Text = "Pune" Or cboCity2.Text = "Vasai" Or _
                    cboCity2.Text = "Virar" Or cboCity2.Text = "Aurangabad" Or _
                    cboCity2.Text = "Solapur" Or cboCity2.Text = "Mira" Or _
                    cboCity2.Text = "Bhayandar" Or cboCity2.Text = "Bhiwandi" Or _
                    cboCity2.Text = "Nizampur" Or cboCity2.Text = "Amravati" Or _
                    cboCity2.Text = "Nanded" Or cboCity2.Text = "Waghala" Or _
                    cboCity2.Text = "Panvel" Or cboCity2.Text = "Sangli" Or _
                    cboCity2.Text = "Akola" Or cboCity2.Text = "Ahmednagar" Or _
                    cboCity2.Text = "Parbhani" Or cboCity2.Text = "Chandrapur" Or _
                    cboCity2.Text = "Dhule" Or cboCity2.Text = "Malegaon" Or _
                    cboCity2.Text = "Jalgaon" Or cboCity2.Text = "Kolhapur" Or _
                    cboCity2.Text = "Nashik" Or cboCity2.Text = "Latur" Then
                            
                        Text12.Text = "ok"
                        
                    Else
                    
                        Picture12.Visible = True
                        
                    End If
                    
                Else
                    
                    'MsgBox "Please select from given cities", vbCritical, "Note"
                    Picture12.Visible = True
                    
                End If
            
                
                'Validating State
                If Trim(cboState2.Text) <> "" Then
                
                    If Trim(cboState2.Text) = "Maharashtra" Then
                
                    'State is ok
                    Text13.Text = "ok"
                    
                    Else
                
                        'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
                        Picture13.Visible = True
                        
                    End If
                
                
                Else
                
                    'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
                    Picture13.Visible = True
                        
                End If
                
                
                'Validating Country
                If Trim(cboCountry2.Text) <> "" Then
                    
                    If Trim(cboCountry2.Text) = "India" Then
                
                        'Country is ok
                        Text14.Text = "ok"
                
                    Else
                    
                        'MsgBox "Please select from given Country", vbCritical, "Note"
                        Picture14.Visible = True
                    
                    End If
                    
                Else
                
                    Picture14.Visible = True
                    
                End If
                
            
            'Residential Postal Code
            If Not IsNumeric(Trim(txtPCode2.Text)) Or Trim(txtPCode2.Text) = "" Then
            
                Picture15.Visible = True
                
            Else
            
                Text15.Text = "ok"
                
            End If
            
            ProgressBar.Value = 50
            
            'Validating Contact
            If Not IsNumeric(Trim(txtContact.Text)) Or Len(Trim(txtContact.Text)) <> 10 Then
            
                Picture16.Visible = True
                'MsgBox "Please enter valid contact", vbCritical, "Invalid contact"
                
            Else
                
                'Contact is ok
                Text16.Text = "ok"
                    
            End If
            
            
            'HomeContact (Dial Code + Number)
            If Not IsNumeric(Trim(txtTelCode.Text)) Or Trim(txtTelCode.Text) = "" Then
            
                Picture17.Visible = True
                'MsgBox "Please enter valid contact", vbCritical, "Invalid contact"
                    
            Else
                
                If Not IsNumeric(Trim(txtHome.Text)) Or Trim(txtHome.Text) = "" Then
                
                    Picture17.Visible = True
                    
                Else
                
                    'Contact is ok
                    Text17.Text = "ok"
                    
                End If
            End If
            
            
            'Occupation
            If Trim(cboOccupation.Text) <> "" Then
            
                If cboOccupation.Text = "Agriculture" Or cboOccupation.Text = "Business" Or _
                cboOccupation.Text = "Medical" Or cboOccupation.Text = "Engineering" Or _
                cboOccupation.Text = "Law Practice" Or _
                cboOccupation.Text = "Government Service" Or _
                cboOccupation.Text = "Public Sector Service" Or _
                cboOccupation.Text = "Private Service" Or cboOccupation.Text = "Teaching" Or _
                cboOccupation.Text = "Other" Then
                
                    Text18.Text = "ok"
                        
                Else
                    
                    Picture18.Visible = True
                        
                End If
                    
            Else
                    
                Picture18.Visible = True
                    
            End If
            
            
            'Email Id + Provider
            If Trim(txtID.Text) = "" Or Trim(cboProvider.Text) = "" Then
            
                Picture19.Visible = True
                'MsgBox "Please re-enter your email id", vbCritical, "Unequal email ids"
                    
            Else
                    
                If Trim(cboProvider.Text) = "Gmail" Or Trim(cboProvider.Text) = "iCloud" Or _
                Trim(cboProvider.Text) = "GMX" Or Trim(cboProvider.Text) = "Outlook" Or _
                Trim(cboProvider.Text) = "Yahoo" Or Trim(cboProvider.Text) = "Aol" Or _
                Trim(cboProvider.Text) = "Zoho" Or Trim(cboProvider.Text) = "Mail" Or _
                Trim(cboProvider.Text) = "Yandex" Or Trim(cboProvider.Text) = "ProtonMail" Then
                    
                    'Email ids are ok
                    Text19.Text = "ok"
                    
                Else
                
                    Picture19.Visible = True
        
                End If
            End If
            
            
            'Criminal & Healthy
            If Trim(cboCriminal.Text) <> "" Then
                
                If cboCriminal.Text = "Yes" Or cboCriminal.Text = "No" Then
                            
                    Text20.Text = "ok"
                        
                Else
                    
                    Picture20.Visible = True
                        
                End If
                    
            Else
                    
                'MsgBox "Please select from given cities", vbCritical, "Note"
                Picture20.Visible = True
                    
            End If
                
            If Trim(cboHealthy.Text) <> "" Then
                
                If cboHealthy.Text = "Yes" Or cboCity2.Text = "No" Then
                            
                    Text21.Text = "ok"
                        
                Else
                    
                    Picture21.Visible = True
                        
                End If
                    
            Else
                    
                'MsgBox "Please select from given cities", vbCritical, "Note"
                Picture21.Visible = True
                    
            End If
            
            
            'Father's Name
            If Not IsAlphabetical(Trim(txtFathers.Text)) Or Trim(txtFathers.Text) = "" Then
            
                'MsgBox "Please re-enter your Fathers name", vbCritical, "Invalid field"
                Picture22.Visible = True
                    
            Else
                    
                'Fathers Name is ok
                Text22.Text = "ok"
                    
            End If
            
            
            'Mother's Name
            If Not IsAlphabetical(Trim(txtFathers.Text)) Or Trim(txtFathers.Text) = "" Then
            
                'MsgBox "Please re-enter your Mothers name", vbCritical, "Invalid field"
                Picture23.Visible = True
                    
            Else
                    
                'Mothers Name is ok
                Text23.Text = "ok"
                    
            End If
            
            ProgressBar.Value = 60
            
            'Father's & Mother's Occupation
            If Trim(cboFathersO.Text) <> "" Then
            
                If cboFathersO.Text = "Agriculture" Or cboFathersO.Text = "Business" Or _
                cboFathersO.Text = "Medical" Or cboFathersO.Text = "Engineering" Or _
                cboFathersO.Text = "Law Practice" Or cboFathersO.Text = "Government Service" Or _
                cboFathersO.Text = "Public Sector Service" Or _
                cboFathersO.Text = "Private Service" Or cboFathersO.Text = "Teaching" Or _
                cboFathersO.Text = "Other" Then
                
                    Text24.Text = "ok"
                        
                Else
                    
                    Picture24.Visible = True
                        
                End If
                    
            Else
                    
                Picture24.Visible = True
                    
            End If
            
            
            If Trim(cboMothersO.Text) <> "" Then
            
                If cboMothersO.Text = "Agriculture" Or cboMothersO.Text = "Business" Or _
                cboMothersO.Text = "Medical" Or cboMothersO.Text = "Engineering" Or _
                cboMothersO.Text = "Law Practice" Or cboMothersO.Text = "Government Service" Or _
                cboMothersO.Text = "Public Sector Service" Or cboMothersO.Text = "Private Service" Or _
                cboMothersO.Text = "Teaching" Or cboMothersO.Text = "Other" Then
                
                    Text25.Text = "ok"
                        
                Else
                    
                    Picture25.Visible = True
                        
                End If
                    
            Else
                    
                Picture25.Visible = True
                    
            End If
            
            
            
            'Siblings
            If Trim(cboSiblings.Text) <> "" Then
                
                If cboSiblings.Text = "Yes" Or cboSiblings.Text = "No" Then
                            
                    Text26.Text = "ok"
                        
                Else
                    
                    Picture26.Visible = True
                        
                End If
                    
            Else
                    
                'MsgBox "Please select from given cities", vbCritical, "Note"
                Picture26.Visible = True
                    
            End If
            
            ProgressBar.Value = 70
            
            'Brothers & Sisters (Numbers)
            If Not IsNumeric(Trim(txtBrothers.Text)) Or Trim(txtBrothers.Text) = "" Then
            
                Picture27.Visible = True
                
            Else
            
                Text27.Text = "ok"
                
            End If
            
            If Not IsNumeric(Trim(txtSisters.Text)) Or Trim(txtSisters.Text) = "" Then
            
                Picture28.Visible = True
                
            Else
            
                Text28.Text = "ok"
                
            End If
            
            ProgressBar.Value = 80
            
            'Highest Education
            If Trim(cboEducation.Text) <> "" Then
            
                If cboEducation.Text = "1st Standard" Or cboEducation.Text = "2nd Standard" Or _
                cboEducation.Text = "3rd Standard" Or cboEducation.Text = "4th Standard" Or _
                cboEducation.Text = "5th Standard" Or cboEducation.Text = "6th Standard" Or _
                cboEducation.Text = "7th Standard" Or cboEducation.Text = "8th Standard" Or _
                cboEducation.Text = "9th Standard" Or cboEducation.Text = "10th Standard" Or _
                cboEducation.Text = "11th Standard" Or cboEducation.Text = "12th Standard" Or _
                cboEducation.Text = "B.E" Or cboEducation.Text = "B.Tech" Or _
                cboEducation.Text = "B.Arch" Or cboEducation.Text = "B.Sc" Or _
                cboEducation.Text = "B.Com" Or cboEducation.Text = "B.B.A" Or _
                cboEducation.Text = "B.C.C.A" Or cboEducation.Text = "M.Tech" Or _
                cboEducation.Text = "M.Sc" Or cboEducation.Text = "M.C.M" Or _
                cboEducation.Text = "M.B.A" Then
                
                    Text29.Text = "ok"
                        
                Else
                    
                    Picture29.Visible = True
                        
                End If
                    
            Else
                    
                Picture29.Visible = True
                    
            End If
                                          
                                           
            'Job Status
            If Trim(cboJob.Text) <> "" Then
                
                If cboJob.Text = "Working" Or cboJob.Text = "Idle" Then
                            
                    Text30.Text = "ok"
                        
                Else
                    
                    Picture30.Visible = True
                        
                End If
                    
            Else
                    
                'MsgBox "Please select from given Job", vbCritical, "Note"
                Picture30.Visible = True
                    
            End If
            
            
            'The following code should only work when the cboJob.text = Working
            
            If Trim(cboJob.Text) = "" Then
            
                Picture31.Visible = True
                Picture32.Visible = True
                Picture33.Visible = True
                Picture34.Visible = True
                Picture35.Visible = True
                Picture36.Visible = True
                
            Else
            
                If cboJob.Text = "Working" Then
                
                    'Employer
                    If Not IsAlphabetical(Trim(txtEmployer.Text)) Or Trim(txtEmployer.Text) = "" Then
                    
                        'MsgBox "Please re-enter your Employer's name", vbCritical, "Invalid field"
                        Picture31.Visible = True
                            
                    Else
                            
                        'Employer name is ok
                        Text31.Text = "ok"
                            
                    End If
                    
                
                    'Employer's Address
                    If Trim(txtEmployerPlotNo.Text) <> "" And _
                    Trim(txtEmployerBuildingName.Text) <> "" And _
                    Trim(txtEmployerLocality.Text) <> "" Then
                    
                        If Not IsNumeric(Trim(txtEmployerPlotNo.Text)) Then
                        
                            Picture32.Visible = True
                            
                        Else
                            
                            Text32.Text = "ok"
                            
                        End If
                        
                        If Not IsAlphabetical(Trim(txtEmployerBuildingName.Text)) Then
                        
                            Picture32.Visible = True
                            
                        Else
                        
                            If Text32.Text = "" Then
                            
                                Text32.Text = ""
                                
                            Else
                                
                                Text32.Text = "ok"
                                
                            End If
                            
                        End If
                        
                        If Not IsAlphabetical(Trim(txtEmployerLocality.Text)) Then
                        
                            Picture32.Visible = True
                            
                        Else
                        
                            If Text32.Text = "" Then
                            
                                Text32.Text = ""
                                
                            Else
                                
                                Text32.Text = "ok"
                                
                            End If
                            
                        End If
                        
                    Else
                        
                        Picture32.Visible = True
                        
                    End If
                    
                    
                    'Employer City State Country PostalCode
                    'Validating City
                        If Trim(cboCity3.Text) <> "" Then
                        
                            If cboCity3.Text = "Mumbai" Or cboCity3.Text = "Navi Mumbai" Or _
                            cboCity3.Text = "Pune" Or cboCity3.Text = "Nagpur" Or _
                            cboCity3.Text = "Thane" Or cboCity3.Text = "Pimpri" Or _
                            cboCity3.Text = "Chinchwad" Or cboCity3.Text = "Nashik" Or _
                            cboCity3.Text = "Kalyan" Or cboCity3.Text = "Dombivali" Or _
                            cboCity3.Text = "Pune" Or cboCity3.Text = "Vasai" Or _
                            cboCity3.Text = "Virar" Or cboCity3.Text = "Aurangabad" Or _
                            cboCity3.Text = "Solapur" Or cboCity3.Text = "Mira" Or _
                            cboCity3.Text = "Bhayandar" Or cboCity3.Text = "Bhiwandi" Or _
                            cboCity3.Text = "Nizampur" Or cboCity3.Text = "Amravati" Or _
                            cboCity3.Text = "Nanded" Or cboCity3.Text = "Waghala" Or _
                            cboCity3.Text = "Panvel" Or cboCity3.Text = "Sangli" Or _
                            cboCity3.Text = "Akola" Or cboCity3.Text = "Ahmednagar" Or _
                            cboCity3.Text = "Parbhani" Or cboCity3.Text = "Chandrapur" Or _
                            cboCity3.Text = "Dhule" Or cboCity3.Text = "Malegaon" Or _
                            cboCity3.Text = "Jalgaon" Or cboCity3.Text = "Kolhapur" Or _
                            cboCity3.Text = "Nashik" Or cboCity3.Text = "Latur" Then
                                    
                                Text33.Text = "ok"
                                
                            Else
                            
                                Picture33.Visible = True
                                
                            End If
                            
                        Else
                            
                            'MsgBox "Please select from given cities", vbCritical, "Note"
                            Picture33.Visible = True
                            
                        End If
                    
                        
                        'Validating State
                        If Trim(cboState3.Text) <> "" Then
                        
                            If Trim(cboState3.Text) = "Maharashtra" Then
                        
                                'State is ok
                                Text34.Text = "ok"
                            
                            Else
                        
                                'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
                                Picture34.Visible = True
                                
                            End If
                        
                        
                        Else
                        
                            'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
                            Picture34.Visible = True
                                
                        End If
                        
                        
                        'Validating Country
                        If Trim(cboCountry3.Text) <> "" Then
                            
                            If Trim(cboCountry3.Text) = "India" Then
                        
                                'Country is ok
                                Text35.Text = "ok"
                        
                            Else
                            
                                'MsgBox "Please select from given Country", vbCritical, "Note"
                                Picture35.Visible = True
                            
                            End If
                            
                        Else
                        
                            Picture35.Visible = True
                            
                        End If
                        
                    
                    'Employer Postal Code
                    If Not IsNumeric(Trim(txtPCode3.Text)) Or Trim(txtPCode3.Text) = "" Then
                    
                        Picture36.Visible = True
                        
                    Else
                    
                        Text36.Text = "ok"
                        
                    End If
                    
                Else
                
                    txtEmployer.Text = ""
                    txtEmployerPlotNo.Text = ""
                    txtEmployerBuildingName.Text = ""
                    txtEmployerLocality.Text = ""
                    cboCity3.Text = ""
                    cboState3.Text = ""
                    cboCountry3.Text = ""
                    txtPCode3.Text = ""
                    
                    Text31.Text = "ok"
                    Text32.Text = "ok"
                    Text33.Text = "ok"
                    Text34.Text = "ok"
                    Text35.Text = "ok"
                    Text36.Text = "ok"
                    
                End If
            End If
            
            ProgressBar.Value = 90
            
            'Marital Status
             If Trim(cboMarital.Text) <> "" Then
                
                If cboMarital.Text = "Married" Or cboMarital.Text = "Single" Or _
                cboMarital.Text = "Divorced" Or cboMarital.Text = "Widowed" Then
                            
                    Text37.Text = "ok"
                        
                Else
                    
                    Picture37.Visible = True
                        
                End If
                    
            Else
                    
                'MsgBox "Please select from given cities", vbCritical, "Note"
                Picture37.Visible = True
                    
            End If
            
            
            'Hobbies & FavoritePlaces
            If Not IsAlphabetical(Trim(txtHobbies.Text)) Or Trim(txtHobbies.Text) = "" Then
            
                    'MsgBox "Please re-enter your Hobbies", vbCritical, "Invalid field"
                    Picture38.Visible = True
                    
            Else
                    
                    'Employer name is ok
                    Text38.Text = "ok"
                    
            End If
            
            If Not IsAlphabetical(Trim(txtFavPlaces.Text)) Or Trim(txtFavPlaces.Text) = "" Then
            
                    'MsgBox "Please re-enter your Employer's name", vbCritical, "Invalid field"
                    Picture39.Visible = True
                    
            Else
                    
                    'Employer name is ok
                    Text39.Text = "ok"
                    
            End If
            
            
            'Validating Image
            If str = "" Then
            
                Picture40.Visible = True
                
            End If
            
            If rs!Picture = "" Then
            
                Picture40.Visible = True
            
            Else
            
                Text40.Text = "ok"
                
            End If
            
            ProgressBar.Value = 100
            ProgressBar.Visible = False
            
            'END OF VALIDATION
                 
        
            'The Ultimate TEST "ALL OK"
            If Text1.Text = "ok" And Text2.Text = "ok" And Text3.Text = "ok" And _
            Text4.Text = "ok" And Text5.Text = "ok" And Text6.Text = "ok" And _
            Text7.Text = "ok" And Text8.Text = "ok" And Text9.Text = "ok" And _
            Text10.Text = "ok" And Text11.Text = "ok" And Text12.Text = "ok" And _
            Text13.Text = "ok" And Text14.Text = "ok" And Text15.Text = "ok" And _
            Text16.Text = "ok" And Text17.Text = "ok" And Text18.Text = "ok" And _
            Text19.Text = "ok" And Text20.Text = "ok" And Text21.Text = "ok" And _
            Text22.Text = "ok" And Text23.Text = "ok" And Text24.Text = "ok" And _
            Text25.Text = "ok" And Text26.Text = "ok" And Text27.Text = "ok" And _
            Text28.Text = "ok" And Text29.Text = "ok" And Text30.Text = "ok" And _
            Text31.Text = "ok" And Text32.Text = "ok" And Text33.Text = "ok" And _
            Text34.Text = "ok" And Text35.Text = "ok" And Text36.Text = "ok" And _
            Text37.Text = "ok" And Text38.Text = "ok" And Text39.Text = "ok" And _
            Text40.Text = "ok" Then
            
                ProgressBar.Visible = True
                ProgressBar.Value = 0
                
                rs.Fields("ID").Value = txtFirstName.Text & "." & txtLastName.Text
            
                rs.Fields("FirstName").Value = txtFirstName.Text
                rs.Fields("LastName").Value = txtLastName.Text
                rs.Fields("Gender").Value = cboGender.Text
                rs.Fields("Picture").Value = str
                rs.Fields("DOB").Value = dtDOB.Value
                rs.Fields("Religion").Value = cboReligion.Text
                rs.Fields("PermanentPlotNo").Value = txtPermanentPlotNo.Text
                rs.Fields("PermanentBuildingName").Value = txtPermanentBuildingName.Text
                rs.Fields("PermanentLocality").Value = txtPermanentLocality.Text
                rs.Fields("PermanentCity").Value = cboCity1.Text
                rs.Fields("PermanentState").Value = cboState1.Text
                rs.Fields("PermanentCountry").Value = cboCountry1.Text
                rs.Fields("PermanentPostalCode").Value = txtPCode1.Text
                
                ProgressBar.Value = 25
                
                rs.Fields("Permanent").Value = txtPermanentPlotNo.Text & ", " & _
                txtPermanentBuildingName.Text & ", " & txtPermanentLocality.Text & _
                ", " & cboCity1.Text & ", " & cboState1.Text & ", " & _
                cboCountry1.Text & ", " & txtPCode1.Text
                
                rs.Fields("ResidentialPlotNo").Value = txtResidentialPlotNo.Text
                rs.Fields("ResidentialBuildingName").Value = txtResidentialBuildingName.Text
                rs.Fields("ResidentialLocality").Value = txtResidentialLocality.Text
                rs.Fields("ResidentialCity").Value = cboCity2.Text
                rs.Fields("ResidentialState").Value = cboState2.Text
                rs.Fields("ResidentialCountry").Value = cboCountry2.Text
                rs.Fields("ResidentialPostalCode").Value = txtPCode2.Text
                
                ProgressBar.Value = 50
                
                rs.Fields("Residential").Value = txtResidentialPlotNo.Text & ", " & _
                txtResidentialBuildingName.Text & ", " & txtResidentialLocality.Text & _
                ", " & cboCity2.Text & ", " & cboState2.Text & ", " & _
                cboCountry2.Text & ", " & txtPCode2.Text
                
                rs.Fields("Contact").Value = txtContact.Text
                rs.Fields("DialCode").Value = txtTelCode.Text
                rs.Fields("HomeContact").Value = txtHome.Text
                rs.Fields("Occupation").Value = cboOccupation.Text
                rs.Fields("Criminal").Value = cboCriminal.Text
                rs.Fields("Healthy").Value = cboHealthy.Text
                rs.Fields("Email").Value = txtID.Text & "@" & cboProvider.Text & ".com"
                rs.Fields("EmailId").Value = txtID.Text
                rs.Fields("EmailProvider").Value = cboProvider.Text
                rs.Fields("FathersName").Value = txtFathers.Text
                rs.Fields("FathersOccupation").Value = cboFathersO.Text
                rs.Fields("MothersName").Value = txtMothers.Text
                rs.Fields("MothersOccupation").Value = cboMothersO.Text
                rs.Fields("Siblings").Value = cboSiblings.Text
                rs.Fields("Brothers").Value = txtBrothers.Text
                rs.Fields("Sisters").Value = txtSisters.Text
                rs.Fields("Education").Value = cboEducation.Text
                rs.Fields("Job").Value = cboJob.Text
                
                ProgressBar.Value = 75
                
                If cboJob.Text = "Working" Then
                
                    rs.Fields("Employer").Value = txtEmployer.Text
                    rs.Fields("EmployerPlotNo").Value = txtEmployerPlotNo.Text
                    rs.Fields("EmployerBuildingName").Value = txtEmployerBuildingName.Text
                    rs.Fields("EmployerLocality").Value = txtEmployerLocality.Text
                    rs.Fields("EmployerCity").Value = cboCity3.Text
                    rs.Fields("EmployerState").Value = cboState3.Text
                    rs.Fields("EmployerCountry").Value = cboCountry3.Text
                    rs.Fields("EmployerPostalCode").Value = txtPCode3.Text
                    
                    rs.Fields("EmployerA").Value = txtEmployerPlotNo.Text & ", " & _
                    txtEmployerBuildingName.Text & ", " & txtEmployerLocality.Text & _
                    ", " & cboCity3.Text & ", " & cboState3.Text & ", " & cboCountry3.Text & _
                    ", " & txtPCode3.Text
                
                End If
                
                rs.Fields("MaritalStatus").Value = cboMarital.Text
                rs.Fields("Hobbies").Value = txtHobbies.Text
                rs.Fields("FavoritePlaces").Value = txtFavPlaces.Text
                   
                rs.Update
                ProgressBar.Value = 100
                ProgressBar.Visible = False
                
                MsgBox "Success", vbInformation, "Date saved"
                txtRecordID.Text = rs!ID
                
                cmdAddNew.Visible = True
                cmdSave.Visible = False
                cmdEdit.Visible = True
                cmdEdit.TabIndex = "58"
                
                rs.Close
                
                HideObjects
                
                Set lb = New Recordset
                lb.Open "Select * from Data", con, adOpenDynamic, adLockOptimistic
                
                ShowLabels
                
                DisplayLB
                
                cmdFirst.Visible = True
                cmdLast.Visible = True
                cmdNext.Visible = True
                cmdPrevious.Visible = True
                
                cmdSave.Visible = False
                cmdEdit.Visible = True
                cmdEdit.TabIndex = "58"
                
                cmdAddNew.Visible = True
                cmdCancel.Visible = False
                cmdAddNew.TabIndex = "2"
                
                cmdDelete.Visible = True
                cmdClear.Visible = False
                cmdDelete.TabIndex = "60"
            
            Else
                
                MsgBox "Please fill correctly", vbCritical, "Incorrect Data"
                
            End If
            
        Else
        
            MsgBox "Please switch format", vbExclamation, "Invalid operation"
            
        End If
    End If
            
End Sub

Sub ClearGeneral()

    'This will be used in place of cmdClear
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdClear.Visible = True Then
        
            Clear
            
        Else
        
            MsgBox "Please switch format", vbExclamation, "Invalid operation"
            
        End If
    End If

End Sub

Sub EditGeneral()

    'This will be used in place of cmdEdit
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdEdit.Visible = True Then
        
            'rs.Close
            lb.Close

            Set rs = New Recordset
            rs.Open "Select * from Data where ID = '" + txtRecordID.Text + " '", _
            con, adOpenDynamic, adLockOptimistic
            
            'Hiding the Record Retrieving Labels
            HideLabels
                
            'Showing the Record Retrieving Textboxes
            ShowObjects
            
            DisplayRS
        
            cmdSave.Visible = True
            cmdEdit.Visible = False
            
            cmdDelete.Visible = False
            cmdClear.Visible = True
            cmdClear.TabIndex = "55"
            
            cmdFirst.Visible = False
            cmdPrevious.Visible = False
            cmdNext.Visible = False
            cmdLast.Visible = False
            
            cmdAddNew.Visible = False
            cmdCancel.Visible = True
            
        Else
        
            MsgBox "You are already editing the file", vbExclamation, "Invalid operation"
        
        End If
    End If

End Sub

Sub CancelGeneral()

    'This will be used in place of cmdCancel
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdCancel.Visible = True Then
            
            Clear
            
            Picture1.Visible = False
            Picture2.Visible = False
            Picture3.Visible = False
            Picture4.Visible = False
            Picture5.Visible = False
            Picture6.Visible = False
            Picture7.Visible = False
            Picture8.Visible = False
            Picture9.Visible = False
            Picture10.Visible = False
            Picture11.Visible = False
            Picture12.Visible = False
            Picture13.Visible = False
            Picture14.Visible = False
            Picture15.Visible = False
            Picture16.Visible = False
            Picture17.Visible = False
            Picture18.Visible = False
            Picture19.Visible = False
            Picture20.Visible = False
            Picture21.Visible = False
            Picture21.Visible = False
            Picture22.Visible = False
            Picture23.Visible = False
            Picture24.Visible = False
            Picture25.Visible = False
            Picture26.Visible = False
            Picture27.Visible = False
            Picture28.Visible = False
            Picture29.Visible = False
            Picture30.Visible = False
            Picture31.Visible = False
            Picture32.Visible = False
            Picture33.Visible = False
            Picture34.Visible = False
            Picture35.Visible = False
            Picture36.Visible = False
            Picture37.Visible = False
            Picture38.Visible = False
            Picture39.Visible = False
            Picture40.Visible = False
         
            cmdCancel.Visible = False
            cmdAddNew.Visible = True
            cmdAddNew.TabIndex = "2"
            
            'Hide Objects
            HideObjects
            
            Set lb = New Recordset
            
            If txtRecordID.Text <> "" Then
            
                lb.Open "Select * from Data", con, adOpenDynamic, adLockOptimistic
                
            Else
            
                lb.Open "Select * from Data", con, adOpenDynamic, adLockOptimistic
                
            End If
            
            ShowLabels
            
            DisplayLB
            
            cmdDelete.Visible = True
            cmdClear.Visible = False
            cmdDelete.TabIndex = "55"
            
            cmdFirst.Visible = True
            cmdPrevious.Visible = True
            cmdNext.Visible = True
            cmdLast.Visible = True
            
            cmdSave.Visible = False
            cmdEdit.Visible = True
            cmdEdit.TabIndex = "58"
            
        Else
        
            MsgBox "Cancel is not available", vbExclamation, "Invalid operation"
            
        End If
    End If

End Sub

Sub FirstGeneral()

    'This will be used in place of cmdFirst
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdFirst.Visible = True Then
    
            lb.MoveFirst
            DisplayLB
                
            txtRecordID.Text = lb!ID
                
        Else
            
            MsgBox "Please switch format", vbExclamation, "Invalid operation"
       
        End If
    End If
    
End Sub

Sub LastGeneral()

    'This will be used in place of cmdLast
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdLast.Visible = True Then
        
            lb.MoveLast
            DisplayLB
            
            txtRecordID.Text = lb!ID
            
        Else
        
            MsgBox "Please switch format", vbExclamation, "Invalid operation"
            
        End If
    End If
    
End Sub

Sub NextGeneral()

    'This will be used in place of cmdNext
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdNext.Visible = True Then
    
            lb.MoveNext
            If Not lb.EOF Then
            
                DisplayLB
                txtRecordID.Text = lb!ID
                
            Else
            
                lb.MoveFirst
                DisplayLB
                
                txtRecordID.Text = lb!ID
                
            End If
            
        Else
        
            MsgBox "Please switch format", vbExclamation, "Invalid operation"
            
        End If
    End If
    
End Sub

Sub PreviousGeneral()

    'This will be used in place of cmdPrevious
    If frmGeneral.Visible = False Then
    
        MsgBox "Please open the form", vbExclamation, "Invalid operation"
        
    Else
    
        If cmdPrevious.Visible = True Then

            lb.MovePrevious
            If lb.BOF Then
            
                lb.MoveLast
                DisplayLB
                
                txtRecordID.Text = lb!ID
                
            Else
            
                DisplayLB
                
                txtRecordID.Text = lb!ID
                
            End If
            
        Else
        
            MsgBox "Please switch format", vbExclamation, "Invalid operation"
            
        End If
    End If
    
End Sub

Sub Current()

    If cmdSave.Visible = False And cmdClear.Visible = False Then
    
        MasterEnvironment.rscmdCurrentReport.Open "Select * from Data where ID = " _
        & "'" + txtRecordID.Text + " '"
        CurrentReport.Refresh
        CurrentReport.Show
        MasterEnvironment.rscmdCurrentReport.Close
        
    Else

        MsgBox "Please switch format", vbExclamation, "Report not available"
            
    End If

End Sub

Sub Complete()

    If cmdSave.Visible = False And cmdClear.Visible = False Then
    
        CompleteDataReport.Show
        
    Else
    
        MsgBox "Please switch format", vbExclamation, "Report not available"
            
    End If

End Sub

Sub ReportSearch()

    If cmdSave.Visible = False And cmdClear.Visible = False Then
    
        If cboField.Text = "All" Then
                
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
                    
        End If
            
        If cboField.Text = "Criminal" Then
              
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where Criminal = " _
            & "'" + cboKeyword.Text + "'"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
                
        End If
            
            
        If cboField.Text = "Healthy" Then
            
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where Healthy =" _
            & " '" + cboKeyword.Text + "'"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
    
        End If
             
        If cboField.Text = "Name" Then
            
            If cboKeyword.Text = "All" Then
                
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where FirstName =" _
                & "'" + txtKeyword.Text + "' or LastName = '" + txtKeyword.Text + "' or" _
                & "FathersName = '" + txtKeyword.Text + "'or MothersName = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                     
            End If
                
            If cboKeyword.Text = "First" Then
                
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where FirstName = " _
                & "'" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                   
            End If
                
            If cboKeyword.Text = "Last" Then
                
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "LastName = '" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
                
            If cboKeyword.Text = "Father" Then
                
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "FathersName = '" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
      
            End If
                
            If cboKeyword.Text = "Mother" Then
              
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "MothersName = '" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
        End If
            
        If cboField.Text = "Gender" Then
            
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "Gender = '" + cboKeyword.Text + " '"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
                    
        End If
            
        If cboField.Text = "DOB" Then
         
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "DOB = '" + txtKeyword.Text + " '"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
    
        End If
            
        If cboField.Text = "Religion" Then
    
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "Religion = '" + cboKeyword.Text + " '"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
    
        End If
            
        If cboField.Text = "Plot No" Then
            
            If cboKeyword.Text = "All" Then
               
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentPlotNo = '" + txtKeyword.Text + " '" _
                & "or ResidentialPlotNo = '" + txtKeyword.Text + " ' or EmployerPlotNo = " _
                & "'" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Permanent" Then
              
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentPlotNo = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
                    
            If cboKeyword.Text = "Residential" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "ResidentialPlotNo = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Employer" Then
     
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "EmployerPlotNo = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
        End If
            
        If cboField.Text = "Building Name" Then
            
            If cboKeyword.Text = "All" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentBuildingName = '" + txtKeyword.Text + " '" _
                & "or ResidentialBuildingName = '" + txtKeyword.Text + " ' or EmployerBuildingName = " _
                & "'" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Permanent" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentBuildingName = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                    
            If cboKeyword.Text = "Residential" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "ResidentialBuildingName = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Employer" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "EmployerBuildingName = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
     
            End If
        End If
            
        If cboField.Text = "Locality" Then
            
            If cboKeyword.Text = "All" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentLocality = '" + txtKeyword.Text + " '" _
                & "or ResidentialLocality = '" + txtKeyword.Text + " ' or EmployerLocality = " _
                & "'" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Permanent" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentLocality = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                    
            If cboKeyword.Text = "Residential" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "ResidentialLocality = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Employer" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "EmployerLocality = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
        End If
            
        If cboField.Text = "City" Then
            
            If cboKeyword.Text = "All" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentCity = '" + txtKeyword.Text + " '" _
                & "or ResidentialCity = '" + txtKeyword.Text + " ' or EmployerCity = " _
                & "'" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Permanent" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentCity = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                    
            If cboKeyword.Text = "Residential" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "ResidentialCity = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Employer" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "EmployerCity = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
        End If
            
        If cboField.Text = "State" Then
            
            If cboKeyword.Text = "All" Then
     
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentState = '" + txtKeyword.Text + " ' " _
                & "or ResidentialState = '" + txtKeyword.Text + " ' or EmployerState = " _
                & "'" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
    
            If cboKeyword.Text = "Permanent" Then
     
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentState = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
    
            If cboKeyword.Text = "Residential" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "ResidentialState = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
                
            If cboKeyword.Text = "Employer" Then
     
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "EmployerState = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
        End If
            
        If cboField.Text = "Country" Then
            
            If cboKeyword.Text = "All" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentCountry = '" + txtKeyword.Text + " ' " _
                & "or ResidentialCountry = '" + txtKeyword.Text + " ' or EmployerCountry = " _
                & "'" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Permanent" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentCountry = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
                    
            If cboKeyword.Text = "Residential" Then
                
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "ResidentialCountry = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Employer" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "EmployerCountry = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
        End If
            
        If cboField.Text = "Postal Code" Then
            
            If cboKeyword.Text = "All" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentPostalCode = '" + txtKeyword.Text + " ' " _
                & "or ResidentialPostalCode = '" + txtKeyword.Text + " ' or EmployerPostalCode = " _
                & "'" + txtKeyword.Text + " '"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
    
            End If
                
            If cboKeyword.Text = "Permanent" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "PermanentPostalCode = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                        
            End If
                    
            If cboKeyword.Text = "Residential" Then
     
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "ResidentialPostalCode = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
                
            If cboKeyword.Text = "Employer" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "EmployerPostalCode = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
     
            End If
        End If
            
            
        If cboField.Text = "Mobile Number" Then
    
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "Contact = '" + txtKeyword.Text + "'"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
    
        End If
            
            
        If cboField.Text = "STD Code" Then
    
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "DialCode = '" + txtKeyword.Text + "'"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
    
        End If
            
        If cboField.Text = "Telephone Number" Then
    
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "HomeContact = '" + txtKeyword.Text + "'"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
     
        End If
            
        If cboField.Text = "Occupation" Then
            
            If cboKeyword.Text = "All" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "Occupation = '" + txtKeyword.Text + "'" _
                & "or FathersOccupation = '" + txtKeyword.Text + "' or MothersOccupation = " _
                & "'" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
        
            End If
                
            If cboKeyword.Text = "Person" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "Occupation = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
     
            End If
                
            If cboKeyword.Text = "Father" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "FathersOccupation = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
                    
            End If
                
            If cboKeyword.Text = "Mother" Then
    
                MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
                & "MothersOccupation = '" + txtKeyword.Text + "'"
                SearchReport.Refresh
                SearchReport.Show
                MasterEnvironment.rscmdSearchReport.Close
       
            End If
        End If
                
            
        If cboField.Text = "Siblings" Then
     
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "Siblings ='" + cboKeyword.Text + "'"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
     
        End If
            
        If cboField.Text = "Brothers" Then
    
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "Brothers ='" + cboKeyword.Text + "'"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
     
        End If
            
        If cboField.Text = "Sisters" Then
    
            MasterEnvironment.rscmdSearchReport.Open "Select * from Data where " _
            & "Sisters ='" + cboKeyword.Text + "'"
            SearchReport.Refresh
            SearchReport.Show
            MasterEnvironment.rscmdSearchReport.Close
    
        End If
        
    Else
    
        MsgBox "Please switch format", vbExclamation, "Report not available"
            
    End If

End Sub

'Following are the sub procedures to be used in and by this form itself
Sub DisplayRS()
    
    'This will display record data in the text boxes
    txtFirstName.Text = rs!FirstName
    txtLastName.Text = rs!LastName
    cboGender.Text = rs!Gender
    
    If rs!Picture <> "" Then
    
        str = rs!Picture
        picResize.Picture = LoadPicture(rs!Picture)
        
        picSelf.PaintPicture picResize.Picture, 0, 0, picSelf.ScaleWidth, _
        picSelf.ScaleHeight, 0, 0, picResize.ScaleWidth, picResize.ScaleHeight, vbSrcCopy
        
    End If
    
    dtDOB.Value = rs!DOB
    cboReligion.Text = rs!Religion
    txtPermanentPlotNo.Text = rs!PermanentPlotNo
    txtPermanentBuildingName.Text = rs!PermanentBuildingName
    txtPermanentLocality.Text = rs!PermanentLocality
    cboCity1.Text = rs!PermanentCity
    cboState1.Text = rs!PermanentState
    cboCountry1.Text = rs!PermanentCountry
    txtPCode1.Text = rs!PermanentPostalCode
    txtResidentialPlotNo.Text = rs!ResidentialPlotNo
    txtResidentialBuildingName.Text = rs!ResidentialBuildingName
    txtResidentialLocality.Text = rs!ResidentialLocality
    cboCity2.Text = rs!ResidentialCity
    cboState2.Text = rs!ResidentialState
    cboCountry2.Text = rs!ResidentialCountry
    txtPCode2.Text = rs!ResidentialPostalCode
    txtContact.Text = rs!Contact
    txtTelCode.Text = rs!DialCode
    txtHome.Text = rs!HomeContact
    cboOccupation.Text = rs!Occupation
    cboCriminal.Text = rs!Criminal
    cboHealthy.Text = rs!Healthy
    txtID.Text = rs!EmailID
    cboProvider.Text = rs!EmailProvider
    txtFathers.Text = rs!FathersName
    cboFathersO.Text = rs!FathersOccupation
    txtMothers.Text = rs!MothersName
    cboMothersO.Text = rs!MothersOccupation
    cboSiblings.Text = rs!Siblings
    txtBrothers.Text = rs!Brothers
    txtSisters.Text = rs!Sisters
    cboEducation.Text = rs!Education
    cboJob.Text = rs!Job
    
    If rs!Job = "Working" Then

        txtEmployer.Text = rs!Employer
        txtEmployerPlotNo.Text = rs!EmployerPlotNo
        txtEmployerBuildingName.Text = rs!EmployerBuildingName
        txtEmployerLocality.Text = rs!EmployerLocality
        cboCity3.Text = rs!EmployerCity
        cboState3.Text = rs!EmployerState
        cboCountry3.Text = rs!EmployerCountry
        txtPCode3.Text = rs!EmployerPostalCode
        
    End If
    
    cboMarital.Text = rs!MaritalStatus
    txtHobbies.Text = rs!Hobbies
    txtFavPlaces.Text = rs!FavoritePlaces

End Sub

Sub DisplayLB()

    'Set lb = New Recordset
    'lb.Open "Select * from Data where ID = '" + txtRecordID.Text + "'", _
    con, adOpenDynamic, adLockOptimistic

    'This will display record data in the labels
    lblFirst.Caption = lb!FirstName
    lblLast.Caption = lb!LastName
    
    lblMyFirstName.Caption = lb!FirstName
    lblMyLastName.Caption = lb!LastName
    lblMyGender.Caption = lb!Gender
    
    If lb.Fields("Picture").Value <> "" Then
    
        str = lb!Picture
        picResize.Picture = LoadPicture(lb!Picture)
        
        picSelf.PaintPicture picResize.Picture, 0, 0, picSelf.ScaleWidth, _
        picSelf.ScaleHeight, 0, 0, picResize.ScaleWidth, picResize.ScaleHeight, vbSrcCopy

    End If
    
    lblMyDOB.Caption = lb!DOB
    lblMyReligion.Caption = lb!Religion
    lblMyPermanentAddress.Caption = lb!Permanent
    lblMyResidentialAddress.Caption = lb!Residential
    lblMyContact.Caption = lb!Contact
    lblMyHomeCode.Caption = lb!DialCode
    lblMyHomeContact.Caption = lb!HomeContact
    lblMyOccupation.Caption = lb!Occupation
    lblMyEmailID.Caption = lb!Email
    
    lblMyCriminal.Caption = lb!Criminal
    lblMyHealthy.Caption = lb!Healthy
        
    lblMyFathersName.Caption = lb!FathersName
    lblMyMothersName.Caption = lb!MothersName
    lblMyFathersOccupation.Caption = lb!FathersOccupation
    lblMyMothersOccupation.Caption = lb!MothersOccupation
    lblMySiblings.Caption = lb!Siblings
    lblMyBrothers.Caption = lb!Brothers
    lblMySisters.Caption = lb!Sisters
        
    lblMyEducationHighest.Caption = lb!Education
    lblMyJobStatus.Caption = lb!Job
    
    If lb!Job = "Working" Then
    
        lblMyEmployer.Caption = lb!Employer
        lblMyEmployersAddress.Caption = lb!EmployerA
        
    End If
    
    lblMyMaritalStatus.Caption = lb!MaritalStatus
    lblMyHobbies.Caption = lb!Hobbies
    lblMyFavouritePlaces.Caption = lb!FavoritePlaces

End Sub

Sub Clear()

    'This will clear all the textboxes
    txtFirstName.Text = ""
    txtLastName.Text = ""
    cboGender.Text = ""
    picSelf.Picture = LoadPicture("")
    dtDOB.Value = "10/05/2005"
    cboReligion.Text = ""
    txtPermanentPlotNo.Text = ""
    txtPermanentBuildingName.Text = ""
    txtPermanentLocality.Text = ""
    cboCity1.Text = ""
    cboState1.Text = ""
    cboCountry1.Text = ""
    txtPCode1.Text = ""
    txtResidentialPlotNo.Text = ""
    txtResidentialBuildingName.Text = ""
    txtResidentialLocality.Text = ""
    cboCity2.Text = ""
    cboState2.Text = ""
    cboCountry2.Text = ""
    txtPCode2.Text = ""
    txtContact.Text = ""
    txtTelCode.Text = ""
    txtHome.Text = ""
    cboOccupation.Text = ""
    cboCriminal.Text = ""
    cboHealthy.Text = ""
    txtID.Text = ""
    cboProvider.Text = ""
    txtFathers.Text = ""
    txtMothers.Text = ""
    cboSiblings.Text = ""
    txtBrothers.Text = ""
    txtSisters.Text = ""
    cboEducation.Text = ""
    cboJob.Text = ""
    txtEmployer.Text = ""
    txtEmployerPlotNo.Text = ""
    txtEmployerBuildingName.Text = ""
    txtEmployerLocality.Text = ""
    cboCity3.Text = ""
    cboState3.Text = ""
    cboCountry3.Text = ""
    txtPCode3.Text = ""
    cboMarital.Text = ""
    txtHobbies.Text = ""
    txtFavPlaces.Text = ""

End Sub

Sub ShowObjects()

    'Show textboxes, combo etc
    txtFirstName.Visible = True
    txtLastName.Visible = True
    cboGender.Visible = True
    dtDOB.Visible = True
    cboReligion.Visible = True
    
    cmdUpload.Visible = True

    txtPermanentPlotNo.Visible = True
    txtPermanentBuildingName.Visible = True
    txtPermanentLocality.Visible = True
    cboCity1.Visible = True
    cboState1.Visible = True
    cboCountry1.Visible = True
    txtPCode1.Visible = True
    txtResidentialPlotNo.Visible = True
    txtResidentialBuildingName.Visible = True
    txtResidentialLocality.Visible = True
    cboCity2.Visible = True
    cboState2.Visible = True
    cboCountry2.Visible = True
    txtPCode2.Visible = True
    txtContact.Visible = True
    txtTelCode.Visible = True
    txtHome.Visible = True
    cboOccupation.Visible = True
    txtID.Visible = True
    cboProvider.Visible = True
        
    cboCriminal.Visible = True
    cboHealthy.Visible = True
        
    txtFathers.Visible = True
    txtMothers.Visible = True
    cboFathersO.Visible = True
    cboMothersO.Visible = True
    cboSiblings.Visible = True
    txtBrothers.Visible = True
    txtSisters.Visible = True
        
    cboEducation.Visible = True
    cboJob.Visible = True
    txtEmployer.Visible = True
    txtEmployerPlotNo.Visible = True
    txtEmployerBuildingName.Visible = True
    txtEmployerLocality.Visible = True
    cboCity3.Visible = True
    cboState3.Visible = True
    cboCountry3.Visible = True
    txtPCode3.Visible = True
    cboMarital.Visible = True
    txtHobbies.Visible = True
    txtFavPlaces.Visible = True
    
    'Show Hidden Object's Labels
    lblCity1.Visible = True
    lblState1.Visible = True
    lblCountry1.Visible = True
    lblPCode1.Visible = True
    
    lblCity2.Visible = True
    lblState2.Visible = True
    lblCountry2.Visible = True
    lblPCode2.Visible = True
    
    lblCity3.Visible = True
    lblState3.Visible = True
    lblCountry3.Visible = True
    lblPCode3.Visible = True
    
    lblAtRate.Visible = True
    lblcom.Visible = True
    
    lblInfo.Visible = True

End Sub

Sub HideObjects()
    
    'Hide textboxes, combo etc
    txtFirstName.Visible = False
    txtLastName.Visible = False
    cboGender.Visible = False
    dtDOB.Visible = False
    cboReligion.Visible = False
    
    cmdUpload.Visible = False
    
    txtPermanentPlotNo.Visible = False
    txtPermanentBuildingName.Visible = False
    txtPermanentLocality.Visible = False
    cboCity1.Visible = False
    cboState1.Visible = False
    cboCountry1.Visible = False
    txtPCode1.Visible = False
    txtResidentialPlotNo.Visible = False
    txtResidentialBuildingName.Visible = False
    txtResidentialLocality.Visible = False
    cboCity2.Visible = False
    cboState2.Visible = False
    cboCountry2.Visible = False
    txtPCode2.Visible = False
    txtContact.Visible = False
    txtTelCode.Visible = False
    txtHome.Visible = False
    cboOccupation.Visible = False
    txtID.Visible = False
    cboProvider.Visible = False
        
    cboCriminal.Visible = False
    cboHealthy.Visible = False
        
    txtFathers.Visible = False
    txtMothers.Visible = False
    cboFathersO.Visible = False
    cboMothersO.Visible = False
    cboSiblings.Visible = False
    txtBrothers.Visible = False
    txtSisters.Visible = False
        
    cboEducation.Visible = False
    cboJob.Visible = False
    txtEmployer.Visible = False
    txtEmployerPlotNo.Visible = False
    txtEmployerBuildingName.Visible = False
    txtEmployerLocality.Visible = False
    cboCity3.Visible = False
    cboState3.Visible = False
    cboCountry3.Visible = False
    txtPCode3.Visible = False
    cboMarital.Visible = False
    txtHobbies.Visible = False
    txtFavPlaces.Visible = False
    
    'Hide Few Object Labels
    lblCity1.Visible = False
    lblState1.Visible = False
    lblCountry1.Visible = False
    lblPCode1.Visible = False
    
    lblCity2.Visible = False
    lblState2.Visible = False
    lblCountry2.Visible = False
    lblPCode2.Visible = False
    
    lblCity3.Visible = False
    lblState3.Visible = False
    lblCountry3.Visible = False
    lblPCode3.Visible = False
    
    lblAtRate.Visible = False
    lblcom.Visible = False
    
    lblInfo.Visible = False
     
End Sub

Sub ShowLabels()
    
    'Shows labels
    lblFirst.Visible = True
    lblLast.Visible = True
    lblMyFirstName.Visible = True
    lblMyLastName.Visible = True
    lblMyGender.Visible = True
    lblMyDOB.Visible = True
    lblMyReligion.Visible = True
    lblMyPermanentAddress.Visible = True
    lblMyResidentialAddress.Visible = True
    lblMyContact.Visible = True
    lblMyHomeCode.Visible = True
    lblMyHomeContact.Visible = True
    lblMyOccupation.Visible = True
    lblMyEmailID.Visible = True
        
    lblMyCriminal.Visible = True
    lblMyHealthy.Visible = True
        
    lblMyFathersName.Visible = True
    lblMyMothersName.Visible = True
    lblMyFathersOccupation.Visible = True
    lblMyMothersOccupation.Visible = True
    lblMySiblings.Visible = True
    lblMyBrothers.Visible = True
    lblMySisters.Visible = True
        
    lblMyEducationHighest.Visible = True
    lblMyJobStatus.Visible = True
    lblMyEmployer.Visible = True
    lblMyEmployersAddress.Visible = True
    lblMyMaritalStatus.Visible = True
    lblMyHobbies.Visible = True
    lblMyFavouritePlaces.Visible = True

End Sub

Sub HideLabels()

    'Hide Labels
    lblFirst.Visible = False
    lblLast.Visible = False
    lblMyFirstName.Visible = False
    lblMyLastName.Visible = False
    lblMyGender.Visible = False
    lblMyDOB.Visible = False
    lblMyReligion.Visible = False
    lblMyPermanentAddress.Visible = False
    lblMyResidentialAddress.Visible = False
    lblMyContact.Visible = False
    lblMyHomeCode.Visible = False
    lblMyHomeContact.Visible = False
    lblMyOccupation.Visible = False
    lblMyEmailID.Visible = False
        
    lblMyCriminal.Visible = False
    lblMyHealthy.Visible = False
        
    lblMyFathersName.Visible = False
    lblMyMothersName.Visible = False
    lblMyFathersOccupation.Visible = False
    lblMyMothersOccupation.Visible = False
    lblMySiblings.Visible = False
    lblMyBrothers.Visible = False
    lblMySisters.Visible = False
        
    lblMyEducationHighest.Visible = False
    lblMyJobStatus.Visible = False
    lblMyEmployer.Visible = False
    lblMyEmployersAddress.Visible = False
    lblMyMaritalStatus.Visible = False
    lblMyHobbies.Visible = False
    lblMyFavouritePlaces.Visible = False

End Sub

Public Sub AfterLogin()

    con.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source=" _
    & "C:\Users\Hp\Desktop\VB6new\MasterDatabase.mdb; Persist Security Info= False"
        
    If txtAfterLogin.Text = "New" Then
    
        rs.Open "Select * from Data", con, adOpenDynamic, adLockOptimistic
    
        'Hiding the Record Retrieving Labels
        HideLabels
        
        'Showing the Record Retrieving Textboxes
        ShowObjects
        
        rs.AddNew
        Clear
        
        'Entering Text in Few address boxes
        txtPermanentPlotNo.Text = "Plot No"
        txtPermanentBuildingName.Text = "Building Name"
        txtPermanentLocality.Text = "Locality"
        txtResidentialPlotNo.Text = "Plot No"
        txtResidentialBuildingName.Text = "Building Name"
        txtResidentialLocality.Text = "Locality"
        txtEmployerPlotNo.Text = "Plot No"
        txtEmployerBuildingName.Text = "Building Name"
        txtEmployerLocality.Text = "Locality"
            
            
        'Altering some properties
        lblInfo.Visible = True
        
        cmdAddNew.Visible = False
        cmdCancel.Visible = False
        cmdOpenExisting.Visible = True
        cmdOpenExisting.TabIndex = "2"
        
        cmdEdit.Visible = False
        cmdSave.Visible = True
        cmdSave.TabIndex = "58"
        
        cmdClear.Visible = True
        cmdDelete.Visible = False
        cmdClear.TabIndex = "60"
        
        cmdFirst.Visible = False
        cmdPrevious.Visible = False
        cmdNext.Visible = False
        cmdLast.Visible = False
        
    End If
    
    If txtAfterLogin.Text = "Existing" Then
    
        lb.Open "Select * from Data", con, adOpenDynamic, adLockOptimistic
    
        'Hiding the Record Retrieving TextBoxes and other objects
        HideObjects
        
        'Showing the Record Retrieving Labels
        ShowLabels
        
        cmdSave.Visible = False
        cmdEdit.Visible = True
        cmdEdit.TabIndex = "58"
        
        cmdAddNew.Visible = True
        cmdCancel.Visible = False
        cmdAddNew.TabIndex = "2"
        
        lb.MoveFirst
        DisplayLB
        
        cmdDelete.Visible = True
        cmdClear.Visible = False
        cmdDelete.TabIndex = "60"
        
        lblInfo.Visible = False
        
        txtRecordID.Text = lb!ID
        
        cboField.Text = "All"
        
    End If

End Sub

Sub CenterChild(Parent As Form, Child As Form)

    Dim iTop As Integer
    Dim iLeft As Integer
    
    If Parent.WindowState <> 0 Then Exit Sub
    
        iTop = ((Parent.Height - Child.Height) \ 2)
        iLeft = ((Parent.Width - Child.Width) \ 2)
        Child.Move iLeft, iTop

End Sub

'Following Lost Focus events are meant to correct input during run time
Private Sub cboCity1_LostFocus()

    Picture7.Visible = False
    If Trim(cboCity1.Text) <> "" Then
        
        If cboCity1.Text = "Mumbai" Or cboCity1.Text = "Navi Mumbai" Or _
        cboCity1.Text = "Pune" Or cboCity1.Text = "Nagpur" Or cboCity1.Text = "Thane" Or _
        cboCity1.Text = "Pimpri" Or cboCity1.Text = "Chinchwad" Or cboCity1.Text = "Nashik" Or _
        cboCity1.Text = "Kalyan" Or cboCity1.Text = "Dombivali" Or cboCity1.Text = "Pune" Or _
        cboCity1.Text = "Vasai" Or cboCity1.Text = "Virar" Or cboCity1.Text = "Aurangabad" Or _
        cboCity1.Text = "Solapur" Or cboCity1.Text = "Mira" Or cboCity1.Text = "Bhayandar" Or _
        cboCity1.Text = "Bhiwandi" Or cboCity1.Text = "Nizampur" Or cboCity1.Text = "Amravati" Or _
        cboCity1.Text = "Nanded" Or cboCity1.Text = "Waghala" Or cboCity1.Text = "Panvel" Or _
        cboCity1.Text = "Sangli" Or cboCity1.Text = "Akola" Or cboCity1.Text = "Ahmednagar" Or _
        cboCity1.Text = "Parbhani" Or cboCity1.Text = "Chandrapur" Or cboCity1.Text = "Dhule" Or _
        cboCity1.Text = "Malegaon" Or cboCity1.Text = "Jalgaon" Or cboCity1.Text = "Kolhapur" Or _
        cboCity1.Text = "Nashik" Or cboCity1.Text = "Latur" Then
                    
        Else
            
            Picture7.Visible = True
                
        End If
    End If
        
End Sub

Private Sub cboCity2_LostFocus()

    Picture12.Visible = False
    If Trim(cboCity2.Text) <> "" Then
        
        If cboCity2.Text = "Mumbai" Or cboCity2.Text = "Navi Mumbai" Or cboCity2.Text = "Pune" Or _
        cboCity2.Text = "Nagpur" Or cboCity2.Text = "Thane" Or cboCity2.Text = "Pimpri" Or _
        cboCity2.Text = "Chinchwad" Or cboCity2.Text = "Nashik" Or cboCity2.Text = "Kalyan" Or _
        cboCity2.Text = "Dombivali" Or cboCity2.Text = "Pune" Or cboCity2.Text = "Vasai" Or _
        cboCity2.Text = "Virar" Or cboCity2.Text = "Aurangabad" Or cboCity2.Text = "Solapur" Or _
        cboCity2.Text = "Mira" Or cboCity2.Text = "Bhayandar" Or cboCity2.Text = "Bhiwandi" Or _
        cboCity2.Text = "Nizampur" Or cboCity2.Text = "Amravati" Or cboCity2.Text = "Nanded" Or _
        cboCity2.Text = "Waghala" Or cboCity2.Text = "Panvel" Or cboCity2.Text = "Sangli" Or _
        cboCity2.Text = "Akola" Or cboCity2.Text = "Ahmednagar" Or cboCity2.Text = "Parbhani" Or _
        cboCity2.Text = "Chandrapur" Or cboCity2.Text = "Dhule" Or cboCity2.Text = "Malegaon" Or _
        cboCity2.Text = "Jalgaon" Or cboCity2.Text = "Kolhapur" Or cboCity2.Text = "Nashik" Or _
        cboCity2.Text = "Latur" Then
        
        Else
            
            Picture12.Visible = True
                
        End If
    End If
    
End Sub

Private Sub cboCity3_LostFocus()

    Picture33.Visible = False
    If Trim(cboCity3.Text) <> "" Then
        
        If cboCity3.Text = "Mumbai" Or cboCity3.Text = "Navi Mumbai" Or cboCity3.Text = "Pune" Or _
        cboCity3.Text = "Nagpur" Or cboCity3.Text = "Thane" Or cboCity3.Text = "Pimpri" Or _
        cboCity3.Text = "Chinchwad" Or cboCity3.Text = "Nashik" Or cboCity3.Text = "Kalyan" Or _
        cboCity3.Text = "Dombivali" Or cboCity3.Text = "Pune" Or cboCity3.Text = "Vasai" Or _
        cboCity3.Text = "Virar" Or cboCity3.Text = "Aurangabad" Or cboCity3.Text = "Solapur" Or _
        cboCity3.Text = "Mira" Or cboCity3.Text = "Bhayandar" Or cboCity3.Text = "Bhiwandi" Or _
        cboCity3.Text = "Nizampur" Or cboCity3.Text = "Amravati" Or cboCity3.Text = "Nanded" Or _
        cboCity3.Text = "Waghala" Or cboCity3.Text = "Panvel" Or cboCity3.Text = "Sangli" Or _
        cboCity3.Text = "Akola" Or cboCity3.Text = "Ahmednagar" Or cboCity3.Text = "Parbhani" Or _
        cboCity3.Text = "Chandrapur" Or cboCity3.Text = "Dhule" Or cboCity3.Text = "Malegaon" Or _
        cboCity3.Text = "Jalgaon" Or cboCity3.Text = "Kolhapur" Or cboCity3.Text = "Nashik" Or _
        cboCity3.Text = "Latur" Then
                
        Else
            
            Picture33.Visible = True
                
        End If
    End If

End Sub

Private Sub cboCountry1_LostFocus()

    Picture9.Visible = False
    If Trim(cboCountry1.Text) <> "" Then
        
        If Trim(cboCountry1.Text) <> "India" Then
            
            'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
            Picture9.Visible = True
                
        End If
    End If
    
End Sub

Private Sub cboCountry2_LostFocus()

    Picture14.Visible = False
    If Trim(cboCountry2.Text) <> "" Then
        
        If Trim(cboCountry2.Text) <> "India" Then
            
            'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
            Picture14.Visible = True
            
        End If
    End If
    
End Sub

Private Sub cboCountry3_LostFocus()

    Picture35.Visible = False
    If Trim(cboCountry3.Text) <> "" Then
        
        If Trim(cboCountry3.Text) <> "India" Then
        
            'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
            Picture35.Visible = True
                
        End If
    End If
    
End Sub

Private Sub cboCriminal_LostFocus()

    Picture20.Visible = False
    If Trim(cboCriminal.Text) <> "" Then
    
        If Trim(cboCriminal.Text) <> "Yes" And Trim(cboCriminal.Text) <> "No" Then
        
            Picture20.Visible = True
            
        End If
    End If

End Sub

Private Sub cboEducation_LostFocus()

    Picture29.Visible = False
    If Trim(cboEducation.Text) <> "" Then
    
        If cboEducation.Text = "1st Standard" Or cboEducation.Text = "2nd Standard" Or _
        cboEducation.Text = "3rd Standard" Or cboEducation.Text = "4th Standard" Or _
        cboEducation.Text = "5th Standard" Or cboEducation.Text = "6th Standard" Or _
        cboEducation.Text = "7th Standard" Or cboEducation.Text = "8th Standard" Or _
        cboEducation.Text = "9th Standard" Or cboEducation.Text = "10th Standard" Or _
        cboEducation.Text = "11th Standard" Or cboEducation.Text = "12th Standard" Or _
        cboEducation.Text = "B.E" Or cboEducation.Text = "B.Tech" Or cboEducation.Text = "B.Arch" Or _
        cboEducation.Text = "B.Sc" Or cboEducation.Text = "B.Com" Or cboEducation.Text = "B.B.A" Or _
        cboEducation.Text = "B.C.C.A" Or cboEducation.Text = "M.Tech" Or cboEducation.Text = "M.Sc" Or _
        cboEducation.Text = "M.C.M" Or cboEducation.Text = "M.B.A" Then
        
        Else
            
            Picture29.Visible = True
                
        End If
    End If

End Sub

Private Sub cboFathersO_LostFocus()

    Picture24.Visible = False
    If Trim(cboFathersO.Text) <> "" Then
    
        If cboFathersO.Text = "Agriculture" Or cboFathersO.Text = "Home Maker" Or _
        cboFathersO.Text = "Business" Or cboFathersO.Text = "Medical" Or _
        cboFathersO.Text = "Engineering" Or cboFathersO.Text = "Law Practice" Or _
        cboFathersO.Text = "Government Service" Or cboFathersO.Text = "Public Sector Service" Or _
        cboFathersO.Text = "Private Service" Or cboFathersO.Text = "Teaching" Or _
        cboFathersO.Text = "Other" Then
                
        Else
            
            Picture24.Visible = True
                
        End If
    End If
    
End Sub

Private Sub cboGender_LostFocus()

    Picture3.Visible = False

    If Trim(cboGender.Text) <> "" Then
    
        If Trim(cboGender.Text) <> "Male" And Trim(cboGender.Text) <> "Female" Then
               
            Picture3.Visible = True
                    
        End If
    End If
        
End Sub

Private Sub cboHealthy_LostFocus()

    Picture21.Visible = False
    If Trim(cboHealthy.Text) <> "" Then
    
        If Trim(cboHealthy.Text) <> "Yes" And Trim(cboHealthy.Text) <> "No" Then
        
            Picture21.Visible = True
            
        End If
    End If
    
End Sub

Private Sub cboJob_LostFocus()

    'This will disable employer fields if job status is idle
    If cboJob.Text = "Working" Then
    
        txtEmployer.Enabled = True
        txtEmployerPlotNo.Enabled = True
        txtEmployerBuildingName.Enabled = True
        txtEmployerLocality.Enabled = True
        cboCity3.Enabled = True
        cboState3.Enabled = True
        cboCountry3.Enabled = True
        txtPCode3.Enabled = True
        
    Else
        
        txtEmployer.Enabled = False
        txtEmployerPlotNo.Enabled = False
        txtEmployerBuildingName.Enabled = False
        txtEmployerLocality.Enabled = False
        cboCity3.Enabled = False
        cboState3.Enabled = False
        cboCountry3.Enabled = False
        txtPCode3.Enabled = False
        
    End If
    
    Picture30.Visible = False
    If Trim(cboJob.Text) <> "" Then
    
        If Trim(cboJob.Text) <> "Working" And Trim(cboJob.Text) <> "Idle" Then
        
            Picture30.Visible = True
            
        End If
    End If

End Sub

Private Sub cboMarital_LostFocus()

    Picture37.Visible = False
    If Trim(cboMarital.Text) <> "" Then
        
        If cboMarital.Text <> "Married" And cboMarital.Text <> "Single" And _
        cboMarital.Text <> "Divorced" And cboMarital.Text <> "Widowed" Then
            
            Picture37.Visible = True
                
        End If
    End If

End Sub

Private Sub cboMothersO_LostFocus()

    Picture25.Visible = False
    If Trim(cboMothersO.Text) <> "" Then
    
        If cboMothersO.Text = "Agriculture" Or cboMothersO.Text = "Home Maker" Or _
        cboMothersO.Text = "Business" Or cboMothersO.Text = "Medical" Or _
        cboMothersO.Text = "Engineering" Or cboMothersO.Text = "Law Practice" Or _
        cboMothersO.Text = "Government Service" Or cboMothersO.Text = "Public Sector Service" Or _
        cboMothersO.Text = "Private Service" Or cboMothersO.Text = "Teaching" Or _
        cboMothersO.Text = "Other" Then
                
        Else
            
            Picture25.Visible = True
                
        End If
    End If

End Sub

Private Sub cboOccupation_LostFocus()

    Picture18.Visible = False
    If Trim(cboOccupation.Text) <> "" Then
    
        If cboOccupation.Text = "Agriculture" Or cboOccupation.Text = "Home Maker" Or _
        cboOccupation.Text = "Business" Or cboOccupation.Text = "Medical" Or _
        cboOccupation.Text = "Engineering" Or cboOccupation.Text = "Law Practice" Or _
        cboOccupation.Text = "Government Service" Or cboOccupation.Text = "Public Sector Service" Or _
        cboOccupation.Text = "Private Service" Or cboOccupation.Text = "Teaching" Or _
        cboOccupation.Text = "Other" Then
        
        Else
            
            Picture18.Visible = True
                
        End If
    End If

End Sub

Private Sub cboProvider_LostFocus()

    If Trim(txtID.Text) <> "" Then
    
        Picture19.Visible = False
        If Trim(cboProvider.Text) = "Gmail" Or Trim(cboProvider.Text) = "iCloud" Or _
        Trim(cboProvider.Text) = "GMX" Or Trim(cboProvider.Text) = "Outlook" Or _
        Trim(cboProvider.Text) = "Yahoo" Or Trim(cboProvider.Text) = "Aol" Or _
        Trim(cboProvider.Text) = "Zoho" Or Trim(cboProvider.Text) = "Mail" Or _
        Trim(cboProvider.Text) = "Yandex" Or Trim(cboProvider.Text) = "ProtonMail" Then
            
        Else
        
            Picture19.Visible = True
            
        End If
    End If

End Sub

Private Sub cboReligion_LostFocus()

    Picture5.Visible = False
    If Trim(cboReligion.Text) <> "" Then

        If Trim(cboReligion.Text) = "Hinduism" Or Trim(cboReligion.Text) = "Islam" Or _
        Trim(cboReligion.Text) = "Christianity" Or Trim(cboReligion.Text) = "Sikhism" Or _
        Trim(cboReligion.Text) = "Buddhism" Or Trim(cboReligion.Text) = "Jainism" Then
            
        Else
            
            Picture5.Visible = True
                
        End If
    End If
    
End Sub

Private Sub cboSiblings_LostFocus()

    If Trim(cboSiblings.Text) <> "" Then
    
        If cboSiblings.Text = "No" Then
        
            txtBrothers.Text = "0"
            txtSisters.Text = "0"
            
            txtBrothers.Enabled = False
            txtSisters.Enabled = False
        
        End If
        
        Picture26.Visible = False
        If cboSiblings.Text <> "Yes" And cboSiblings.Text <> "No" Then
        
            Picture26.Visible = True
            
        End If
    End If

End Sub

Private Sub cboState1_LostFocus()

    Picture8.Visible = False
    If Trim(cboState1.Text) <> "" Then
        
        If Trim(cboState1.Text) <> "Maharashtra" Then
        
            'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
            Picture8.Visible = True
                
        End If
            
    End If
    
End Sub

Private Sub cboState2_LostFocus()

    Picture13.Visible = False
    If Trim(cboState2.Text) <> "" Then
        
        If Trim(cboState2.Text) <> "Maharashtra" Then
       
            'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
            Picture13.Visible = True
                
        End If
    End If
    
End Sub

Private Sub cboState3_LostFocus()

    Picture34.Visible = False
    If Trim(cboState3.Text) <> "" Then
        
        If Trim(cboState3.Text) <> "Maharashtra" Then
            
            'MsgBox "This database is for India(Maharashtra) only", vbCritical, "Note"
            Picture34.Visible = True
                
        End If
    End If

End Sub

Private Sub txtResidentialPlotNo_LostFocus()

    Picture11.Visible = False
    If Trim(txtResidentialPlotNo.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtPermanentPlotNo.Text)) Then
            
            Picture11.Visible = True
            'MsgBox "Plot number should contain only number", vbExclamation, "invalid Plot Number"
                
        End If
    End If

End Sub

Private Sub txtSisters_LostFocus()

    Picture28.Visible = False
    If Trim(txtSisters.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtSisters.Text)) Then
        
            Picture28.Visible = True
            
        End If
    End If

End Sub

Private Sub txtTelCode_LostFocus()

    Picture17.Visible = False
    If Trim(txtTelCode.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtTelCode.Text)) Then
        
            Picture17.Visible = True
            'MsgBox "Please enter valid contact", vbCritical, "Invalid contact"
                
        End If
    End If
    
End Sub

Private Sub dtDOB_LostFocus()

    Picture4.Visible = False
    If dtDOB.Value = "12/20/2016" Then
        
        If MsgBox("Is your DOB, 12/20/2016?", vbInformation + vbYesNo, "Note") = vbNo Then
        
            MsgBox "Please cboose your DOB", vbExclamation, "Invalid field"
            Picture4.Visible = True
            
        Else
        
             Text4.Text = "ok"

        End If
        
    Else
    
        Text4.Text = "ok"
        
    End If
    
End Sub

Private Sub txtBrothers_LostFocus()

    Picture27.Visible = False
    If Trim(txtBrothers.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtBrothers.Text)) Then
        
            Picture27.Visible = True
                 
        End If
    End If

End Sub

Private Sub txtContact_LostFocus()

    Picture16.Visible = False
    If Trim(txtContact.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtContact.Text)) Or Len(Trim(txtContact.Text)) <> 10 Then
    
        Picture16.Visible = True
        'MsgBox "Please enter valid contact", vbCritical, "Invalid contact"
      
        End If
    End If

End Sub

Private Sub txtEmployer_LostFocus()

    Picture31.Visible = False
    If Trim(txtEmployer.Text) <> "" Then
        
        If Not IsAlphabetical(Trim(txtEmployer.Text)) Then
        
            'MsgBox "Please re-enter your Employer's name", vbCritical, "Invalid field"
            Picture31.Visible = True
                    
        End If
    End If

End Sub

Private Sub txtEmployerBuildingName_LostFocus()

     If Trim(txtEmployerBuildingName.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtEmployerBuildingName.Text)) Then
            
            Picture32.Visible = True
            'MsgBox "Building name should contain only alphabets", vbExclamation, "Invalid Building Name"
        
        End If
    End If
    
End Sub

Private Sub txtEmployerLocality_LostFocus()

    If Trim(txtEmployerLocality.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtEmployerLocality.Text)) Then
            
            Picture32.Visible = True
                
        End If
    End If
    
End Sub


Private Sub txtEmployerPlotNo_LostFocus()

    Picture32.Visible = False
    If Trim(txtEmployerPlotNo.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtEmployerPlotNo.Text)) Then
            
            Picture32.Visible = True
            'MsgBox "Plot number should contain only number", vbExclamation, "invalid Plot Number"
                
        End If
    End If

End Sub

Private Sub txtFathers_LostFocus()

    Picture22.Visible = False
    If Trim(txtFathers.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtFathers.Text)) Then
        
            'MsgBox "Please re-enter your first name", vbCritical, "Invalid field"
            Picture22.Visible = True
    
        End If
    End If

End Sub

Private Sub txtFavPlaces_LostFocus()

    Picture39.Visible = False
    If Trim(txtFavPlaces.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtFavPlaces.Text)) Then
                
            'MsgBox "Please re-enter your Employer's name", vbCritical, "Invalid field"
            Picture39.Visible = True
                
        End If
    End If

End Sub

Private Sub txtFirstName_LostFocus()

    Picture1.Visible = False

    If Trim(txtFirstName.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtFirstName.Text)) Then
            'MsgBox "Please re-enter your first name", vbCritical, "Invalid field"
            Picture1.Visible = True
    
        End If
    End If

End Sub

Private Sub txtHobbies_LostFocus()

    Picture38.Visible = False
    If Trim(txtHobbies.Text) <> "" Then

        If Not IsAlphabetical(Trim(txtHobbies.Text)) Then
            
            'MsgBox "Please re-enter your Hobbies", vbCritical, "Invalid field"
            Picture38.Visible = True
                
        End If
    End If
    
End Sub

Private Sub txtHome_LostFocus()

    If IsNumeric(Trim(txtTelCode.Text)) Then
    
        Picture17.Visible = False
        If Trim(txtHome.Text) <> "" Then
        
            If Not IsNumeric(Trim(txtHome.Text)) Then
                
                Picture17.Visible = True
    
            End If
        End If
    End If
        
End Sub

Private Sub txtLastName_LostFocus()

    Picture2.Visible = False

    If Trim(txtLastName.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtLastName.Text)) Then
                Picture2.Visible = True
                
        End If
    End If
    
End Sub

Private Sub txtMothers_LostFocus()

     Picture23.Visible = False
     If Trim(txtMothers.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtMothers.Text)) Then
        
            'MsgBox "Please re-enter your first name", vbCritical, "Invalid field"
            Picture23.Visible = True

        End If
    End If

End Sub

Private Sub txtPCode1_LostFocus()

    Picture10.Visible = False
    If Trim(txtPCode1.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtPCode1.Text)) Then
        
            Picture10.Visible = True
            
        End If
    End If

End Sub

Private Sub txtPCode2_LostFocus()

    Picture15.Visible = False
    If Trim(txtPCode2.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtPCode2.Text)) Then
        
            Picture15.Visible = True
            
        End If
    End If
    
End Sub

Private Sub txtPCode3_LostFocus()

    Picture36.Visible = False
    If Trim(txtPCode3.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtPCode3.Text)) Then
        
            Picture36.Visible = True
            
        End If
    End If
    
End Sub


Private Sub txtPermanentBuildingName_LostFocus()

    If Trim(txtPermanentBuildingName.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtPermanentBuildingName.Text)) Then
            
            Picture6.Visible = True
            'MsgBox "Building name should contain only alphabets", vbExclamation, "Invalid Building Name"
        
        End If
    End If

End Sub



Private Sub txtPermanentLocality_LostFocus()

     If Trim(txtPermanentLocality.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtPermanentLocality.Text)) Then
            
            Picture6.Visible = True
                
        End If
    End If
    
End Sub

Private Sub txtPermanentPlotNo_LostFocus()

    Picture6.Visible = False
    If Trim(txtPermanentPlotNo.Text) <> "" Then
    
        If Not IsNumeric(Trim(txtPermanentPlotNo.Text)) Then
            
            Picture6.Visible = True
            'MsgBox "Plot number should contain only number", vbExclamation, "invalid Plot Number"
                
        End If
    End If
        
End Sub

Private Sub txtResidentialBuildingName_LostFocus()

    If Trim(txtResidentialBuildingName.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtPermanentBuildingName.Text)) Then
            
            Picture11.Visible = True
            'MsgBox "Building name should contain only alphabets", vbExclamation, "Invalid Building Name"
        
        End If
    End If
    
End Sub

Private Sub txtResidentialLocality_LostFocus()

    If Trim(txtPermanentLocality.Text) <> "" Then
    
        If Not IsAlphabetical(Trim(txtPermanentLocality.Text)) Then
            
            Picture11.Visible = True
                
        End If
    End If
    
End Sub

'Following Got Focus events are meant to support respective Lost Focus events
Private Sub txtEmployerBuildingName_GotFocus()

    If Trim(txtEmployerBuildingName.Text) = "Building Name" Then
    
        txtEmployerBuildingName.Text = ""
            
    End If

End Sub

Private Sub txtEmployerLocality_GotFocus()

    If Trim(txtEmployerLocality.Text) = "Locality" Then
    
        txtEmployerLocality.Text = ""
            
    End If

End Sub

Private Sub txtEmployerPlotNo_GotFocus()

    If Trim(txtEmployerPlotNo.Text) = "Plot No" Then
    
        txtEmployerPlotNo.Text = ""
            
    End If

End Sub

Private Sub txtPermanentBuildingName_GotFocus()

    If Trim(txtPermanentBuildingName.Text) = "Building Name" Then
    
        txtPermanentBuildingName.Text = ""
            
    End If

End Sub

Private Sub txtPermanentLocality_GotFocus()

    If Trim(txtPermanentLocality.Text) = "Locality" Then
    
        txtPermanentLocality.Text = ""
            
    End If

End Sub

Private Sub txtPermanentPlotNo_GotFocus()

    If Trim(txtPermanentPlotNo.Text) = "Plot No" Then
    
        txtPermanentPlotNo.Text = ""
            
    End If

End Sub

Private Sub txtResidentialLocality_GotFocus()

    If Trim(txtResidentialLocality.Text) = "Locality" Then
    
        txtResidentialLocality.Text = ""
            
    End If

End Sub
Private Sub txtResidentialPlotNo_GotFocus()

    If Trim(txtResidentialPlotNo.Text) = "Plot No" Then
    
        txtResidentialPlotNo.Text = ""
            
    End If

End Sub

Private Sub txtResidentialBuildingName_GotFocus()

    If Trim(txtResidentialBuildingName.Text) = "Building Name" Then
    
        txtResidentialBuildingName.Text = ""
            
    End If

End Sub

'Following Click events are meant to support respective Lost Focus events
Private Sub txtEmployerBuildingName_Click()

    If Trim(txtEmployerBuildingName.Text) = "Building Name" Then
    
        txtEmployerBuildingName.Text = ""
            
    End If
    
End Sub

Private Sub txtEmployerLocality_Click()

    If Trim(txtEmployerLocality.Text) = "Locality" Then
    
        txtEmployerLocality.Text = ""
            
    End If

End Sub

Private Sub txtEmployerPlotNo_Click()

    If Trim(txtEmployerPlotNo.Text) = "Plot No" Then
    
        txtEmployerPlotNo.Text = ""
            
    End If

End Sub


Private Sub txtPermanentBuildingName_Click()

     If Trim(txtPermanentBuildingName.Text) = "Building Name" Then
    
        txtPermanentBuildingName.Text = ""
            
    End If

End Sub

Private Sub txtPermanentLocality_Click()

    If Trim(txtPermanentLocality.Text) = "Locality" Then
    
        txtPermanentLocality.Text = ""
            
    End If

End Sub

Private Sub txtPermanentPlotNo_Click()

    If Trim(txtPermanentPlotNo.Text) = "Plot No" Then
    
        txtPermanentPlotNo.Text = ""
            
    End If

End Sub



Private Sub txtResidentialBuildingName_Click()

    If Trim(txtResidentialBuildingName.Text) = "Building Name" Then
    
        txtResidentialBuildingName.Text = ""
            
    End If

End Sub

Private Sub txtResidentialLocality_Click()

    If Trim(txtResidentialLocality.Text) = "Locality" Then
    
        txtResidentialLocality.Text = ""
            
    End If

End Sub

Private Sub txtResidentialPlotNo_Click()

    If Trim(txtResidentialPlotNo.Text) = "Plot No" Then
    
        txtResidentialPlotNo.Text = ""
            
    End If
    
End Sub

'Following two are functions used in this form to validate
Private Function IsAlphabetical(TestString As String) As Boolean

    Dim sTemp As String
    Dim iLen As Integer
    Dim iCtr As Integer
    Dim sChar As String

    'This returns true if all characters are alphabets else false
    
    sTemp = TestString
    iLen = Len(sTemp)
    
    If iLen > 0 Then
    
        For iCtr = 1 To iLen
        sChar = Mid(sTemp, iCtr, 1)
        
        If Not sChar Like "[A-Za-z' '',']" Then Exit Function
        
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

'From now onwards all the command button's code is written
Private Sub cmdFirst_Click()

    FirstGeneral
    
End Sub

Private Sub cmdLast_Click()

    LastGeneral
    
End Sub

Private Sub cmdNext_Click()

    NextGeneral
    
End Sub

Private Sub cmdPrevious_Click()

    PreviousGeneral
    
End Sub

Private Sub cmdSave_Click()
        
    SaveGeneral
    
End Sub

Private Sub cmdAddNew_Click()

    AddNewGeneral
    
End Sub

Private Sub cmdCancel_Click()

    CancelGeneral
    
End Sub

Private Sub cmdClear_Click()

    Clear
    
End Sub

Private Sub cmdDelete_Click()

    DeleteGeneral
    
End Sub

Private Sub cmdEdit_Click()
    
    EditGeneral
    
End Sub

Private Sub cmdOpenExisting_Click()

    'Invisible Correction Picture boxes
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
    Picture5.Visible = False
    Picture6.Visible = False
    Picture7.Visible = False
    Picture8.Visible = False
    Picture9.Visible = False
    Picture10.Visible = False
    Picture11.Visible = False
    Picture12.Visible = False
    Picture13.Visible = False
    Picture14.Visible = False
    Picture15.Visible = False
    Picture16.Visible = False
    Picture17.Visible = False
    Picture18.Visible = False
    Picture19.Visible = False
    Picture20.Visible = False
    Picture21.Visible = False
    Picture21.Visible = False
    Picture22.Visible = False
    Picture23.Visible = False
    Picture24.Visible = False
    Picture25.Visible = False
    Picture26.Visible = False
    Picture27.Visible = False
    Picture28.Visible = False
    Picture29.Visible = False
    Picture30.Visible = False
    Picture31.Visible = False
    Picture32.Visible = False
    Picture33.Visible = False
    Picture34.Visible = False
    Picture35.Visible = False
    Picture36.Visible = False
    Picture37.Visible = False
    Picture38.Visible = False
    Picture39.Visible = False
    Picture40.Visible = False
    
    cboField.Text = "All"
    'rs.Close
    
    lb.Open "Select * from Data", con, adOpenDynamic, adLockOptimistic
    
    'Hiding the Record Retrieving TextBoxes and other objects
    HideObjects
        
    'Showing the Record Retrieving Labels
    ShowLabels
        
    cmdSave.Visible = False
    cmdEdit.Visible = True
    cmdEdit.TabIndex = "58"
        
    cmdAddNew.Visible = True
    cmdCancel.Visible = False
    cmdOpenExisting.Visible = False
    cmdAddNew.TabIndex = "2"
        
    lb.MoveFirst
    DisplayLB
        
    cmdDelete.Visible = True
    cmdClear.Visible = False
    cmdDelete.TabIndex = "60"
    
    cmdFirst.Visible = True
    cmdLast.Visible = True
    cmdPrevious.Visible = True
    cmdNext.Visible = True
        
    lblInfo.Visible = False
        
    txtRecordID.Text = lb!ID
        
End Sub

Private Sub cboField_Click()

    'This section deals with the value cboField holds and the action to be performed for each value
    cboKeyword.Clear
    txtKeyword.Enabled = True
    
    txtSearch.Text = ""
    
    lblSearchDOB.Visible = False
    If cboField.Text = "DOB" Then
    
        lblSearchDOB.Visible = True
    
    End If
    
    If cboField.Text = "All" Then
    
        cboKeyword.Enabled = False
        txtKeyword.Enabled = False
        
    End If

    If cboField.Text = "Occupation" Then
    
        cboKeyword.Enabled = True
        
        cboKeyword.AddItem "All"
        cboKeyword.AddItem "Person"
        cboKeyword.AddItem "Father"
        cboKeyword.AddItem "Mother"
        
    End If
    
    If cboField.Text = "Plot No" Or cboField.Text = "Building Name" Or cboField.Text = "Locality" _
    Or cboField.Text = "City" Or cboField.Text = "State" Or cboField.Text = "Country" Or _
    cboField.Text = "Postal Code" Then
    
        cboKeyword.Enabled = True
        
        cboKeyword.AddItem "All"
        cboKeyword.AddItem "Permanent"
        cboKeyword.AddItem "Residential"
        cboKeyword.AddItem "Employer"
        
    End If
    
    If cboField.Text = "Criminal" Or cboField.Text = "Healthy" Then
    
        cboKeyword.Enabled = True
        txtKeyword.Enabled = False
        
        cboKeyword.AddItem "Yes"
        cboKeyword.AddItem "No"
        
    End If
    
    If cboField.Text = "Name" Then
    
        cboKeyword.Enabled = True
       
        cboKeyword.AddItem "All"
        cboKeyword.AddItem "First"
        cboKeyword.AddItem "Last"
        cboKeyword.AddItem "Father"
        cboKeyword.AddItem "Mother"
        
    End If
     
    If cboField.Text = "Siblings" Then
    
        cboKeyword.Enabled = True
        txtKeyword.Enabled = False
        
        cboKeyword.AddItem "Yes"
        cboKeyword.AddItem "No"
        
    End If
    
    If cboField.Text = "Gender" Then
    
        cboKeyword.Enabled = True
        txtKeyword.Enabled = False
        
        cboKeyword.AddItem "Male"
        cboKeyword.AddItem "Female"
        
    End If
    
    If cboField.Text = "Religion" Then
    
        cboKeyword.Enabled = True
        txtKeyword.Enabled = False
        
        cboKeyword.AddItem "Hinduism"
        cboKeyword.AddItem "Islam "
        cboKeyword.AddItem "Christianity "
        cboKeyword.AddItem "Sikhism "
        cboKeyword.AddItem "Buddhism "
        cboKeyword.AddItem "Jainism"
        
    End If

End Sub

Private Sub cmdSearch_Click()
    
    Dim b As Integer
    Dim str As String
    
    b = 1
    str = ""
    txtSearch.Text = ""
    
    If cmdSave.Visible = True Or cmdClear.Visible = True Then
    
        MsgBox "Please switch format", vbExclamation, "Search not available"
        
    Else
    
        If cboField.Text <> "" Then
        
            'lb.Close
            Set lb = New Recordset
            
            If cboField.Text = "All" Then
                
                lb.Open "Select * from Data", con, adOpenDynamic, adLockOptimistic
                DisplayLB
                txtSearch.Text = "ok"
                cboType.Text = "Search"
                MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                "Search Report available"
                    
            End If
            
            If cboField.Text = "Criminal" Then
                
                lb.Open "Select * from Data where Criminal = '" + cboKeyword.Text + " '", _
                con, adOpenDynamic, adLockOptimistic
                
                If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                    
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            
            If cboField.Text = "Healthy" Then
            
                lb.Open "Select * from Data where Healthy = '" + cboKeyword.Text + " '", _
                con, adOpenDynamic, adLockOptimistic
                
                If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                    
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            
            If cboField.Text = "Name" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where FirstName = '" + txtKeyword.Text + "' or " _
                    & "LastName = '" + txtKeyword.Text + "' or FathersName = '" + txtKeyword.Text + "' " _
                    & "or MothersName = '" + txtKeyword.Text + "'", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                        
                    End If
                End If
                
                If cboKeyword.Text = "First" Then
                
                    lb.Open "Select * from Data where FirstName = '" + txtKeyword.Text + " '", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Last" Then
                
                    lb.Open "Select * from Data where LastName = '" + txtKeyword.Text + " '", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Father" Then
                
                    lb.Open "Select * from Data where FathersName = '" + txtKeyword.Text + " '", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Mother" Then
                
                    lb.Open "Select * from Data where MothersName = '" + txtKeyword.Text + " '", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
            
            If cboField.Text = "Gender" Then
            
                lb.Open "Select * from Data where Gender = '" + cboKeyword.Text + " '", _
                con, adOpenDynamic, adLockOptimistic
                
                If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                      
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            If cboField.Text = "DOB" Then
            
                lb.Open "Select * from Data where DOB = '" + txtKeyword.Text + " '", _
                con, adOpenDynamic, adLockOptimistic
                
                If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                        
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            If cboField.Text = "Religion" Then
            
                lb.Open "Select * from Data where Religion = '" + cboKeyword.Text + " '", _
                con, adOpenDynamic, adLockOptimistic
                
                If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                       
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            If cboField.Text = "Plot No" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where PermanentPlotNo = '" + txtKeyword.Text + " '" _
                    & "or ResidentialPlotNo = '" + txtKeyword.Text + " ' or EmployerPlotNo = " _
                    & "'" + txtKeyword.Text + " '", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Permanent" Then
                
                    lb.Open "Select * from Data where PermanentPlotNo = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                    
                If cboKeyword.Text = "Residential" Then
                
                    lb.Open "Select * from Data where ResidentialPlotNo = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Employer" Then
                
                    lb.Open "Select * from Data where EmployerPlotNo = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
            
            If cboField.Text = "Building Name" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where PermanentBuildingName = '" + txtKeyword.Text + " '" _
                    & "or ResidentialBuildingName = '" + txtKeyword.Text + " ' or EmployerBuildingName = " _
                    & "'" + txtKeyword.Text + " '", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Permanent" Then
                
                    lb.Open "Select * from Data where PermanentBuildingName = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                    
                If cboKeyword.Text = "Residential" Then
                
                    lb.Open "Select * from Data where ResidentialBuildingName = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Employer" Then
                
                    lb.Open "Select * from Data where EmployerBuildingName = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
            
            If cboField.Text = "Locality" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where PermanentLocality = '" + txtKeyword.Text + " '" _
                    & "or ResidentialLocality = '" + txtKeyword.Text + " ' or EmployerLocality = " _
                    & "'" + txtKeyword.Text + " '", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Permanent" Then
                
                    lb.Open "Select * from Data where PermanentLocality = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                    
                If cboKeyword.Text = "Residential" Then
                
                    lb.Open "Select * from Data where ResidentialLocality = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Employer" Then
                
                    lb.Open "Select * from Data where EmployerLocality = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
            
            If cboField.Text = "City" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where PermanentCity = '" + txtKeyword.Text + " '" _
                    & "or ResidentialCity = '" + txtKeyword.Text + " ' or EmployerCity = " _
                    & "'" + txtKeyword.Text + " '", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Permanent" Then
                
                    lb.Open "Select * from Data where PermanentCity = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                    
                If cboKeyword.Text = "Residential" Then
                
                    lb.Open "Select * from Data where ResidentialCity = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Employer" Then
                
                    lb.Open "Select * from Data where EmployerCity = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
            
            If cboField.Text = "State" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where PermanentState = '" + txtKeyword.Text + " ' " _
                    & "or ResidentialState = '" + txtKeyword.Text + " ' or EmployerState = " _
                    & "'" + txtKeyword.Text + " '", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                     
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Permanent" Then
                
                    lb.Open "Select * from Data where PermanentState = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                    
                If cboKeyword.Text = "Residential" Then
                
                    lb.Open "Select * from Data where ResidentialState = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Employer" Then
                
                    lb.Open "Select * from Data where EmployerState = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
            
            If cboField.Text = "Country" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where PermanentCountry = '" + txtKeyword.Text + " ' " _
                    & "or ResidentialCountry = '" + txtKeyword.Text + " ' or EmployerCountry = " _
                    & "'" + txtKeyword.Text + " '", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Permanent" Then
                
                    lb.Open "Select * from Data where PermanentCountry = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                    
                If cboKeyword.Text = "Residential" Then
                
                    lb.Open "Select * from Data where ResidentialCountry = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Employer" Then
                
                    lb.Open "Select * from Data where EmployerCountry = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
            
            If cboField.Text = "Postal Code" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where PermanentPostalCode = '" + txtKeyword.Text + " ' " _
                    & "or ResidentialPostalCode = '" + txtKeyword.Text + " ' or EmployerPostalCode = " _
                    & "'" + txtKeyword.Text + " '", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Permanent" Then
                
                    lb.Open "Select * from Data where PermanentPostalCode = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                    
                If cboKeyword.Text = "Residential" Then
                
                    lb.Open "Select * from Data where ResidentialPostalCode = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Employer" Then
                
                    lb.Open "Select * from Data where EmployerPostalCode = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
            
            
            If cboField.Text = "Mobile Number" Then
            
                lb.Open "Select * from Data where Contact = '" + txtKeyword.Text + "'", _
                con, adOpenDynamic, adLockOptimistic
                
                If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                      
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            
            If cboField.Text = "STD Code" Then
            
            lb.Open "Select * from Data where DialCode = '" + txtKeyword.Text + "'", _
            con, adOpenDynamic, adLockOptimistic
                
                If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                      
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            If cboField.Text = "Telephone Number" Then
            
            lb.Open "Select * from Data where HomeContact = '" + txtKeyword.Text + "'", _
            con, adOpenDynamic, adLockOptimistic
                
                If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                      
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            If cboField.Text = "Occupation" Then
            
                If cboKeyword.Text = "All" Then
                
                    lb.Open "Select * from Data where Occupation = '" + txtKeyword.Text + "'" _
                    & "or FathersOccupation = '" + txtKeyword.Text + "' or MothersOccupation = " _
                    & "'" + txtKeyword.Text + "'", con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Person" Then
                
                    lb.Open "Select * from Data where Occupation = '" + txtKeyword.Text + "'", _
                    con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                If cboKeyword.Text = "Father" Then
                
                lb.Open "Select * from Data where FathersOccupation = '" + txtKeyword.Text + "'", _
                con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
                
                
                If cboKeyword.Text = "Mother" Then
            
                lb.Open "Select * from Data where MothersOccupation = '" + txtKeyword.Text + "'", _
                con, adOpenDynamic, adLockOptimistic
                    
                    If lb.EOF Then
                
                        MsgBox "Nothing found", vbInformation, "Search complete"
                        txtSearch.Text = "no"
                    
                    Else
                
                        DisplayLB
                        txtSearch.Text = "ok"
                        MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                        cboType.Text = "Search"
                        MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                        "Search Report available"
                    
                    End If
                End If
            End If
                
            
            If cboField.Text = "Siblings" Then
            
                lb.Open "Select * from Data where Siblings ='" + cboKeyword.Text + "'", _
                con, adOpenDynamic, adLockOptimistic
                
                 If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                        
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            If cboField.Text = "Brothers" Then
            
                lb.Open "Select * from Data where Brothers ='" + cboKeyword.Text + "'", _
                con, adOpenDynamic, adLockOptimistic
                
                 If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                       
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
            If cboField.Text = "Sisters" Then
            
                lb.Open "Select * from Data where Sisters ='" + cboKeyword.Text + "'", _
                con, adOpenDynamic, adLockOptimistic
                
                 If lb.EOF Then
                
                    MsgBox "Nothing found", vbInformation, "Search complete"
                    txtSearch.Text = "no"
                       
                Else
                
                    DisplayLB
                    txtSearch.Text = "ok"
                    MsgBox "Record(s) found", vbInformation, "Click > or < to view"
                    cboType.Text = "Search"
                    MsgBox "Click on Report to view Search Report", vbInformation + vbOKOnly, _
                    "Search Report available"
                    
                End If
            End If
            
        Else
       
            MsgBox "Empty fields", vbInformation, "Note"
            
        End If
    End If
        
End Sub

Private Sub cmdReport_Click()
 
    If cboType.Text <> "" And cmdSave.Visible = False And cmdClear.Visible = False Then
    
        If cboType.Text = "Current" Then
        
            MasterEnvironment.rscmdCurrentReport.Open "Select * from Data where ID = " _
            & "'" + txtRecordID.Text + " '"
            CurrentReport.Refresh
            CurrentReport.Show
            MasterEnvironment.rscmdCurrentReport.Close
            
        End If
    
        If cboType.Text = "Complete" Then
        
            MasterEnvironment.rscmdCompleteReport.Open "Select * from Data"
            CompleteDataReport.Refresh
            CompleteDataReport.Show
            MasterEnvironment.rscmdCompleteReport.Close
            
            
        End If
            
        If cboType.Text = "Search" Then
            
            If txtSearch.Text = "ok" Then
            
                ReportSearch
              
            Else
                
                If txtSearch.Text = "no" Then
                
                    MsgBox "Search Report unavailable", vbInformation, "No record(s) found"
                    
                Else
                
                    MsgBox "Please search first", vbInformation, "Note"
                    
                End If
            End If
        End If
            
    Else
        
        If Trim(cboType.Text) = "" Then
        
            MsgBox "Please choose report type", vbExclamation, "Invalid command"
            
        Else
    
            MsgBox "Please switch format", vbExclamation, "Report not available"
            
        End If
    End If

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

Private Sub Form_Activate()

    AfterLogin

End Sub

Private Sub Form_Load()
    
    frmAfterLogin.ProgressBar.Visible = True
    frmAfterLogin.ProgressBar.Value = 0
    
    cboKeyword.Enabled = False
    cboField.Enabled = True
    txtKeyword.Enabled = True
    
    txtSearch.Text = ""
    
    'Hide Everything
    HideObjects
    HideLabels
    
    'Adding Type of Reports
    cboType.AddItem "Current"
    cboType.AddItem "Complete"
    cboType.AddItem "Search"
    
    
    'Adding Type of Search in Field
    cboField.AddItem "All"
    cboField.AddItem "Criminal"
    cboField.AddItem "Healthy"
    cboField.AddItem "Name"
    cboField.AddItem "Gender"
    cboField.AddItem "DOB"
    cboField.AddItem "Religion"
    cboField.AddItem "Plot Number"
    cboField.AddItem "Building Name"
    cboField.AddItem "Locality"
    cboField.AddItem "City"
    cboField.AddItem "State"
    cboField.AddItem "Country"
    cboField.AddItem "Postal Code"
    cboField.AddItem "Mobile Number"
    cboField.AddItem "STD Code"
    cboField.AddItem "Telephone Number"
    cboField.AddItem "Occupation"
    cboField.AddItem "Siblings"
    cboField.AddItem "Brothers"
    cboField.AddItem "Sisters"
    cboField.AddItem "Education"
    cboField.AddItem "Job Status"
    cboField.AddItem "Marital Status"
    
    frmAfterLogin.ProgressBar.Value = 10
    
    'Adding items in the combo boxes
    'Criminal and Healthy
        cboCriminal.AddItem "Yes"
        cboCriminal.AddItem "No"
        
        cboHealthy.AddItem "Yes"
        cboHealthy.AddItem "No"
        
    frmAfterLogin.ProgressBar.Value = 20
        
    'Gender
        cboGender.AddItem "Male"
        cboGender.AddItem "Female"
        
    'Religion
        cboReligion.AddItem "Hinduism"
        cboReligion.AddItem "Islam "
        cboReligion.AddItem "Christianity "
        cboReligion.AddItem "Sikhism "
        cboReligion.AddItem "Buddhism "
        cboReligion.AddItem "Jainism"
        
    frmAfterLogin.ProgressBar.Value = 30
    
    'Permanent
    'City 1
        cboCity1.AddItem "Mumbai"
        cboCity1.AddItem "Navi Mumbai"
        cboCity1.AddItem "Pune"
        cboCity1.AddItem "Nagpur"
        cboCity1.AddItem "Thane"
        cboCity1.AddItem "Pimpri"
        cboCity1.AddItem "Chinchwad"
        cboCity1.AddItem "Nashik"
        cboCity1.AddItem "Kalyan"
        cboCity1.AddItem "Dombivali"
        cboCity1.AddItem "Vasai"
        cboCity1.AddItem "Virar"
        cboCity1.AddItem "Aurangabad"
        cboCity1.AddItem "Solapur"
        cboCity1.AddItem "Mira"
        cboCity1.AddItem "Bhayandar"
        cboCity1.AddItem "Bhiwandi"
        cboCity1.AddItem "Nizampur"
        cboCity1.AddItem "Amravati"
        cboCity1.AddItem "Nanded"
        cboCity1.AddItem "Waghala"
        cboCity1.AddItem "Panvel"
        cboCity1.AddItem "Sangli"
        cboCity1.AddItem "Akola"
        cboCity1.AddItem "Ahmednagar"
        cboCity1.AddItem "Parbhani"
        cboCity1.AddItem "Chandrapur"
        cboCity1.AddItem "Dhule"
        cboCity1.AddItem "Malegaon"
        cboCity1.AddItem "Jalgaon"
        cboCity1.AddItem "Kolhapur"
        cboCity1.AddItem "Nashik"
        cboCity1.AddItem "Latur"

    'State 1
        cboState1.AddItem "Maharashtra"
        
    frmAfterLogin.ProgressBar.Value = 40
    
    'Country 1
        cboCountry1.AddItem "India"
        
    frmAfterLogin.ProgressBar.Value = 50
    
    'Residential
    'City 2
        cboCity2.AddItem "Mumbai"
        cboCity2.AddItem "Navi Mumbai"
        cboCity2.AddItem "Pune"
        cboCity2.AddItem "Nagpur"
        cboCity2.AddItem "Thane"
        cboCity2.AddItem "Pimpri"
        cboCity2.AddItem "Chinchwad"
        cboCity2.AddItem "Nashik"
        cboCity2.AddItem "Kalyan"
        cboCity2.AddItem "Dombivali"
        cboCity2.AddItem "Vasai"
        cboCity2.AddItem "Virar"
        cboCity2.AddItem "Aurangabad"
        cboCity2.AddItem "Solapur"
        cboCity2.AddItem "Mira"
        cboCity2.AddItem "Bhayandar"
        cboCity2.AddItem "Bhiwandi"
        cboCity2.AddItem "Nizampur"
        cboCity2.AddItem "Amravati"
        cboCity2.AddItem "Nanded"
        cboCity2.AddItem "Waghala"
        cboCity2.AddItem "Panvel"
        cboCity2.AddItem "Sangli"
        cboCity2.AddItem "Akola"
        cboCity2.AddItem "Ahmednagar"
        cboCity2.AddItem "Parbhani"
        cboCity2.AddItem "Chandrapur"
        cboCity2.AddItem "Dhule"
        cboCity2.AddItem "Malegaon"
        cboCity2.AddItem "Jalgaon"
        cboCity2.AddItem "Kolhapur"
        cboCity2.AddItem "Nashik"
        cboCity2.AddItem "Latur"

    'State 2
        cboState2.AddItem "Maharashtra"
        
     frmAfterLogin.ProgressBar.Value = 60
     
    'Country 2
        cboCountry2.AddItem "India"
       
     frmAfterLogin.ProgressBar.Value = 70
    
    'Employers
    'City 3
        cboCity3.AddItem "Mumbai"
        cboCity3.AddItem "Navi Mumbai"
        cboCity3.AddItem "Pune"
        cboCity3.AddItem "Nagpur"
        cboCity3.AddItem "Thane"
        cboCity3.AddItem "Pimpri"
        cboCity3.AddItem "Chinchwad"
        cboCity3.AddItem "Nashik"
        cboCity3.AddItem "Kalyan"
        cboCity3.AddItem "Dombivali"
        cboCity3.AddItem "Vasai"
        cboCity3.AddItem "Virar"
        cboCity3.AddItem "Aurangabad"
        cboCity3.AddItem "Solapur"
        cboCity3.AddItem "Mira"
        cboCity3.AddItem "Bhayandar"
        cboCity3.AddItem "Bhiwandi"
        cboCity3.AddItem "Nizampur"
        cboCity3.AddItem "Amravati"
        cboCity3.AddItem "Nanded"
        cboCity3.AddItem "Waghala"
        cboCity3.AddItem "Panvel"
        cboCity3.AddItem "Sangli"
        cboCity3.AddItem "Akola"
        cboCity3.AddItem "Ahmednagar"
        cboCity3.AddItem "Parbhani"
        cboCity3.AddItem "Chandrapur"
        cboCity3.AddItem "Dhule"
        cboCity3.AddItem "Malegaon"
        cboCity3.AddItem "Jalgaon"
        cboCity3.AddItem "Kolhapur"
        cboCity3.AddItem "Nashik"
        cboCity3.AddItem "Latur"

    'State 3
        cboState3.AddItem "Maharashtra"
    
    'Country 3
        cboCountry3.AddItem "India"
        
    'Email Service Providers
        cboProvider.AddItem "Gmail"
        cboProvider.AddItem "Outlook"
        cboProvider.AddItem "Yahoo"
        cboProvider.AddItem "Aol"
        cboProvider.AddItem "Zoho"
        cboProvider.AddItem "Mail"
        cboProvider.AddItem "Yandex"
        cboProvider.AddItem "ProtonMail"
        cboProvider.AddItem "GMX"
        cboProvider.AddItem "iCloud"
        
    frmAfterLogin.ProgressBar.Value = 80
        
    'Occuption
        cboOccupation.AddItem "Agriculture"
        cboOccupation.AddItem "Business"
        cboOccupation.AddItem "Medical"
        cboOccupation.AddItem "Engineering"
        cboOccupation.AddItem "Law Practice"
        cboOccupation.AddItem "Government Service"
        cboOccupation.AddItem "Public Sector Service"
        cboOccupation.AddItem "Private Service"
        cboOccupation.AddItem "Teaching"
        cboOccupation.AddItem "Home Maker"
        cboOccupation.AddItem "Other"
        
    'Father's Occupation
        cboFathersO.AddItem "Agriculture"
        cboFathersO.AddItem "Business"
        cboFathersO.AddItem "Medical"
        cboFathersO.AddItem "Engineering"
        cboFathersO.AddItem "Law Practice"
        cboFathersO.AddItem "Government Service"
        cboFathersO.AddItem "Public Sector Service"
        cboFathersO.AddItem "Private Service"
        cboFathersO.AddItem "Teaching"
        cboFathersO.AddItem "Home Maker"
        cboFathersO.AddItem "Other"
    
    'Mother's Occupation
        cboMothersO.AddItem "Agriculture"
        cboMothersO.AddItem "Business"
        cboMothersO.AddItem "Medical"
        cboMothersO.AddItem "Engineering"
        cboMothersO.AddItem "Law Practice"
        cboMothersO.AddItem "Government Service"
        cboMothersO.AddItem "Public Sector Service"
        cboMothersO.AddItem "Private Service"
        cboMothersO.AddItem "Teaching"
        cboMothersO.AddItem "Home Maker"
        cboMothersO.AddItem "Other"
        
    frmAfterLogin.ProgressBar.Value = 90
        
    'Siblings
        cboSiblings.AddItem "Yes"
        cboSiblings.AddItem "No"
    
    'Marital status
        cboMarital.AddItem "Married"
        cboMarital.AddItem "Single"
        cboMarital.AddItem "Divorced"
        cboMarital.AddItem "Widowed"
        
    'Job Status
        cboJob.AddItem "Working"
        cboJob.AddItem "Idle"
        
    'Education status
        cboEducation.AddItem "M.B.A"
        cboEducation.AddItem "M.C.M"
        cboEducation.AddItem "M.Sc"
        cboEducation.AddItem "M.Tech"
        cboEducation.AddItem "B.C.C.A"
        cboEducation.AddItem "B.B.A"
        cboEducation.AddItem "B.Com"
        cboEducation.AddItem "B.Sc"
        cboEducation.AddItem "B.Arch"
        cboEducation.AddItem "B.Tech"
        cboEducation.AddItem "B.E"
        cboEducation.AddItem "12th "
        cboEducation.AddItem "11th"
        cboEducation.AddItem "10th"
        cboEducation.AddItem "9th"
        cboEducation.AddItem "8th"
        cboEducation.AddItem "7th"
        cboEducation.AddItem "6th"
        cboEducation.AddItem "5th"
        cboEducation.AddItem "4th"
        cboEducation.AddItem "3rd"
        cboEducation.AddItem "2nd"
        cboEducation.AddItem "1st"
        
    frmAfterLogin.ProgressBar.Value = 100
    
    str = ""
    
    'Adding the wrong entre image in boxes pic 1 to 12
        Picture1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture3.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture4.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture5.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture6.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture7.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture8.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture9.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture10.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture11.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture12.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture13.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture14.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture15.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture16.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture17.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture18.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture19.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture20.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture21.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture22.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture23.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture24.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture25.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture26.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture27.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture28.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture29.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture30.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture31.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture32.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture33.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture34.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture35.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture36.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture37.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture38.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture39.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        Picture40.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\WrongEntry.jpg")
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    con.Close

End Sub

