VERSION 5.00
Begin VB.Form frmAboutMaster 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Master 1.0"
   ClientHeight    =   4320
   ClientLeft      =   -165
   ClientTop       =   390
   ClientWidth     =   5550
   Icon            =   "frmAboutMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
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
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   14
      Top             =   2040
      Width           =   1575
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   3840
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picResize 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   3840
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Forms10: 11.0.1111.000"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   3960
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "VBA: Retail 1.1.0001"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   3960
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0001"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5400
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAboutMaster.frx":7FC7
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   5295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number:  00000000000000000000"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft Windows"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Shubhanshu Aman Gunjan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   120
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This product is licensed to:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1987-1988 Microsoft Corp."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "For 32-bit Windows System"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "General Database Management System"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3375
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFName 
      BackStyle       =   0  'Transparent
      Caption         =   "Master 1.0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAboutMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdok_Click()

    Set frmAboutMaster = Nothing
    Unload Me
    
End Sub

Private Sub Form_Load()

    'Loading logo and resizing it
    picResize.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\24.jpg")
    picLogo.PaintPicture picResize.Picture, 0, 0, picLogo.ScaleWidth, _
    picLogo.ScaleHeight, 0, 0, picResize.ScaleWidth, picResize.ScaleHeight, vbSrcCopy
    

End Sub
