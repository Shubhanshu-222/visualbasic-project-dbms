VERSION 5.00
Begin VB.Form frmAboutCreators 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Creators of Master 1.0"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5430
   Icon            =   "frmAboutCreators.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   5430
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
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   20
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox picGroup 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   18
      Top             =   4200
      Width           =   5175
   End
   Begin VB.PictureBox picShubhanshu 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picAman 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   1920
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picGunjan 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   3720
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox picResize3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   3720
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picResize1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picResize2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   1920
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picResize4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   19
      Top             =   4200
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rashtrasant Tukdoji Maharaj Nagpur University"
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
      TabIndex        =   17
      Top             =   3840
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dr. Ambedkar Institute of Management Studies and Research"
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
      TabIndex        =   16
      Top             =   3600
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pursuing B.Com (Computer Application)"
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
      TabIndex        =   15
      Top             =   3360
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5280
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Database Manager"
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
      Left            =   3720
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Software Designer"
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
      Left            =   1920
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Software Programmer"
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
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Raut"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gunjan"
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
      Left            =   3720
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sawarkar"
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
      Left            =   1920
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aman"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bansod"
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
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shubhanshu"
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
      Top             =   1920
      Width           =   1455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAboutCreators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdok_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Loading Creators Images and resizing it
    
    picResize1.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Shubhanshu.jpg")
    picShubhanshu.PaintPicture picResize1.Picture, 0, 0, picShubhanshu.ScaleWidth, _
    picShubhanshu.ScaleHeight, 0, 0, picResize1.ScaleWidth, picResize1.ScaleHeight, vbSrcCopy
    
    picResize2.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Aman.jpg")
    picAman.PaintPicture picResize2.Picture, 0, 0, picAman.ScaleWidth, _
    picAman.ScaleHeight, 0, 0, picResize2.ScaleWidth, picResize2.ScaleHeight, vbSrcCopy
    
    picResize3.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Gunjan.jpg")
    picGunjan.PaintPicture picResize3.Picture, 0, 0, picGunjan.ScaleWidth, _
    picGunjan.ScaleHeight, 0, 0, picResize3.ScaleWidth, picResize3.ScaleHeight, vbSrcCopy
    
    'Resiing Group Image
    picResize4.Picture = LoadPicture("C:\Users\Hp\Desktop\VB6new\Pictures\Group.jpg")
    picGroup.PaintPicture picResize4.Picture, 0, 0, picGroup.ScaleWidth, _
    picGroup.ScaleHeight, 0, 0, picResize4.ScaleWidth, picResize4.ScaleHeight, vbSrcCopy
    
End Sub
