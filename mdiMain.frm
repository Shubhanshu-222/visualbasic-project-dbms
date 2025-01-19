VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Welcome to Master Window"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   9840
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   19560
      Top             =   9600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7FC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":80D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8273
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8385
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":87D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8C29
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":8D3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":918D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":95DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9A31
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":9E83
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A2D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":A727
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarMain 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   2310
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Caps Lock on"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NUM"
            Object.ToolTipText     =   "Num Lock on"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "INS"
            Object.ToolTipText     =   "Active Insert"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2/16/2017"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "1:40 AM"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar MainToolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add new"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "First record"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous record"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next record"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Last record"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Current report"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Complete report"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search Report"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAddNew 
         Caption         =   "&Add New"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Ca&ncel"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMove 
      Caption         =   "&Move"
      Begin VB.Menu mnuFirst 
         Caption         =   "&First"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Nex&t"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "&Previous"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuLast 
         Caption         =   "&Last"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuCurrentRecord 
         Caption         =   "Cu&rrent Record"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuCompleteRecord 
         Caption         =   "Co&mplete Record(s)"
         Checked         =   -1  'True
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSearchReport 
         Caption         =   "Search Report"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuMe 
         Caption         =   "My Acco&unt"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaster 
         Caption         =   "A&bout Master 1.0"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuAboutCreators 
         Caption         =   "About &Creators"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuExit1 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MainToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Index = 1 Then
    
        frmGeneral.AddNewGeneral
        
    End If
    
    If Button.Index = 2 Then
    
        frmGeneral.DeleteGeneral
        
    End If
    
    If Button.Index = 4 Then
    
        frmGeneral.SaveGeneral
        
    End If
    
    If Button.Index = 5 Then
    
        frmGeneral.ClearGeneral
        
    End If
    
    If Button.Index = 7 Then
    
        frmGeneral.EditGeneral
        
    End If
    
    If Button.Index = 8 Then
    
        frmGeneral.CancelGeneral
        
    End If
    
    If Button.Index = 10 Then
    
        frmGeneral.FirstGeneral
        
    End If
    
    If Button.Index = 11 Then
    
        frmGeneral.NextGeneral
        
    End If
    
    If Button.Index = 12 Then
    
        frmGeneral.PreviousGeneral
        
    End If
    
    If Button.Index = 13 Then
    
        frmGeneral.LastGeneral
        
    End If
    
    If Button.Index = 15 Then
    
        frmGeneral.Current
        
    End If
    
    If Button.Index = 16 Then
    
        frmGeneral.Complete
        
    End If
    
    If Button.Index = 17 Then
    
        If frmGeneral.txtSearch.Text = "ok" Then
        
            frmGeneral.ReportSearch
            
        End If
        
        If frmGeneral.txtSearch.Text = "no" Then
        
            MsgBox "Search Report unavailable", vbInformation, "No record(s) found"
            
        End If
        
        If frmGeneral.txtSearch.Text = "" Then
        
            MsgBox "Please search first", vbInformation, "Note"
        
        End If
        
    End If
    
End Sub

Private Sub mnuAboutCreators_Click()

    frmAboutCreators.Show
    
End Sub

Private Sub mnuAddNew_Click()

    frmGeneral.AddNewGeneral

End Sub

Private Sub mnuCancel_Click()

    frmGeneral.CancelGeneral
    
End Sub

Private Sub mnuClear_Click()

    frmGeneral.ClearGeneral

End Sub

Private Sub mnuCompleteRecord_Click()
    
    frmGeneral.Complete
    
End Sub

Private Sub mnuCurrentRecord_Click()

    frmGeneral.Current
    
End Sub

Private Sub mnuDelete_Click()

    frmGeneral.DeleteGeneral
    
End Sub

Private Sub mnuEdit_Click()
    
    frmGeneral.EditGeneral

End Sub

Private Sub mnuExit_Click()
    
    If MsgBox("Are you sure?", vbCritical + vbYesNo, "Master") = vbYes Then
    
        Unload Me
        
    End If
    
End Sub

Private Sub mnuExit1_Click()
    
    If MsgBox("Are you sure?", vbCritical + vbYesNo, "Master") = vbYes Then
    
        Unload Me
        
    End If
    
End Sub

Private Sub mnuFirst_Click()

    frmGeneral.FirstGeneral
    
End Sub

Private Sub mnuLast_Click()

   frmGeneral.LastGeneral
   
End Sub

Private Sub mnuNext_Click()

    frmGeneral.NextGeneral
    
End Sub

Private Sub mnuPrevious_Click()

    frmGeneral.PreviousGeneral
    
End Sub

Private Sub mnuMaster_Click()

    frmAboutMaster.Show

End Sub

Private Sub mnuMe_Click()

    frmAboutMe.Show
    frmAboutMe.txtPrimary.Text = frmGeneral.txtPrimary.Text
    
End Sub

Private Sub mnuSave_Click()

    frmGeneral.SaveGeneral

End Sub

Private Sub MDIForm_Load()

    frmGeneral.Show
    frmGeneral.CenterChild mdiMain, frmGeneral
    frmGeneral.ScaleLeft = "0"
    frmGeneral.ScaleTop = "0"
    
End Sub

Private Sub mnuSearchReport_Click()

    If frmGeneral.txtSearch.Text = "ok" Then
        
        frmGeneral.ReportSearch
            
    End If
        
    If frmGeneral.txtSearch.Text = "no" Then
        
        MsgBox "Search Report unavailable", vbInformation, "No record(s) found"
            
    End If
        
    If frmGeneral.txtSearch.Text = "" Then
        
        MsgBox "Please search first", vbInformation, "Note"
        
    End If
    
End Sub
