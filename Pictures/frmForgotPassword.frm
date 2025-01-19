VERSION 5.00
Begin VB.Form frmdE 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alternate login method"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmdE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdCheck_Click()
 
    If rs.Fields("Answer") = txtAns.Text Then
    
        mdiMain.Show
        frmGeneral.txtPrimary.Text = txtPrimary.Text
    
        Set frmForgotPassword = Nothing
        Unload Me
    Else
    
        MsgBox "Please try again", vbCritical, "Incorrect Answer"
    
    End If

End Sub

Private Sub cmdSubmit_Click()
    
    If txtUsername.Text = "" Then
        MsgBox "Enter you Username", vbCritical, "Blank field(s)"
    Else
        rs.Open "Select * from Security where Username = '" + txtUsername.Text + "'", con, rspenDynamic, adLockOptimistic
        
            If rs.Fields("DOB").Value = txtDOB.Text And rs.Fields("Contact").Value = txtContact.Text Then
                txtPrimary.Text = rs!Primary

                lblSecurityQues.Visible = True
                lblEnterAns.Visible = True
                
                lblQues.Visible = True
                lblQues.Caption = rs!Security
                
                txtAns.Visible = True
                
                cmdCheck.Visible = True
            Else
                
                MsgBox "Please try again.", vbCritical, "Incorrect Field(s)"
                rs.Close
                
            End If
        End If
    
End Sub

Private Sub Form_Load()
    
    con.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source= C:\Users\Hp\Desktop\VB6new\Security.mdb; Persist Security Info= False"
    
    lblSecurityQues.Visible = False
    lblEnterAns.Visible = False
    
    lblQues.Visible = False
    txtAns.Visible = False
    
    cmdCheck.Visible = False

End Sub
