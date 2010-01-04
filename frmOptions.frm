VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Change Password"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton cmdSaveChanges 
      Caption         =   "Save Changes"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblNewPassword 
      Caption         =   "New Password 2:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblNewPassword 
      Caption         =   "New Password 1:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblCurrentPassword 
      Caption         =   "Current Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblToChange 
      Caption         =   "To change your password, first enter your current password and then enter your new password twice to avoid mistakes."
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdCancel_Click
' Parameters  :
' Description :       Cancel change password
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdCancel_Click_Err
    '</EhHeader>
    Unload Me
    'frmMain.Visible = True
    
    '<EhFooter>
    Exit Sub

cmdCancel_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmOptions.cmdCancel_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdSaveChanges_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdSaveChanges_Click
' Parameters  :
' Description :       Save changes to password
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdSaveChanges_Click_Err
    '</EhHeader>

    If Me.txtOldPassword.Text <> "" And Me.txtNewPassword(0) <> "" And _
            Me.txtNewPassword(1) <> "" Then
    
        If Me.txtNewPassword(0) = Me.txtNewPassword(1) Then

            Dim rs As New ADODB.recordset
            rs.Open "SELECT F_PASSWORD FROM T_AUTHUSERS WHERE F_USERNAME='" & _
                    strUsername & "'", con, adOpenDynamic, adLockOptimistic
            
            If rs!F_PASSWORD = Me.txtOldPassword Then
                Dim strSQL As String
            
                strSQL = strSQL + "UPDATE T_AUTHUSERS "
                strSQL = strSQL + "SET F_PASSWORD='"
                strSQL = strSQL + Me.txtNewPassword(0) & "'"
                strSQL = strSQL + " WHERE F_USERNAME='"
                strSQL = strSQL + strUsername & "'"
            
                con.Execute (strSQL)
                'Execute the update
            
                MsgBox "Your password has been changed to '" & _
                        Me.txtNewPassword(0) & "'. Write it down!", , _
                        "Password Changed"
                'Let the user know everything went okay
            
                Unload Me
                'Unload this form
            
                frmMain.Visible = True
                'Show the main form
            
            Else
                'The password didn't match
                MsgBox _
                        "The current password you entered does not match the one on file.", , _
                        "Password Error"
                'Let the user know the error
                        
            End If
        
        Else
            'New pass 1 and new pass 2
        
            Dim mbrPassword As VbMsgBoxResult
            'New msgbox result
        
            mbrPassword = MsgBox( _
                    "The New Password 1 did not match New Password 2." & _
                    "Would you like to unmask the text boxes?", vbYesNo, _
                    "Password Error")
            'Let the user know the passwords didnt match
            'ask if they want to unmask the password boxes
        
            If mbrPassword = vbYes Then
                'If they do want to unmask them
            
                Me.txtNewPassword(0).PasswordChar = ""
                Me.txtNewPassword(1).PasswordChar = ""
                Me.txtOldPassword.PasswordChar = ""
                'Do it...
                
            End If
        
        End If
    
    Else
        'They didn't enter all data
    
        MsgBox "You must fill in all fields.", , "Error"
        'Let the user know
        
    End If

    '<EhFooter>
    Exit Sub

cmdSaveChanges_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmOptions.cmdSaveChanges_Click " & "at line " _
            & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       Form_Unload
' Parameters  :       Cancel (Integer)
' Description :       Close this form and perform any needed cleanup
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>
    frmMain.Visible = True
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmOptions.Form_Unload " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

