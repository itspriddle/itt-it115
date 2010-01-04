VERSION 5.00
Begin VB.Form frmUsersEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit User Info"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmUsersEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSaveChanges 
      Caption         =   "Save Changes"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox cmbSecurity 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblUsernameWarning 
      Caption         =   "* Usernames are not case sensitive."
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblMSG 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblSecurity 
      Caption         =   "Security Level:"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblUsername 
      Caption         =   "* Username:"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmUsersEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdBack_Click
' Parameters  :
' Description :       Go Back
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdBack_Click_Err
    '</EhHeader>
    Unload Me
    'Unload this frame
    
    'frmUsersView.Visible = True
    'Load the last frame
    
    'con.Close
    'Close the DB Connection
    
    '<EhFooter>
    Exit Sub

cmdBack_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersEdit.cmdBack_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdSaveChanges_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdSaveChanges_Click
' Parameters  :
' Description :       Save Changes
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdSaveChanges_Click_Err
    '</EhHeader>
    Dim strSQL As String

    'Create a string for the SQL statement
    If txtFirstName <> "" And txtLastName <> "" And txtUsername <> "" And _
            txtPassword <> "" And cmbSecurity.Text <> "" Then

        'Make sure they entered data into the textboxes
        If cmbSecurity.Text = "User" Or cmbSecurity.Text = "Manager" Then
            'Make sure they chose Admin or User
            strSQL = "UPDATE T_AUTHUSERS "
            strSQL = strSQL + "SET F_LNAME='" & txtLastName.Text & "', "
            strSQL = strSQL + "F_FNAME='" & txtFirstName.Text & "', "
            strSQL = strSQL + "F_USERNAME='" & txtUsername.Text & "', "
            strSQL = strSQL + "F_PASSWORD='" & txtPassword.Text & "', "
            strSQL = strSQL + "F_SECURITYLVL='" & cmbSecurity.Text & "' "
            strSQL = strSQL + "WHERE "
            strSQL = strSQL + "F_ID=" & txtID.Text & ""
                      
            'Execute the SQL statement
            con.Execute (strSQL)
    
            'Give them a success message
            MsgBox "Info changed.", , "User Added!"
        Else
            'Let them know their mistake
            MsgBox "Choose either User or Manager for security level.", , _
                    "Error!"
        End If
        
    Else
        'Let them know their mistake
    End If

    '<EhFooter>
    Exit Sub

cmdSaveChanges_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersEdit.cmdSaveChanges_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
'-----------------------------------------------------------------------------
' Procedure   :       Form_Load
' Parameters  :
' Description :       Perform at load time...
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>

    cmbSecurity.AddItem ("User")
    cmbSecurity.AddItem ("Manager")
    'Add these values to the dropdown
    
    lblMSG.Caption = "The following information has been found for " & _
            frmUsersView.txtFirstName & " " & frmUsersView.txtLastName & "."
    'Set the lbl to the values from frmUsersView
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersEdit.Form_Load " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       Form_Unload
' Parameters  :       Cancel (Integer)
' Description :       Close the form and perform any needed cleanup
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>
    frmUsersView.Visible = True
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersEdit.Form_Unload " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

