VERSION 5.00
Begin VB.Form frmUsersAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmUsersAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
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
   Begin VB.CommandButton cmdUsersAdd 
      Caption         =   "Add User"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
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
      TabIndex        =   13
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblMSG 
      Caption         =   "Fill in each of the fields to add a new employee to the database.  Make sure to select either User or Manager."
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblSecurity 
      Caption         =   "Security Level:"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblUsername 
      Caption         =   "* Username:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmUsersAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdBack_Click
' Parameters  :
' Description :       Back to frmUsersView
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdBack_Click_Err
    '</EhHeader>
    Unload Me
    'Unload this form
    
    'frmUsersView.Visible = True
    'Show the main Users Form
    
    '<EhFooter>
    Exit Sub

cmdBack_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersAdd.cmdBack_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdUsersAdd_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdUsersAdd_Click
' Parameters  :
' Description :       Add a new user to T_AUTHUSERS
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdUsersAdd_Click_Err
    '</EhHeader>
    Dim strSQLAdd As String

    'Create a string called strSQL to hold
    'the SQL statement
    If txtFirstName <> "" And txtLastName <> "" And txtUsername <> "" And _
            txtPassword <> "" Then

        'If they filled in ALL of the text boxes
        If cmbSecurity.Text = "User" Or cmbSecurity.Text = "Manager" Then
            'If they chose User or Manager
        
            Dim rs2 As New recordset
            'Create a new recordset
        
            rs2.Open "SELECT * FROM T_AUTHUSERS WHERE F_USERNAME='" & _
                    txtUsername & "'", con, adOpenDynamic, adLockOptimistic
            'Make sure there are no duplicate usernames.
        
            If rs2.RecordCount <> 0 Then
                'If the count isn't 0 then someone else already has that name.
        
                MsgBox "The Username " & txtUsername & _
                        " is already taken.  Please choose another.", , _
                        "Username Taken!"
                'Give the user a message box asking them to choose a different username
            
            Else
                'They entered a unique username
        
                'Begin the SQL statement
                'Insert the info from each text box into the appropriate
                'DB table field
                strSQLAdd = strSQLAdd + "INSERT INTO T_AUTHUSERS (F_FNAME,"
                strSQLAdd = strSQLAdd + "F_LNAME,"
                strSQLAdd = strSQLAdd + "F_USERNAME,"
                strSQLAdd = strSQLAdd + "F_PASSWORD,"
                strSQLAdd = strSQLAdd + "F_SECURITYLVL)"
                strSQLAdd = strSQLAdd + "VALUES ('" & txtFirstName.Text & _
                        "','" & txtLastName.Text & "','" & LCase$( _
                        txtUsername.Text) & "','" & txtPassword.Text & "','" _
                        & cmbSecurity.Text & "')"
            
                'Send the statement to the DB,
                'add the info
                con.Execute (strSQLAdd)
            
                'Let the user know everything went OK, but lets get sassy shall we?
            
                Dim mbxYesNo As VbMsgBoxResult
                'Create a var to hold the result of what the user clicks, Yes or No
                'By doing this we can perform different actions for what gets clicked.
            
                mbxYesNo = MsgBox(Me.txtFirstName.Text & " " & _
                        Me.txtLastName.Text & _
                        " has been added to the database with the username '" _
                        & Me.txtUsername.Text & "' and password '" & _
                        Me.txtPassword.Text & _
                        "'.  Be sure to write these down!" & vbNewLine & _
                        vbNewLine & "Would you like to add another user?", _
                        vbYesNo, "User Added")
                'Print a message box telling the User all of the info that was added to the
                'database.  Ask them if they would like to add another user.
            
                If mbxYesNo = vbYes Then
                    'If they clicked yes
            
                    rs2.Close
                    'Close the recordset we used to check the username
                    'since we'll need to use it again
                
                    txtFirstName.Text = ""
                    txtLastName.Text = ""
                    txtUsername.Text = ""
                    txtPassword.Text = ""
                    'Clear the contents of the text boxes
                
                Else
                    'They are done adding users
            
                    Unload Me
                    'Unload this form, we're done with it for now.
                
                    frmUsersView.Visible = True
                    'Show the main user form
                
                    frmUsersView.rs.Requery
                    'Requery the recordset on the main users form
                    'so that the recently added users are in the
                    'list.
                
                End If
                 
            End If
        
        Else
            'They didn't choose User or Admin
            MsgBox "Choose either User or Manager for security level.", , _
                    "Error!"
            'Let them know their mistake
        End If
        
    Else
        'They didn't fill in all required text boxes
        MsgBox _
                "You must add data to all text boxes to add a new employee to the database.", , _
                "Error!"
        'Let them know their mistake
    End If

    '<EhFooter>
    Exit Sub

cmdUsersAdd_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersAdd.cmdUsersAdd_Click " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
'-----------------------------------------------------------------------------
' Procedure   :       Form_Load
' Parameters  :
' Description :       Perform these actions at load time
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>

    cmbSecurity.AddItem ("User")
    cmbSecurity.AddItem ("Manager")
    'Add these values to the dropdown..
    
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersAdd.Form_Load " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       Form_Unload
' Parameters  :       Cancel (Integer)
' Description :       Used to close this form and any needed cleanup
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>
    frmMain.Visible = True
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersAdd.Form_Unload " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

