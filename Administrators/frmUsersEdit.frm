VERSION 5.00
Begin VB.Form frmUsersEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit User Info"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
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
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   12
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
      ItemData        =   "frmUsersEdit.frx":0000
      Left            =   1920
      List            =   "frmUsersEdit.frx":0002
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
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3855
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
      Caption         =   "Username:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmUsersEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public con As New ADODB.Connection
'Create a var called con as a DB connection
Public rs As New ADODB.Recordset
'Create a var called rs as the returned recordset

Private Sub cmdBack_Click()
    Unload Me
    'Unload this frame
    
    frmUsersView.Visible = True
    'Load the last frame
    
    con.Close
    'Close the DB Connection
    
End Sub

Private Sub cmdSaveChanges_Click()
    Dim strSQL As String
    'Create a string for the SQL statement
    If txtFirstName <> "" And txtLastName <> "" And txtUsername <> "" And txtPassword <> "" And cmbSecurity.Text <> "" Then
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
        MsgBox "Info changed."
        Else
            'Let them know their mistake
            MsgBox "Choose either User or Admin for security level."
        End If
        
    Else
        'Let them know their mistake
        MsgBox "You must add data to all text boxes to add a new employee to the database."
    End If
End Sub

Private Sub Form_Load()
    
    
    con.CursorLocation = adUseClient
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\restaurant.mdb"
    con.Open
    rs.Open "SELECT * FROM T_AUTHUSERS", con, adOpenDynamic, adLockOptimistic
    
    cmbSecurity.AddItem ("User")
    cmbSecurity.AddItem ("Manager")
    
    lblMSG.Caption = "The following information has been found for " & rs!F_FNAME & " " & rs!F_LNAME & "."
    
End Sub
