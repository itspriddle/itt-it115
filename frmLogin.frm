VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1470
   ClientLeft      =   2835
   ClientTop       =   3585
   ClientWidth     =   7095
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   868.524
   ScaleMode       =   0  'User
   ScaleWidth      =   6661.821
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Clock_Timer 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   4530
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3735
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   5340
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4530
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   3345
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   3345
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "frmLogin.frx":1272
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public login As New ADODB.Connection

Public Sub Clock_Timer_Timer()
'-----------------------------------------------------------------------------
' Procedure   :       Clock_Timer_Timer
' Parameters  :
' Description :       Used for the clock on the login prompt
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    lblTime.Caption = TimeValue(Time)
End Sub

Public Sub cmdOK_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdOK_Click
' Parameters  :
' Description :       Login, show admin frame if user is admin, insert
'                     tablerow into T_INOUT with name, date, and time in
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdOK_Click_Err
    '</EhHeader>
    Dim rsUsers As New ADODB.recordset
    rsUsers.Open "SELECT * FROM T_AUTHUSERS WHERE F_USERNAME='" & txtUsername _
            & "'", login, adOpenDynamic, adLockOptimistic
    
    If rsUsers.RecordCount > 0 Then
        'If there was a user with that username
    
        If txtPassword = rsUsers!F_PASSWORD Then
            'If they input the correct password
        
            Dim strSQL As String
            'New string for some SQL
            
            strSQL = strSQL + "INSERT INTO T_INOUT "
            strSQL = strSQL + "(F_FNAME, F_LNAME, F_DATE, F_PUNCHEDIN) "
            strSQL = strSQL + "VALUES ("
            strSQL = strSQL + "'" & rsUsers!F_FNAME & "',"
            strSQL = strSQL + "'" & rsUsers!F_LNAME & "',"
            strSQL = strSQL + "'" & Date & "',"
            strSQL = strSQL + "'" & TimeValue(Time) & "'"
            
            strSQL = strSQL + ")"
            'Set the SQL for T_INOUT
            
            login.Execute (strSQL)
            'Execute the SQL
        
            Dim rsID As New ADODB.recordset
            
            rsID.Open "SELECT * FROM T_INOUT ORDER BY F_ID DESC", login, _
                    adOpenDynamic, adLockOptimistic
            
            intLoginID = rsID!F_ID
            LoginDate = Date
            strUsername = rsUsers!F_USERNAME
            strFirstName = rsUsers!F_FNAME
            strLastName = rsUsers!F_LNAME
            
            rsID.Close
            
            frmMain.Visible = True
            'Show the main form
            
            If rsUsers!F_SECURITYLVL = "Manager" Then
                frmMain.mnuAdmin.Visible = True
                'Show the Admin menu
                
                frmMain.fraAdmin.Visible = True
                'Show the admin frame
                                
            ElseIf rsUsers!F_SECURITYLVL = "User" Then
                frmMain.mnuAdmin.Visible = False
                'Hide the Admin Menu
                
                frmMain.fraAdmin.Visible = False
                'Hide the Admin frame
                
            End If

            Me.Hide
            'Hide the login form, they're done with it
            
        Else
            'The password was wrong
            MsgBox "Invalid Password, try again!", , "Login"
            'Let them know...
            
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
            
        End If

    Else
        'The username wasn't found
        MsgBox "There was no user found with that Username, try again!", , _
                "Login"
        'Let them know
        txtUsername.SetFocus
        
    End If

    rsUsers.Close

    '<EhFooter>
    Exit Sub

cmdOK_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmLogin.cmdOK_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdCancel_Click
' Parameters  :
' Description :       Don't login
'-----------------------------------------------------------------------------
    'Close the program
    '<EhHeader>
    On Error GoTo cmdCancel_Click_Err
    '</EhHeader>
    End
    '<EhFooter>
    Exit Sub

cmdCancel_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmLogin.cmdCancel_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
'-----------------------------------------------------------------------------
' Procedure   :       Form_Load
' Parameters  :
' Description :       Perform at load time
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>

    login.CursorLocation = adUseClient
    'Set the location
    
    login.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
            & App.Path & "\restaurant.mdb;" & _
            "Jet OLEDB:Database Password=password"
    'Set the connection and the path to DB
    
    login.Open
    'Establish the connection
    
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    MsgBox Err.Description & vbCrLf & "in RestaurantMenu.frmLogin.Form_Load " _
            & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtPassword_GotFocus()
'-----------------------------------------------------------------------------
' Procedure   :       txtPassword_GotFocus
' Parameters  :
' Description :       The users focuses on txtPassword.  Used for a tooltip
'                     warning if Caps Lock is ON
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo txtPassword_GotFocus_Err
    '</EhHeader>
    
    Call getCapsStatus
    'Determine if the capslock key is on

    If intKeyState = 1 Then txtPassword.ToolTipText = _
            "The Caps Lock key is on. Passwords are case sensitive!"
    'If it is then let the user know in a tool tip
    
    '<EhFooter>
    Exit Sub

txtPassword_GotFocus_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmLogin.txtPassword_GotFocus " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

