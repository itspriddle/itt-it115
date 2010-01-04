VERSION 5.00
Begin VB.Form frmUsersView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Users"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4200
   Icon            =   "frmUsersView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddUser 
      Caption         =   "Add User"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdDeleteUser 
      Caption         =   "Delete User"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   3960
      Width           =   975
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
   Begin VB.CommandButton cmdEditEntry 
      Caption         =   "Edit Entry"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevUser 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtFirstName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtLastName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox cmbSecurity 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtUsername 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblFirstName 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblLastName 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblMSG 
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblSecurity 
      Caption         =   "Security Level:"
      Height          =   495
      Left            =   720
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username:"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   1920
      Width           =   975
   End
   Begin VB.Menu mnuSort 
      Caption         =   "Sort"
      Begin VB.Menu mnuSortFirstNameABC 
         Caption         =   "By First Name (ABC)"
      End
      Begin VB.Menu mnuSortFirstNameZYX 
         Caption         =   "By First Name (ZYX)"
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortLNameABC 
         Caption         =   "By Last Name (ABC)"
      End
      Begin VB.Menu mnuSortLNameZYX 
         Caption         =   "By Last Name (ZYX)"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortUsernameABC 
         Caption         =   "By Username (ABC)"
      End
      Begin VB.Menu mnuSortUsernameZYX 
         Caption         =   "By Username (ZYX)"
      End
   End
End
Attribute VB_Name = "frmUsersView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public con As New ADODB.Connection 'Create a var called con as a DB connection

Public rs As New ADODB.recordset 'Create a var called rs as the returned recordset

Public Function FillTextboxesUsers()
'-----------------------------------------------------------------------------
' Procedure   :       FillTextboxesUsers
' Parameters  :
' Description :       Fill the user textboxes
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo FillTextboxesUsers_Err
    '</EhHeader>

    If Not rs.EOF And Not rs.BOF Then
        'If the current position in the recordset isnt the beginning
        'and isnt the end
        txtFirstName = rs!F_FNAME & ""
        txtLastName = rs!F_LNAME & ""
        txtUsername = rs!F_USERNAME & ""
        txtPassword = rs!F_PASSWORD & ""
        cmbSecurity.Text = rs!F_SECURITYLVL & ""
        txtID = rs!F_ID & ""
        'Set the above txtBoxes to the values from the DB
        
    End If

    '<EhFooter>
    Exit Function

FillTextboxesUsers_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.FillTextboxesUsers " & "at line " _
            & Erl
    Resume Next
    '</EhFooter>
End Function

Private Sub cmdAddUser_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdAddUser_Click
' Parameters  :
' Description :       Add a new user
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdAddUser_Click_Err
    '</EhHeader>
    Me.Hide
    'Hide this form for now, we'll be back
    
    frmUsersAdd.Visible = True
    'Show the add users Form
    '<EhFooter>
    Exit Sub

cmdAddUser_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.cmdAddUser_Click " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdBack_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdBack_Click
' Parameters  :
' Description :       Back to main
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdBack_Click_Err
    '</EhHeader>
    
    Unload Me
    'Unload this form
    
    'frmMain.Visible = True
    'Show the Manager's form

    '<EhFooter>
    Exit Sub

cmdBack_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.cmdBack_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdDeleteUser_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdDeleteUser_Click
' Parameters  :
' Description :       Delete the selected user
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdDeleteUser_Click_Err
    '</EhHeader>
    Dim mbrDelete As VbMsgBoxResult
    mbrDelete = MsgBox("Are you sure you want to delete this user?", vbYesNo, _
            "Confirm Delete")

    If mbrDelete = vbYes Then
        If Me.txtUsername.Text <> frmLogin.txtUsername.Text Then
            'Delete the user
            
            Dim rs2 As New ADODB.recordset
        
            rs2.Open "DELETE FROM T_AUTHUSERS WHERE F_USERNAME='" & _
                    txtUsername.Text & "'", con, adOpenDynamic, _
                    adLockOptimistic
        
            'MsgBox "User deleted." & rs2!F_USERNAME
            rs.Requery
        
            'rs.MoveNext
            FillTextboxesUsers
        
        Else
            MsgBox "You cannot delete yourself!", vbOKOnly, "Can't Delete"
        End If

    Else
        'Cancel

    End If

    '<EhFooter>
    Exit Sub

cmdDeleteUser_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.cmdDeleteUser_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdEditEntry_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdEditEntry_Click
' Parameters  :
' Description :       Edit the selected entry
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdEditEntry_Click_Err
    '</EhHeader>
    Load frmUsersEdit

    'Load the Edit Frame
    With frmUsersEdit
        'Set the textboxes.text to the values from this frames boxes
        .txtFirstName = frmUsersView.txtFirstName
        .txtLastName = frmUsersView.txtLastName
        .txtUsername = frmUsersView.txtUsername
        .txtPassword = frmUsersView.txtPassword
        .cmbSecurity.Text = frmUsersView.cmbSecurity.Text
        .txtID = frmUsersView.txtID
        
        .lblMSG.Caption = "The following information has been found for " & _
                frmUsersView.txtFirstName & " " & frmUsersView.txtLastName & _
                "."
        'Set the lblcaption to the above..
    End With
  
    Me.Hide
    'Unload the current frame
    
    frmUsersEdit.Visible = True
    'Show the edit frame
    
    'Close the DB connection
    '<EhFooter>
    Exit Sub

cmdEditEntry_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.cmdEditEntry_Click " & "at line " _
            & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdNext_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdNext_Click
' Parameters  :
' Description :       View the next users data
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdNext_Click_Err
    '</EhHeader>

    If Not rs.EOF Then
        'If we're not at the end of the recordset
    
        rs.MoveNext
        'Move to the next record
        
        If rs.EOF Then
            'If we are at the end of the recordset...
        
            MsgBox "At the end of the list."
            'Let the user know
            
        Else
            'Not at the end yet...
            
            FillTextboxesUsers
            'Fill the text boxes with the data from the DB
            
            lblMSG.Caption = "The following information has been found for " _
                    & rs!F_FNAME & " " & rs!F_LNAME & "."
        End If

    Else
        'We're at the end of the recordset...
    
        MsgBox "At the end of the list."
        'So let the user know
    End If

    '<EhFooter>
    Exit Sub

cmdNext_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.cmdNext_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdPrevUser_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdPrevUser_Click
' Parameters  :
' Description :       View the previous user's data
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdPrevUser_Click_Err
    '</EhHeader>

    If Not rs.BOF Then
        'If we're not at the beginning of the recordset
    
        rs.MovePrevious
        'Move to the previous record
        
        If rs.BOF Then
            'If we are...
        
            MsgBox "At the beginning of the list."
            'Let the user know
            
        Else
            'We're not at the beginning yet...
        
            FillTextboxesUsers
            'Fill the text boxes with the data from the DB
            
            lblMSG.Caption = "The following information has been found for " _
                    & rs!F_FNAME & " " & rs!F_LNAME & "."
            
        End If

    Else
        'We're at the beginning of the recordset
    
        MsgBox "At the beginning of the list."
        'So let the user know
    End If

    '<EhFooter>
    Exit Sub

cmdPrevUser_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.cmdPrevUser_Click " & "at line " _
            & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
'-----------------------------------------------------------------------------
' Procedure   :       Form_Load
' Parameters  :
' Description :       Perform at load time...
'-----------------------------------------------------------------------------
    'Start the DB Connection
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
    
    con.CursorLocation = adUseClient
    'Set the cursor location
    
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & _
            App.Path & "\restaurant.mdb;" & _
            "Jet OLEDB:Database Password=password"
    'Set the connection and password
    
    con.Open
    'Open the connection
    
    rs.Open "SELECT * FROM T_AUTHUSERS", con, adOpenDynamic, adLockOptimistic
    'Select all records from AUTHUSERS
    
    cmbSecurity.AddItem ("User")
    'Add the item User to the security lvl dropdown
    
    cmbSecurity.AddItem ("Manager")
    'Add the Item Manager to the security lvl dropdown
    
    FillTextboxesUsers
    'Fill the text boxes with the data from the DB
    
    lblMSG.Caption = "The following information has been found for: " & _
            vbNewLine & vbNewLine & rs!F_FNAME & " " & rs!F_LNAME
    'Change the caption of lblMSG to say:
    '"The following information has been found for <FIRST NAME> <LAST NAME>.
    
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.Form_Load " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       Form_Unload
' Parameters  :       Cancel (Integer)
' Description :       Unload this form, perform any needed cleanup
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>
    frmMain.Visible = True
    con.Close
    'Close the DB connection
    
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.Form_Unload " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSortFirstNameABC_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuSortFirstNameABC_Click
' Parameters  :
' Description :       Sort by First Name Ascending
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuSortFirstNameABC_Click_Err
    '</EhHeader>
    rs.Close
    'Close the current recordset
    
    rs.Open "SELECT * FROM T_AUTHUSERS ORDER BY F_FNAME", con, adOpenDynamic, _
            adLockOptimistic
    'Start a new one ordered by First Names
    
    FillTextboxesUsers
    'Fill the text boxes
    
    Me.lblFirstName.FontBold = True
    'Make First Name: bold
    
    Me.lblLastName.FontBold = False
    'But Make sure the rest aren't
    
    Me.lblUserName.FontBold = False
    'But make sure the rest aren't
    
    Me.txtFirstName.Top = 1200   'Make these appear on top
    Me.lblFirstName.Top = 1200
    Me.txtFirstName.TabIndex = 1
    
    Me.txtLastName.Top = 1560   'And these 2nd from top
    Me.lblLastName.Top = 1560
    Me.txtLastName.TabIndex = 2
    
    Me.lblUserName.Top = 1920  'And these 3rd from top
    Me.txtUsername.Top = 1920
    Me.txtUsername.TabIndex = 3
    
    Me.lblPassword.Top = 2280   'And these 4th from top
    Me.txtPassword.Top = 2280
    Me.txtPassword.TabIndex = 4
    
    '<EhFooter>
    Exit Sub

mnuSortFirstNameABC_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.mnuSortFirstNameABC_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSortFirstNameZYX_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuSortFirstNameZYX_Click
' Parameters  :
' Description :       Sort by First Name Descending
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuSortFirstNameZYX_Click_Err
    '</EhHeader>
    rs.Close
    'Close the current rs
    
    rs.Open "SELECT * FROM T_AUTHUSERS ORDER BY F_FNAME DESC", con, _
            adOpenDynamic, adLockOptimistic
    'Start a new one ordered by First Names Z-A
    
    FillTextboxesUsers
    'Fill the textboxes
    
    Me.lblFirstName.FontBold = True
    'Make the First name labels bold
    
    Me.lblLastName.FontBold = False
    'But make sure the rest aren't
    
    Me.lblUserName.FontBold = False
    'But make sure the rest aren't
    
    Me.txtFirstName.Top = 1200   'Make these appear on top
    Me.lblFirstName.Top = 1200
    Me.txtFirstName.TabIndex = 1
    
    Me.txtLastName.Top = 1560   'And these 2nd from top
    Me.lblLastName.Top = 1560
    Me.txtFirstName.TabIndex = 2
    
    Me.lblUserName.Top = 1920  'And these 3rd from top
    Me.txtUsername.Top = 1920
    Me.txtFirstName.TabIndex = 3
    
    Me.lblPassword.Top = 2280   'And these 4th from top
    Me.txtPassword.Top = 2280
    Me.txtFirstName.TabIndex = 4
    
    '<EhFooter>
    Exit Sub

mnuSortFirstNameZYX_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.mnuSortFirstNameZYX_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSortLNameABC_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuSortLNameABC_Click
' Parameters  :
' Description :       Sort by Last Name Ascending
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuSortLNameABC_Click_Err
    '</EhHeader>
    rs.Close
    'Close the current rs
    
    rs.Open "SELECT * FROM T_AUTHUSERS ORDER BY F_LNAME", con, adOpenDynamic, _
            adLockOptimistic
    'Open a new one ordered by Last Name A-Z
    
    FillTextboxesUsers
    'Fill the text boxes
    
    Me.lblLastName.FontBold = True
    'Make the Last Name Label bold
    
    Me.lblFirstName.FontBold = False
    'But make sure the rest aren't
    
    Me.lblUserName.FontBold = False
    'But make sure the rest aren't
    
    Me.txtLastName.Top = 1200   'Make these appear on top
    Me.lblLastName.Top = 1200
    Me.txtFirstName.TabIndex = 1
    
    Me.txtFirstName.Top = 1560   'And these 2nd from top
    Me.lblFirstName.Top = 1560
    Me.txtFirstName.TabIndex = 2
    
    Me.lblUserName.Top = 1920  'And these 3rd from top
    Me.txtUsername.Top = 1920
    Me.txtFirstName.TabIndex = 3
    
    Me.lblPassword.Top = 2280   'And these 4th from top
    Me.txtPassword.Top = 2280
    Me.txtFirstName.TabIndex = 4
    
    '<EhFooter>
    Exit Sub

mnuSortLNameABC_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.mnuSortLNameABC_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSortLNameZYX_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuSortLNameZYX_Click
' Parameters  :
' Description :       Sort by Last Name Descending
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuSortLNameZYX_Click_Err
    '</EhHeader>
    rs.Close
    'Close the current rs
    
    rs.Open "SELECT * FROM T_AUTHUSERS ORDER BY F_LNAME DESC", con, _
            adOpenDynamic, adLockOptimistic
    'Open a new one ordered by Last Name Z-A
    
    FillTextboxesUsers
    'Fill the textboxes
    
    Me.lblLastName.FontBold = True
    'Make the Last Name label bold
    
    Me.lblFirstName.FontBold = False
    'But make sure the rest aren't
    
    Me.lblUserName.FontBold = False
    'But make sure the rest aren't
    
    Me.txtLastName.Top = 1200   'Make these appear on top
    Me.lblLastName.Top = 1200
    Me.txtFirstName.TabIndex = 1
    
    Me.txtFirstName.Top = 1560   'And these 2nd from top
    Me.lblFirstName.Top = 1560
    Me.txtFirstName.TabIndex = 2
    
    Me.lblUserName.Top = 1920  'And these 3rd from top
    Me.txtUsername.Top = 1920
    Me.txtFirstName.TabIndex = 3
    
    Me.lblPassword.Top = 2280   'And these 4th from top
    Me.txtPassword.Top = 2280
    Me.txtFirstName.TabIndex = 4
    '<EhFooter>
    Exit Sub

mnuSortLNameZYX_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.mnuSortLNameZYX_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSortUsernameABC_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuSortUsernameABC_Click
' Parameters  :
' Description :       Sort by Username Ascending
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuSortUsernameABC_Click_Err
    '</EhHeader>

    rs.Close
    'Close the current rs
    
    rs.Open "SELECT * FROM T_AUTHUSERS ORDER BY F_USERNAME", con, _
            adOpenDynamic, adLockOptimistic
    'Open a new one ordered by Username A-Z
    
    FillTextboxesUsers
    'Fill the textboxes
    
    Me.lblUserName.FontBold = True
    'Make the Username label bold
    
    Me.lblLastName.FontBold = False
    'But make sure the rest arent
    
    Me.lblFirstName.FontBold = False
    'But make sure the rest arent

    Me.txtUsername.Top = 1200   'Make these appear on top
    Me.lblUserName.Top = 1200
    Me.txtFirstName.TabIndex = 1
    
    Me.txtPassword.Top = 1560   'And these 2nd from top
    Me.lblPassword.Top = 1560
    Me.txtFirstName.TabIndex = 2
    
    Me.lblFirstName.Top = 1920  'And these 3rd from top
    Me.txtFirstName.Top = 1920
    Me.txtFirstName.TabIndex = 3
    
    Me.lblLastName.Top = 2280   'And these 4th from top
    Me.txtLastName.Top = 2280
    Me.txtFirstName.TabIndex = 4
    '<EhFooter>
    Exit Sub

mnuSortUsernameABC_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.mnuSortUsernameABC_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuSortUsernameZYX_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuSortUsernameZYX_Click
' Parameters  :
' Description :       Sort by Username Descending
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuSortUsernameZYX_Click_Err
    '</EhHeader>
    rs.Close
    'Close the current rs
    
    rs.Open "SELECT * FROM T_AUTHUSERS ORDER BY F_USERNAME DESC", con, _
            adOpenDynamic, adLockOptimistic
    'Open a new one ordered by Username Z-A
    
    FillTextboxesUsers
    'Fill the textboxes

    Me.lblUserName.FontBold = True
    'Make the Username label bold
    
    Me.lblLastName.FontBold = False
    'But make sure the rest arent
    
    Me.lblFirstName.FontBold = False
    'But make sure the rest arent
    
    Me.txtUsername.Top = 1200   'Make these appear on top
    Me.lblUserName.Top = 1200
    Me.txtFirstName.TabIndex = 1
    
    Me.txtPassword.Top = 1560   'And these 2nd from top
    Me.lblPassword.Top = 1560
    Me.txtFirstName.TabIndex = 2
    
    Me.lblFirstName.Top = 1920  'And these 3rd from top
    Me.txtFirstName.Top = 1920
    Me.txtFirstName.TabIndex = 3
    
    Me.lblLastName.Top = 2280   'And these 4th from top
    Me.txtLastName.Top = 2280
    Me.txtFirstName.TabIndex = 4
    
    '<EhFooter>
    Exit Sub

mnuSortUsernameZYX_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmUsersView.mnuSortUsernameZYX_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

