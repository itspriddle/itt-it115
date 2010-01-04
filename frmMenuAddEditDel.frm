VERSION 5.00
Begin VB.Form frmMenuAddEditDel 
   Caption         =   "View / Add / Delete Item"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   Icon            =   "frmMenuAddEditDel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAlcoholic 
      Caption         =   "Alcoholic"
      Enabled         =   0   'False
      Height          =   855
      Index           =   0
      Left            =   2400
      TabIndex        =   34
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton optAlcoholic 
         Caption         =   "No"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optAlcoholic 
         Caption         =   "Yes"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox txtFoodID 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraAddItem 
      Caption         =   "Add Item"
      Height          =   4815
      Left            =   4680
      TabIndex        =   24
      Top             =   1800
      Width           =   4335
      Begin VB.Frame fraAlcoholic 
         Caption         =   "Alcoholic"
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton optAlcoholic 
            Caption         =   "Yes"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton optAlcoholic 
            Caption         =   "No"
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   6
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Return to main menu"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Save Item"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         ToolTipText     =   "Save new item to menu"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.ComboBox cmbAddItemCategory 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select new items category"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtAddItemDescription 
         Height          =   1620
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Enter item description here"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtAddItemPrice 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Enter item price (example: 12.99)"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtAddItemName 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Add item name (example: Jo Mommas House Salad)"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   4320
         Y1              =   3945
         Y2              =   3945
      End
      Begin VB.Label lblAddCategory 
         Caption         =   "Category:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblItemPrice 
         Caption         =   "Price:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblItemDescription 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblItemName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000016&
         BorderWidth     =   2
         Index           =   1
         X1              =   4320
         X2              =   0
         Y1              =   3960
         Y2              =   3960
      End
   End
   Begin VB.Frame fraItemProperties 
      Caption         =   "View / Edit / Delete Item: "
      Height          =   4815
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   4335
      Begin VB.ListBox lstFoodPrice 
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   3510
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox lstFoodDesc 
         Height          =   450
         Left            =   2280
         TabIndex        =   14
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtEditFoodPrice 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   3480
         Width           =   1815
      End
      Begin VB.ListBox lstFoodID 
         Height          =   645
         Left            =   480
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeleteItem 
         Caption         =   "Delete Item"
         Height          =   375
         Left            =   2280
         TabIndex        =   19
         ToolTipText     =   "Delete item from database"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton cmdEditItem 
         Caption         =   "Save Changes"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Save Changes"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox txtItemDesc 
         Height          =   900
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ListBox lstFoodName 
         Height          =   1620
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cmbCategory 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Select a category to list items"
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lblPrice 
         Caption         =   "Price"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblDirectionsAddEdit 
         Height          =   735
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblItemDesc 
         Caption         =   "Description:"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   4320
         Y1              =   3945
         Y2              =   3945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000016&
         BorderWidth     =   2
         Index           =   0
         X1              =   4320
         X2              =   0
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label lblCategory 
         Caption         =   "Category:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.Label lblDirectionsOverall 
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Top             =   480
      Width           =   8895
   End
   Begin VB.Label lblViewAddDeleteItem 
      Caption         =   "View / Add / Delete Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMenuAddEditDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public con As New ADODB.Connection 'Create a var called con as a DB connection

Public rs As New ADODB.recordset 'Create a var called rs as the returned recordset

Public Sub cmbCategory_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmbCategory_Click
' Parameters  :
' Description :       The user clicks on cmbCategory; used to populate lstboxes, etc
'-----------------------------------------------------------------------------
    'When the user clicks a different category
    '<EhHeader>
    On Error GoTo cmbCategory_Click_Err
    '</EhHeader>

    If cmbCategory.Text = "Beverages" Then

        Me.fraAlcoholic(0).Visible = True
    
        Me.txtItemDesc.Visible = False
    
        Me.lblItemDesc.Visible = False
    
    Else
    
        Me.fraAlcoholic(0).Visible = False
    
        Me.txtItemDesc.Visible = True
    
        Me.lblItemDesc.Visible = True
    
    End If

    Select Case cmbCategory.Text
            'Depending on which one they choose...
    
        Case "Breakfast"
            RetrieveFoodsAdmin (0)
            
        Case "Lunch"
            RetrieveFoodsAdmin (1)

        Case "Dinner"
            RetrieveFoodsAdmin (2)

        Case "Beverages"
            RetrieveFoodsAdmin (3)
            
        Case "Appetizers"
            RetrieveFoodsAdmin (4)

        Case "Desserts"
            RetrieveFoodsAdmin (5)

    End Select
    
    '<EhFooter>
    Exit Sub

cmbCategory_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.cmbCategory_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub lstFoodName_Click()
'-----------------------------------------------------------------------------
' Procedure   :       lstFoodName_Click
' Parameters  :
' Description :       The user clicks on lstFoodName; used to populate other lstboxes and txtboxes
'-----------------------------------------------------------------------------
    'When a user clicks a category on lstFoodName
    '<EhHeader>
    On Error GoTo lstFoodName_Click_Err
    '</EhHeader>

    txtItemDesc.Text = ""
    'Clear the itemDesc textbox
    
    txtItemDesc.Text = Me.lstFoodDesc.List(Me.lstFoodName.ListIndex)
    'Fill the item description textbox
    
    Me.txtEditFoodPrice.Text = Me.lstFoodPrice.List(Me.lstFoodName.ListIndex)
    'And the price textbox
    
    If Me.cmbCategory.Text = "Beverages" Then
        Me.fraAlcoholic(0).Enabled = True
        Me.optAlcoholic(0).Enabled = True
        Me.optAlcoholic(1).Enabled = True
        
        If Me.lstFoodDesc.List(Me.lstFoodName.ListIndex) = "Alcoholic" Then
            Me.optAlcoholic(0).Value = True
        Else
            Me.optAlcoholic(1).Value = True
        End If
        
    End If

    '<EhFooter>
    Exit Sub

lstFoodName_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.lstFoodName_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmbAddItemCategory_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmbAddItemCategory_Click
' Parameters  :
' Description :       The user clicks on cmbAddItemCategory
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmbAddItemCategory_Click_Err
    '</EhHeader>
    
    If cmbAddItemCategory.Text = "Beverages" Then

        Me.fraAlcoholic(1).Visible = True
    
        Me.txtAddItemDescription.Visible = False
    
        Me.lblItemDescription.Visible = False
    
    Else
    
        Me.fraAlcoholic(1).Visible = False
    
        Me.txtAddItemDescription.Visible = True
        Me.lblItemDescription.Visible = True
    
    End If

    '<EhFooter>
    Exit Sub

cmbAddItemCategory_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.cmbAddItemCategory_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdAddItem_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdAddItem_Click
' Parameters  :
' Description :       Add a new item to the DB
'-----------------------------------------------------------------------------
    'Add an item to the db
    '<EhHeader>
    On Error GoTo cmdAddItem_Click_Err
    '</EhHeader>

    If Me.cmbAddItemCategory.Text <> "" Then
        'If they selected an item category
    
        If Me.txtAddItemName <> "" And Me.txtAddItemPrice <> "" Then
            'And filled in all fields
        
            If Me.cmbAddItemCategory.Text <> "Beverages" And _
                    Me.txtAddItemDescription.Text = "" Then
                'If its not beverages and they didnt enter a desc
                    
                MsgBox "You must enter a description for this item.", , "Error"
                'Give them an error message
                
                Exit Sub 'Exit this procedure
            
            End If
            
            If IsCurrency(Me.txtAddItemPrice.Text) = False Then Exit Sub
            
            Dim strSQL As String
            
            strSQL = strSQL + "INSERT INTO T_" & UCase$(cmbAddItemCategory.Text)
            strSQL = strSQL + " (F_ITEM, "
            
            If Me.cmbAddItemCategory.Text = "Beverages" Then
                strSQL = strSQL + "F_ALCOHOLIC, "
            Else
                strSQL = strSQL + "F_DESC, "
            End If
            
            strSQL = strSQL + "F_PRICE)"
            strSQL = strSQL + " VALUES ("
            strSQL = strSQL + "'" & Me.txtAddItemName.Text & "',"
            
            If Me.cmbAddItemCategory.Text = "Beverages" Then
                'If its beverages
            
                If Me.optAlcoholic(2) = True Then
                    'If its alcoholic
                
                    strSQL = strSQL + "TRUE,"
                    
                ElseIf Me.optAlcoholic(3) = True Then
                    'Its not alcoholic
                
                    strSQL = strSQL + "FALSE,"
                    
                End If
                
            Else
                'Its not beverages
            
                strSQL = strSQL + "'" & Me.txtAddItemDescription.Text & "',"
                
            End If
            
            strSQL = strSQL + "'" & Me.txtAddItemPrice.Text & "'"
            strSQL = strSQL + ")"
            
            con.Execute (strSQL)
            
            Dim mbrAdded As VbMsgBoxResult
            
            mbrAdded = MsgBox(Me.txtAddItemName.Text & _
                    " was added to the Menu at $" & Me.txtAddItemPrice.Text & _
                    " a serving." & vbNewLine & vbNewLine & _
                    "Item Description:" & vbNewLine & _
                    Me.txtAddItemDescription.Text & vbNewLine & vbNewLine & _
                    "Would you like to add another?", vbYesNo, "Item Added")
            'Ask if they'd like to add another user
                
            If mbrAdded = vbYes Then
                'If they do
            
                Me.txtAddItemName.Text = ""
                Me.txtAddItemDescription.Text = ""
                Me.txtAddItemPrice.Text = ""
                'Clear the textboxes
                
            Else
                'They're done
                
                Unload Me
                'Unload this form
                
            End If
        
        Else
            'All the data wasn't there
        
            MsgBox _
                    "You must enter data into all fields to add a new item to the menu.", , _
                    "Can't Save"
            'let them know
                    
        End If

    Else
        MsgBox "You must choose a category.", , "Can't Save"
    End If
    
    '<EhFooter>
    Exit Sub

cmdAddItem_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.cmdAddItem_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdCancel_Click
' Parameters  :
' Description :       Leave this form
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdCancel_Click_Err
    '</EhHeader>

    Unload Me
    
    '<EhFooter>
    Exit Sub

cmdCancel_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.cmdCancel_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdDeleteItem_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdDeleteItem_Click
' Parameters  :
' Description :       Delete an item from the DB
'-----------------------------------------------------------------------------
    'Delete an item from the DB
    '<EhHeader>
    On Error GoTo cmdDeleteItem_Click_Err
    '</EhHeader>

    If Me.lstFoodName.Text <> "" Then
        'If they selected an item to delete
    
        Dim strSQL As String
        'Create a new string
        
        strSQL = strSQL + "DELETE FROM T_" & UCase$(cmbCategory.Text)
        strSQL = strSQL + " WHERE F_ID=" & Me.lstFoodID.List( _
                Me.lstFoodName.ListIndex) & ";"
        'For some delete SQL
        
        Dim mbrDelete As VbMsgBoxResult
        'Make a new var as a message box result
        
        mbrDelete = MsgBox("Are you sure you want to delete " & _
                Me.lstFoodName.Text & "?", vbYesNo, "Verify Delete")
            
        If mbrDelete = vbYes Then
        
            con.Execute (strSQL)
            Me.lstFoodDesc.RemoveItem (Me.lstFoodName.ListIndex)
            Me.lstFoodID.RemoveItem (Me.lstFoodName.ListIndex)
            Me.lstFoodName.RemoveItem (Me.lstFoodName.ListIndex)
            Me.txtEditFoodPrice = ""
            Me.txtFoodID = ""
            Me.txtItemDesc = ""
            
        End If
        
    Else
        MsgBox "You must select an item to delete.", , "Select an Item"
    End If

    '<EhFooter>
    Exit Sub

cmdDeleteItem_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.cmdDeleteItem_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdEditItem_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdEditItem_Click
' Parameters  :
' Description :       Update DB with edited info
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdEditItem_Click_Err
    '</EhHeader>

    If Me.lstFoodName.Text <> "" Then
        'If they selected an item
    
        If IsCurrency(Me.txtEditFoodPrice.Text) = False Then Exit Sub
    
        Dim strSQL As String
        'Create a new sting
    
        strSQL = strSQL + "UPDATE T_" & UCase$(cmbCategory.Text)
        strSQL = strSQL + " SET F_ITEM='" & Me.lstFoodName.Text & "', "
    
        If cmbCategory.Text = "Beverages" Then
            'If its beverages....
    
            strSQL = strSQL + " F_ALCOHOLIC="
        
            If Me.lstFoodDesc.List(Me.lstFoodName.ListIndex) = "Alcoholic" Then
                'If its alcoholic
        
                strSQL = strSQL + "Yes"
            
            Else
                'Its not alcoholic
        
                strSQL = strSQL + "No"
            
            End If
        
            strSQL = strSQL + ", "
        
        Else
            'Not Beverages
    
            strSQL = strSQL + " F_DESC='" & Me.txtItemDesc.Text & "', "
    
        End If
    
        strSQL = strSQL + " F_PRICE='" & Me.txtEditFoodPrice.Text & "'"
        strSQL = strSQL + " WHERE F_ID="
        strSQL = strSQL + Me.lstFoodID.List(Me.lstFoodName.ListIndex) & ";"
        'For some SQL to Update the selected item
    
        'MsgBox strSQL
    
        con.Execute (strSQL)
        'Execute the SQL
    
        'MsgBox strSQL
    
        MsgBox "The item was successfully updated.", , "Item Updated"
        'Let the user know everything went ok
    
        'Something here to refresh the listboxes would be good...
    
    Else
        'They havent made a selection

        MsgBox "You must select an item to edit.", , "Select an Item"
        'So let the user know
    
    End If

    '<EhFooter>
    Exit Sub

cmdEditItem_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.cmdEditItem_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
'-----------------------------------------------------------------------------
' Procedure   :       Form_Load
' Parameters  :
' Description :       Events to perform at load time
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Form_Load_Err
    '</EhHeader>
    Me.lblDirectionsOverall.Caption = _
            "To View items just click on a Category under View / Edit / Delete Item. To Edit or Delete that item, highlight it and click the Edit or Delete button." _
            & vbNewLine & vbNewLine & _
            "To add a new item, enter the item's details under Add Item, and the Save Item button."
    'Set this caption..
    
    Me.lblDirectionsAddEdit.Caption = _
            "Choose a category below to view any items associated with it.  Double click an item name to rename it."
    'And set this one too...
    
    con.CursorLocation = adUseClient
    'Set the cursor location
    
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & _
            App.Path & "\restaurant.mdb;" & _
            "Jet OLEDB:Database Password=password"
    'Set the connection and password
    
    con.Open
    'Open the connection
    
    cmbCategory.AddItem ("Breakfast")
    cmbCategory.AddItem ("Lunch")
    cmbCategory.AddItem ("Dinner")
    cmbCategory.AddItem ("Beverages")
    cmbCategory.AddItem ("Appetizers")
    cmbCategory.AddItem ("Desserts")
    'Populate the dropdown
    
    cmbAddItemCategory.AddItem ("Breakfast")
    cmbAddItemCategory.AddItem ("Lunch")
    cmbAddItemCategory.AddItem ("Dinner")
    cmbAddItemCategory.AddItem ("Beverages")
    cmbAddItemCategory.AddItem ("Appetizers")
    cmbAddItemCategory.AddItem ("Desserts")
    'Populate the dropdown

    '<EhFooter>
    Exit Sub

Form_Load_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.Form_Load " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       Form_Unload
' Parameters  :       Cancel (Integer)
' Description :       Used to unload this form and do any needed cleanup
'-----------------------------------------------------------------------------
    'This form is unloaded
    '<EhHeader>
    On Error GoTo Form_Unload_Err
    '</EhHeader>

    frmMain.Visible = True
    'Show the main window
    
    con.Close
    'Close the connection
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.Form_Unload " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub lstFoodName_DblClick()
'-----------------------------------------------------------------------------
' Procedure   :       lstFoodName_DblClick
' Parameters  :
' Description :       Used to change the Food Name
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo lstFoodName_DblClick_Err
    '</EhHeader>
    
    Dim strNewName As String
    
    strNewName = InputBox("Insert the new name:", "New Name")
    
    If strNewName <> "" Then
        Dim intListIndex As Integer
        
        intListIndex = Me.lstFoodName.ListIndex
        
        Me.lstFoodName.List(intListIndex) = strNewName
    
    End If

    '<EhFooter>
    Exit Sub

lstFoodName_DblClick_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.lstFoodName_DblClick " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub txtItemDesc_GotFocus()
'-----------------------------------------------------------------------------
' Procedure   :       txtItemDesc_GotFocus
' Parameters  :
' Description :       txtItemDesc is focused on
'-----------------------------------------------------------------------------
    'User focuses on txtItemDesc
    '<EhHeader>
    On Error GoTo txtItemDesc_GotFocus_Err
    '</EhHeader>

    If Me.cmbCategory.Text = "Beverages" Then Me.txtItemDesc.ToolTipText = _
            "Enter Alcoholic or Not Alcoholic"
    'If they're using beverages create a tooltip that says...
    
    '<EhFooter>
    Exit Sub

txtItemDesc_GotFocus_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMenuAddEditDel.txtItemDesc_GotFocus " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

