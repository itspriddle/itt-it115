Attribute VB_Name = "modRest"
Option Explicit
'-----------------------------------------------------------------------------
' Module      :       modRest.bas
' Description :       Used for public functions, vars, etc for
'                     RestaurantMenu.vbp
'-----------------------------------------------------------------------------



Public Declare Function GetKeyState _
               Lib "user32" (ByVal nVirtKey As Long) As Integer

Public intKeyState As Integer

'----------------------------------------------------------------------------
' These Public vars are used to set up the database connection
' when we need to.
'----------------------------------------------------------------------------
Public con As New ADODB.Connection

Public rs As New ADODB.recordset
    
'----------------------------------------------------------------------------
' These are used to keep track of who is logged in and what their
' username is.  Later on, we'll use this to make sure some dummy
' doesn't delete themselves from the DB
'----------------------------------------------------------------------------

Public strUsername As String, strFirstName As String, strLastName As String, _
        LoginDate As Date, intLoginID As Integer

Public Function AddToOrder(intIndex As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       AddToOrder
' Parameters  :       intIndex (Integer)
' Description :       Add to an existing order
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo AddToOrder_Err
    '</EhHeader>

    If frmMain.lstFood(intIndex).Text <> "" Then

        'Call IsInt
        If IsInt(frmMain.txtQuantity(intIndex)) = False Then Exit Function
        
        Dim curTotal As Currency
    
        If frmMain.txtMainSubTotal.Text <> "" Then
            curTotal = frmMain.txtMainSubTotal
        Else
            curTotal = 0
        End If
    
        Dim intColPos As Integer, curPrice As Currency, strFoodName As String
        'Create a new integer and a new currency
    
        intColPos = InStr(frmMain.lstFood(intIndex).Text, ":")
        'Set this to the position of ':' in the string
    
        curPrice = Mid$(frmMain.lstFood(intIndex).Text, intColPos + 2, Len( _
                frmMain.lstFood(intIndex).Text) - (intColPos + 1))
        'Set the current item's price
    
        strFoodName = Left$(frmMain.lstFood(intIndex).Text, intColPos - 1)
    
        curTotal = curTotal + (curPrice * frmMain.txtQuantity(intIndex).Text)
        'Add the current item's price to the grand total
    
        frmMain.lstMainOrderSummary.AddItem (frmMain.txtQuantity( _
                intIndex).Text & "; " & frmMain.lstFood(intIndex).Text & _
                " ea.")
        'Add QTY; Food Name to the listbox
    
        frmMain.txtMainSubTotal.Text = curTotal
        'Change the textbox to show the subtotal
    
        frmMain.lstFood(intIndex).Clear
        'Clear the listbox from this frame..
    
        If frmMain.txtCookingInstructions(intIndex).Text <> "" Then
            frmMain.lstMainOrderDirections.AddItem (strFoodName & ": " & _
                    frmMain.txtCookingInstructions(intIndex).Text)
            'Add the cooking instructions to a hidden listbox for access later
        
        Else
            frmMain.lstMainOrderDirections.AddItem (strFoodName & ": Normal")
            'Add the cooking instructions to a hidden listbox for access later
        
        End If
    
        frmMain.txtTotal.Text = frmMain.txtMainSubTotal.Text * 1.0725
        'Set the total value and add some sales tax

        frmMain.txtTotal.Text = Round(frmMain.txtTotal.Text, 2)
        frmMain.txtMainSubTotal.Text = Round(frmMain.txtMainSubTotal.Text, 2)
    
        Select Case intIndex
                'Hide the appropriate frame...
    
            Case 0
                frmMain.fraBreakfast.Visible = False

                'Hide this frame
            Case 1
                frmMain.fraLunch.Visible = False

                'Hide this frame
            Case 2
                frmMain.fraDinner.Visible = False

                'Hide this frame
            Case 3
                frmMain.fraBeverages.Visible = False

                'Hide this frame
            Case 4
                frmMain.fraAppetizers.Visible = False

                'Hide this frame
            Case 5
                frmMain.fraDesserts.Visible = False
                'Hide this frame

        End Select
    
        frmMain.fraMain.Visible = True
        'Show the main frame
    Else
        MsgBox "You must select an item to add to the order.", , _
                "No Item Selected"
    End If

    '<EhFooter>
    Exit Function

AddToOrder_Err:
    MsgBox Err.Description & vbCrLf & "in RestaurantMenu.modRest.AddToOrder " _
            & "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Function ClearOrder()
'-----------------------------------------------------------------------------
' Procedure   :       ClearOrder
' Parameters  :
' Description :       Clear the current order
'-----------------------------------------------------------------------------

    '<EhHeader>
    On Error GoTo ClearOrder_Err
    '</EhHeader>

    frmMain.lstMainOrderSummary.Clear
    'Clear the order summary
    
    frmMain.lstMainOrderDirections.Clear
    'Clear the directions
    
    frmMain.txtTotal.Text = ""
    'Clear the total
    
    frmMain.txtMainSubTotal.Text = ""
    'Clear the subtotal
    
    frmMain.txtMainOrderInstructions.Text = ""
    'Clear any instructions that might be showing

    '<EhFooter>
    Exit Function

ClearOrder_Err:
    MsgBox Err.Description & vbCrLf & "in RestaurantMenu.modRest.ClearOrder " _
            & "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Function getCapsStatus()
'-----------------------------------------------------------------------------
' Procedure   :       getCapsStatus
' Parameters  :
' Description :       Determine if the Caps Lock key is ON or OFF
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo getCapsStatus_Err
    '</EhHeader>
    
    intKeyState = GetKeyState(vbKeyCapital)
    
    '<EhFooter>
    Exit Function

getCapsStatus_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.modRest.getCapsStatus " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Function IsCurrency(strNumber As String) As Boolean
'-----------------------------------------------------------------------------
' Procedure   :       IsCurrency
' Parameters  :       strNumber (String)
' Description :       Check if a string is currency (doesn't contain letters
'                     or ,)
'
'                     Returns TRUE or FALSE
'-----------------------------------------------------------------------------

    '<EhHeader>
    On Error GoTo IsCurrency_Err
    '</EhHeader>

    If IsNumeric(strNumber) Then
        
        IsCurrency = True
        
    Else
        MsgBox "You must insert a numeric value for Price."
        
        IsCurrency = False
        
    End If

    '<EhFooter>
    Exit Function

IsCurrency_Err:
    MsgBox Err.Description & vbCrLf & "in RestaurantMenu.modRest.IsCurrency " _
            & "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Function IsInt(strNumber As String) As Boolean
'-----------------------------------------------------------------------------
' Procedure   :       IsInt
' Parameters  :       strNumber (String)
' Description :       Check if a string is an integer
'
'                     Returns TRUE or FALSE
'-----------------------------------------------------------------------------

    '<EhHeader>
    On Error GoTo IsInt_Err
    '</EhHeader>

    If IsNumeric(strNumber) Then
            
        IsInt = True
    
        If InStr(strNumber, ".") Then IsInt = False
        If InStr(strNumber, ",") Then IsInt = False

    Else
        IsInt = False
        
    End If
    
    If IsInt = False Then
        MsgBox "You must enter a whole number for Quantity."
    End If

    '<EhFooter>
    Exit Function

IsInt_Err:
    MsgBox Err.Description & vbCrLf & "in RestaurantMenu.modRest.IsInt " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Function Logout()
'-----------------------------------------------------------------------------
' Procedure   :       Logout
' Parameters  :
' Description :       Logout, set punchout time, show login prompt
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo Logout_Err
    '</EhHeader>

    'Close the main form and show the login prompt
    
    Dim strSQL As String
    
    strSQL = strSQL + "UPDATE T_INOUT "
    strSQL = strSQL + "SET F_PUNCHEDOUT='" & TimeValue(Time) & "'"
    strSQL = strSQL + " WHERE F_FNAME='" & strFirstName & "'"
    strSQL = strSQL + " AND F_LNAME='" & strLastName & "'"
    strSQL = strSQL + " AND F_DATE='" & LoginDate & "'"
    strSQL = strSQL + " AND F_ID=" & intLoginID & ""
    'MsgBox strSQL
    con.Execute (strSQL)
    
    frmLogin.txtUsername.Text = ""
    'Clear the username
    
    frmLogin.txtPassword.Text = ""
    'Clear the Password
    
    strUsername = ""
    'Clear the Username we stored earlier
    
    strFirstName = ""
    'And the First Name
    
    strLastName = ""
    'And last Name

    frmLogin.Visible = True
    'Show the login prompt
    
    con.Close

    '<EhFooter>
    Exit Function

Logout_Err:
    MsgBox Err.Description & vbCrLf & "in RestaurantMenu.modRest.Logout " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Function RetrieveFoods(intIndex As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       RetrieveFoods
' Parameters  :       intIndex (Integer)
' Description :       Retrieve Foods for General Use
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo RetrieveFoods_Err
    '</EhHeader>

    frmMain.lstFood(intIndex).Clear
    'Clear the main food listbox

    Dim strTable As String, recordset As New ADODB.recordset
    'Create some new vars
    
    Select Case intIndex
            'Switch strTable depending on intIndex
    
        Case 0
            strTable = "Breakfast"

        Case 1
            strTable = "Lunch"

        Case 2
            strTable = "Dinner"

        Case 3
            strTable = "Beverages"

        Case 4
            strTable = "Appetizers"

        Case 5
            strTable = "Desserts"
    
    End Select

    recordset.Open "SELECT * FROM T_" & UCase$(strTable), con, adOpenDynamic, _
            adLockOptimistic
    'open a new recordset
    
    With recordset
        While Not .EOF
            frmMain.lstFood(intIndex).AddItem (!F_ITEM & ": $" & !F_PRICE)
            'Add Item: $Price to the listbox
            
            .MoveNext
            'Move to the next item in the recordset
        Wend
        'Loop
        
    End With
    
    recordset.Close
    'Close the recordset so we can use rs again

    '<EhFooter>
    Exit Function

RetrieveFoods_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.modRest.RetrieveFoods " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

Public Function RetrieveFoodsAdmin(intIndex As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       RetrieveFoodsAdmin
' Parameters  :       intIndex (Integer)
' Description :       Retrieve Foods for Administration
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo RetrieveFoodsAdmin_Err
    '</EhHeader>
    
    Dim strTable As String, recordset As New ADODB.recordset
    'Create some new vars
    
    Select Case intIndex
            'Switch strTable by intIndex
    
        Case 0
            strTable = "Breakfast"

        Case 1
            strTable = "Lunch"

        Case 2
            strTable = "Dinner"

        Case 3
            strTable = "Beverages"

        Case 4
            strTable = "Appetizers"

        Case 5
            strTable = "Desserts"
    
    End Select

    If intIndex = 3 Then
        'If we need beverages we have to do some different SQL
    
        frmMenuAddEditDel.lstFoodName.Clear
        'Clear the listbox contents...
            
        frmMenuAddEditDel.lstFoodDesc.Clear
            
        frmMenuAddEditDel.lstFoodID.Clear
            
        frmMenuAddEditDel.lstFoodPrice.Clear
            
        frmMenuAddEditDel.txtItemDesc.Text = ""
        'Clear the itemdesc textbox
        
        recordset.Open "SELECT F_ALCOHOLIC, F_ID, F_ITEM, F_PRICE FROM T_" & _
                UCase$(strTable) & "", con, adOpenDynamic, adLockOptimistic
        'Select the data from the beverages table
            
        Dim strAlcoholic As String
            
        With recordset
            'With the recordset
            
            While Not .EOF
                'While we're not at the end of the recordset
                
                If recordset!F_ALCOHOLIC = True Then
                    strAlcoholic = "Alcoholic"
                Else
                    strAlcoholic = "Not Alcoholic"
                        
                End If
                
                frmMenuAddEditDel.lstFoodName.AddItem (recordset!F_ITEM)
                'Add the item to the listbox
                    
                frmMenuAddEditDel.lstFoodPrice.AddItem (recordset!F_PRICE)
                    
                frmMenuAddEditDel.lstFoodDesc.AddItem (strAlcoholic)
                    
                frmMenuAddEditDel.lstFoodID.AddItem (recordset!F_ID)
                    
                .MoveNext
                'Move to the next item in the recordset
            Wend
            'LOOP
                
        End With
            
        frmMenuAddEditDel.lblDirectionsAddEdit.Caption = "You are viewing " & _
                strTable & _
                ".  Choose an item from the left panel to view the description of the item."
        'Change this lbl's caption...

        recordset.Close
        'Close the recordset so we can use rs again if we need to
    Else
        'Its not beverages

        frmMenuAddEditDel.lstFoodDesc.Clear
            
        frmMenuAddEditDel.lstFoodID.Clear
            
        frmMenuAddEditDel.lstFoodPrice.Clear
            
        frmMenuAddEditDel.lstFoodName.Clear
        'Clear the listbox contents...
        
        frmMenuAddEditDel.txtItemDesc.Text = ""
        'Clear the itemdesc textbox
        
        recordset.Open "SELECT * FROM T_" & UCase$(strTable) & "", con, _
                adOpenDynamic, adLockOptimistic
        'Select the data from the table
        
        With recordset
            'With the recordset
        
            While Not .EOF
                'While we're not at the end of the recordset
            
                frmMenuAddEditDel.lstFoodName.AddItem (!F_ITEM)
                'Add the item to the listbox
                
                frmMenuAddEditDel.lstFoodPrice.AddItem (!F_PRICE)
                'Add the item price to the listbox
                
                frmMenuAddEditDel.lstFoodID.AddItem (!F_ID)
                'Add the item ID to the listbox
                
                frmMenuAddEditDel.lstFoodDesc.AddItem (!F_DESC)
                'Add the item desc to the listbox
                
                .MoveNext
                'Move to the next item in the recordset
            Wend
            'LOOP
            
        End With
        
        frmMenuAddEditDel.lblDirectionsAddEdit.Caption = "You are viewing " & _
                strTable & _
                ".  Choose an item from the left panel to view the description of the item."
        'Change this lbl's caption...
        
        recordset.Close
        'Close the recordset
        
    End If

    '<EhFooter>
    Exit Function

RetrieveFoodsAdmin_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.modRest.RetrieveFoodsAdmin " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Function

