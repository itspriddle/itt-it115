VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Menu"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   765
   ClientWidth     =   9465
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   0
      Top             =   -120
   End
   Begin VB.Frame fraLogout 
      Height          =   1215
      Left            =   6960
      TabIndex        =   58
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   85
         ToolTipText     =   "Log out"
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   86
         ToolTipText     =   "Log out"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "BLANK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Currently logged in as"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtOrderID 
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fraAdmin 
      Caption         =   "Administrative Options"
      Height          =   975
      Left            =   360
      TabIndex        =   43
      Top             =   5880
      Width           =   4335
      Begin VB.CommandButton cmdEditMenu 
         Caption         =   "Edit Order Menus"
         Height          =   375
         Left            =   2280
         TabIndex        =   84
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdEmployees 
         Caption         =   "Edit Employee Data"
         Height          =   375
         Left            =   240
         TabIndex        =   83
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   360
      TabIndex        =   21
      Top             =   4800
      Width           =   4335
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print Receipt"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdEditOrder 
         Caption         =   "Add to Existing Order"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Click here if you need to add more items to a party's order."
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   5175
      Left            =   4800
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
      Begin VB.CommandButton cmdShowFood 
         Caption         =   "Desserts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   5
         Left            =   2280
         Picture         =   "frmMain.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowFood 
         Caption         =   "Appetizers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   4
         Left            =   2280
         Picture         =   "frmMain.frx":16A9
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowFood 
         Caption         =   "Beverages"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   3
         Left            =   2280
         Picture         =   "frmMain.frx":1A21
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowFood 
         Caption         =   "Dinner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   2
         Left            =   240
         Picture         =   "frmMain.frx":1DAA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowFood 
         Caption         =   "Lunch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   1
         Left            =   240
         Picture         =   "frmMain.frx":2335
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowFood 
         Caption         =   "Breakfast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":26B8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "Order Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   42
      Top             =   1440
      Width           =   4335
      Begin VB.CommandButton cmdChangeDirections 
         Caption         =   "Save Changes"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   3000
         TabIndex        =   53
         ToolTipText     =   "Click here to save changes to the Cooking Instructions."
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtMainOrderInstructions 
         Height          =   855
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   51
         Top             =   600
         Width           =   1455
      End
      Begin VB.ListBox lstMainOrderDirections 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1155
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdMainRemoveSelected 
         Caption         =   "Remove Item"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   49
         ToolTipText     =   "Click here to remove the selected item from the order."
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtMainSubTotal 
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton cmdMainCancelOrder 
         Caption         =   "Cancel Order"
         Height          =   375
         Left            =   2640
         TabIndex        =   60
         ToolTipText     =   "Click here to cancel this order."
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdMainSendOrder 
         Caption         =   "Send Order"
         Height          =   375
         Left            =   240
         TabIndex        =   59
         ToolTipText     =   "Click here to send this order."
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ListBox lstMainOrderSummary 
         Height          =   840
         Left            =   240
         TabIndex        =   47
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblSpecialInstructions 
         Caption         =   "Special Instructions"
         Height          =   255
         Left            =   2640
         TabIndex        =   82
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblMainTotal 
         Caption         =   "Total"
         Height          =   255
         Left            =   2640
         TabIndex        =   54
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblMainSubTotal 
         Caption         =   "Sub Total"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblMainOrderSummary 
         Caption         =   "Current Items"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraDesserts 
      Caption         =   "Desserts: Add Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   69
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ListBox lstFood 
         Height          =   1620
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtCookingInstructions 
         Height          =   1155
         Index           =   5
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Index           =   5
         Left            =   3360
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelOrder 
         Caption         =   "Cancel Order"
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   14
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddToOrder 
         Caption         =   "Add To Order"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblSelectItem 
         Caption         =   "Select Item"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   75
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblServingInstructions 
         Caption         =   "Special Instructions"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   74
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   255
         Index           =   5
         Left            =   2640
         TabIndex        =   73
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.Frame fraAppetizers 
      Caption         =   "Appetizers: Add Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   68
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ListBox lstFood 
         Height          =   1620
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtCookingInstructions 
         Height          =   1155
         Index           =   4
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Index           =   4
         Left            =   3360
         TabIndex        =   17
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelOrder 
         Caption         =   "Cancel Order"
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   19
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddToOrder 
         Caption         =   "Add To Order"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblSelectItem 
         Caption         =   "Select Item"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   72
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblServingInstructions 
         Caption         =   "Special Instructions"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   71
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   255
         Index           =   4
         Left            =   2640
         TabIndex        =   70
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.Frame fraBeverages 
      Caption         =   "Beverages: Add Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   64
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ListBox lstFood 
         Height          =   1620
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtCookingInstructions 
         Height          =   1155
         Index           =   3
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Index           =   3
         Left            =   3360
         TabIndex        =   23
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelOrder 
         Caption         =   "Cancel Order"
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   25
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddToOrder 
         Caption         =   "Add To Order"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblSelectItem 
         Caption         =   "Select Item"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   67
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblServingInstructions 
         Caption         =   "Special Instructions"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   66
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   255
         Index           =   3
         Left            =   2640
         TabIndex        =   65
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.Frame fraDinner 
      Caption         =   "Dinner: Add Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   41
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ListBox lstFood 
         Height          =   1620
         Index           =   2
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtCookingInstructions 
         Height          =   1155
         Index           =   2
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   28
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelOrder 
         Caption         =   "Cancel Order"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   30
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddToOrder 
         Caption         =   "Add To Order"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblSelectItem 
         Caption         =   "Select Item"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   81
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblServingInstructions 
         Caption         =   "Special Instructions"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   80
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   79
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.Frame fraLunch 
      Caption         =   "Lunch: Add Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   63
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ListBox lstFood 
         Height          =   1620
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtCookingInstructions 
         Height          =   1155
         Index           =   1
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   33
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelOrder 
         Caption         =   "Cancel Order"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   35
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddToOrder 
         Caption         =   "Add To Order"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblSelectItem 
         Caption         =   "Select Item"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   78
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblServingInstructions 
         Caption         =   "Special Instructions"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   77
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   76
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.Frame fraBreakfast 
      Caption         =   "Breakfast: Add Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   360
      TabIndex        =   44
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdAddToOrder 
         Caption         =   "Add To Order"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   39
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancelOrder 
         Caption         =   "Cancel Order"
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   40
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Index           =   0
         Left            =   3360
         TabIndex        =   38
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCookingInstructions 
         Height          =   1155
         Index           =   0
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   1455
      End
      Begin VB.ListBox lstFood 
         Height          =   1620
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblQuantity 
         Caption         =   "Quantity"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   48
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblServingInstructions 
         Caption         =   "Special Instructions"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   46
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblSelectItem 
         Caption         =   "Select Item"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1320
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   1350
      Left            =   480
      Picture         =   "frmMain.frx":2959
      Top             =   -30
      Width           =   3135
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4560
      Picture         =   "frmMain.frx":51A5
      Top             =   480
      Width           =   1740
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Logout"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "Order"
      Begin VB.Menu mnuCurrentOrders 
         Caption         =   "Current Orders"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBreakfast 
         Caption         =   "Breakfast Menu"
      End
      Begin VB.Menu mnuLunch 
         Caption         =   "Lunch Menu"
      End
      Begin VB.Menu mnuDinner 
         Caption         =   "Dinner Menu"
      End
      Begin VB.Menu mnuBeverage 
         Caption         =   "Beverage Menu"
      End
      Begin VB.Menu mnuAppetizers 
         Caption         =   "Appetizer Menu"
      End
      Begin VB.Menu mnuDessert 
         Caption         =   "Dessert Menu"
      End
   End
   Begin VB.Menu mnuSchedule 
      Caption         =   "Schedule"
      Visible         =   0   'False
      Begin VB.Menu mnuViewSchedule 
         Caption         =   "View Schedule"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administration"
      Begin VB.Menu mnuAdminViewUsers 
         Caption         =   "View Users"
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuAdminEditMenu 
         Caption         =   "Edit Menu"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
      End
   End
   Begin VB.Menu mnuClock 
      Caption         =   "[TIME]"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rs As New ADODB.recordset
'Create a var called rs as the returned recordset

Public Sub cmdClose_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdClose_Click
' Parameters  :
' Description :       Logout
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdClose_Click_Err
    '</EhHeader>
    Unload Me
    '<EhFooter>
    Exit Sub

cmdClose_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdClose_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub cmdPrint_Click()
    MsgBox "Printing not yet implemented."
End Sub

Public Sub cmdShowFood_Click(Index As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       cmdShowFood_Click
' Parameters  :       Index (Integer)
' Description :       Show food based on button clicked (by index 0-5)
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdShowFood_Click_Err
    '</EhHeader>

    fraMain.Visible = False             'Hide all frames first
    fraBreakfast.Visible = False        'and we'll just the one we
    fraLunch.Visible = False            'need next
    fraDinner.Visible = False
    fraBeverages.Visible = False
    fraAppetizers.Visible = False
    fraDesserts.Visible = False

    Select Case Index
            'Depending on which button they click...
    
        Case 0
            'Show Breakfast
        
            fraBreakfast.Visible = True
        
        Case 1
            'Show Lunch
        
            fraLunch.Visible = True
        
        Case 2
            'Show Dinner
        
            fraDinner.Visible = True
        
        Case 3
            'Show Beverages
        
            fraBeverages.Visible = True
        
        Case 4
            'Show Appetizers
        
            fraAppetizers.Visible = True
        
        Case 5
            'Show Desserts
        
            fraDesserts.Visible = True
        
    End Select
    
    RetrieveFoods (Index)
    'Fill in any lstboxes or txtboxes

    '<EhFooter>
    Exit Sub

cmdShowFood_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdShowFood_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdAddToOrder_Click(Index As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       cmdAddToOrder_Click
' Parameters  :       Index (Integer)
' Description :       Add to existing order
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdAddToOrder_Click_Err
    '</EhHeader>

    Select Case Index
    
        Case 0
            AddToOrder (Index)

        Case 1
            AddToOrder (Index)

        Case 2
            AddToOrder (Index)

        Case 3
            AddToOrder (Index)

        Case 4
            AddToOrder (Index)

        Case 5
            AddToOrder (Index)
    End Select

    '<EhFooter>
    Exit Sub

cmdAddToOrder_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdAddToOrder_Click " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdCancelOrder_Click(Index As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       cmdCancelOrder_Click
' Parameters  :       Index (Integer)
' Description :       Cancel items from menu order
'
'                     This cancel button is for the ones that show in
'                     Breakfast, Lunch, etc. NOT the main cancel button.
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdCancelOrder_Click_Err
    '</EhHeader>

    'IE: if you have items selected in the Breakfast window
    'and click cancel, all items from Breakfast are removed
    'but any on the main order window are retained

    Select Case Index

            'Depending on which Cancel button they press
            'Hide the open frame
        Case 0
            fraBreakfast.Visible = False
        
        Case 1
            fraLunch.Visible = False
        
        Case 2
            fraDinner.Visible = False
        
        Case 3
            fraBeverages.Visible = False
        
        Case 4
            fraAppetizers.Visible = False
        
        Case 5
            fraDesserts.Visible = False
    
    End Select

    fraMain.Visible = True
    'And make the main one visible again
    '<EhFooter>
    Exit Sub

cmdCancelOrder_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdCancelOrder_Click " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdChangeDirections_Click(Index As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       cmdChangeDirections_Click
' Parameters  :       Index (Integer)
' Description :       Save any changes made to the Instructions window
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdChangeDirections_Click_Err
    '</EhHeader>

    If Me.lstMainOrderSummary.ListCount > 0 Then
        'If theres at least 1 item in the Order
    
        If Me.lstMainOrderSummary.Text <> "" Then
            'If they have selected an item
        
            Me.lstMainOrderDirections.RemoveItem ( _
                    Me.lstMainOrderSummary.ListIndex)
            'Remove the selected item from the Directions listbox
            
            Me.lstMainOrderDirections.AddItem Me.txtMainOrderInstructions, _
                    Me.lstMainOrderSummary.ListIndex
            'Add what they typed into the directions txtbox inplace of what was
            'just removed
            
        Else
            'There isn't an item selected
        
            MsgBox "You must select an item first!", , "Can't Save"
            'So let the user know.
            
        End If

    Else
        'The order is blank
    
        MsgBox "There is nothing to edit.", , "Can't Save"
        'So let the user know
        
    End If

    '<EhFooter>
    Exit Sub

cmdChangeDirections_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdChangeDirections_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Public Sub cmdEditMenu_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdEditMenu_Click
' Parameters  :
' Description :       Edit the Menus
'
'                     IE: Add/Edit/Delete items from the restaurants menu
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdEditMenu_Click_Err
    '</EhHeader>

    Me.Hide
    'Hide this form
    
    frmMenuAddEditDel.Visible = True
    'Show the Menu Administration form
    
    '<EhFooter>
    Exit Sub

cmdEditMenu_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdEditMenu_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdEditOrder_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdEditOrder_Click
' Parameters  :
' Description :       Add to an existing order
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdEditOrder_Click_Err
    '</EhHeader>

    Dim strInputOrderNum As String, rsOrderNum As New ADODB.recordset
    'Create a new string and a new recordset
    
    strInputOrderNum = InputBox("Enter the order number:", "Enter Order Number")
    'Set that string equal to the value entered from an inputbox
    
    If strInputOrderNum <> "" Then
        'If they gave us an order number
    
        rsOrderNum.Open "SELECT * FROM T_ORDERS WHERE F_ID=" & _
                strInputOrderNum, con, adOpenDynamic, adLockOptimistic
        'Open a recordset corresponding with the order number just entered
        
        If rsOrderNum.RecordCount < 1 Then
            'If that recordset returned 0 results...
            MsgBox "There was no order found with that order number.", _
                    vbOKOnly, "No Order"
            'Tell the user they entered the wrong order number
            
        Else
            'The record count is 1 so we can edit the order
        
            If Me.lstMainOrderSummary.ListCount > 0 Then
                'If theres anything already in the order window
                'we want to add it to the existing order
            
                Me.txtTotal = Me.txtTotal + rsOrderNum!F_TOTAL
                'So add the totals
                
                Me.txtMainSubTotal = Me.txtMainSubTotal + Round(( _
                        rsOrderNum!F_TOTAL / 1.0725), 2)
                'And the subtotals
                
            Else
                'They haven't entered anything yet so we dont need to add it
                'to the existing order
            
                Me.txtTotal = rsOrderNum!F_TOTAL
                'So just set the total = to the one from DB
            
                Me.txtMainSubTotal = Round((rsOrderNum!F_TOTAL / 1.0725), 2)
                'And do the same for the subtotal, but remove the tax
                
            End If
            
            Dim strOrder As String, strInstructions As String
            'Create some new strings
            
            strOrder = rsOrderNum!F_ORDER
            'Set strorder to the order from the DB
            
            strInstructions = rsOrderNum!F_INSTRUCTIONS
            'And the instructions to those from the DB
            
            Dim strItems() As String, strAryInstructions() As String
            'Dim some new arrays
            
            strItems() = Split(strOrder, ",")
            'And fill them
            
            strAryInstructions() = Split(strInstructions, "||")
            'Fill this one too
            
            Dim i As Integer, ii As Integer
            'Counter time
            
            For i = 0 To UBound(strItems())
                'For 0 to the number of items in strItems
            
                Me.lstMainOrderSummary.AddItem (strItems(i))
                'Add the current item to the order summary
                
            Next 'Loop
            
            For ii = 0 To UBound(strAryInstructions())
                'For 0 to the number of items in strAryInstructions
            
                Me.lstMainOrderDirections.AddItem (strAryInstructions(ii))
                'Add the current item to the directions
                
            Next 'Loop
            
            Me.txtOrderID.Text = rsOrderNum!F_ID
            'And set the hidden ID to the ID
            'We'll use this later for the UPDATE statement
            
        End If
        
        rsOrderNum.Close
        'Close the recordset

    End If

    '<EhFooter>
    Exit Sub

cmdEditOrder_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdEditOrder_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdEmployees_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdEmployees_Click
' Parameters  :
' Description :       Show Employee Administration
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdEmployees_Click_Err
    '</EhHeader>
    Me.Hide
    frmUsersView.Visible = True
    
    '<EhFooter>
    Exit Sub

cmdEmployees_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdEmployees_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdMainCancelOrder_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdMainCancelOrder_Click
' Parameters  :
' Description :       Cancel the current order
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdMainCancelOrder_Click_Err
    '</EhHeader>

    Call ClearOrder
    'Clear the order
    
    MsgBox "Order cancelled.", , "Canceled"
    'Let the user know in case the blank
    'text boxes weren't enough of a clue..
    
    '<EhFooter>
    Exit Sub

cmdMainCancelOrder_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdMainCancelOrder_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdMainRemoveSelected_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdMainRemoveSelected_Click
' Parameters  :
' Description :       Remove the selected item from the current order
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdMainRemoveSelected_Click_Err
    '</EhHeader>

    If Me.lstMainOrderSummary.ListCount > 0 Then
        'If there is at least one item in the Order Summary
    
        If Me.lstMainOrderSummary.Text <> "" Then
            'If they have something selected
        
            Dim intColPos As Integer, intQuantity As Integer
            'Create some new integers
             
            intColPos = InStr(Me.lstMainOrderSummary.Text, ":")
            'Set this int to the position of ':'
    
            intQuantity = Left$(Me.lstMainOrderSummary.Text, InStr( _
                    Me.lstMainOrderSummary.Text, ";") - 1)
            'And this to the Quantity
             
            Dim strSubPrice As String, curSubPrice As Currency
            'Create some more vars
             
            strSubPrice = Mid$(Me.lstMainOrderSummary.Text, intColPos + 3, ( _
                    Len(Me.lstMainOrderSummary.Text) - (intColPos + 6)))
            'Get the item price from the middle of the lstbox item
             
            curSubPrice = strSubPrice
            'Change str into cur for some math
             
            Me.txtMainSubTotal.Text = Me.txtMainSubTotal.Text - Round(( _
                    curSubPrice * intQuantity), 2)
            'Subtract total price of items being removed from the current subtotal
            'and do some rounding
             
            Me.txtTotal.Text = Me.txtTotal.Text - Round(((curSubPrice * _
                    intQuantity) * 1.0725), 2)
            'And do the same with the total
                     
            Me.lstMainOrderDirections.RemoveItem ( _
                    Me.lstMainOrderSummary.ListIndex)
            'Remove the item from the Order Directions listbox
             
            Me.lstMainOrderSummary.RemoveItem (Me.lstMainOrderSummary.ListIndex)
            'Remove the item from the Order Summary listbox
            
            If Me.lstMainOrderSummary.ListCount < 1 Then
                'If there aren't any items in the Order Summary
             
                Me.txtMainSubTotal.Text = "0"
                'Set the subtotal to 0
                 
                Me.txtTotal.Text = "0"
                'And the total to 0
                 
            End If
             
            Me.txtMainOrderInstructions.Text = ""
            'Clear the order instructions
             
        Else
            'Nothing selected to remove
        
            MsgBox "You must select an item to remove", , "Select Item"
            'Let the user know
            
        End If

    Else
        'Nothing at all to remove
    
        MsgBox "There is nothing to remove.", , "Can't Remove Item"
        'Let the user know
        
    End If
        
    '<EhFooter>
    Exit Sub

cmdMainRemoveSelected_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdMainRemoveSelected_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdMainSendOrder_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdMainSendOrder_Click
' Parameters  :
' Description :       Send the order to the DB (or kitchen...)
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdMainSendOrder_Click_Err
    '</EhHeader>

    If Me.lstMainOrderSummary.ListCount > 0 Then
        'If there is at least 1 item in the Order Summary
        
        If Me.txtOrderID <> "" Then
            'If txtOrderID has a value we're adding to an order
        
            If IsCurrency(Me.txtMainSubTotal.Text) = False Then Exit Sub
            
            If IsCurrency(Me.txtTotal.Text) = False Then Exit Sub
        
            Dim strSQLUpdate As String
            'Need a new sql string
            
            strSQLUpdate = strSQLUpdate + "UPDATE T_ORDERS "
            strSQLUpdate = strSQLUpdate + "SET F_TOTAL='" & Round( _
                    Me.txtTotal.Text, 2)
            strSQLUpdate = strSQLUpdate + "', F_ORDER='"
            
            Dim z As Integer, zz As Integer
            'And a few ints for counters
            
            For z = 0 To Me.lstMainOrderSummary.ListCount - 1
                'For each item in the listbox
            
                strSQLUpdate = strSQLUpdate + Me.lstMainOrderSummary.List(z)
                'Add it to the SQL string
                 
                If z <> Me.lstMainOrderSummary.ListCount - 1 Then
                    'If its not the last entry in the listbox add a comma
                    'to the SQL string, we don't need one if its the last
            
                    strSQLUpdate = strSQLUpdate + ","
                
                End If

            Next
            
            strSQLUpdate = strSQLUpdate + "', F_INSTRUCTIONS='"
            'More SQL
            
            For zz = 0 To Me.lstMainOrderSummary.ListCount - 1
                'And another loop adding the order directions to the listbox
            
                strSQLUpdate = strSQLUpdate + Me.lstMainOrderDirections.List( _
                        zz) & "||"
                
            Next
            
            strSQLUpdate = strSQLUpdate + "' "
            strSQLUpdate = strSQLUpdate + "WHERE F_ID=" & Me.txtOrderID.Text
            'Finish up the SQL
            
            'MsgBox strSQLUpdate 'Print the SQL (FOR TESTING PURPOSES)
            
            MsgBox "Order sent!", , "Order Sent"
            
            con.Execute (strSQLUpdate)
            'Execute the SQL
            
            Me.txtOrderID.Text = ""
            'And reset txtOrderID
        
            Call ClearOrder
            'Clear the order
            
        Else
            'We're adding a new order to the DB
        
            Dim strSQL As String
            'Create a new string for some SQL
            
            strSQL = strSQL + _
                    "INSERT INTO T_ORDERS (F_TOTAL, F_ORDER, F_INSTRUCTIONS) "
            'Start the SQL statement
            
            strSQL = strSQL + "VALUES ("
            'More sql...
            
            Dim strTotal As String
            'Create a new string
            
            strTotal = Round(Me.txtTotal.Text, 2)
            'And set it's value = to txtTotal.Text
            
            strSQL = strSQL + "'" & strTotal & "','"
            'Add 'total', to the SQL string
            
            Dim i As Integer
            'Counter time...
            
            For i = 0 To Me.lstMainOrderSummary.ListCount - 1
                'For 0 to the number of items in Order Summary
            
                strSQL = strSQL + Me.lstMainOrderSummary.List(i)
                'Add the current item
                
                If i <> Me.lstMainOrderSummary.ListCount - 1 Then
                    'If its not the last entry in the listbox add a comma
                    'to the SQL string, we don't need one if its the last
                
                    strSQL = strSQL + ","
                    
                End If
                
            Next

            'Loop
            
            strSQL = strSQL + "','"
    
            Dim x As Integer
            'Counter time...
            
            For x = 0 To Me.lstMainOrderDirections.ListCount - 1
                'For 0 to the total number of items in Order Directions
    
                strSQL = strSQL + Me.lstMainOrderDirections.List(x) & "||"
                'add it to the SQL string
    
            Next

            'Loop
    
            strSQL = strSQL + "') "
            'Close the SQL command
            
            'MsgBox strSQL 'Print the SQL (FOR TESTING PURPOSES)
            
            con.Execute (strSQL)
            'Execute SQL
            
            Dim findID As New ADODB.recordset
            'Create a new rs
            
            findID.Open "SELECT F_ID FROM T_ORDERS ORDER BY F_ID DESC", con, _
                    adOpenDynamic, adLockOptimistic
            'Pull the last ID
            
            MsgBox "Order sent!  The order number is " & findID!F_ID, , _
                    "Order sent!"
            'tell the user what that ID is so they can edit it later if needed
            
            findID.Close
            'close the recordset
            
            Call ClearOrder
            'Clear the order so we can start a new one
        End If
        
    Else
        'The order is blank
    
        MsgBox "The current order is blank.", , "Can't Send Order"
        'Let the user know...
        
    End If

    '<EhFooter>
    Exit Sub

cmdMainSendOrder_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdMainSendOrder_Click " & "at line " _
            & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub cmdOptions_Click()
'-----------------------------------------------------------------------------
' Procedure   :       cmdOptions_Click
' Parameters  :
' Description :       Options (right now just change password)
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo cmdOptions_Click_Err
    '</EhHeader>
    Me.Hide
    
    frmOptions.Visible = True
    
    '<EhFooter>
    Exit Sub

cmdOptions_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.cmdOptions_Click " & "at line " & Erl
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

    lblUserName.Caption = strUsername
    'Set the username
    
    con.CursorLocation = adUseClient
    'Set the cursor location
    
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & _
            App.Path & "\restaurant.mdb;" & _
            "Jet OLEDB:Database Password=password"
    'Set the connection and password
    
    con.Open
    'Open the connection
    '<EhFooter>
    Exit Sub

Form_Load_Err:
    MsgBox Err.Description & vbCrLf & "in RestaurantMenu.frmMain.Form_Load " _
            & "at line " & Erl
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
    Call Logout
    '<EhFooter>
    Exit Sub

Form_Unload_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.Form_Unload " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub lstFood_Click(Index As Integer)
'-----------------------------------------------------------------------------
' Procedure   :       lstFood_Click
' Parameters  :       Index (Integer)
' Description :       User clicks on lstFood
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo lstFood_Click_Err
    '</EhHeader>
    Me.txtCookingInstructions(Index).Text = ""
    'Clear the cooking instructions textbox
    'so the user can enter new ones if needed
    
    Me.txtQuantity(Index).Text = "1"
    'Set the Quantity to 1
    
    '<EhFooter>
    Exit Sub

lstFood_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.lstFood_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub lstMainOrderSummary_Click()
'-----------------------------------------------------------------------------
' Procedure   :       lstMainOrderSummary_Click
' Parameters  :
' Description :       Used to change details per item
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo lstMainOrderSummary_Click_Err
    '</EhHeader>

    Me.txtMainOrderInstructions.Text = Me.lstMainOrderDirections.List( _
            Me.lstMainOrderSummary.ListIndex)
    'Change the directions to whats in the listbox when the user clicks an
    'item
    
    '<EhFooter>
    Exit Sub

lstMainOrderSummary_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.lstMainOrderSummary_Click " & _
            "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuAdminAddUser_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuAdminAddUser_Click
' Parameters  :
' Description :       Show the Add User Form
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuAdminAddUser_Click_Err
    '</EhHeader>

    Me.Hide

    frmUsersAdd.Visible = True

    '<EhFooter>
    Exit Sub

mnuAdminAddUser_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuAdminAddUser_Click " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuAbout_Click()
    Me.Hide
    frmAbout.Visible = True
End Sub

Private Sub mnuAdminEditMenu_Click()
    Me.cmdEditMenu_Click
End Sub

Private Sub mnuAdminViewUsers_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuAdminViewUsers_Click
' Parameters  :
' Description :       Show the view users form
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuAdminViewUsers_Click_Err
    '</EhHeader>

    Me.Hide

    frmUsersView.Visible = True

    '<EhFooter>
    Exit Sub

mnuAdminViewUsers_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuAdminViewUsers_Click " & "at line " _
            & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuAppetizers_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuAdminViewUsers_Click
' Parameters  :
' Description :       Show the view users form
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuAppetizers_Click_Err
    '</EhHeader>
    Me.cmdShowFood_Click (4)
    '<EhFooter>
    Exit Sub

mnuAppetizers_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuAppetizers_Click " & "at line " & _
            Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuBeverage_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuBeverage_Click
' Parameters  :
' Description :       Show the Beverages Menu
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuBeverage_Click_Err
    '</EhHeader>
    Me.cmdShowFood_Click (3)
    '<EhFooter>
    Exit Sub

mnuBeverage_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuBeverage_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuBreakfast_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuBreakfast_Click
' Parameters  :
' Description :       Show the Breakfast menu
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuBreakfast_Click_Err
    '</EhHeader>
    Me.cmdShowFood_Click (0)
    '<EhFooter>
    Exit Sub

mnuBreakfast_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuBreakfast_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuCurrentOrders_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuCurrentOrders_Click
' Parameters  :
' Description :       Show current orders
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuCurrentOrders_Click_Err
    '</EhHeader>
    Call ClearOrder

    '<EhFooter>
    Exit Sub

mnuCurrentOrders_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuCurrentOrders_Click " & "at line " _
            & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuDessert_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuDessert_Click
' Parameters  :
' Description :       Show the Dessert Menu
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuDessert_Click_Err
    '</EhHeader>
    Me.cmdShowFood_Click (5)
    '<EhFooter>
    Exit Sub

mnuDessert_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuDessert_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuDinner_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuDinner_Click
' Parameters  :
' Description :       Show the Dinner Menu
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuDinner_Click_Err
    '</EhHeader>
    Me.cmdShowFood_Click (2)
    '<EhFooter>
    Exit Sub

mnuDinner_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuDinner_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub mnuFileClose_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuFileClose_Click
' Parameters  :
' Description :       Logout
'-----------------------------------------------------------------------------
    Call Logout
End Sub

Private Sub mnuFilePrint_Click()
    Me.cmdPrint_Click
End Sub

Private Sub mnuLunch_Click()
'-----------------------------------------------------------------------------
' Procedure   :       mnuLunch_Click
' Parameters  :
' Description :       Show the Lunch Menu
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error GoTo mnuLunch_Click_Err
    '</EhHeader>
    Me.cmdShowFood_Click (1)
    '<EhFooter>
    Exit Sub

mnuLunch_Click_Err:
    MsgBox Err.Description & vbCrLf & _
            "in RestaurantMenu.frmMain.mnuLunch_Click " & "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Timer1_Timer()
'-----------------------------------------------------------------------------
' Procedure   :       Timer1_Timer
' Parameters  :
' Description :       Used for clock on frmMain (in menu bar: mnuClock)
'-----------------------------------------------------------------------------
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    mnuClock.Caption = Trim$(Left$(TimeValue(Left$(Time, 4)), 4)) & " " & _
            Right$(Time, 2)
    
End Sub

