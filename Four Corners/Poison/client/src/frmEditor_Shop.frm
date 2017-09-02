VERSION 5.00
Begin VB.Form frmEditor_Shop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Shop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   460
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Shop Properties"
      Height          =   6135
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox chkEdit 
         Caption         =   "Edit Values"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtNumCosts 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   4560
         TabIndex        =   28
         Text            =   "0"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox txtCostValue2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   26
         Text            =   "1"
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem2 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2040
         Width           =   3135
      End
      Begin VB.HScrollBar scrlShopType 
         Height          =   255
         Left            =   2760
         Max             =   8
         TabIndex        =   22
         Top             =   840
         Width           =   2415
      End
      Begin VB.CommandButton cmdDeleteTrade 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   2640
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBuy 
         Height          =   255
         Left            =   120
         Max             =   1000
         Min             =   1
         TabIndex        =   19
         Top             =   840
         Value           =   100
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox lstTradeItem 
         Height          =   2760
         ItemData        =   "frmEditor_Shop.frx":3332
         Left            =   120
         List            =   "frmEditor_Shop.frx":334E
         TabIndex        =   11
         Top             =   3000
         Width           =   5055
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtItemValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   9
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Num of costs:"
         Height          =   255
         Left            =   3480
         TabIndex        =   29
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   27
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label lblShopType 
         Caption         =   "Shop Type: None"
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblBuy 
         AutoSize        =   -1  'True
         Caption         =   "Buy Rate: 100%"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   13
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shop List"
      Height          =   6135
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   5640
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   6480
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call ShopEditorOk
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ShopEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
Dim Index As Long
Dim tmpPos As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    tmpPos = lstTradeItem.ListIndex
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = cmbItem.ListIndex
        .ItemValue = Val(txtItemValue.text)
        .CostItem = cmbCostItem.ListIndex
        .CostValue = Val(txtCostValue.text)
        .CostItem2 = cmbCostItem2.ListIndex
        .CostValue2 = Val(txtCostValue2.text)
        .NumCosts = txtNumCosts.text
    End With
    UpdateShopTrade tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdUpdate_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDeleteTrade_Click()
Dim Index As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = 0
        .ItemValue = 0
        .CostItem = 0
        .CostValue = 0
    End With
    Call UpdateShopTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDeleteTrade_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ShopEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstTradeItem_Click()
Dim Index As Long
Dim tmpPos As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
   
If chkEdit.Value = 1 Then

If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        cmbItem.ListIndex = .Item
        txtItemValue.text = .ItemValue
        cmbCostItem.ListIndex = .CostItem
        txtCostValue.text = .CostValue
        cmbCostItem2.ListIndex = .CostItem2
        txtCostValue2.text = .CostValue2
        txtNumCosts.text = .NumCosts
    End With
    
Else
    Exit Sub
End If
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkEdit_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBuy_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblBuy.Caption = "Buy Rate: " & scrlBuy.Value & "%"
    Shop(EditorIndex).BuyRate = scrlBuy.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlBuy_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlShopType_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

Select Case scrlShopType.Value

        Case 0
                lblShopType.Caption = "Shop Type: Shop"
        Case 1
                lblShopType.Caption = "Shop Type: Furnace"
        Case 2
                lblShopType.Caption = "Shop Type: Cooking"
        Case 3
                lblShopType.Caption = "Shop Type: Fletching"
        Case 4
                lblShopType.Caption = "Shop Type: Crafting"
        Case 5
                lblShopType.Caption = "Shop Type: Anvil"
        Case 6
                lblShopType.Caption = "Shop Type: Potion Brewing"
        Case 7
                lblShopType.Caption = "Shop Type: Water Source"
        Case 8
                lblShopType.Caption = "Shop Type: Pottery Oven"
        Case Else
                lblShopType.Caption = "Shop Type: None"
End Select

Shop(EditorIndex).ShopType = scrlShopType.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlBuy_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
