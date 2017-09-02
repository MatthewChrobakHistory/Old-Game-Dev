Attribute VB_Name = "modAccountEditor"
Option Explicit

Public EditUserIndex As Byte

Public Sub AddInfo(ByVal Text As String)

frmAccountEditor.lblInfo.Caption = Text

End Sub

Public Sub AccountEditorInit(ByVal index As Byte)
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    .FrameAccountDetails.Visible = True
    .txtUserName.Text = Trim$(Player(index).Name)
    .txtPassword.Text = Trim$(Player(index).Password)
    .txtLogin.Text = Trim$(Player(index).Login)
    .txtAccess.Text = Trim$(Player(index).Access)
    .cmbClass.ListIndex = Player(index).Class - 1
    .txtLevel.Text = Player(index).Level
    .txtSprite.Text = Player(index).Sprite
    .txtPoints.Text = Player(index).POINTS
    .txtXP.Text = Player(index).exp
    'skills
    .frameSkills.Visible = True
    .txtWoodcutting.Text = Player(index).WoodcuttingXP
    .txtMining.Text = Player(index).MiningXP
    .txtFishing.Text = Player(index).FishingXP
    .txtSmithing.Text = Player(index).SmithingXP
    .txtCooking.Text = Player(index).CookingXP
    .txtCrafting.Text = Player(index).CraftingXP
    .txtFletching.Text = Player(index).FletchingXP
    .txtPotionBrewing.Text = Player(index).PotionBrewingXP
    
    .txtDaily.Text = Player(index).DailyValue
    
    'bank
    .frameBank.Visible = True
    For i = 1 To 99
        If Bank(EditUserIndex).Item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Bank(EditUserIndex).Item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).Item(i).Value)
    Next
    .lstBank.ListIndex = 0
    
    'inventory
    .frameInventory.Visible = True
    For i = 1 To 35
        If Player(index).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Player(index).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & Player(index).Inv(i).Value)
    Next
    .lstInventory.ListIndex = 0
    
    'equipment
    .lblHelm.Caption = "Helm: "
    .lblBody.Caption = "Body: "
    .lblLegs.Caption = "Legs: "
    .lblShield.Caption = "Shield: "
    .lblWeapon.Caption = "Weapon: "
    If GetPlayerEquipment(index, Helmet) > 0 Then .lblHelm.Caption = "Helm: " & Trim$(Item(GetPlayerEquipment(index, Helmet)).Name)
    If GetPlayerEquipment(index, Armor) > 0 Then .lblBody.Caption = "Armor: " & Trim$(Item(GetPlayerEquipment(index, Armor)).Name)
    If GetPlayerEquipment(index, Legs) > 0 Then .lblLegs.Caption = "Legs: " & Trim$(Item(GetPlayerEquipment(index, Legs)).Name)
    If GetPlayerEquipment(index, Shield) > 0 Then .lblShield.Caption = "Shield: " & Trim$(Item(GetPlayerEquipment(index, Shield)).Name)
    If GetPlayerEquipment(index, Weapon) > 0 Then .lblWeapon.Caption = "Weapon: " & Trim$(Item(GetPlayerEquipment(index, Shield)).Name)
        
End With

End Sub

Public Sub BankEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    .lstBank.Clear
    For i = 1 To 99 '99 bank space
        If Bank(EditUserIndex).Item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Bank(EditUserIndex).Item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).Item(i).Value)
    Next
    .lstBank.ListIndex = 0
End With

End Sub

Public Sub SaveEditPlayer(ByVal index As Byte)

With Player(index)
    .Name = frmAccountEditor.txtUserName.Text
    .Password = frmAccountEditor.txtPassword.Text
    .Login = frmAccountEditor.txtLogin.Text
    .Access = frmAccountEditor.txtAccess.Text
    .Class = frmAccountEditor.cmbClass.ListIndex + 1
    .Level = frmAccountEditor.txtLevel.Text
    .Sprite = frmAccountEditor.txtSprite.Text
    '.exp = frmAccountEditor.txtXP.Text
    .POINTS = frmAccountEditor.txtPoints.Text
    'skills
    .WoodcuttingXP = frmAccountEditor.txtWoodcutting.Text
    .MiningXP = frmAccountEditor.txtMining.Text
    .FishingXP = frmAccountEditor.txtFishing.Text
    .SmithingXP = frmAccountEditor.txtSmithing.Text
    .CookingXP = frmAccountEditor.txtCooking.Text
    .CraftingXP = frmAccountEditor.txtCrafting.Text
    .FletchingXP = frmAccountEditor.txtFletching.Text
    .PotionBrewingXP = frmAccountEditor.txtPotionBrewing.Text
    
    .DailyValue = frmAccountEditor.txtDaily.Text
End With

Call CheckPlayerLevelUp(EditUserIndex)
Call SendPlayerData(index)

Call PlayerMsg(index, "Your account was edited by an admin!", Pink)

End Sub

Public Sub InvEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    'inventory
    .lstInventory.Clear
    .frameInventory.Visible = True
    For i = 1 To 35
        If Player(EditUserIndex).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Player(EditUserIndex).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & Player(EditUserIndex).Inv(i).Value)
    Next
    .lstInventory.ListIndex = 0
End With

End Sub
