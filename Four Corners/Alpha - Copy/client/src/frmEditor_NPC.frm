VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11085
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
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   739
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraBoss 
      Caption         =   "Boss Info"
      Height          =   1935
      Left            =   8400
      TabIndex        =   70
      Top             =   6240
      Width           =   2655
      Begin VB.HScrollBar scrlRewardItem 
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtMap 
         Height          =   270
         Left            =   600
         TabIndex        =   73
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtX 
         Height          =   270
         Left            =   600
         TabIndex        =   72
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtY 
         Height          =   270
         Left            =   600
         TabIndex        =   71
         Text            =   "0"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblItemReward 
         Caption         =   "Reward: None"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Map:"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   32
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   31
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   30
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC Properties"
      Height          =   7815
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtSpeed 
         Height          =   270
         Left            =   3960
         TabIndex        =   69
         Text            =   "1"
         Top             =   3600
         Width           =   855
      End
      Begin VB.Frame FraRequiredTools 
         Caption         =   "Required Tools"
         Height          =   3855
         Left            =   5040
         TabIndex        =   57
         Top             =   2160
         Width           =   2295
         Begin VB.HScrollBar scrlWeapon 
            Height          =   255
            Left            =   120
            Max             =   3
            TabIndex        =   67
            Top             =   3480
            Width           =   2055
         End
         Begin VB.HScrollBar scrlShield 
            Height          =   255
            Left            =   120
            Max             =   3
            TabIndex        =   66
            Top             =   2760
            Width           =   2055
         End
         Begin VB.HScrollBar scrlLegs 
            Height          =   255
            Left            =   120
            Max             =   3
            TabIndex        =   63
            Top             =   2040
            Width           =   2055
         End
         Begin VB.HScrollBar scrlArmor 
            Height          =   255
            Left            =   120
            Max             =   3
            TabIndex        =   61
            Top             =   1320
            Width           =   2055
         End
         Begin VB.HScrollBar scrlHelmet 
            Height          =   255
            Left            =   120
            Max             =   3
            TabIndex        =   59
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label LblWeapon 
            AutoSize        =   -1  'True
            Caption         =   "Weapon Required: None"
            Height          =   180
            Left            =   240
            TabIndex        =   65
            Top             =   3120
            Width           =   1785
         End
         Begin VB.Label LblShield 
            AutoSize        =   -1  'True
            Caption         =   "Shield Required: None"
            Height          =   180
            Left            =   240
            TabIndex        =   64
            Top             =   2400
            Width           =   1665
         End
         Begin VB.Label LblLegs 
            AutoSize        =   -1  'True
            Caption         =   "Legs Required: None"
            Height          =   180
            Left            =   120
            TabIndex        =   62
            Top             =   1680
            Width           =   1560
         End
         Begin VB.Label LblArmor 
            AutoSize        =   -1  'True
            Caption         =   "Armor Required: None"
            Height          =   180
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   1665
         End
         Begin VB.Label lblHelmet 
            AutoSize        =   -1  'True
            Caption         =   "Helmet Required: None"
            Height          =   180
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   1740
         End
      End
      Begin VB.Frame fraSpell 
         Caption         =   "Spells"
         Height          =   1815
         Left            =   5040
         TabIndex        =   52
         Top             =   240
         Width           =   1935
         Begin VB.HScrollBar scrlSpellNum 
            Height          =   255
            Left            =   120
            Max             =   5
            Min             =   1
            TabIndex        =   54
            Top             =   360
            Value           =   1
            Width           =   1695
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   53
            Top             =   1200
            Value           =   1
            Width           =   1695
         End
         Begin VB.Label lblSpellName 
            AutoSize        =   -1  'True
            Caption         =   "Spell: None"
            Height          =   180
            Left            =   120
            TabIndex        =   56
            Top             =   720
            Width           =   870
         End
         Begin VB.Label lblSpellNum 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   55
            Top             =   960
            Width           =   555
         End
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   50
         Text            =   "0"
         Top             =   3600
         Width           =   855
      End
      Begin VB.Frame fraDrop 
         Caption         =   "Drop"
         Height          =   2295
         Left            =   120
         TabIndex        =   41
         Top             =   5400
         Width           =   4815
         Begin VB.TextBox txtChance 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   45
            Text            =   "0"
            Top             =   840
            Width           =   1815
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   44
            Top             =   1560
            Width           =   3495
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   43
            Top             =   1920
            Width           =   3495
         End
         Begin VB.HScrollBar scrlDrop 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   42
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Chance:"
            Height          =   180
            Left            =   120
            TabIndex        =   49
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   630
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   1560
            Width           =   555
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   46
            Top             =   1920
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   120
            X2              =   4680
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   3240
         TabIndex        =   36
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Top             =   2880
         Width           =   1575
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   3240
         Width           =   2175
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   22
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   21
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1320
         List            =   "frmEditor_NPC.frx":3345
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1680
         Width           =   3615
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   18
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   600
         Width           =   3975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   4815
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1680
            Max             =   255
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   3240
            Max             =   255
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            Max             =   255
            TabIndex        =   8
            Top             =   840
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1680
            Max             =   255
            TabIndex        =   7
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   15
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   14
            Top             =   480
            Width           =   435
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Will: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   12
            Top             =   1080
            Width           =   480
         End
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "        Speed         (4 = default)"
         Height          =   495
         Left            =   2760
         TabIndex        =   68
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Rate (in seconds)"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   2640
         TabIndex        =   37
         Top             =   2880
         Width           =   465
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   2640
         TabIndex        =   24
         Top             =   2520
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7440
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   8040
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SpellIndex As Long
Private DropIndex As Byte

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.Max = NumCharacters
    scrlAnimation.Max = MAX_ANIMATIONS
    DropIndex = scrlDrop.Value
    scrlDrop.Max = MAX_NPC_DROPS
    scrlDrop.Min = 1
    fraDrop.Caption = "Drop - " & DropIndex
    scrlRewardItem.Max = MAX_ITEMS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnimation.Caption = "Anim: " & sString
    Npc(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlArmor_Change()
    Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlArmor.Value
        Case 0
            Name = "None"
        Case 1
            Name = "Hatchet"
        Case 2
            Name = "Rod"
        Case 3
            Name = "Pickaxe"
        Case 4
            Name = "Fishing Net"
        Case 5
            Name = "Shovel"
    End Select

    LblArmor.Caption = "Armor Required: " & Name
    
    Npc(EditorIndex).ArmorReq = scrlArmor.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLegs_Change()
    Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlLegs.Value
        Case 0
            Name = "None"
        Case 1
            Name = "Hatchet"
        Case 2
            Name = "Rod"
        Case 3
            Name = "Pickaxe"
        Case 4
            Name = "Fishing Net"
        Case 5
            Name = "Shovel"
    End Select

    LblLegs.Caption = "Legs Required: " & Name
    
    Npc(EditorIndex).LegsReq = scrlLegs.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRewardItem_Change()

If scrlRewardItem.Value = 0 Then
    lblItemReward.Caption = "Item: None"
    Exit Sub
End If

lblItemReward.Caption = "Item: " & Item(scrlRewardItem.Value).Name
Npc(EditorIndex).RewardItem = scrlRewardItem.Value

End Sub

Private Sub scrlShield_Change()
    Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlShield.Value
        Case 0
            Name = "None"
        Case 1
            Name = "Hatchet"
        Case 2
            Name = "Rod"
        Case 3
            Name = "Pickaxe"
        Case 4
            Name = "Fishing Net"
        Case 5
            Name = "Shovel"
    End Select

    LblShield.Caption = "Shield Required: " & Name
    
    Npc(EditorIndex).ShieldReq = scrlShield.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    lblSpellNum.Caption = "Num: " & scrlSpell.Value
    If scrlSpell.Value > 0 Then
        lblSpellName.Caption = "Spell: " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.Caption = "Spell: None"
    End If
    Npc(EditorIndex).Spell(SpellIndex) = scrlSpell.Value
End Sub

Private Sub scrlSpellNum_Change()
    SpellIndex = scrlSpellNum.Value
    fraSpell.Caption = "Spell - " & SpellIndex
    scrlSpell.Value = Npc(EditorIndex).Spell(SpellIndex)
End Sub

Private Sub scrlDrop_Change()
    DropIndex = scrlDrop.Value
    fraDrop.Caption = "Drop - " & DropIndex
    txtChance.text = Npc(EditorIndex).DropChance(DropIndex)
    scrlNum.Value = Npc(EditorIndex).DropItem(DropIndex)
    scrlValue.Value = Npc(EditorIndex).DropItemValue(DropIndex)
    
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Call EditorNpc_BltSprite
    Npc(EditorIndex).Sprite = scrlSprite.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.Value
    Npc(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNum.Caption = "Num: " & scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.Value).Name)
    End If
    
    Npc(EditorIndex).DropItem(DropIndex) = scrlNum.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(Index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Will: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    Npc(EditorIndex).Stat(Index) = scrlStat(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHelmet_Change()
    Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlHelmet.Value
        Case 0
            Name = "None"
        Case 1
            Name = "Hatchet"
        Case 2
            Name = "Rod"
        Case 3
            Name = "Pickaxe"
        Case 4
            Name = "Fishing Net"
        Case 5
            Name = "Shovel"
    End Select

    lblHelmet.Caption = "Helmet Required: " & Name
    
    Npc(EditorIndex).HelmetReq = scrlHelmet.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.Value
    Npc(EditorIndex).DropItemValue(DropIndex) = scrlValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlWeapon_Change()
    Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlWeapon.Value
        Case 0
            Name = "None"
        Case 1
            Name = "Hatchet"
        Case 2
            Name = "Rod"
        Case 3
            Name = "Pickaxe"
        Case 4
            Name = "Fishing Net"
        Case 5
            Name = "Shovel"
    End Select

    LblWeapon.Caption = "Weapon Required: " & Name
    
    Npc(EditorIndex).WeaponReq = scrlWeapon.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAttackSay_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).AttackSay = txtAttackSay.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChance_Change()
    On Error GoTo chanceErr
    
    If Not IsNumeric(txtChance.text) And Not Right$(txtChance.text, 1) = "%" And Not InStr(1, txtChance.text, "/") > 0 And Not InStr(1, txtChance.text, ".") Then
        txtChance.text = "0"
        Npc(EditorIndex).DropChance(DropIndex) = 0
        Exit Sub
    End If
    
    If Right$(txtChance.text, 1) = "%" Then
        txtChance.text = Left(txtChance.text, Len(txtChance.text) - 1) / 100
    ElseIf InStr(1, txtChance.text, "/") > 0 Then
        Dim i() As String
        i = Split(txtChance.text, "/")
        txtChance.text = Int(i(0) / i(1) * 1000) / 1000
    End If
    
    If txtChance.text > 1 Or txtChance.text < 0 Then
        'Err.Description = "Value must be between 0 and 1!"
        'GoTo chanceErr
    End If
    
    Npc(EditorIndex).DropChance(DropIndex) = txtChance.text
    Exit Sub
    
chanceErr:
    MsgBox "Invalid entry for chance! " & Err.Description
    txtChance.text = "0"
    Npc(EditorIndex).DropChance(DropIndex) = 0
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtDamage.text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.text) Then Npc(EditorIndex).Damage = Val(txtDamage.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtEXP.text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.text) Then Npc(EditorIndex).EXP = Val(txtEXP.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtHP.text) > 0 Then Exit Sub
    If IsNumeric(txtHP.text) Then Npc(EditorIndex).HP = Val(txtHP.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtLevel.text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.text) Then Npc(EditorIndex).Level = Val(txtLevel.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMap_Change()
If IsNumeric(txtMap.text) = False Then txtMap.text = 0

Npc(EditorIndex).Map = txtMap.text
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtSpawnSecs.text) > 0 Then Exit Sub
    Npc(EditorIndex).SpawnSecs = Val(txtSpawnSecs.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Npc(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Npc(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtSpeed.text) > 0 Then Exit Sub
    Npc(EditorIndex).Speed = Val(txtSpeed.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpeed_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtX_Change()
If IsNumeric(txtX.text) = False Then txtX.text = 1

Npc(EditorIndex).x = txtX.text
End Sub

Private Sub txtY_Change()
If IsNumeric(txtY.text) = False Then txtY.text = 1

Npc(EditorIndex).Y = txtY.text
End Sub
