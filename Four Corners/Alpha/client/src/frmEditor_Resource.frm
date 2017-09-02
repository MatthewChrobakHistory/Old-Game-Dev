VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13440
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
   Icon            =   "frmEditor_Resource.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   896
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraSkill 
      Caption         =   "Skill Info"
      Height          =   4335
      Left            =   8520
      TabIndex        =   42
      Top             =   2520
      Width           =   4815
      Begin VB.Frame Frame4 
         Caption         =   "Skill Requirements"
         Height          =   2415
         Left            =   120
         TabIndex        =   49
         Top             =   1800
         Width           =   4455
         Begin VB.HScrollBar scrlMining 
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   2040
            Width           =   4095
         End
         Begin VB.HScrollBar scrlFishing 
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1320
            Width           =   4095
         End
         Begin VB.HScrollBar scrlWoodcutting 
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label lblMining 
            Caption         =   "Mining: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   1680
            Width           =   3135
         End
         Begin VB.Label lblFishing 
            Caption         =   "Fishing: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label lblWoodcutting 
            Caption         =   "Woodcutting: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.HScrollBar scrlRewardXP 
         Height          =   255
         Left            =   2160
         TabIndex        =   47
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Xp rewarded in?"
         Height          =   1095
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   4575
         Begin VB.CheckBox chkMXP 
            Caption         =   "Mining"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkFXP 
            Caption         =   "Fishing"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkWcXP 
            Caption         =   "Woodcutting"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Label lblRewardXP 
         Caption         =   "XP Rewarded: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   1935
      End
   End
   Begin VB.Frame fraDrop 
      Caption         =   "Drop"
      Height          =   2295
      Left            =   8520
      TabIndex        =   33
      Top             =   120
      Width           =   4815
      Begin VB.HScrollBar scrlDrop 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   37
         Top             =   240
         Value           =   1
         Width           =   4575
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   36
         Top             =   1920
         Width           =   3495
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   35
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   34
         Text            =   "0"
         Top             =   840
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   120
         X2              =   4680
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Value: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Chance:"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   28
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resource Properties"
      Height          =   7575
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.HScrollBar scrlHealthMax 
         Height          =   255
         Left            =   2640
         Max             =   255
         TabIndex        =   56
         Top             =   5520
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   7080
         Width           =   3975
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   29
         Top             =   6720
         Width           =   4815
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   1920
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   2640
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   23
         Top             =   2280
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Resource.frx":3332
         Left            =   960
         List            =   "frmEditor_Resource.frx":3342
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4320
         Width           =   4815
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   9
         Top             =   4920
         Width           =   4815
      End
      Begin VB.HScrollBar scrlHealth 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   8
         Top             =   5520
         Width           =   2295
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   7
         Top             =   2280
         Width           =   2280
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   6
         Top             =   6120
         Width           =   4815
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtMessage2 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   6480
         Width           =   1260
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   25
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1470
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Item Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   4080
         Width           =   1440
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   4680
         Width           =   1530
      End
      Begin VB.Label lblHealth 
         AutoSize        =   -1  'True
         Caption         =   "Health: 0 - 0"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Width           =   2100
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DropIndex As Byte

Private Sub chkFXP_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkFXP.Value = 0 Then
        Resource(EditorIndex).FXP = False
    Else
        Resource(EditorIndex).FXP = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkFXP", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkMXP_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkMXP.Value = 0 Then
        Resource(EditorIndex).MXP = False
    Else
        Resource(EditorIndex).MXP = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkMXP", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkWcXP_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkWcXP.Value = 0 Then
        Resource(EditorIndex).WcXP = False
    Else
        Resource(EditorIndex).WcXP = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkWcXP", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).ResourceType = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearResource EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlReward.Max = MAX_ITEMS
    DropIndex = scrlDrop.Value
    scrlDrop.Max = MAX_NPC_DROPS
    scrlDrop.Min = 1
    fraDrop.Caption = "Drop - " & DropIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ResourceEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFishing_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFishing.Caption = "Fishing: " & scrlFishing.Value
    Resource(EditorIndex).FReq = scrlFishing.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFishing_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHealthMax_Change()
lblHealth.Caption = "Health: " & scrlHealth.Value & " - " & scrlHealthMax.Value
Resource(EditorIndex).health_max = scrlHealthMax.Value
End Sub

Private Sub scrlMining_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMining.Caption = "Mining: " & scrlMining.Value
    Resource(EditorIndex).MReq = scrlMining.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMining_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRewardXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRewardXP.Caption = "XP Rewarded: " & scrlRewardXP.Value
    Resource(EditorIndex).RewardXP = scrlRewardXP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRewardXP_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ResourceEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnim.Caption = "Animation: " & sString
    Resource(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDrop_Change()
    DropIndex = scrlDrop.Value
    fraDrop.Caption = "Drop - " & DropIndex
    txtChance.text = Resource(EditorIndex).DropChance(DropIndex)
    scrlNum.Value = Resource(EditorIndex).DropItem(DropIndex)
    scrlValue.Value = Resource(EditorIndex).DropItemValue(DropIndex)
End Sub

Private Sub scrlExhaustedPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.Value
    EditorResource_BltSprite
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlExhaustedPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlHealth_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblHealth.Caption = "Health: " & scrlHealth.Value & " - " & scrlHealthMax.Value
Resource(EditorIndex).health_min = scrlHealth.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlHealth_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub scrlNormalPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.Value
    EditorResource_BltSprite
    Resource(EditorIndex).ResourceImage = scrlNormalPic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNormalPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    
    Resource(EditorIndex).DropItem(DropIndex) = scrlNum.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRespawn_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRespawn.Caption = "Respawn Time (Seconds): " & scrlRespawn.Value
    Resource(EditorIndex).RespawnTime = scrlRespawn.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRespawn_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlReward.Value > 0 Then
        lblReward.Caption = "Item Reward: " & Trim$(Item(scrlReward.Value).Name)
    Else
        lblReward.Caption = "Item Reward: None"
    End If
    
    Resource(EditorIndex).ItemReward = scrlReward.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlReward_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub scrlTool_Change()
    Dim Name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlTool.Value
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

    lblTool.Caption = "Tool Required: " & Name
    
    Resource(EditorIndex).ToolRequired = scrlTool.Value
    
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
    Resource(EditorIndex).DropItemValue(DropIndex) = scrlValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlWoodcutting_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblWoodcutting.Caption = "Woodcutting: " & scrlWoodcutting.Value
    Resource(EditorIndex).WcReq = scrlWoodcutting.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlWoodcutting_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChance_Change()
    On Error GoTo chanceErr
    
    If Not IsNumeric(txtChance.text) And Not Right$(txtChance.text, 1) = "%" And Not InStr(1, txtChance.text, "/") > 0 And Not InStr(1, txtChance.text, ".") Then
        txtChance.text = "0"
        Resource(EditorIndex).DropChance(DropIndex) = 0
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
    
    Resource(EditorIndex).DropChance(DropIndex) = txtChance.text
    Exit Sub
    
chanceErr:
    MsgBox "Invalid entry for chance! " & Err.Description
    txtChance.text = "0"
    Npc(EditorIndex).DropChance(DropIndex) = 0
End Sub

Private Sub txtMessage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).SuccessMessage = Trim$(txtMessage.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Resource(EditorIndex).EmptyMessage = Trim$(txtMessage2.text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage2_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Resource(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Resource(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
