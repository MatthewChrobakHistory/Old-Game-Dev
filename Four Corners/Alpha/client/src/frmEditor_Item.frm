VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   643
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   882
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   107
      Top             =   4800
      Visible         =   0   'False
      Width           =   9855
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   132
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   1320
         List            =   "frmEditor_Item.frx":3348
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   360
         Width           =   4815
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   130
         Top             =   840
         Width           =   1815
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   129
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5280
         Max             =   255
         TabIndex        =   128
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   127
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   126
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   4560
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   125
         Top             =   840
         Value           =   100
         Width           =   1575
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   120
         ScaleHeight     =   72
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   400
         TabIndex        =   124
         Top             =   1920
         Width           =   6000
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   5040
         TabIndex        =   123
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Frame fraPaperdoll 
         Caption         =   "MALE"
         Height          =   2775
         Left            =   6240
         TabIndex        =   108
         Top             =   240
         Width           =   3495
         Begin VB.HScrollBar scrlPaperdoll6 
            Height          =   255
            Left            =   1680
            TabIndex        =   114
            Top             =   2400
            Width           =   1695
         End
         Begin VB.HScrollBar scrlPaperdoll5 
            Height          =   255
            Left            =   1680
            TabIndex        =   113
            Top             =   960
            Width           =   1695
         End
         Begin VB.HScrollBar scrlPaperdoll4 
            Height          =   255
            Left            =   1680
            TabIndex        =   112
            Top             =   2040
            Width           =   1695
         End
         Begin VB.HScrollBar scrlPaperdoll3 
            Height          =   255
            Left            =   1680
            TabIndex        =   111
            Top             =   600
            Width           =   1695
         End
         Begin VB.HScrollBar scrlPaperdoll2 
            Height          =   255
            Left            =   1680
            TabIndex        =   110
            Top             =   1680
            Width           =   1695
         End
         Begin VB.HScrollBar scrlPaperdoll1 
            Height          =   255
            Left            =   1680
            TabIndex        =   109
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblPaperdoll6 
            AutoSize        =   -1  'True
            Caption         =   "Female 2h: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   122
            Top             =   2400
            Width           =   990
         End
         Begin VB.Label lblPaperdoll5 
            AutoSize        =   -1  'True
            Caption         =   "Male 2h: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   121
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label lblPaperdoll4 
            AutoSize        =   -1  'True
            Caption         =   "Female Shield: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   120
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Label lblPaperdoll3 
            AutoSize        =   -1  'True
            Caption         =   "Male Shield: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   119
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label lblPaperdoll2 
            AutoSize        =   -1  'True
            Caption         =   "Female Standing: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   118
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblPaperdoll1 
            AutoSize        =   -1  'True
            Caption         =   "Male Standing: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label5 
            Caption         =   "--------------------------------"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label6 
            Caption         =   "FEMALE"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   1320
            Width           =   3015
         End
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   141
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   140
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   139
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2160
         TabIndex        =   138
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4440
         TabIndex        =   137
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   136
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   135
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   3240
         TabIndex        =   134
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblPaperdoll 
         Caption         =   "Paperdoll: 0"
         Height          =   255
         Left            =   3960
         TabIndex        =   133
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Consume Data"
      Height          =   3615
      Left            =   3360
      TabIndex        =   95
      Top             =   4920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   106
         Top             =   3000
         Width           =   3495
      End
      Begin VB.CheckBox chkInstant 
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   5000
         TabIndex        =   99
         Top             =   600
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   5000
         TabIndex        =   98
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   5000
         TabIndex        =   97
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   96
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   180
         Left            =   120
         TabIndex        =   104
         Top             =   2760
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   103
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   102
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   101
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         Caption         =   "Cast Spell: None"
         Height          =   180
         Left            =   120
         TabIndex        =   100
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "XP Rewarded"
      Height          =   2655
      Left            =   11520
      TabIndex        =   72
      Top             =   2160
      Width           =   1695
      Begin VB.HScrollBar scrlPotionBrewingRew 
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   2400
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCraftingRew 
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1920
         Width           =   1455
      End
      Begin VB.HScrollBar scrlFletchingRew 
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1440
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCookingRew 
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   960
         Width           =   1455
      End
      Begin VB.HScrollBar scrlSmithingRew 
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblPotionBrewingRew 
         Caption         =   "Potion Brew: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblCraftingRew 
         Caption         =   "Crafting: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblFletchingRew 
         Caption         =   "Fletching: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCookingRew 
         Caption         =   "Cooking: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblSmithingRew 
         Caption         =   "Smithing: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Req to make"
      Height          =   2655
      Left            =   9720
      TabIndex        =   63
      Top             =   2160
      Width           =   1695
      Begin VB.HScrollBar scrlPotionBrewing 
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   2400
         Width           =   1455
      End
      Begin VB.HScrollBar scrlSmithing 
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCooking 
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   960
         Width           =   1455
      End
      Begin VB.HScrollBar scrlFletching 
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCrafting 
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblPotionBrewing 
         Caption         =   "Potion Brew: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblSmithing 
         Caption         =   "Smithing: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCooking 
         Caption         =   "Cooking: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblFletching 
         Caption         =   "Fletching: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCrafting 
         Caption         =   "Crafting: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Skill Requirements"
      Height          =   2175
      Left            =   9720
      TabIndex        =   56
      Top             =   0
      Width           =   3495
      Begin VB.HScrollBar scrlPotionBrewingXP 
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   1920
         Width           =   1575
      End
      Begin VB.HScrollBar scrlCraftingXP 
         Height          =   255
         Left            =   1800
         TabIndex        =   88
         Top             =   1920
         Width           =   1575
      End
      Begin VB.HScrollBar scrlFletchingXP 
         Height          =   255
         Left            =   1800
         TabIndex        =   83
         Top             =   1440
         Width           =   1575
      End
      Begin VB.HScrollBar scrlCookingXP 
         Height          =   255
         Left            =   1800
         TabIndex        =   82
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrlSmithingXP 
         Height          =   255
         Left            =   1800
         TabIndex        =   81
         Top             =   480
         Width           =   1575
      End
      Begin VB.HScrollBar scrlMining 
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1440
         Width           =   1575
      End
      Begin VB.HScrollBar scrlFishing 
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   960
         Width           =   1575
      End
      Begin VB.HScrollBar scrlWoodcutting 
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblPotionBrewingXP 
         Caption         =   "Potion Brew: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblCraftingXP 
         Caption         =   "Crafting: 0"
         Height          =   255
         Left            =   1800
         TabIndex        =   87
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblFletchingXP 
         Caption         =   "Fletching: 0"
         Height          =   255
         Left            =   1800
         TabIndex        =   86
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblCookingXP 
         Caption         =   "Cooking: 0"
         Height          =   255
         Left            =   1800
         TabIndex        =   85
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblSmithingXP 
         Caption         =   "Smithing: 0"
         Height          =   255
         Left            =   1800
         TabIndex        =   84
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblMining 
         Caption         =   "Mining: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblFishing 
         Caption         =   "Fishing: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblWoodcutting 
         Caption         =   "Woodcutting: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Frame fraProjectiles 
      Caption         =   "Projectiles"
      Height          =   900
      Left            =   3360
      TabIndex        =   46
      Top             =   7920
      Visible         =   0   'False
      Width           =   6240
      Begin VB.HScrollBar scrlProjectileDamage 
         Height          =   255
         Left            =   4170
         TabIndex        =   54
         Top             =   525
         Width           =   1470
      End
      Begin VB.HScrollBar scrlProjectileSpeed 
         Height          =   255
         Left            =   4170
         TabIndex        =   52
         Top             =   180
         Width           =   1470
      End
      Begin VB.HScrollBar scrlProjectileRange 
         Height          =   255
         Left            =   1080
         TabIndex        =   50
         Top             =   525
         Width           =   1470
      End
      Begin VB.HScrollBar scrlProjectilePic 
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label lblProjectileDamage 
         Caption         =   "Damage: 0"
         Height          =   225
         Left            =   3000
         TabIndex        =   53
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label lblProjectilesSpeed 
         Caption         =   "Speed: 0"
         Height          =   225
         Left            =   3000
         TabIndex        =   51
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lblProjectileRange 
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   150
         TabIndex        =   49
         Top             =   540
         Width           =   765
      End
      Begin VB.Label lblProjectilePiC 
         BackStyle       =   0  'Transparent
         Caption         =   "Pic: 0"
         Height          =   270
         Left            =   150
         TabIndex        =   47
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox ChkTwoh 
         Caption         =   "2h Sword"
         Height          =   180
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   1215
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   44
         Top             =   2760
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   42
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtDesc 
         Height          =   1455
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":337E
         Left            =   4200
         List            =   "frmEditor_Item.frx":338B
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33B4
         Left            =   120
         List            =   "frmEditor_Item.frx":33D6
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   45
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   43
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   41
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   38
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   31
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   3600
      Width           =   6255
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   225
      TabIndex        =   5
      Top             =   9135
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   8910
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   8520
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   32
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   33
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub ChkTwoh_Click()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ChkTwoh.Value = 0 Then
        Item(EditorIndex).istwohander = False
    Else
        Item(EditorIndex).istwohander = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkTwoh", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPic.Max = NumItems
    scrlAnim.Max = MAX_ANIMATIONS
    scrlPaperdoll.Max = NumPaperdolls
    scrlPaperdoll1.Max = NumPaperdolls
    scrlPaperdoll2.Max = NumPaperdolls
    scrlPaperdoll3.Max = NumPaperdolls
    scrlPaperdoll4.Max = NumPaperdolls
    scrlPaperdoll5.Max = NumPaperdolls
    scrlPaperdoll6.Max = NumPaperdolls
    scrlItem.Max = MAX_ITEMS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub scrlCookingXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblCookingXP.Caption = "Cooking: " & scrlCookingXP.Value
    Item(EditorIndex).EqCoXP = scrlCookingXP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCookingXP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCraftingXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblCraftingXP.Caption = "Crafting: " & scrlCraftingXP.Value
    Item(EditorIndex).EqCrXP = scrlCraftingXP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCraftingXP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFletchingXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFletchingXP.Caption = "Fletching: " & scrlFletchingXP.Value
    Item(EditorIndex).EqFlXP = scrlFletchingXP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFletchingXP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItem_Change()

    If scrlItem.Value = 0 Then
        lblItem.Caption = "Item: None"
        Item(EditorIndex).ConsumeItem = scrlItem.Value
        Exit Sub
    End If
    
    lblItem.Caption = "Item: " & Item(scrlItem.Value).Name
    Item(EditorIndex).ConsumeItem = scrlItem.Value
    
End Sub

Private Sub scrlPotionBrewing_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblPotionBrewing.Caption = "Potion Brew: " & scrlPotionBrewing.Value
Item(EditorIndex).PBXP = scrlPotionBrewing.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlPotionBrewing_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub scrlPotionBrewingRew_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblPotionBrewingRew.Caption = "Potion Brew: " & scrlPotionBrewingRew.Value
Item(EditorIndex).PBRew = scrlPotionBrewingRew.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlPotionBrewingRew_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub scrlPotionBrewingXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblPotionBrewingXP.Caption = "Potion Brew: " & scrlPotionBrewingXP.Value
    Item(EditorIndex).EqPBXP = scrlPotionBrewingXP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPotionBrewingXP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSmithing_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblSmithing.Caption = "Smithing: " & scrlSmithing.Value
Item(EditorIndex).SmXP = scrlSmithing.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlSmithing_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlCooking_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblCooking.Caption = "Cooking: " & scrlCooking.Value
Item(EditorIndex).CoXP = scrlCooking.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlCooking_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlFletching_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler
lblFletching.Caption = "Fletching: " & scrlFletching.Value
Item(EditorIndex).FlXP = scrlFletching.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlFletching_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlCrafting_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblCrafting.Caption = "Crafting: " & scrlCrafting.Value
Item(EditorIndex).CrXP = scrlCrafting.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlCrafting_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlSmithingRew_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblSmithingRew.Caption = "Smithing: " & scrlSmithingRew.Value
Item(EditorIndex).SmRew = scrlSmithingRew.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlSmithingRew_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlCookingRew_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblCookingRew.Caption = "Cooking: " & scrlCookingRew.Value
Item(EditorIndex).CoRew = scrlCookingRew.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlCookingRew_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlFletchingRew_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblFletchingRew.Caption = "Fletching: " & scrlFletchingRew.Value
Item(EditorIndex).FlRew = scrlFletchingRew.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlFletchingRew_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub
Private Sub scrlCraftingRew_Change()
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo errorhandler

lblCraftingRew.Caption = "Crafting: " & scrlCraftingRew.Value
Item(EditorIndex).CrRew = scrlCraftingRew.Value

' Error handler
Exit Sub
errorhandler:
HandleError "scrlCraftingRew_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
Err.Clear
Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If (cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
        ChkTwoh.Visible = True
        fraProjectiles.Visible = True
    Else
        ChkTwoh.Visible = False
        fraProjectiles.Visible = False
    End If

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
        'scrlDamage_Change
    Else
        fraEquipment.Visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub




Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access Req: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).Name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    Item(EditorIndex).Data2 = scrlDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFishing_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFishing.Caption = "Fishing: " & scrlFishing.Value
    Item(EditorIndex).FXP = scrlFishing.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFishing_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMining_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblMining.Caption = "Mining: " & scrlMining.Value
    Item(EditorIndex).MXP = scrlMining.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMining_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).paperdoll = scrlPaperdoll.Value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll1_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll1.Caption = "Male Standing: " & scrlPaperdoll1.Value
    Item(EditorIndex).Paperdoll1 = scrlPaperdoll1.Value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll2.Caption = "Female Standing: " & scrlPaperdoll2.Value
    Item(EditorIndex).Paperdoll2 = scrlPaperdoll2.Value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll3_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll3.Caption = "Male Shield: " & scrlPaperdoll3.Value
    Item(EditorIndex).Paperdoll3 = scrlPaperdoll3.Value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll4_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll4.Caption = "Female Shield: " & scrlPaperdoll4.Value
    Item(EditorIndex).Paperdoll4 = scrlPaperdoll4.Value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll5_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll5.Caption = "Male 2h: " & scrlPaperdoll5.Value
    Item(EditorIndex).Paperdoll5 = scrlPaperdoll5.Value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll6_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll6.Caption = "Female 2h: " & scrlPaperdoll6.Value
    Item(EditorIndex).Paperdoll6 = scrlPaperdoll6.Value
    Call EditorItem_BltPaperdoll
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
    Call EditorItem_BltItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.Value
    Item(EditorIndex).Price = scrlPrice.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileDamage.Caption = "Damage: " & scrlProjectileDamage.Value
    Item(EditorIndex).ProjecTile.Damage = scrlProjectileDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePiC.Caption = "Pic: " & scrlProjectilePic.Value
    Item(EditorIndex).ProjecTile.Pic = scrlProjectilePic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ProjecTile
Private Sub scrlProjectileRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.Value
    Item(EditorIndex).ProjecTile.Range = scrlProjectileRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilesSpeed.Caption = "Speed: " & scrlProjectileSpeed.Value
    Item(EditorIndex).ProjecTile.Speed = scrlProjectileSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSmithingXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSmithingXP.Caption = "Smithing: " & scrlSmithingXP.Value
    Item(EditorIndex).EqSmXP = scrlSmithingXP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSmithingXP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).Speed = scrlSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Will: "
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Will: "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.Value).Name)) > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.Caption = "Name: None"
    End If
    
    lblSpell.Caption = "Spell: " & scrlSpell.Value
    
    Item(EditorIndex).Data1 = scrlSpell.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlWoodcutting_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblWoodcutting.Caption = "Woodcuting: " & scrlWoodcutting.Value
    Item(EditorIndex).WcXP = scrlWoodcutting.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlWoodcutting_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

