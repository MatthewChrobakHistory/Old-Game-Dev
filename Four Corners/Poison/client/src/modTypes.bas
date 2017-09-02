Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Doors(1 To MAX_DOORS) As DoorRec

' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As ButtonRec
Public MainButton(1 To MAX_MAINBUTTONS) As ButtonRec
Public Party As PartyRec

' options
Public Options As OptionsRec

' Type recs
Private Type OptionsRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    Sound As Byte
    Debug As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    num As Long
    Value As Long
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    spellnum As Long
    Timer As Long
    FramePointer As Long
End Type

' projectiles
Public Type ProjectileRec
    TravelTime As Long
    Direction As Long
    x As Long
    Y As Long
    Pic As Long
    Range As Long
    Damage As Long
    Speed As Long
End Type

Public Type DoorRec
    Name As String * NAME_LENGTH
    DoorType As Long
    
    WarpMap As Long
    WarpX As Long
    WarpY As Long
    
    UnlockType As Long
    key As Long
    Switch As Long
    
    state As Long
End Type

Private Type PlayerRec
    ' General
    Name As String
    Class As Long
    Sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    Map As Long
    x As Byte
    Y As Byte
    Dir As Byte
    
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    
    ' projectiles
    ProjecTile(1 To MAX_PLAYER_PROJECTILES) As ProjectileRec
    PlayerDoors(1 To MAX_DOORS) As DoorRec
    
    WoodcuttingXP As Long
    MiningXP As Long
    FishingXP As Long
    DailyValue As Long
    SmithingXP As Long
    CookingXP As Long
    FletchingXP As Long
    CraftingXP As Long
    PotionBrewingXP As Long
    SecOutofCombat As Byte
End Type

Private Type TileDataRec
    x As Long
    Y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * NAME_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    paperdoll As Long
    Paperdoll1 As Long
    Paperdoll2 As Long
    Paperdoll3 As Long
    Paperdoll4 As Long
    Paperdoll5 As Long
    Paperdoll6 As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    
    'projectile
    ProjecTile As ProjectileRec
    istwohander As Boolean
    WcXP As Long
    FXP As Long
    MXP As Long
    SmXP As Long
    CoXP As Long
    FlXP As Long
    CrXP As Long
    PBXP As Long
    
    SmRew As Long
    CoRew As Long
    FlRew As Long
    CrRew As Long
    PBRew As Long
    EqSmXP As Long
    EqCoXP As Long
    EqFlXP As Long
    EqCrXP As Long
    EqPBXP As Long
    ConsumeItem As Long
End Type

Private Type MapItemRec
    playerName As String
    num As Long
    Value As Long
    Frame As Byte
    x As Byte
    Y As Byte
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropItemValue(1 To MAX_NPC_DROPS) As Integer
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    ' Npc Spells
    Spell(1 To MAX_NPC_SPELLS) As Long
    'Tool Requirements
    HelmetReq As Long
    ArmorReq As Long
    LegsReq As Long
    WeaponReq As Long
    ShieldReq As Long
    Speed As Byte
    
    Map As Long
    x As Byte
    Y As Byte
    message As String
    RewardItem As Long
End Type

Private Type MapNpcRec
    num As Long
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    x As Byte
    Y As Byte
    Dir As Byte
    ' Client use only
    XOffset As Long
    YOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    ' Npc spells
    SpellTimer(1 To MAX_NPC_SPELLS) As Long
    Heals As Integer
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
    CostItem2 As Long
    CostValue2 As Long
    NumCosts As Byte
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    ShopType As Byte
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 255
    Sound As String * NAME_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    x As Long
    Y As Long
    Dir As Byte
    Vital As Long
    VitalMax As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    trans As Long
End Type

Private Type TempTileRec
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    x As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health_min As Long
    health_max As Long
    RespawnTime As Long
    Walkthrough As Boolean
    Animation As Long
    DropChance(1 To MAX_RES_DROPS) As Double
    DropItem(1 To MAX_RES_DROPS) As Byte
    DropItemValue(1 To MAX_RES_DROPS) As Integer
    RewardXP As Long
    WcXP As Boolean
    FXP As Boolean
    MXP As Boolean
    'requirements for XP
    WcReq As Long
    FReq As Long
    MReq As Long
End Type

Private Type ActionMsgRec
    message As String
    Created As Long
    Type As Long
    color As Long
    Scroll As Long
    x As Long
    Y As Long
    Timer As Long
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Long
    x As Long
    Y As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    x As Long
    Y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    Timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    FrameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type ButtonRec
    fileName As String
    state As Byte
End Type
