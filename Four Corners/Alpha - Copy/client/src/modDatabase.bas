Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim fileName As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    fileName = App.Path & "\data files\logs\errors.txt"
    Open fileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal fileName As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & fileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(fileName)) > 0 Then
            FileExist = True
        End If
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(Header, Var, Value, File)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim fileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    fileName = App.Path & "\Data Files\config.ini"
    
    Call PutVar(fileName, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(fileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(fileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(fileName, "Options", "SavePass", Str(Options.SavePass))
    Call PutVar(fileName, "Options", "IP", Options.IP)
    Call PutVar(fileName, "Options", "Port", Str(Options.Port))
    Call PutVar(fileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(fileName, "Options", "Music", Str(Options.Music))
    Call PutVar(fileName, "Options", "Sound", Str(Options.Sound))
    Call PutVar(fileName, "Options", "Debug", Str(Options.Debug))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim fileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    fileName = App.Path & "\Data Files\config.ini"
    
    If Not FileExist(fileName, True) Then
        Options.Game_Name = "Eclipse Origins"
        Options.Password = vbNullString
        Options.SavePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.Sound = 1
        Options.Debug = 0
        SaveOptions
    Else
        Options.Game_Name = GetVar(fileName, "Options", "Game_Name")
        Options.Username = GetVar(fileName, "Options", "Username")
        Options.Password = GetVar(fileName, "Options", "Password")
        Options.SavePass = Val(GetVar(fileName, "Options", "SavePass"))
        Options.IP = GetVar(fileName, "Options", "IP")
        Options.Port = Val(GetVar(fileName, "Options", "Port"))
        Options.MenuMusic = GetVar(fileName, "Options", "MenuMusic")
        Options.Music = GetVar(fileName, "Options", "Music")
        Options.Sound = GetVar(fileName, "Options", "Sound")
        Options.Debug = GetVar(fileName, "Options", "Debug")
    End If
    
    ' show in GUI
    If Options.Music = 0 Then
        frmMain.optMOff.Value = True
    Else
        frmMain.optMOn.Value = True
    End If
    
    If Options.Sound = 0 Then
        frmMain.optSOff.Value = True
    Else
        frmMain.optSOn.Value = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim fileName As String
Dim f As Long
Dim x As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    fileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT

    f = FreeFile
    Open fileName For Binary As #f
    Put #f, , Map.Name
    Put #f, , Map.Music
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    Put #f, , Map.MaxX
    Put #f, , Map.MaxY

    For x = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Put #f, , Map.Tile(x, Y)
        Next

        DoEvents
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , Map.Npc(x)
    Next

    Close #f
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim fileName As String
Dim f As Long
Dim x As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    fileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    f = FreeFile
    Open fileName For Binary As #f
    Get #f, , Map.Name
    Get #f, , Map.Music
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    Get #f, , Map.MaxX
    Get #f, , Map.MaxY
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For x = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            Get #f, , Map.Tile(x, Y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , Map.Npc(x)
    Next

    Close #f
    ClearTempTile
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "\tilesets\" & i & GFX_EXT)
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend
    
    If NumTileSets = 0 Then Exit Sub
    
    ReDim DDS_Tileset(1 To NumTileSets)
    ReDim DDSD_Tileset(1 To NumTileSets)
    ReDim TilesetTimer(1 To NumTileSets)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckCharacters()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend
    
    If NumCharacters = 0 Then Exit Sub

    ReDim DDS_Character(1 To NumCharacters)
    ReDim DDSD_Character(1 To NumCharacters)
    ReDim CharacterTimer(1 To NumCharacters)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend
    
    If NumPaperdolls = 0 Then Exit Sub

    ReDim DDS_Paperdoll(1 To NumPaperdolls)
    ReDim DDSD_Paperdoll(1 To NumPaperdolls)
    ReDim PaperdollTimer(1 To NumPaperdolls)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend
    
    If NumAnimations = 0 Then Exit Sub

    ReDim DDS_Animation(1 To NumAnimations)
    ReDim DDSD_Animation(1 To NumAnimations)
    ReDim AnimationTimer(1 To NumAnimations)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Items\" & i & GFX_EXT)
        NumItems = NumItems + 1
        i = i + 1
    Wend
    
    If NumItems = 0 Then Exit Sub

    ReDim DDS_Item(1 To NumItems)
    ReDim DDSD_Item(1 To NumItems)
    ReDim ItemTimer(1 To NumItems)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectiles
Public Sub CheckProjectiles()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Projectiles\" & i & GFX_EXT)
        NumProjectiles = NumProjectiles + 1
        i = i + 1
    Wend
    
    If NumProjectiles = 0 Then Exit Sub

    ReDim DDS_Projectile(1 To NumProjectiles)
    ReDim DDSD_Projectile(1 To NumProjectiles)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Resources\" & i & GFX_EXT)
        NumResources = NumResources + 1
        i = i + 1
    Wend
    
    If NumResources = 0 Then Exit Sub

    ReDim DDS_Resource(1 To NumResources)
    ReDim DDSD_Resource(1 To NumResources)
    ReDim ResourceTimer(1 To NumResources)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpellIcons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "SpellIcons\" & i & GFX_EXT)
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend
    
    If NumSpellIcons = 0 Then Exit Sub

    ReDim DDS_SpellIcon(1 To NumSpellIcons)
    ReDim DDSD_SpellIcon(1 To NumSpellIcons)
    ReDim SpellIconTimer(1 To NumSpellIcons)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1

    While FileExist(GFX_PATH & "Faces\" & i & GFX_EXT)
        NumFaces = NumFaces + 1
        i = i + 1
    Wend
    
    If NumFaces = 0 Then Exit Sub

    ReDim DDS_Face(1 To NumFaces)
    ReDim DDSD_Face(1 To NumFaces)
    ReDim FaceTimer(1 To NumFaces)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Player(index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimInstance(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(index)), LenB(AnimInstance(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).Name = vbNullString
    Animation(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(index)), LenB(Npc(index)))
    Npc(index).Name = vbNullString
    Npc(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projec tile
Sub ClearProjectile(ByVal index As Long, ByVal PlayerProjectile As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Player(index).ProjecTile(PlayerProjectile)
        .Direction = 0
        .Pic = 0
        .TravelTime = 0
        .x = 0
        .Y = 0
        .Range = 0
        .Damage = 0
        .Speed = 0
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearProjectile", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(index)), LenB(MapItem(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(index)), LenB(MapNpc(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Name = Name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(index).Class
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Class = ClassNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Sprite = Sprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Level = Level
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(index).EXP
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDailyValue(ByVal index As Long, ByVal DailyValue As Long)
        ' If debug mode, handle error then exit out
        If Options.Debug = 1 Then On Error GoTo errorhandler
        If index > MAX_PLAYERS Then Exit Sub
        Player(index).DailyValue = DailyValue
   
        ' Error handler
        Exit Sub
errorhandler:
        HandleError "SetPlayerDailyValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
End Sub
Sub SetPlayerSecOutofCombat(ByVal index As Long, ByVal SecOutofCombat As Byte)
If index > MAX_PLAYERS Then Exit Sub
    Player(index).SecOutofCombat = SecOutofCombat
End Sub

Sub SetPlayerMiningXP(ByVal index As Long, ByVal MiningXP As Long)

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).MiningXP = MiningXP
    
End Sub

Sub SetPlayerFishingXP(ByVal index As Long, ByVal FishingXP As Long)

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).FishingXP = FishingXP
    
End Sub

Sub SetPlayerWoodcuttingXP(ByVal index As Long, ByVal WoodcuttingXP As Long)

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).WoodcuttingXP = WoodcuttingXP
    
End Sub

Sub SetPlayerSmithingXP(ByVal index As Long, ByVal SmithingXP As Long)
If index > MAX_PLAYERS Then Exit Sub
Player(index).SmithingXP = SmithingXP
End Sub
Sub SetPlayerCookingXP(ByVal index As Long, ByVal CookingXP As Long)
If index > MAX_PLAYERS Then Exit Sub
Player(index).CookingXP = CookingXP
End Sub
Sub SetPlayerFletchingXP(ByVal index As Long, ByVal FletchingXP As Long)
If index > MAX_PLAYERS Then Exit Sub
Player(index).FletchingXP = FletchingXP
End Sub
Sub SetPlayerCraftingXP(ByVal index As Long, ByVal CraftingXP As Long)
If index > MAX_PLAYERS Then Exit Sub
Player(index).CraftingXP = CraftingXP
End Sub
Sub SetPlayerPotionBrewingXP(ByVal index As Long, ByVal PotionBrewingXP As Long)
If index > MAX_PLAYERS Then Exit Sub
Player(index).PotionBrewingXP = PotionBrewingXP
End Sub


Sub SetPlayerExp(ByVal index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).EXP = EXP
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Access = Access
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).PK = PK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(index).MaxVital(Vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(index).Stat(Stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal index As Long, Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(index).Stat(Stat) = Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).POINTS = POINTS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or index <= 0 Then Exit Function
    GetPlayerMap = Player(index).Map
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Map = MapNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).x = x
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).Y
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Y = Y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Dir = Dir
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    If invslot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invslot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invslot As Long, ByVal itemnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).num = itemnum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invslot).Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).Value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).Equipment(EquipmentSlot) = InvNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearDoor(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Doors(index)), LenB(Doors(index)))
    Doors(index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearDoor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearDoors()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOORS
        Call ClearDoor(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
