Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal mapNum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Long

    ' Check for subscript out of range
    If itemnum < 1 Or itemnum > MAX_ITEMS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapNum)
    Call SpawnItemSlot(i, itemnum, ItemVal, mapNum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemnum >= 0 Then
            If itemnum <= MAX_ITEMS Then
                MapItem(mapNum, i).Num = itemnum
                MapItem(mapNum, i).Value = ItemVal
                MapItem(mapNum, i).x = x
                MapItem(mapNum, i).y = y
                Set Buffer = New clsBuffer
                Buffer.WriteLong SSpawnItem
                Buffer.WriteLong i
                Buffer.WriteLong itemnum
                Buffer.WriteLong ItemVal
                Buffer.WriteLong x
                Buffer.WriteLong y
                SendDataToMap mapNum, Buffer.ToArray()
                Set Buffer = Nothing
            End If
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapNum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(mapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapNum).Tile(x, y).Data1, 1, mapNum, x, y)
                Else
                    Call SpawnItem(Map(mapNum).Tile(x, y).Data1, Map(mapNum).Tile(x, y).Data2, mapNum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal mapNum As Long)
    Dim Buffer As clsBuffer
    Dim npcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or mapNum <= 0 Or mapNum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapNum).Npc(mapNpcNum)

    If npcNum > 0 Then
    
        MapNpc(mapNum).Npc(mapNpcNum).Num = npcNum
        MapNpc(mapNum).Npc(mapNpcNum).target = 0
        MapNpc(mapNum).Npc(mapNpcNum).targetType = 0 ' clear
        
        MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
        MapNpc(mapNum).Npc(mapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(npcNum, Vitals.MP)
        
        MapNpc(mapNum).Npc(mapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(mapNum).MaxX
            For y = 0 To Map(mapNum).MaxY
                If Map(mapNum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(mapNum).Tile(x, y).Data1 = mapNpcNum Then
                        MapNpc(mapNum).Npc(mapNpcNum).x = x
                        MapNpc(mapNum).Npc(mapNpcNum).y = y
                        MapNpc(mapNum).Npc(mapNpcNum).Dir = Map(mapNum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(mapNum).MaxX)
                y = Random(0, Map(mapNum).MaxY)
    
                If x > Map(mapNum).MaxX Then x = Map(mapNum).MaxX
                If y > Map(mapNum).MaxY Then y = Map(mapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapNum, x, y) Then
                    MapNpc(mapNum).Npc(mapNpcNum).x = x
                    MapNpc(mapNum).Npc(mapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(mapNum).MaxX
                For y = 0 To Map(mapNum).MaxY

                    If NpcTileIsOpen(mapNum, x, y) Then
                        MapNpc(mapNum).Npc(mapNpcNum).x = x
                        MapNpc(mapNum).Npc(mapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).Num
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).Dir
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals mapNum, mapNpcNum
    End If

End Sub

Public Function NpcTileIsOpen(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(mapNum).Npc(LoopI).x = x Then
                If MapNpc(mapNum).Npc(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapNum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapNum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapNum).Npc(mapNpcNum).x
    y = MapNpc(mapNum).Npc(mapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(mapNum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).Npc(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapNum).Npc(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).Npc(i).Num > 0) And (MapNpc(mapNum).Npc(i).x = MapNpc(mapNum).Npc(mapNpcNum).x) And (MapNpc(mapNum).Npc(i).y = MapNpc(mapNum).Npc(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapNum).MaxY Then
                n = Map(mapNum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).Npc(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapNum).Npc(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).Npc(i).Num > 0) And (MapNpc(mapNum).Npc(i).x = MapNpc(mapNum).Npc(mapNpcNum).x) And (MapNpc(mapNum).Npc(i).y = MapNpc(mapNum).Npc(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(mapNum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).Npc(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapNum).Npc(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).Npc(i).Num > 0) And (MapNpc(mapNum).Npc(i).x = MapNpc(mapNum).Npc(mapNpcNum).x - 1) And (MapNpc(mapNum).Npc(i).y = MapNpc(mapNum).Npc(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapNum).MaxX Then
                n = Map(mapNum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = MapNpc(mapNum).Npc(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapNum).Npc(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(mapNum).Npc(i).Num > 0) And (MapNpc(mapNum).Npc(i).x = MapNpc(mapNum).Npc(mapNpcNum).x + 1) And (MapNpc(mapNum).Npc(i).y = MapNpc(mapNum).Npc(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapNum).Tile(MapNpc(mapNum).Npc(mapNpcNum).x, MapNpc(mapNum).Npc(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapNum).Npc(mapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNpc(mapNum).Npc(mapNpcNum).y = MapNpc(mapNum).Npc(mapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapNum).Npc(mapNpcNum).y = MapNpc(mapNum).Npc(mapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapNum).Npc(mapNpcNum).x = MapNpc(mapNum).Npc(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapNum).Npc(mapNpcNum).x = MapNpc(mapNum).Npc(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap mapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal mapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNum <= 0 Or mapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapNum).Npc(mapNpcNum).Dir = Dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapNum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = mapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapNum As Long)
    Dim y As Long
    Dim x As Long
    TempTile(mapNum).DoorTimer = 0
    ReDim TempTile(mapNum).DoorOpen(0 To Map(mapNum).MaxX, 0 To Map(mapNum).MaxY)

    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY
            TempTile(mapNum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapNum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY

            If Map(mapNum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapNum).ResourceData(0 To Resource_Count)
                ResourceCache(mapNum).ResourceData(Resource_Count).x = x
                ResourceCache(mapNum).ResourceData(Resource_Count).y = y
                ResourceCache(mapNum).ResourceData(Resource_Count).cur_health = Resource(Map(mapNum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        PlayerMsg index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function
