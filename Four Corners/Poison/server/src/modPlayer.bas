Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal index As Long)
    If Not IsPlaying(index) Then
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendDoors(index)
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    SendEXP index
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
    
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' leave party.
        Party_PlayerLeave index

        ' save and clear data.
        Call SavePlayer(index)
        Call SaveBank(index)
        Call ClearBank(index)

        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(index)
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    Dim Legs As Long
    Dim Shield As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, Helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If
    
    If Legs > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If
    
    If Shield > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Shield).Data2
    End If
    
End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If X > Map(mapnum).MaxX Then X = Map(mapnum).MaxX
    If Y > Map(mapnum).MaxY Then Y = Map(mapnum).MaxY
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    
    ' if same map then just send their co-ordinates
    If mapnum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
    End If
    
    ' clear target
    TempPlayer(index).target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If OldMap <> mapnum Then
        Call SendLeaveMap(index, OldMap)
    End If

    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = mapnum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    TempPlayer(index).GettingMap = YES
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong mapnum
    Buffer.WriteLong Map(mapnum).Revision
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, mapnum As Long
    Dim X As Long, Y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, amount As Long
    Dim DoorNum As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Moved = NO
    mapnum = GetPlayerMap(index)
    
Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "You need the right kind of key to open this door. (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "You need to activate a switch to open this door. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpMap
                                                    X = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpX
                                                    Y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, X, Y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = YES) Then
                                    Call SetPlayerY(index, GetPlayerY(index) - 1)
                                    SendPlayerMove index, movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Up).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(mapnum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "You need the right kind of key to open this door. (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "You need to activate a switch to open this door. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpMap
                                                    X = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpX
                                                    Y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, X, Y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = YES) Then
                                    Call SetPlayerY(index, GetPlayerY(index) + 1)
                                    SendPlayerMove index, movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Down).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "You need the right kind of key to open this door. (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "You need to activate a switch to open this door. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpMap
                                                    X = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpX
                                                    Y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, X, Y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = YES) Then
                                    Call SetPlayerX(index, GetPlayerX(index) - 1)
                                    SendPlayerMove index, movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, NewMapX, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
            
        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(mapnum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
                                    If Player(index).PlayerDoors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).state = 0 Then
                                            
                                            
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).DoorType = 0 Then
                                                If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).UnlockType = 0 Then
                                                    PlayerMsg index, "You need the right kind of key to open this door. (" & Trim$(Item(Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).key).Name) & ")", BrightRed
                                                ElseIf Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).UnlockType = 1 Then
                                                    PlayerMsg index, "You need to activate a switch to open this door. ", BrightRed
                                                Else
                                                    mapnum = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpMap
                                                    X = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpX
                                                    Y = Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).WarpY
                                                    Call PlayerWarp(index, mapnum, X, Y)
                                                    Moved = YES
                                                End If
                                                End If
                                            
                                            PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                            Exit Sub
                                    Else
                                            If Doors(Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Data1).DoorType = 1 Then
                                                PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
                                                Exit Sub
                                            End If
                                    End If
                                End If
                                
                                ' Check to see if the tile is a key and if it is check if its opened
                                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = YES) Then
                                    Call SetPlayerX(index, GetPlayerX(index) + 1)
                                    SendPlayerMove index, movement, sendToSelf
                                    Moved = YES
                                End If
                            
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Right).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
            
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            mapnum = .Data1
            X = .Data2
            Y = .Data3
            Call PlayerWarp(index, mapnum, X, Y)
            Moved = YES
        End If
    
         ' Check to see if the tile is a door tile
        If .Type = TILE_TYPE_DOOR Then
            DoorNum = .Data1
            
            If Player(index).PlayerDoors(DoorNum).state = 1 Then
                mapnum = Doors(DoorNum).WarpMap
                X = Doors(DoorNum).WarpX
                Y = Doors(DoorNum).WarpY
                Call PlayerWarp(index, mapnum, X, Y)
                Moved = YES
            End If
            
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            X = .Data1
            Y = .Data2
    
            If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                SendMapKey index, X, Y, 1
                Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            X = .Data1
            If X > 0 Then ' shop exists?
                If Len(Trim$(Shop(X).Name)) > 0 Then ' name exists?
                    SendOpenShop index, X
                    TempPlayer(index).InShop = X ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            amount = .Data2
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(index), "+" & amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + amount
                PlayerMsg index, "You feel rejuvinating forces flowing through your body.", BrightGreen
                Call SendVital(index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, HP) - amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "You're killed by a trap.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - amount
                PlayerMsg index, "You're injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        'Checks for sprite tile
        If .Type = TILE_TYPE_SPRITE Then
            amount = .Data1
            Call SetPlayerSprite(index, amount)
            Call SendPlayerData(index)
            Moved = YES
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove index, MOVING_WALKING, GetPlayerDir(index)
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If

End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim itemnum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(index, i)

        If itemnum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(itemnum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(itemnum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(itemnum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Legs
                
                    If Item(itemnum).Type <> ITEM_TYPE_LEGS Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If Item(itemnum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = itemnum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = itemnum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal itemnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = itemnum Then
            If Item(itemnum).Type = ITEM_TYPE_CURRENCY Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim itemnum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    itemnum = GetPlayerInvItemNum(index, invSlot)

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal index As Long, ByVal itemnum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or itemnum <= 0 Or itemnum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, itemnum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, itemnum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        If sendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim mapnum As Long
    Dim Msg As String

    If Not IsPlaying(index) Then Exit Sub
    mapnum = GetPlayerMap(index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, i).Num > 0) And (MapItem(mapnum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).X = GetPlayerX(index)) Then
                    If (MapItem(mapnum, i).Y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(mapnum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(index, n, MapItem(mapnum, i).Num)
    
                            If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(mapnum, i).Value)
                                Msg = MapItem(mapnum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                            End If
    
                            ' Erase item from the map
                            ClearMapItem i, mapnum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim mapnum As Long

    mapnum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(mapnum, mapItemNum).playerName = vbNullString Or MapItem(mapnum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(index))

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), i).X = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).Y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = GetTickCount + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).canDespawn = True
                MapItem(GetPlayerMap(index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(index, invNum)).Type = ITEM_TYPE_CURRENCY Then

                    ' Check if its more then they have and if so drop it all
                    If amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                    Else
                        MapItem(GetPlayerMap(index), i).Value = amount
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).canDespawn)
            Else
                Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 3)
        Call SetPlayerExp(index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            PlayerMsg index, "You have gained " & level_count & "level!", Brown
            If Player(index).Level >= 100 Then
                GlobalMsg GetPlayerName(index) & " is now level " & Player(index).Level & "!", BrightGreen
            End If
        Else
            'plural
            PlayerMsg index, "You have gained " & level_count & "levels!", Brown
            If Player(index).Level >= 100 Then
                GlobalMsg GetPlayerName(index) & " is now level " & Player(index).Level & "!", BrightGreen
            End If
        End If
        SendEXP index
        SendPlayerData index
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim$(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    If IsPlaying(index) = False Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(index) + 1) ^ 3 - (6 * (GetPlayerLevel(index) + 1) ^ 2) + 17 * (GetPlayerLevel(index) + 1) - 12)
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    Player(index).exp = exp
End Sub

Function GetPlayerSecOutofCombat(ByVal index As Long) As Byte
If index > MAX_PLAYERS Then Exit Function
    GetPlayerSecOutofCombat = Player(index).SecOutofCombat
End Function

Function GetPlayerMiningXP(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerMiningXP = Player(index).MiningXP
End Function

Function GetPlayerWoodcuttingXP(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerWoodcuttingXP = Player(index).WoodcuttingXP
End Function

Function GetPlayerFishingXP(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerFishingXP = Player(index).FishingXP
End Function

Function GetPlayerDailyValue(ByVal index As Long) As Long
If index > MAX_PLAYERS Then Exit Function
GetPlayerDailyValue = Player(index).DailyValue
End Function

Function GetPlayerSmithingXP(ByVal index As Long) As Long
If index > MAX_PLAYERS Then Exit Function
GetPlayerSmithingXP = Player(index).SmithingXP
End Function
Function GetPlayerCookingXP(ByVal index As Long) As Long
If index > MAX_PLAYERS Then Exit Function
GetPlayerCookingXP = Player(index).CookingXP
End Function
Function GetPlayerFletchingXP(ByVal index As Long) As Long
If index > MAX_PLAYERS Then Exit Function
GetPlayerFletchingXP = Player(index).FletchingXP
End Function
Function GetPlayerCraftingXP(ByVal index As Long) As Long
If index > MAX_PLAYERS Then Exit Function
GetPlayerCraftingXP = Player(index).CraftingXP
End Function
Function GetPlayerPotionBrewingXP(ByVal index As Long) As Long
If index > MAX_PLAYERS Then Exit Function
GetPlayerPotionBrewingXP = Player(index).PotionBrewingXP
End Function

Function GetPlayerAccess(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    If IsPlaying(index) = False Then Exit Sub
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Player(index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal Stat As Stats) As Long
    Dim X As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
    
    X = Player(index).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(index).Equipment(i) > 0 Then
            If Item(Player(index).Equipment(i)).Add_Stat(Stat) > 0 Then
                X = X + Item(Player(index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = X
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal Stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)

    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(index).Map = mapnum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).X
End Function

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)
    Player(index).X = X
End Sub

Function GetPlayerY(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).Y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)
    Player(index).Y = Y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Inv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal itemnum As Long)
    Player(index).Inv(invSlot).Num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(invSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal spellslot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellslot As Long, ByVal SpellNum As Long)
    Player(index).Spell(spellslot) = SpellNum
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long

    If index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = invNum
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)

'Drop inventory items
    For i = 1 To MAX_INV
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
    Next

    For i = 1 To MAX_INV
    PlayerMapDropItem index, i, GetPlayerInvItemValue(index, i)
    Next


    'Send all equiped items to the inventory to be dumped.
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(index, i) > 0 Then
            PlayerMapDropItem index, GetPlayerEquipment(index, i), 0
        End If
       
        'Send Weapon
        GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
        SetPlayerEquipment index, 0, Weapon
        'Send Armor
        GiveInvItem index, GetPlayerEquipment(index, Armor), 0
        SetPlayerEquipment index, 0, Armor
        'Send Shield
        GiveInvItem index, GetPlayerEquipment(index, Shield), 0
        SetPlayerEquipment index, 0, Shield
        'Send Helmet
        GiveInvItem index, GetPlayerEquipment(index, Helmet), 0
        SetPlayerEquipment index, 0, Helmet
        'Send Legs
        GiveInvItem index, GetPlayerEquipment(index, Legs), 0
        SetPlayerEquipment index, 0, Legs
        
               
    Next
    
    For i = 1 To MAX_INV
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
        If Item(i).BindType > 0 Then
            TakeInvItem index, i, 0
        End If
    Next

    'Drop *equipped* inventory items
    For i = 1 To MAX_INV
        PlayerMapDropItem index, i, 0
    Next
    
    Call PlayerMsg(index, "You have died.", BrightRed)

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
        Call SendPlayerData(index)
    End If

End Sub

Sub CheckResource(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim damage As Long
    Dim n As Long
    
    If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(index)).Tile(X, Y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count

            If ResourceCache(GetPlayerMap(index)).ResourceData(i).X = X Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).Y = Y Then
                    Resource_num = i
                End If
            End If
        
        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Data3 <> Resource(Resource_index).ToolRequired Then
                    If Resource(Resource_index).ToolRequired > 0 Then
                        PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
                        Exit Sub
                    End If
                End If

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If
                    
                    ' check to see if experienced
                    If Resource(Resource_index).FReq > 0 Then
                        If Player(index).FishingXP < Resource(Resource_index).FReq Then
                            PlayerMsg index, "You need " & Resource(Resource_index).FReq - Player(index).FishingXP & " more experience in fishing to access this resource.", Red
                            Exit Sub
                        End If
                    End If
                    
                    If Resource(Resource_index).MReq > 0 Then
                        If Player(index).MiningXP < Resource(Resource_index).MReq Then
                            PlayerMsg index, "You need " & Resource(Resource_index).FReq - Player(index).MiningXP & " more experience in mining to access this resource.", Red
                            Exit Sub
                        End If
                    End If
                    
                    If Resource(Resource_index).WcReq > 0 Then
                        If Player(index).WoodcuttingXP < Resource(Resource_index).WcReq Then
                            PlayerMsg index, "You need " & Resource(Resource_index).WcReq - Player(index).WoodcuttingXP & " more experience in woodcutting to access this resource.", Red
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).X
                        rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).Y
                        
                        damage = Item(GetPlayerEquipment(index, Weapon)).Data2
                    
                        ' check if damage is more than health
                        If damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - damage <= 0 Then
                                SendActionMsg GetPlayerMap(index), "-" & ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(index), Resource_num
                                ' send message if it exists
                                If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                    
    For n = 1 To MAX_RES_DROPS
        If Resource(Resource_index).DropItem(n) = 0 Then Exit For
        If Rnd <= Resource(Resource_index).DropChance(n) Then
            Call SpawnItem(Resource(Resource_index).DropItem(n), Resource(Resource_index).DropItemValue(n), GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(index, "You notice something falls to the floor.", Magenta)
        End If
    Next
                                
                                End If
                                ' carry on
                                GiveInvItem index, Resource(Resource_index).ItemReward, 1
                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                                
                                If Resource(Resource_index).RewardXP > 0 Then
                                    If Resource(Resource_index).FXP = True Then
                                        Player(index).FishingXP = Player(index).FishingXP + Resource(Resource_index).RewardXP
                                    End If
                                    If Resource(Resource_index).MXP = True Then
                                        Player(index).MiningXP = Player(index).MiningXP + Resource(Resource_index).RewardXP
                                    End If
                                    If Resource(Resource_index).WcXP = True Then
                                        Player(index).WoodcuttingXP = Player(index).WoodcuttingXP + Resource(Resource_index).RewardXP
                                    End If
                                    SendPlayerData index
                                End If
                                
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - damage
                                SendActionMsg GetPlayerMap(index), "-" & damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                            End If
                            ' send the sound
                            SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                        End If
                    End If

            Else
                PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(index).Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
    Bank(index).Item(BankSlot).Num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal amount As Long)
Dim BankSlot

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, invSlot)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal amount As Long)
Dim invSlot

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Public Sub KillPlayer(ByVal index As Long)

    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, X As Long, Y As Long, itemnum As Long, b As Long, j As Long

    For j = 1 To MAX_INV
    Next
    
    b = FindOpenInvSlot(index, j)

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, invNum)).Data2
        itemnum = GetPlayerInvItemNum(index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(itemnum).Type
            Case ITEM_TYPE_ARMOR
            
                ' skill requirements
                If GetPlayerWoodcuttingXP(index) < Item(itemnum).WcXP Then
                    PlayerMsg index, "You need " & Item(itemnum).WcXP - Player(index).WoodcuttingXP & " more experience in woodcutting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFishingXP(index) < Item(itemnum).FXP Then
                    PlayerMsg index, "You need " & Item(itemnum).FXP - Player(index).FishingXP & " more experience in fishing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerMiningXP(index) < Item(itemnum).MXP Then
                    PlayerMsg index, "You need " & Item(itemnum).MXP - Player(index).MiningXP & " more experience in mining in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerSmithingXP(index) < Item(itemnum).EqSmXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqSmXP - Player(index).SmithingXP & " more experience in smithing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCookingXP(index) < Item(itemnum).EqCoXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCoXP - Player(index).CookingXP & " more experience in cooking in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFletchingXP(index) < Item(itemnum).EqFlXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqFlXP - Player(index).FletchingXP & " more experience in fletching in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCraftingXP(index) < Item(itemnum).EqCrXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCrXP - Player(index).CraftingXP & " more experience in crafting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerPotionBrewingXP(index) < Item(itemnum).EqPBXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqPBXP - Player(index).CraftingXP & " more experience in potion brewing in order to wear this.", Red
                    Exit Sub
                End If
                    
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                If GetPlayerEquipment(index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(index, Armor)
                End If

                SetPlayerEquipment index, itemnum, Armor
                TakeInvItem index, itemnum, 0

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_WEAPON
            
                ' skill requirements
                If GetPlayerWoodcuttingXP(index) < Item(itemnum).WcXP Then
                    PlayerMsg index, "You need " & Item(itemnum).WcXP - Player(index).WoodcuttingXP & " more experience in woodcutting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFishingXP(index) < Item(itemnum).FXP Then
                    PlayerMsg index, "You need " & Item(itemnum).FXP - Player(index).FishingXP & " more experience in fishing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerMiningXP(index) < Item(itemnum).MXP Then
                    PlayerMsg index, "You need " & Item(itemnum).MXP - Player(index).MiningXP & " more experience in mining in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerSmithingXP(index) < Item(itemnum).EqSmXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqSmXP - Player(index).SmithingXP & " more experience in smithing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCookingXP(index) < Item(itemnum).EqCoXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCoXP - Player(index).CookingXP & " more experience in cooking in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFletchingXP(index) < Item(itemnum).EqFlXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqFlXP - Player(index).FletchingXP & " more experience in fletching in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCraftingXP(index) < Item(itemnum).EqCrXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCrXP - Player(index).CraftingXP & " more experience in crafting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerPotionBrewingXP(index) < Item(itemnum).EqPBXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqPBXP - Player(index).CraftingXP & " more experience in potion brewing in order to wear this.", Red
                    Exit Sub
                End If
                
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                If Item(itemnum).istwohander = True Then
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        If GetPlayerEquipment(index, Weapon) > 0 Then
                            If b < 1 Then
                                Call PlayerMsg(index, "You have no room in your inventory.", BrightRed)
                                Exit Sub
                                
                                End If
                            End If
                        End If
                    End If

                If GetPlayerEquipment(index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(index, Weapon)
                End If

                SetPlayerEquipment index, itemnum, Weapon
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                           
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                If Item(itemnum).istwohander = True Then
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        GiveInvItem index, GetPlayerEquipment(index, Shield), 0
                        SetPlayerEquipment index, 0, Shield
                    End If
                    'two handed sprite
                    If Player(index).Sprite = 1 Or Player(index).Sprite = 2 Then
                        Player(index).Sprite = 3
                    End If
                    If Player(index).Sprite = 4 Or Player(index).Sprite = 5 Then
                        Player(index).Sprite = 6
                    End If
                    If Player(index).Sprite = 7 Or Player(index).Sprite = 8 Then
                        Player(index).Sprite = 9
                    End If
                    If Player(index).Sprite = 10 Or Player(index).Sprite = 11 Then
                        Player(index).Sprite = 12
                    End If
                    If Player(index).Sprite = 13 Or Player(index).Sprite = 14 Then
                        Player(index).Sprite = 15
                    End If
                    If Player(index).Sprite = 16 Or Player(index).Sprite = 17 Then
                        Player(index).Sprite = 18
                    End If
                Else
                    
                    'single handed
                    
                    If GetPlayerEquipment(index, Shield) > 0 Then
                        If Player(index).Sprite = 1 Or Player(index).Sprite = 3 Then
                            Player(index).Sprite = 2
                        End If
                        If Player(index).Sprite = 4 Or Player(index).Sprite = 6 Then
                            Player(index).Sprite = 5
                        End If
                        If Player(index).Sprite = 7 Or Player(index).Sprite = 9 Then
                            Player(index).Sprite = 8
                        End If
                        If Player(index).Sprite = 10 Or Player(index).Sprite = 12 Then
                            Player(index).Sprite = 11
                        End If
                        If Player(index).Sprite = 13 Or Player(index).Sprite = 15 Then
                            Player(index).Sprite = 14
                        End If
                        If Player(index).Sprite = 16 Or Player(index).Sprite = 18 Then
                            Player(index).Sprite = 17
                        End If
                    End If
                    
                    If GetPlayerEquipment(index, Shield) = 0 Then
                        If Player(index).Sprite = 2 Or Player(index).Sprite = 3 Then
                            Player(index).Sprite = 1
                        End If
                        If Player(index).Sprite = 5 Or Player(index).Sprite = 6 Then
                            Player(index).Sprite = 4
                        End If
                        If Player(index).Sprite = 8 Or Player(index).Sprite = 9 Then
                            Player(index).Sprite = 7
                        End If
                        If Player(index).Sprite = 11 Or Player(index).Sprite = 12 Then
                            Player(index).Sprite = 10
                        End If
                        If Player(index).Sprite = 14 Or Player(index).Sprite = 15 Then
                            Player(index).Sprite = 13
                        End If
                        If Player(index).Sprite = 17 Or Player(index).Sprite = 18 Then
                            Player(index).Sprite = 16
                        End If
                    End If
                End If
     
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendPlayerData(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_HELMET
            
                ' skill requirements
                If GetPlayerWoodcuttingXP(index) < Item(itemnum).WcXP Then
                    PlayerMsg index, "You need " & Item(itemnum).WcXP - Player(index).WoodcuttingXP & " more experience in woodcutting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFishingXP(index) < Item(itemnum).FXP Then
                    PlayerMsg index, "You need " & Item(itemnum).FXP - Player(index).FishingXP & " more experience in fishing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerMiningXP(index) < Item(itemnum).MXP Then
                    PlayerMsg index, "You need " & Item(itemnum).MXP - Player(index).MiningXP & " more experience in mining in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerSmithingXP(index) < Item(itemnum).EqSmXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqSmXP - Player(index).SmithingXP & " more experience in smithing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCookingXP(index) < Item(itemnum).EqCoXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCoXP - Player(index).CookingXP & " more experience in cooking in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFletchingXP(index) < Item(itemnum).EqFlXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqFlXP - Player(index).FletchingXP & " more experience in fletching in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCraftingXP(index) < Item(itemnum).EqCrXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCrXP - Player(index).CraftingXP & " more experience in crafting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerPotionBrewingXP(index) < Item(itemnum).EqPBXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqPBXP - Player(index).CraftingXP & " more experience in potion brewing in order to wear this.", Red
                    Exit Sub
                End If
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(index, Helmet)
                End If

                SetPlayerEquipment index, itemnum, Helmet
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            
            Case ITEM_TYPE_LEGS
            
                ' skill requirements
                If GetPlayerWoodcuttingXP(index) < Item(itemnum).WcXP Then
                    PlayerMsg index, "You need " & Item(itemnum).WcXP - Player(index).WoodcuttingXP & " more experience in woodcutting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFishingXP(index) < Item(itemnum).FXP Then
                    PlayerMsg index, "You need " & Item(itemnum).FXP - Player(index).FishingXP & " more experience in fishing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerMiningXP(index) < Item(itemnum).MXP Then
                    PlayerMsg index, "You need " & Item(itemnum).MXP - Player(index).MiningXP & " more experience in mining in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerSmithingXP(index) < Item(itemnum).EqSmXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqSmXP - Player(index).SmithingXP & " more experience in smithing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCookingXP(index) < Item(itemnum).EqCoXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCoXP - Player(index).CookingXP & " more experience in cooking in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFletchingXP(index) < Item(itemnum).EqFlXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqFlXP - Player(index).FletchingXP & " more experience in fletching in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCraftingXP(index) < Item(itemnum).EqCrXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCrXP - Player(index).CraftingXP & " more experience in crafting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerPotionBrewingXP(index) < Item(itemnum).EqPBXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqPBXP - Player(index).CraftingXP & " more experience in potion brewing in order to wear this.", Red
                    Exit Sub
                End If
                
                 ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Legs) > 0 Then
                    tempItem = GetPlayerEquipment(index, Legs)
                End If

                SetPlayerEquipment index, itemnum, Legs
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            
            Case ITEM_TYPE_SHIELD
            
                ' skill requirements
                If GetPlayerWoodcuttingXP(index) < Item(itemnum).WcXP Then
                    PlayerMsg index, "You need " & Item(itemnum).WcXP - Player(index).WoodcuttingXP & " more experience in woodcutting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFishingXP(index) < Item(itemnum).FXP Then
                    PlayerMsg index, "You need " & Item(itemnum).FXP - Player(index).FishingXP & " more experience in fishing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerMiningXP(index) < Item(itemnum).MXP Then
                    PlayerMsg index, "You need " & Item(itemnum).MXP - Player(index).MiningXP & " more experience in mining in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerSmithingXP(index) < Item(itemnum).EqSmXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqSmXP - Player(index).SmithingXP & " more experience in smithing in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCookingXP(index) < Item(itemnum).EqCoXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCoXP - Player(index).CookingXP & " more experience in cooking in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFletchingXP(index) < Item(itemnum).EqFlXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqFlXP - Player(index).FletchingXP & " more experience in fletching in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCraftingXP(index) < Item(itemnum).EqCrXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCrXP - Player(index).CraftingXP & " more experience in crafting in order to wear this.", Red
                    Exit Sub
                End If
                
                If GetPlayerPotionBrewingXP(index) < Item(itemnum).EqPBXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqPBXP - Player(index).CraftingXP & " more experience in potion brewing in order to wear this.", Red
                    Exit Sub
                End If
            
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(index, Shield)
                End If

                SetPlayerEquipment index, itemnum, Shield
                TakeInvItem index, itemnum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                If GetPlayerEquipment(index, Weapon) > 0 Then
                    If Item(GetPlayerEquipment(index, Weapon)).istwohander = True Then
                        GiveInvItem index, GetPlayerEquipment(index, Weapon), 0
                        SetPlayerEquipment index, 0, Weapon
                    End If
                End If
                
                If Player(index).Sprite = 1 Or Player(index).Sprite = 3 Then
                    Player(index).Sprite = 2
                End If
                If Player(index).Sprite = 4 Or Player(index).Sprite = 6 Then
                    Player(index).Sprite = 5
                End If
                If Player(index).Sprite = 7 Or Player(index).Sprite = 9 Then
                    Player(index).Sprite = 8
                End If
                If Player(index).Sprite = 10 Or Player(index).Sprite = 12 Then
                    Player(index).Sprite = 11
                End If
                If Player(index).Sprite = 13 Or Player(index).Sprite = 15 Then
                    Player(index).Sprite = 14
                End If
                If Player(index).Sprite = 16 Or Player(index).Sprite = 18 Then
                    Player(index).Sprite = 17
                End If
                
                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                Call SendPlayerData(index)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            ' consumable
            Case ITEM_TYPE_CONSUME
            
            
            
            
            If GetPlayerWoodcuttingXP(index) < Item(itemnum).WcXP Then
                    PlayerMsg index, "You need " & Item(itemnum).WcXP - Player(index).WoodcuttingXP & " more experience in woodcutting in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFishingXP(index) < Item(itemnum).FXP Then
                    PlayerMsg index, "You need " & Item(itemnum).FXP - Player(index).FishingXP & " more experience in fishing in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerMiningXP(index) < Item(itemnum).MXP Then
                    PlayerMsg index, "You need " & Item(itemnum).MXP - Player(index).MiningXP & " more experience in mining in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerSmithingXP(index) < Item(itemnum).EqSmXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqSmXP - Player(index).SmithingXP & " more experience in smithing in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCookingXP(index) < Item(itemnum).EqCoXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCoXP - Player(index).CookingXP & " more experience in cooking in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFletchingXP(index) < Item(itemnum).EqFlXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqFlXP - Player(index).FletchingXP & " more experience in fletching in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCraftingXP(index) < Item(itemnum).EqCrXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCrXP - Player(index).CraftingXP & " more experience in crafting in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerPotionBrewingXP(index) < Item(itemnum).EqPBXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqPBXP - Player(index).CraftingXP & " more experience in potion brewing in order to use this.", Red
                    Exit Sub
                End If
                
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(itemnum).AddHP > 0 Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Item(itemnum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(itemnum).AddMP > 0 Then
                    Player(index).Vital(Vitals.MP) = Player(index).Vital(Vitals.MP) + Item(itemnum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(itemnum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(itemnum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(itemnum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(invNum).Num, 0)
                
                If Item(itemnum).ConsumeItem <> 0 Then
                    GiveInvItem index, Item(itemnum).ConsumeItem, 1
                End If
                    
                    
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_KEY
            
            
            
            
            If GetPlayerWoodcuttingXP(index) < Item(itemnum).WcXP Then
                    PlayerMsg index, "You need " & Item(itemnum).WcXP - Player(index).WoodcuttingXP & " more experience in woodcutting in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFishingXP(index) < Item(itemnum).FXP Then
                    PlayerMsg index, "You need " & Item(itemnum).FXP - Player(index).FishingXP & " more experience in fishing in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerMiningXP(index) < Item(itemnum).MXP Then
                    PlayerMsg index, "You need " & Item(itemnum).MXP - Player(index).MiningXP & " more experience in mining in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerSmithingXP(index) < Item(itemnum).EqSmXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqSmXP - Player(index).SmithingXP & " more experience in smithing in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCookingXP(index) < Item(itemnum).EqCoXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCoXP - Player(index).CookingXP & " more experience in cooking in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFletchingXP(index) < Item(itemnum).EqFlXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqFlXP - Player(index).FletchingXP & " more experience in fletching in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCraftingXP(index) < Item(itemnum).EqCrXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCrXP - Player(index).CraftingXP & " more experience in crafting in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerPotionBrewingXP(index) < Item(itemnum).EqPBXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqPBXP - Player(index).CraftingXP & " more experience in potion brewing in order to use this.", Red
                    Exit Sub
                End If
                
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(index)
                    Case DIR_UP

                        If GetPlayerY(index) > 0 Then
                            X = GetPlayerX(index)
                            Y = GetPlayerY(index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(index) < Map(GetPlayerMap(index)).MaxY Then
                            X = GetPlayerX(index)
                            Y = GetPlayerY(index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(index) > 0 Then
                            X = GetPlayerX(index) - 1
                            Y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(index) < Map(GetPlayerMap(index)).MaxX Then
                            X = GetPlayerX(index) + 1
                            Y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If itemnum = Map(GetPlayerMap(index)).Tile(X, Y).Data1 Then
                        TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                        TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        SendMapKey index, X, Y, 1
                        Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, X, Y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).Tile(X, Y).Data2 = 1 Then
                            Call TakeInvItem(index, itemnum, 0)
                            Call PlayerMsg(index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
            Case ITEM_TYPE_SPELL
            
            
            
            
                If GetPlayerWoodcuttingXP(index) < Item(itemnum).WcXP Then
                    PlayerMsg index, "You need " & Item(itemnum).WcXP - Player(index).WoodcuttingXP & " more experience in woodcutting in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFishingXP(index) < Item(itemnum).FXP Then
                    PlayerMsg index, "You need " & Item(itemnum).FXP - Player(index).FishingXP & " more experience in fishing in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerMiningXP(index) < Item(itemnum).MXP Then
                    PlayerMsg index, "You need " & Item(itemnum).MXP - Player(index).MiningXP & " more experience in mining in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerSmithingXP(index) < Item(itemnum).EqSmXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqSmXP - Player(index).SmithingXP & " more experience in smithing in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCookingXP(index) < Item(itemnum).EqCoXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCoXP - Player(index).CookingXP & " more experience in cooking in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerFletchingXP(index) < Item(itemnum).EqFlXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqFlXP - Player(index).FletchingXP & " more experience in fletching in order to use this.", Red
                    Exit Sub
                End If
                
                If GetPlayerCraftingXP(index) < Item(itemnum).EqCrXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqCrXP - Player(index).CraftingXP & " more experience in crafting in order to use this.", Red
                    Exit Sub
                End If
            
                If GetPlayerPotionBrewingXP(index) < Item(itemnum).EqPBXP Then
                    PlayerMsg index, "You need " & Item(itemnum).EqPBXP - Player(index).CraftingXP & " more experience in potion brewing in order to use this.", Red
                    Exit Sub
                End If
                
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(itemnum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(itemnum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(itemnum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(itemnum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(itemnum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(itemnum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Call SetPlayerSpell(index, i, n)
                                    Call SendAnimation(GetPlayerMap(index), Item(itemnum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, itemnum, 0)
                                    Call PlayerMsg(index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                Else
                                    Call PlayerMsg(index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, itemnum
        End Select
    End If
End Sub

Sub CheckDoor(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    Dim Door_num As Long
    Dim i As Long
    Dim n As Long
    Dim key As Long
    Dim tmpIndex As Long
    
    If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
        Door_num = Map(GetPlayerMap(index)).Tile(X, Y).Data1



        If Door_num > 0 Then
            If Doors(Door_num).DoorType = 0 Then
                If Player(index).PlayerDoors(Door_num).state = 0 Then
                    If Doors(Door_num).UnlockType = 0 Then
                        For i = 1 To MAX_INV
                            key = GetPlayerInvItemNum(index, i)
                            If Doors(Door_num).key = key Then
                                TakeInvItem index, key, 1
                                If TempPlayer(index).inParty > 0 Then
                                    For n = 1 To MAX_PARTY_MEMBERS
                                        tmpIndex = Party(TempPlayer(index).inParty).Member(n)
                                        If tmpIndex > 0 Then
                                            Player(tmpIndex).PlayerDoors(Door_num).state = 1
                                            SendPlayerData (tmpIndex)
                                            If index <> tmpIndex Then
                                                PlayerMsg tmpIndex, "A member of your party unlocked a door.", BrightBlue
                                            Else
                                                PlayerMsg tmpIndex, "You used a key to unlock the door.", BrightBlue
                                            End If
                                        End If
                                    Next
                                    
                                Else
                                    Player(index).PlayerDoors(Door_num).state = 1
                                    PlayerMsg index, "You used a key to unlock the door.", BrightBlue
                                    SendPlayerData (index)
                                End If
                                Exit Sub
                            End If
                        Next
                        PlayerMsg index, "You do not have the right key to unlock the door.", BrightBlue
                    ElseIf Doors(Door_num).UnlockType = 1 Then
                        If Doors(Door_num).state = 0 Then
                            PlayerMsg index, "You have not fliped the right switch to unlock this door.", BrightBlue
                        End If
                    ElseIf Doors(Door_num).UnlockType = 2 Then
                        PlayerMsg index, "This door has no lock.", BrightBlue
                    End If
                    
                Else
                    PlayerMsg index, "This door is already unlocked.", BrightBlue
                End If
            ElseIf Doors(Door_num).DoorType = 1 Then
                If Player(index).PlayerDoors(Door_num).state = 0 Then
                                If TempPlayer(index).inParty > 0 Then
                                    For n = 1 To MAX_PARTY_MEMBERS
                                        tmpIndex = Party(TempPlayer(index).inParty).Member(n)
                                        If tmpIndex > 0 Then
                                            Player(tmpIndex).PlayerDoors(Door_num).state = 1
                                            Player(tmpIndex).PlayerDoors(Doors(Door_num).Switch).state = 1
                                            SendPlayerData (tmpIndex)
                                            If index <> tmpIndex Then
                                                PlayerMsg tmpIndex, "A member of your party filped a switch on and unlocked a door.", BrightBlue
                                            Else
                                                PlayerMsg tmpIndex, "You filp the switch on and unlocked a door.", BrightBlue
                                            End If
                                        End If
                                    Next
                                Else
                                    Player(index).PlayerDoors(Door_num).state = 1
                                    Player(index).PlayerDoors(Doors(Door_num).Switch).state = 1
                                    PlayerMsg index, "You filp the switch on and unlocked a door.", BrightBlue
                                    SendPlayerData (index)
                                End If
                    
                Else
                                If TempPlayer(index).inParty > 0 Then
                                    For n = 1 To MAX_PARTY_MEMBERS
                                        tmpIndex = Party(TempPlayer(index).inParty).Member(n)
                                        If tmpIndex > 0 Then
                                            Player(tmpIndex).PlayerDoors(Door_num).state = 0
                                            Player(tmpIndex).PlayerDoors(Doors(Door_num).Switch).state = 0
                                            SendPlayerData (tmpIndex)
                                            If index <> tmpIndex Then
                                                PlayerMsg tmpIndex, "A member of your party filped a switch off and locked a door.", BrightBlue
                                            Else
                                                PlayerMsg tmpIndex, "You filp the switch off and locked a door.", BrightBlue
                                            End If
                                        End If
                                    Next
                                Else
                                    Player(index).PlayerDoors(Door_num).state = 0
                                    Player(index).PlayerDoors(Doors(Door_num).Switch).state = 0
                                    PlayerMsg index, "You filp the switch off and locked a door.", BrightBlue
                                    SendPlayerData (index)
                                End If
                End If
            End If
        End If
    End If
End Sub
