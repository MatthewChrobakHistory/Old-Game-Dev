Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, X As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For X = 1 To MAX_DOTS
                        HandleDoT_Player i, X
                        HandleHoT_Player i, X
                    Next
                End If
            Next
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).state > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            For i = 1 To MAX_PLAYERS
                If Player(i).SecOutofCombat > 0 Then
                    Player(i).SecOutofCombat = Player(i).SecOutofCombat - 1
                    SendPlayerData (i)
                End If
            Next
            tmr1000 = GetTickCount + 1000
        End If
        
        ' projectiles
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                For X = 1 To MAX_PLAYER_PROJECTILES
                    If TempPlayer(i).ProjecTile(X).Pic > 0 Then
                        ' handle the projec tile
                        HandleProjecTile i, X
                    End If
                Next
            End If
        Next

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If

        ' Checks to save players every 5 seconds - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 5000
        End If

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
    Dim X As Long
    Dim Y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For Y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(Y) Then

            ' Clear out unnecessary junk
            For X = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(X, Y)
            Next

            ' Spawn the items
            Call SpawnMapItems(Y)
            Call SendMapItemsToAll(Y)
        End If

        DoEvents
    Next

End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, X As Long, mapnum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, damage As Long, DistanceX As Long, DistanceY As Long, npcNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean

    For mapnum = 1 To MAX_MAPS
    
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(mapnum, i).Num > 0 Then
                If MapItem(mapnum, i).playerName <> vbNullString Then
                    ' make item public?
                    If MapItem(mapnum, i).playerTimer < GetTickCount Then
                        ' make it public
                        MapItem(mapnum, i).playerName = vbNullString
                        MapItem(mapnum, i).playerTimer = 0
                        ' send updates to everyone
                        SendMapItemsToAll mapnum
                    End If
                    ' despawn item?
                    If MapItem(mapnum, i).canDespawn Then
                        If MapItem(mapnum, i).despawnTimer < GetTickCount Then
                            ' despawn it
                            ClearMapItem i, mapnum
                            ' send updates to everyone
                            SendMapItemsToAll mapnum
                        End If
                    End If
                End If
            End If
        Next
        
        '  Close the doors
        If TickCount > TempTile(mapnum).DoorTimer + 5000 Then
            For x1 = 0 To Map(mapnum).MaxX
                For y1 = 0 To Map(mapnum).MaxY
                    If Map(mapnum).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(mapnum).DoorOpen(x1, y1) = YES Then
                        TempTile(mapnum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap mapnum, x1, y1, 0
                    End If
                Next
            Next
        End If
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(mapnum).Npc(i).Num > 0 Then
                For X = 1 To MAX_DOTS
                    HandleDoT_Npc mapnum, i, X
                    HandleHoT_Npc mapnum, i, X
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(mapnum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(mapnum).Resource_Count
                Resource_index = Map(mapnum).Tile(ResourceCache(mapnum).ResourceData(i).X, ResourceCache(mapnum).ResourceData(i).Y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(mapnum).ResourceData(i).ResourceState = 1 Or ResourceCache(mapnum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(mapnum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(mapnum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(mapnum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(mapnum).ResourceData(i).cur_health = RAND(Resource(Resource_index).health_min, Resource(Resource_index).health_max)
                            SendResourceCacheToMap mapnum, i
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(mapnum) = YES Then
            TickCount = GetTickCount
            
            For X = 1 To MAX_MAP_NPCS
                npcNum = MapNpc(mapnum).Npc(X).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).Npc(X) > 0 And MapNpc(mapnum).Npc(X).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(mapnum).Npc(X).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = mapnum And MapNpc(mapnum).Npc(X).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        n = Npc(npcNum).Range
                                        DistanceX = MapNpc(mapnum).Npc(X).X - GetPlayerX(i)
                                        DistanceY = MapNpc(mapnum).Npc(X).Y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                If Len(Trim$(Npc(npcNum).AttackSay)) > 0 Then
                                                    Call PlayerMsg(i, Trim$(Npc(npcNum).Name) & " says: " & Trim$(Npc(npcNum).AttackSay), SayColor)
                                                End If
                                                MapNpc(mapnum).Npc(X).targetType = 1 ' player
                                                MapNpc(mapnum).Npc(X).target = i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).Npc(X) > 0 And MapNpc(mapnum).Npc(X).Num > 0 Then
                    If MapNpc(mapnum).Npc(X).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(mapnum).Npc(X).StunTimer + (MapNpc(mapnum).Npc(X).StunDuration * 1000) Then
                            MapNpc(mapnum).Npc(X).StunDuration = 0
                            MapNpc(mapnum).Npc(X).StunTimer = 0
                        End If
                    Else
                            
                        target = MapNpc(mapnum).Npc(X).target
                        targetType = MapNpc(mapnum).Npc(X).targetType
    
                        ' Check to see if its time for the npc to walk
                        If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If targetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(target) And GetPlayerMap(target) = mapnum Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(target)
                                        TargetX = GetPlayerX(target)
                                    Else
                                        MapNpc(mapnum).Npc(X).targetType = 0 ' clear
                                        MapNpc(mapnum).Npc(X).target = 0
                                    End If
                                End If
                            
                            ElseIf targetType = 2 Then 'npc
                                
                                If target > 0 Then
                                    
                                    If MapNpc(mapnum).Npc(target).Num > 0 Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = MapNpc(mapnum).Npc(target).Y
                                        TargetX = MapNpc(mapnum).Npc(target).X
                                    Else
                                        MapNpc(mapnum).Npc(X).targetType = 0 ' clear
                                        MapNpc(mapnum).Npc(X).target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                
                                i = Int(Rnd * 5)
    
                                ' Lets move the npc
                                Select Case i
                                    Case 0
    
                                        ' Up
                                        If MapNpc(mapnum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_UP) Then
                                                Call NpcMove(mapnum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapnum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_DOWN) Then
                                                Call NpcMove(mapnum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapnum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_LEFT) Then
                                                Call NpcMove(mapnum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapnum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_RIGHT) Then
                                                Call NpcMove(mapnum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 1
    
                                        ' Right
                                        If MapNpc(mapnum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_RIGHT) Then
                                                Call NpcMove(mapnum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapnum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_LEFT) Then
                                                Call NpcMove(mapnum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapnum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_DOWN) Then
                                                Call NpcMove(mapnum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapnum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_UP) Then
                                                Call NpcMove(mapnum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 2
    
                                        ' Down
                                        If MapNpc(mapnum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_DOWN) Then
                                                Call NpcMove(mapnum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapnum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_UP) Then
                                                Call NpcMove(mapnum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapnum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_RIGHT) Then
                                                Call NpcMove(mapnum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(mapnum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_LEFT) Then
                                                Call NpcMove(mapnum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 3
    
                                        ' Left
                                        If MapNpc(mapnum).Npc(X).X > TargetX And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_LEFT) Then
                                                Call NpcMove(mapnum, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(mapnum).Npc(X).X < TargetX And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_RIGHT) Then
                                                Call NpcMove(mapnum, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(mapnum).Npc(X).Y > TargetY And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_UP) Then
                                                Call NpcMove(mapnum, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(mapnum).Npc(X).Y < TargetY And Not DidWalk Then
                                            If CanNpcMove(mapnum, X, DIR_DOWN) Then
                                                Call NpcMove(mapnum, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                End Select
    
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(mapnum).Npc(X).X - 1 = TargetX And MapNpc(mapnum).Npc(X).Y = TargetY Then
                                        If MapNpc(mapnum).Npc(X).Dir <> DIR_LEFT Then
                                            Call NpcDir(mapnum, X, DIR_LEFT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(mapnum).Npc(X).X + 1 = TargetX And MapNpc(mapnum).Npc(X).Y = TargetY Then
                                        If MapNpc(mapnum).Npc(X).Dir <> DIR_RIGHT Then
                                            Call NpcDir(mapnum, X, DIR_RIGHT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(mapnum).Npc(X).X = TargetX And MapNpc(mapnum).Npc(X).Y - 1 = TargetY Then
                                        If MapNpc(mapnum).Npc(X).Dir <> DIR_UP Then
                                            Call NpcDir(mapnum, X, DIR_UP)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(mapnum).Npc(X).X = TargetX And MapNpc(mapnum).Npc(X).Y + 1 = TargetY Then
                                        If MapNpc(mapnum).Npc(X).Dir <> DIR_DOWN Then
                                            Call NpcDir(mapnum, X, DIR_DOWN)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
    
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
    
                                            If CanNpcMove(mapnum, X, i) Then
                                                Call NpcMove(mapnum, X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
    
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(mapnum, X, i) Then
                                        Call NpcMove(mapnum, X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).Npc(X) > 0 And MapNpc(mapnum).Npc(X).Num > 0 Then
                    target = MapNpc(mapnum).Npc(X).target
                    targetType = MapNpc(mapnum).Npc(X).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                    
                        If targetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = mapnum Then
                                TryNpcAttackPlayer X, target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(mapnum).Npc(X).target = 0
                                MapNpc(mapnum).Npc(X).targetType = 0 ' clear
                            End If
                    End If
                End If
                
                                            ' Spell Casting
                                For i = 1 To MAX_NPC_SPELLS
                                    If Npc(npcNum).Spell(i) > 0 Then
                                        If MapNpc(mapnum).Npc(X).SpellTimer(i) + (Spell(Npc(npcNum).Spell(i)).CastTime * 1000) < GetTickCount Then
                                            NpcSpellPlayer X, target, i
                                        End If
                                    End If
                                Next
                            End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(mapnum).Npc(X).stopRegen Then
                    If MapNpc(mapnum).Npc(X).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(mapnum).Npc(X).Vital(Vitals.HP) > 0 Then
                            MapNpc(mapnum).Npc(X).Vital(Vitals.HP) = MapNpc(mapnum).Npc(X).Vital(Vitals.HP) + GetNpcVitalRegen(npcNum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(mapnum).Npc(X).Vital(Vitals.HP) > GetNpcMaxVital(npcNum, Vitals.HP) Then
                                MapNpc(mapnum).Npc(X).Vital(Vitals.HP) = GetNpcMaxVital(npcNum, Vitals.HP)
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(mapnum).Npc(X).Num = 0 And Map(mapnum).Npc(X) > 0 Then
                    If TickCount > MapNpc(mapnum).Npc(X).SpawnWait + (Npc(Map(mapnum).Npc(X)).SpawnSecs * 1000) Then
                        Call SpawnNpc(X, mapnum)
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Public Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        'Call TextAdd("Saving all online players...")

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
                Call SaveBank(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 300
    If Secs Mod 30 = 0 Or Secs <= 30 Then
        Call GlobalMsg("Server shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated server shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1
    

    If Secs <= 0 Then
        Call GlobalMsg("Server is now being shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub
