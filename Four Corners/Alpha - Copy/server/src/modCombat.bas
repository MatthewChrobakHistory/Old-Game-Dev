Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
                Case 2 ' Mage
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
            End Select
        Case MP
            Select Case GetPlayerClass(index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, intelligence) / 2)) * 5 + 25
                Case 2 ' Mage
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, intelligence) / 2)) * 5 + 25
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, intelligence) / 2)) * 5 + 25
            End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerStat(index, Stats.willpower) * 0.8) + 6
        Case MP
            i = (GetPlayerStat(index, Stats.willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) * Item(weaponNum).Data2 + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) + (GetPlayerLevel(index) / 5)
    End If

End Function

Function GetPlayerDef(ByVal index As Long) As Long
    Dim DefNum As Long
    Dim Def As Long
    
    GetPlayerDef = 0
    Def = 0
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        DefNum = GetPlayerEquipment(index, Armor)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        DefNum = GetPlayerEquipment(index, Helmet)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Legs) > 0 Then
        DefNum = GetPlayerEquipment(index, Legs)
        Def = Def + Item(DefNum).Data2
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        DefNum = GetPlayerEquipment(index, Shield)
        Def = Def + Item(DefNum).Data2
    End If
    
   If Not GetPlayerEquipment(index, Armor) > 0 And Not GetPlayerEquipment(index, Helmet) > 0 And Not GetPlayerEquipment(index, Shield) > 0 Then
        GetPlayerDef = 0.085 * GetPlayerStat(index, Endurance) + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerDef = 0.085 * GetPlayerStat(index, Endurance) * Def + (GetPlayerLevel(index) / 5)
    End If
    

End Function

Function GetNpcMaxVital(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = Npc(npcNum).HP
        Case MP
            GetNpcMaxVital = 30 + (Npc(npcNum).Stat(intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Npc(npcNum).Stat(Stats.willpower) * 0.8) + 6
        Case MP
            i = (Npc(npcNum).Stat(Stats.willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal npcNum As Long) As Long
    GetNpcDamage = 0.085 * 5 * Npc(npcNum).Stat(Stats.Strength) * Npc(npcNum).damage + (Npc(npcNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long
Dim shieldNum As Long

    CanPlayerBlock = False

    rate = 0

    If GetPlayerEquipment(index, Shield) > 0 Then
        shieldNum = GetPlayerEquipment(index, Shield)
        rate = Item(shieldNum).Data2 / 43.3
        rndNum = RAND(1, 100)
        If rndNum <= rate Then
            CanPlayerBlock = True
        Else
            CanPlayerBlock = False
        End If
    Else
        rate = 0.5
        rndNum = RAND(1, 100)
        If rndNum <= rate Then
            CanPlayerBlock = True
        Else
            CanPlayerBlock = False
        End If
    End If
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(index, Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(index, Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(index, Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = Npc(npcNum).Stat(Stats.Agility) / 52.08
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = Npc(npcNum).Stat(Stats.Agility) / 83.3
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = Npc(npcNum).Stat(Stats.Strength) * 0.25
    rndNum = RAND(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal MapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim damage As Long

    damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, MapNpcNum) Then
    
        mapnum = GetPlayerMap(index)
        npcNum = MapNpc(mapnum).Npc(MapNpcNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcNum) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (MapNpc(mapnum).Npc(MapNpcNum).X * 32), (MapNpc(mapnum).Npc(MapNpcNum).Y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (MapNpc(mapnum).Npc(MapNpcNum).X * 32), (MapNpc(mapnum).Npc(MapNpcNum).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(MapNpcNum)
        damage = damage - blockAmount
        
        ' take away armour
        damage = damage - RAND(1, (Npc(npcNum).Stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            damage = damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
            
        If damage > 0 Then
            Call PlayerAttackNpc(index, MapNpcNum, damage)
        Else
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapnum).Npc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If npcNum > 0 Then
                If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    If GetPlayerEquipment(attacker, Helmet) > 0 Then
        If Npc(npcNum).HelmetReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Helmet)).Data3 <> Npc(npcNum).HelmetReq Then
                Call PlayerMsg(attacker, "You should use a more appropriate helmet before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Armor) > 0 Then
        If Npc(npcNum).ArmorReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Armor)).Data3 <> Npc(npcNum).ArmorReq Then
                Call PlayerMsg(attacker, "You should use more appropriate body armor before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Legs) > 0 Then
        If Npc(npcNum).LegsReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Legs)).Data3 <> Npc(npcNum).LegsReq Then
                Call PlayerMsg(attacker, "You should use a more appropriate leg armor before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Shield) > 0 Then
        If Npc(npcNum).ShieldReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Shield)).Data3 <> Npc(npcNum).ShieldReq Then
                Call PlayerMsg(attacker, "You should use a more appropriate shield before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        If Npc(npcNum).WeaponReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Weapon)).Data3 <> Npc(npcNum).WeaponReq Then
                Call PlayerMsg(attacker, "You should use a more appropriate weapon before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Weapon) = 0 Then
        If Npc(npcNum).WeaponReq > 0 Then
            Call PlayerMsg(attacker, "You need a special weapon to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
    If GetPlayerEquipment(attacker, Shield) = 0 Then
        If Npc(npcNum).ShieldReq > 0 Then
            Call PlayerMsg(attacker, "You need a special type of shield to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
    If GetPlayerEquipment(attacker, Helmet) = 0 Then
        If Npc(npcNum).HelmetReq > 0 Then
            Call PlayerMsg(attacker, "You need a special type of helmet to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
    If GetPlayerEquipment(attacker, Armor) = 0 Then
        If Npc(npcNum).HelmetReq > 0 Then
            Call PlayerMsg(attacker, "You need a special type of body armor to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
    If GetPlayerEquipment(attacker, Legs) = 0 Then
        If Npc(npcNum).HelmetReq > 0 Then
            Call PlayerMsg(attacker, "You need a special type of leg armor to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If npcNum > 0 And GetTickCount > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    NpcX = MapNpc(mapnum).Npc(MapNpcNum).X
                    NpcY = MapNpc(mapnum).Npc(MapNpcNum).Y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapnum).Npc(MapNpcNum).X
                    NpcY = MapNpc(mapnum).Npc(MapNpcNum).Y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapnum).Npc(MapNpcNum).X + 1
                    NpcY = MapNpc(mapnum).Npc(MapNpcNum).Y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapnum).Npc(MapNpcNum).X - 1
                    NpcY = MapNpc(mapnum).Npc(MapNpcNum).Y
            End Select

            If NpcX = GetPlayerX(attacker) Then
                If NpcY = GetPlayerY(attacker) Then
                    If Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
    If GetPlayerEquipment(attacker, Helmet) > 0 Then
        If Npc(npcNum).HelmetReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Helmet)).Data3 <> Npc(npcNum).HelmetReq Then
                Call PlayerMsg(attacker, "You should use a more appropriate helmet before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Armor) > 0 Then
        If Npc(npcNum).ArmorReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Armor)).Data3 <> Npc(npcNum).ArmorReq Then
                Call PlayerMsg(attacker, "You should use more appropriate body armor before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Legs) > 0 Then
        If Npc(npcNum).LegsReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Legs)).Data3 <> Npc(npcNum).LegsReq Then
                Call PlayerMsg(attacker, "You should use a more appropriate leg armor before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Shield) > 0 Then
        If Npc(npcNum).ShieldReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Shield)).Data3 <> Npc(npcNum).ShieldReq Then
                Call PlayerMsg(attacker, "You should use a more appropriate shield before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        If Npc(npcNum).WeaponReq > 0 Then
            If Item(GetPlayerEquipment(attacker, Weapon)).Data3 <> Npc(npcNum).WeaponReq Then
                Call PlayerMsg(attacker, "You should use a more appropriate weapon before attacking this creature.", BrightRed)
                Exit Function
            End If
        End If
    End If
    
    If GetPlayerEquipment(attacker, Weapon) = 0 Then
        If Npc(npcNum).WeaponReq > 0 Then
            Call PlayerMsg(attacker, "You need a special weapon to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
    If GetPlayerEquipment(attacker, Shield) = 0 Then
        If Npc(npcNum).ShieldReq > 0 Then
            Call PlayerMsg(attacker, "You need a special type of shield to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
    If GetPlayerEquipment(attacker, Helmet) = 0 Then
        If Npc(npcNum).HelmetReq > 0 Then
            Call PlayerMsg(attacker, "You need a special type of helmet to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
    If GetPlayerEquipment(attacker, Armor) = 0 Then
        If Npc(npcNum).HelmetReq > 0 Then
            Call PlayerMsg(attacker, "You need a special type of body armor to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
    If GetPlayerEquipment(attacker, Legs) = 0 Then
        If Npc(npcNum).LegsReq > 0 Then
            Call PlayerMsg(attacker, "You need a special type of leg armor to attack this NPC!", Yellow)
            Exit Function
        End If
    End If
    
                        CanPlayerAttackNpc = True
                    Else
                        If Len(Trim$(Npc(npcNum).AttackSay)) > 0 Then
                            PlayerMsg attacker, Trim$(Npc(npcNum).Name) & ": " & Trim$(Npc(npcNum).AttackSay), White
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, ByVal damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim Def As Long
    Dim mapnum As Long
    Dim npcNum As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapnum).Npc(MapNpcNum).Num
    
    ' projectiles
    If npcNum < 1 Then Exit Sub
    Name = Trim$(Npc(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount
    
    Player(attacker).SecOutofCombat = 10
    SendPlayerData (attacker)

    If damage >= MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapnum).Npc(MapNpcNum).X * 32), (MapNpc(mapnum).Npc(MapNpcNum).Y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound attacker, MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y)
            End If
        End If
        
        'check to see if npc is a boss npc
        If Npc(npcNum).Map <> 0 And Npc(npcNum).X <> 0 And Npc(npcNum).Y <> 0 Then
            Call PlayerWarp(attacker, Npc(npcNum).Map, Npc(npcNum).X, Npc(npcNum).Y)
            If Npc(npcNum).RewardItem <> 0 Then
                If FindOpenInvSlot(attacker, Npc(npcNum).RewardItem) <> 0 Then
                    GiveInvItem attacker, Npc(npcNum).RewardItem, 1
                Else
                    Call SpawnItem(Npc(npcNum).RewardItem, 1, Player(attacker).Map, Player(attacker).X, Player(attacker).Y)
                End If
            End If
        End If

        ' Calculate exp to give attacker
        exp = Npc(npcNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, exp, attacker
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, exp
        End If
        
        'Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If Npc(npcNum).DropItem(n) = 0 Then Exit For

        If Rnd <= Npc(npcNum).DropChance(n) Then
            Call SpawnItem(Npc(npcNum).DropItem(n), Npc(npcNum).DropItemValue(n), mapnum, MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y)
        End If
    Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).Npc(MapNpcNum).Num = 0
        MapNpc(mapnum).Npc(MapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(mapnum).Npc(MapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).Npc(MapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong MapNpcNum
        SendDataToMap mapnum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = mapnum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = MapNpcNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) - damage

        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & damage, BrightRed, 1, (MapNpc(mapnum).Npc(MapNpcNum).X * 32), (MapNpc(mapnum).Npc(MapNpcNum).Y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound attacker, MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(mapnum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(mapnum).Npc(MapNpcNum).targetType = 1 ' player
        MapNpc(mapnum).Npc(MapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).Npc(i).Num = MapNpc(mapnum).Npc(MapNpcNum).Num Then
                    MapNpc(mapnum).Npc(i).target = attacker
                    MapNpc(mapnum).Npc(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapnum).Npc(MapNpcNum).stopRegen = True
        MapNpc(mapnum).Npc(MapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNpcNum, mapnum, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Npc mapnum, MapNpcNum, SpellNum, attacker
            End If
        End If
        
        SendMapNpcVitals mapnum, MapNpcNum
    End If

    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetTickCount
    End If
    
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long)
Dim mapnum As Long, npcNum As Long, blockAmount As Long, damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNpcNum, index) Then
        mapnum = GetPlayerMap(index)
        npcNum = MapNpc(mapnum).Npc(MapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (Player(index).X * 32), (Player(index).Y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (Player(index).X * 32), (Player(index).Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetNpcDamage(npcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(index)
        damage = damage - blockAmount
        
        ' take away armour
        damage = damage - RAND(1, (GetPlayerStat(index, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(npcNum) Then
            damage = damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (MapNpc(mapnum).Npc(MapNpcNum).X * 32), (MapNpc(mapnum).Npc(MapNpcNum).Y * 32)
        End If

        damage = damage - GetPlayerDef(index)

        If damage > 0 Then
            Call NpcAttackPlayer(MapNpcNum, index, damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(index)
    npcNum = MapNpc(mapnum).Npc(MapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapnum).Npc(MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(mapnum).Npc(MapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If npcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(mapnum).Npc(MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(mapnum).Npc(MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(mapnum).Npc(MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(mapnum).Npc(MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(mapnum).Npc(MapNpcNum).Y) And (GetPlayerX(index) + 1 = MapNpc(mapnum).Npc(MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(mapnum).Npc(MapNpcNum).Y) And (GetPlayerX(index) - 1 = MapNpc(mapnum).Npc(MapNpcNum).X) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal damage As Long)
    Dim Name As String
    Dim exp As Long
    Dim mapnum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong MapNpcNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(mapnum).Npc(MapNpcNum).stopRegen = True
    MapNpc(mapnum).Npc(MapNpcNum).stopRegenTimer = GetTickCount
    
    Player(Victim).SecOutofCombat = 10

    If damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapnum).Npc(MapNpcNum).Num
        
        ' kill player
        KillPlayer Victim
        
        ' Player is dead

        ' Set NPC target to 0
        MapNpc(mapnum).Npc(MapNpcNum).target = 0
        MapNpc(mapnum).Npc(MapNpcNum).targetType = 0
        
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - damage)
        Call SendVital(Victim, Vitals.HP)
        Call SendAnimation(mapnum, Npc(MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        
        'XPOISONX
        If Npc(MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num).IsPoison = True Then
            If Player(Victim).PoisonTick = 0 Then Call PlayerMsg(Victim, "You have been poisoned!", BrightRed)
            Player(Victim).PoisonTick = Npc(MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num).PoisonTick
            Player(Victim).PoisonDamage = Npc(MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num).PoisonDamage
        End If
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(mapnum).Npc(MapNpcNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        ' If the player is a noob
        If Player(Victim).Level < 20 Then
            If GetPlayerVital(Victim, Vitals.HP) < GetPlayerMaxVital(Victim, Vitals.HP) / 2 Then
                If GetPlayerVital(Victim, Vitals.HP) > 0 Then
                    Call PlayerMsg(Victim, "Your health is low! Eat some food to regain health!", Yellow)
                End If
            End If
        End If
        
    End If

End Sub

Sub NpcSpellPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, SpellSlotNum As Long)
    Dim mapnum As Long
    Dim i As Long
    Dim n As Long
    Dim SpellNum As Long
    Dim Buffer As clsBuffer
    Dim InitDamage As Long
    Dim damage As Long
    Dim MaxHeals As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    If SpellSlotNum <= 0 Or SpellSlotNum > MAX_NPC_SPELLS Then Exit Sub

    ' The Variables
    mapnum = GetPlayerMap(Victim)
    SpellNum = Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).Spell(SpellSlotNum)
    
    ' Send this packet so they can see the person attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong MapNpcNum
    SendDataToMap mapnum, Buffer.ToArray()
    Set Buffer = Nothing
    
    ' CoolDown Time
    If MapNpc(mapnum).Npc(MapNpcNum).SpellTimer(SpellSlotNum) > GetTickCount Then Exit Sub
    
    ' Spell Types
        Select Case Spell(SpellNum).Type
            ' AOE Healing Spells
            Case SPELL_TYPE_HEALHP
            ' Make sure an npc waits for the spell to cooldown
            MaxHeals = 1 + Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).Stat(Stats.intelligence) \ 25
            If MapNpc(mapnum).Npc(MapNpcNum).Heals >= MaxHeals Then Exit Sub
                If MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) <= Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).HP * 0.3 Then
                    If Spell(SpellNum).IsAoE Then
                        For i = 1 To MAX_MAP_NPCS
                            If MapNpc(mapnum).Npc(i).Num > 0 Then
                                If MapNpc(mapnum).Npc(i).Vital(Vitals.HP) > 0 Then
                                    If isInRange(Spell(SpellNum).AoE, MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y, MapNpc(mapnum).Npc(i).X, MapNpc(mapnum).Npc(i).Y) Then
                                        InitDamage = Spell(SpellNum).Vital + (Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).Stat(Stats.intelligence) / 2)
                    
                                        MapNpc(mapnum).Npc(i).Vital(Vitals.HP) = MapNpc(mapnum).Npc(i).Vital(Vitals.HP) + InitDamage
                                        SendActionMsg mapnum, "+" & InitDamage, BrightGreen, 1, (MapNpc(mapnum).Npc(i).X * 32), (MapNpc(mapnum).Npc(i).Y * 32)
                                        Call SendAnimation(mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
                    
                                        If MapNpc(mapnum).Npc(i).Vital(Vitals.HP) > Npc(MapNpc(mapnum).Npc(i).Num).HP Then
                                            MapNpc(mapnum).Npc(i).Vital(Vitals.HP) = Npc(MapNpc(mapnum).Npc(i).Num).HP
                                        End If
                    
                                        MapNpc(mapnum).Npc(MapNpcNum).Heals = MapNpc(mapnum).Npc(MapNpcNum).Heals + 1
                    
                                        MapNpc(mapnum).Npc(MapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(SpellNum).CDTime * 1000
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Next
                    Else
                    ' Non AOE Healing Spells
                        InitDamage = Spell(SpellNum).Vital + (Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).Stat(Stats.intelligence) / 2)
                    
                        MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) + InitDamage
                        SendActionMsg mapnum, "+" & InitDamage, BrightGreen, 1, (MapNpc(mapnum).Npc(MapNpcNum).X * 32), (MapNpc(mapnum).Npc(MapNpcNum).Y * 32)
                        Call SendAnimation(mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
                    
                        If MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) > Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).HP Then
                            MapNpc(mapnum).Npc(MapNpcNum).Vital(Vitals.HP) = Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).HP
                        End If
                    
                        MapNpc(mapnum).Npc(MapNpcNum).Heals = MapNpc(mapnum).Npc(MapNpcNum).Heals + 1
                    
                        MapNpc(mapnum).Npc(MapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(SpellNum).CDTime * 1000
                        Exit Sub
                    End If
                End If
                
            ' AOE Damaging Spells
            Case SPELL_TYPE_DAMAGEHP
            ' Make sure an npc waits for the spell to cooldown
                If Spell(SpellNum).IsAoE Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = mapnum Then
                                If isInRange(Spell(SpellNum).AoE, MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    InitDamage = Spell(SpellNum).Vital + (Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).Stat(Stats.intelligence) / 2)
                                    damage = InitDamage - Player(i).Stat(Stats.willpower)
                                        If damage <= 0 Then
                                            SendActionMsg GetPlayerMap(i), "RESIST!", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                                            Exit Sub
                                        Else
                                            NpcAttackPlayer MapNpcNum, i, damage
                                            SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, MapNpcNum
                                            MapNpc(mapnum).Npc(MapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(SpellNum).CDTime * 1000
                                            Exit Sub
                                        End If
                                End If
                            End If
                        End If
                    Next
                ' Non AoE Damaging Spells
                Else
                    If isInRange(Spell(SpellNum).Range, MapNpc(mapnum).Npc(MapNpcNum).X, MapNpc(mapnum).Npc(MapNpcNum).Y, GetPlayerX(Victim), GetPlayerY(Victim)) Then
                    InitDamage = Spell(SpellNum).Vital + (Npc(MapNpc(mapnum).Npc(MapNpcNum).Num).Stat(Stats.intelligence) / 2)
                    damage = InitDamage - Player(Victim).Stat(Stats.willpower)
                        If damage <= 0 Then
                            SendActionMsg GetPlayerMap(Victim), "RESIST!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
                            Exit Sub
                        Else
                            NpcAttackPlayer MapNpcNum, Victim, damage
                            SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Victim
                            MapNpc(mapnum).Npc(MapNpcNum).SpellTimer(SpellSlotNum) = GetTickCount + Spell(SpellNum).CDTime * 1000
                            Exit Sub
                        End If
                    End If
                End If
            End Select
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal Victim As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim damage As Long

    damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, Victim) Then
    
        mapnum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim) Then
            SendActionMsg mapnum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(Victim) Then
            SendActionMsg mapnum, "Parry!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(Victim)
        damage = damage - blockAmount
        
        ' take away armour
        damage = damage - RAND(1, (GetPlayerStat(Victim, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        damage = RAND(1, damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            damage = damage * 1.5
            SendActionMsg mapnum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        damage = damage - GetPlayerDef(Victim)

        If damage > 0 Then
            Call PlayerAttackPlayer(attacker, Victim, damage)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

' projectiles
Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    If Not IsSpell And Not IsProjectile Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    If Not IsSpell And Not IsProjectile Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(Victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(Victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal Victim As Long, ByVal damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(Victim) = False Or damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount
    
    Player(attacker).SecOutofCombat = 10
    Player(Victim).SecOutofCombat = 10
    SendPlayerData (attacker)
    SendPlayerData (Victim)

    If damage >= GetPlayerVital(Victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(Victim), "-" & GetPlayerVital(Victim, Vitals.HP), BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum
        
        ' Player is dead
        ' Calculate exp to give attacker
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = Victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(Victim).inParty > 0 Then SendPartyVitals TempPlayer(Victim).inParty, Victim
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum
        
        SendActionMsg GetPlayerMap(Victim), "-" & damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer Victim, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Player Victim, SpellNum, attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    SpellNum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)
    
    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
        ' Transformation
    With Spell(SpellNum)
        If .Type = SPELL_TYPE_TRANFORMATION Then
            Call SetPlayerSprite(index, Spell(index).Trans)
            Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        End If
    End With
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    target = TempPlayer(index).target
    Range = Spell(SpellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).Npc(target).X, MapNpc(mapnum).Npc(target).Y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg mapnum, "Casting " & Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        TempPlayer(index).spellBuffer.Spell = spellslot
        TempPlayer(index).spellBuffer.Timer = GetTickCount
        TempPlayer(index).spellBuffer.target = TempPlayer(index).target
        TempPlayer(index).spellBuffer.tType = TempPlayer(index).targetType
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim X As Long, Y As Long
    Dim damage As Long
    
    Dim Buffer As clsBuffer
    Dim SpellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then Exit Sub

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    ' set the vital
    Vital = RAND(Spell(SpellNum).Vital, Spell(SpellNum).VitalMax)
    AoE = Spell(SpellNum).AoE
    Range = Spell(SpellNum).Range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    PlayerWarp index, Spell(SpellNum).Map, Spell(SpellNum).X, Spell(SpellNum).Y
                    SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(index)
                Y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(target)
                    Y = GetPlayerY(target)
                Else
                    X = MapNpc(mapnum).Npc(target).X
                    Y = MapNpc(mapnum).Npc(target).Y
                End If
                
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, Vital, SpellNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).Npc(i).Num > 0 Then
                            If MapNpc(mapnum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(mapnum).Npc(i).X, MapNpc(mapnum).Npc(i).Y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, Vital, SpellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                    
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, Vital, SpellNum
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).Npc(i).Num > 0 Then
                            If MapNpc(mapnum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, X, Y, MapNpc(mapnum).Npc(i).X, MapNpc(mapnum).Npc(i).Y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, SpellNum, mapnum
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(target)
                Y = GetPlayerY(target)
            Else
                X = MapNpc(mapnum).Npc(target).X
                Y = MapNpc(mapnum).Npc(target).Y
            End If
                
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                PlayerMsg index, "Target not in range.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer index, target, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc index, target, Vital, SpellNum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, target, True) Then
                                SpellPlayer_Effect VitalType, increment, target, Vital, SpellNum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, target, Vital, SpellNum
                        End If
                    Else
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, target, True) Then
                                SpellNpc_Effect VitalType, increment, target, Vital, SpellNum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, target, Vital, SpellNum, mapnum
                        End If
                    End If
            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(SpellNum).CDTime * 1000)
        Call SendCooldown(index, spellslot)
        SendActionMsg mapnum, Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal damage As Long, ByVal SpellNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg GetPlayerMap(index), sSymbol & damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
        
        If increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Player index, SpellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - damage
        End If
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal damage As Long, ByVal SpellNum As Long, ByVal mapnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation mapnum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, index
        SendActionMsg mapnum, sSymbol & damage, Colour, ACTIONMSG_SCROLL, MapNpc(mapnum).Npc(index).X * 32, MapNpc(mapnum).Npc(index).Y * 32
        
        ' send the sound
        SendMapSound index, MapNpc(mapnum).Npc(index).X, MapNpc(mapnum).Npc(index).Y, SoundEntity.seSpell, SpellNum
        
        If increment Then
            MapNpc(mapnum).Npc(index).Vital(Vital) = MapNpc(mapnum).Npc(index).Vital(Vital) + damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Npc mapnum, index, SpellNum
            End If
        ElseIf Not increment Then
            MapNpc(mapnum).Npc(index).Vital(Vital) = MapNpc(mapnum).Npc(index).Vital(Vital) - damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).Npc(index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal SpellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(mapnum).Npc(index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Player(index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(index).X * 32, Player(index).Y * 32
                Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Spell(.Spell).Vital
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal dotNum As Long)
    With MapNpc(mapnum).Npc(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerAttackNpc .Caster, index, Spell(.Spell).Vital, , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal hotNum As Long)
    With MapNpc(mapnum).Npc(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg mapnum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapnum).Npc(index).X * 32, MapNpc(mapnum).Npc(index).Y * 32
                MapNpc(mapnum).Npc(index).Vital(Vitals.HP) = MapNpc(mapnum).Npc(index).Vital(Vitals.HP) + Spell(.Spell).Vital
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal mapnum As Long, ByVal SpellNum As Long)
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).Npc(index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(mapnum).Npc(index).StunTimer = GetTickCount
    End If
End Sub

