Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
    Dim FrameTime As Long
    Dim Tick As Long
    Dim TickFPS As Long
    Dim FPS As Long
    Dim i As Long
    Dim WalkTimer As Long
    Dim tmr25 As Long
    Dim tmr100 As Long
    Dim tmr10000 As Long

    ' *** Start GameLoop ***
    Do While InGame
        Tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = Tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = Tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < Tick Then

            ' characters
            If NumCharacters > 0 Then
                For i = 1 To NumCharacters    'Check to unload surfaces
                    If CharacterTimer(i) > 0 Then 'Only update surfaces in use
                        If CharacterTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i)))
                            Set DDS_Character(i) = Nothing
                            CharacterTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' Paperdolls
            If NumPaperdolls > 0 Then
                For i = 1 To NumPaperdolls    'Check to unload surfaces
                    If PaperdollTimer(i) > 0 Then 'Only update surfaces in use
                        If PaperdollTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i)))
                            Set DDS_Paperdoll(i) = Nothing
                            PaperdollTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' animations
            If NumAnimations > 0 Then
                For i = 1 To NumAnimations    'Check to unload surfaces
                    If AnimationTimer(i) > 0 Then 'Only update surfaces in use
                        If AnimationTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i)))
                            Set DDS_Animation(i) = Nothing
                            AnimationTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Items
            If NumItems > 0 Then
                For i = 1 To NumItems    'Check to unload surfaces
                    If ItemTimer(i) > 0 Then 'Only update surfaces in use
                        If ItemTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i)))
                            Set DDS_Item(i) = Nothing
                            ItemTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Resources
            If NumResources > 0 Then
                For i = 1 To NumResources    'Check to unload surfaces
                    If ResourceTimer(i) > 0 Then 'Only update surfaces in use
                        If ResourceTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i)))
                            Set DDS_Resource(i) = Nothing
                            ResourceTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' spell icons
            If NumSpellIcons > 0 Then
                For i = 1 To NumSpellIcons    'Check to unload surfaces
                    If SpellIconTimer(i) > 0 Then 'Only update surfaces in use
                        If SpellIconTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i)))
                            Set DDS_SpellIcon(i) = Nothing
                            SpellIconTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' faces
            If NumFaces > 0 Then
                For i = 1 To NumFaces    'Check to unload surfaces
                    If FaceTimer(i) > 0 Then 'Only update surfaces in use
                        If FaceTimer(i) < Tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i)))
                            Set DDS_Face(i) = Nothing
                            FaceTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' check ping
            Call GetPing
            Call DrawPing
            tmr10000 = Tick + 10000
        End If

        If tmr25 < Tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if trade timed out
            If TradeRequest Then
                If TradeTimer < Tick Then
                    AddText "You took too long to decide. Please try again.", BrightRed
                    TradeRequest = False
                    TradeTimer = 0
                End If
            End If
            
            ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i)).CDTime * 1000) < Tick Then
                                SpellCD(i) = 0
                                BltPlayerSpells
                                BltHotbar
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < Tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < Tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = Tick + 250
            End If
            
            ' Update inv animation
            If NumItems > 0 Then
                If tmr100 < Tick Then
                    BltAnimatedInvItems
                    tmr100 = Tick + 100
                End If
            End If
            
            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            
            tmr25 = Tick + 25
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < Tick Then

            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To Npc_HighIndex
                If Map.Npc(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i

            WalkTimer = Tick + 30 ' edit this value to change WalkTimer
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        DoEvents

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < Tick + 15
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < Tick Then
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

    Loop

    frmMain.Visible = False

    If isLogging Then
        isLogging = False
        frmMain.picScreen.Visible = False
        frmMenu.Visible = True
        GettingMap = True
        StopMidi
        PlayMidi Options.MenuMusic
    Else
        ' Shutdown the game
        frmLoad.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
        Case MOVING_WALKING: MovementSpeed = ((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).YOffset = Player(Index).YOffset - MovementSpeed
            If Player(Index).YOffset < 0 Then Player(Index).YOffset = 0
        Case DIR_DOWN
            Player(Index).YOffset = Player(Index).YOffset + MovementSpeed
            If Player(Index).YOffset > 0 Then Player(Index).YOffset = 0
        Case DIR_LEFT
            Player(Index).XOffset = Player(Index).XOffset - MovementSpeed
            If Player(Index).XOffset < 0 Then Player(Index).XOffset = 0
        Case DIR_RIGHT
            Player(Index).XOffset = Player(Index).XOffset + MovementSpeed
            If Player(Index).XOffset > 0 Then Player(Index).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).XOffset >= 0) And (Player(Index).YOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        Else
            If (Player(Index).XOffset <= 0) And (Player(Index).YOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 1 Then
                    Player(Index).Step = 3
                Else
                    Player(Index).Step = 1
                End If
            End If
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).YOffset < 0 Then MapNpc(MapNpcNum).YOffset = 0
                
            Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).YOffset > 0 Then MapNpc(MapNpcNum).YOffset = 0
                
            Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).XOffset < 0 Then MapNpc(MapNpcNum).XOffset = 0
                
            Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).XOffset > 0 Then MapNpc(MapNpcNum).XOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum).Moving > 0 Then
            If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Then
                If (MapNpc(MapNpcNum).XOffset >= 0) And (MapNpc(MapNpcNum).YOffset >= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            Else
                If (MapNpc(MapNpcNum).XOffset <= 0) And (MapNpc(MapNpcNum).YOffset <= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 1 Then
                        MapNpc(MapNpcNum).Step = 3
                    Else
                        MapNpc(MapNpcNum).Step = 1
                    End If
                End If
            End If
        End If
    End If
End Sub

Sub CheckMapGetItem()
    Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount
            Buffer.WriteLong CMapGetItem
            SendData Buffer.ToArray()
        End If
    End If

    Set Buffer = Nothing
End Sub

Public Sub CheckAttack()
    Dim Buffer As clsBuffer
    Dim attackspeed As Long

    If ControlDown Then
    
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With

                Set Buffer = New clsBuffer
                Buffer.WriteLong CAttack
                SendData Buffer.ToArray()
                Set Buffer = Nothing
            End If
        End If
    End If

End Sub

Function IsTryingToMove() As Boolean

    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If

End Function

Function CanMove() As Boolean
    Dim d As Long
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        frmMain.picCover.Visible = False
        frmMain.picBank.Visible = False
    End If

    d = GetPlayerDir(MyIndex)

    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
    Dim x As Long
    Dim y As Long
    Dim i As Long
    CheckDirection = False
    
    ' check directional blocking
    If isDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, Direction + 1) Then
        CheckDirection = True
        Exit Function
    End If

    Select Case Direction
        Case DIR_UP
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            x = GetPlayerX(MyIndex) - 1
            y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            x = GetPlayerX(MyIndex) + 1
            y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If Map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the key door is open or not
    If Map.Tile(x, y).Type = TILE_TYPE_KEY Then

        ' This actually checks if its open or not
        If TempTile(x, y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
    ' Check to see if a player is already on that tile
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If GetPlayerX(i) = x Then
                If GetPlayerY(i) = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next i

    ' Check to see if a npc is already on that tile
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If MapNpc(i).x = x Then
                If MapNpc(i).y = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next

End Function

Sub CheckMovement()

    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If

            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

            If Player(MyIndex).XOffset = 0 Then
                If Player(MyIndex).YOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End If
End Sub

Sub PlayerSearch(ByVal CurX As Long, ByVal CurY As Long)
    Dim Buffer As clsBuffer

    If isInBounds Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CSearch
        Buffer.WriteLong CurX
        Buffer.WriteLong CurY
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    End If

End Sub

Sub SendTradeRequest(ByVal CurX As Long, ByVal CurY As Long)
    Dim Buffer As clsBuffer

    If isInBounds Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CTradeRequest
        Buffer.WriteLong CurX
        Buffer.WriteLong CurY
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    End If

End Sub

Public Function isInBounds()

    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

End Function

Public Sub UpdateDrawMapName()
    
    DrawMapNameX = Camera.Left + ((MAX_MAPX + 1) * PIC_X / 2) - getWidth(TexthDC, Trim$(Map.Name))
    DrawMapNameY = Camera.top + 1

    Select Case Map.Moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = QBColor(BrightRed)
        Case MAP_MORAL_SAFE
            DrawMapNameColor = QBColor(White)
        Case Else
            DrawMapNameColor = QBColor(White)
    End Select

End Sub

Public Sub UseItem()

    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
End Sub

Public Sub ForgetSpell(ByVal spellslot As Long)
    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellslot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellslot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellslot) > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong spellslot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If
End Sub

Public Sub CastSpell(ByVal spellslot As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellslot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellslot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellslot)).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellslot)).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellslot) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong spellslot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                SpellBuffer = spellslot
                SpellBufferTimer = GetTickCount
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If

End Sub

Sub ClearTempTile()
    Dim x As Long
    Dim y As Long
    ReDim TempTile(0 To Map.MaxX, 0 To Map.MaxY)

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            TempTile(x, y).DoorOpen = NO
        Next
    Next

End Sub

Public Sub DevMsg(ByVal text As String, ByVal color As Byte)

    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(text, color)
        End If
    End If

    Debug.Print text
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If

End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If

End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String

    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If

End Function

Sub DrawPing()
    Dim PingToDraw As String
    PingToDraw = Ping

    Select Case Ping
        Case -1
            PingToDraw = "Syncing"
        Case 0 To 5
            PingToDraw = "Local"
    End Select

    frmMain.lblPing.Caption = PingToDraw
End Sub

Public Sub UpdateSpellWindow(ByVal spellnum As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long

    ' check for off-screen
    If y + frmMain.picSpellDesc.height > frmMain.ScaleHeight Then
        y = frmMain.ScaleHeight - frmMain.picSpellDesc.height
    End If
    
    With frmMain
        .picSpellDesc.top = y
        .picSpellDesc.Left = x
        .picSpellDesc.Visible = True
        
        If LastSpellDesc = spellnum Then Exit Sub
        
        .lblSpellName.Caption = Trim$(Spell(spellnum).Name)
        .lblSpellDesc.Caption = Trim$(Spell(spellnum).Desc)
        BltSpellDesc spellnum
    End With
End Sub

Public Sub UpdateDescWindow(ByVal itemnum As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long
Dim FirstLetter As String * 1
Dim Name As String

    FirstLetter = LCase$(Left$(Trim$(Item(itemnum).Name), 1))
   
    If FirstLetter = "$" Then
        Name = (Mid$(Trim$(Item(itemnum).Name), 2, Len(Trim$(Item(itemnum).Name)) - 1))
    Else
        Name = Trim$(Item(itemnum).Name)
    End If
    
    ' check for off-screen
    If y + frmMain.picItemDesc.height > frmMain.ScaleHeight Then
        y = frmMain.ScaleHeight - frmMain.picItemDesc.height
    End If
    
    ' set z-order
    frmMain.picItemDesc.ZOrder (0)

    With frmMain
        .picItemDesc.top = y
        .picItemDesc.Left = x
        .picItemDesc.Visible = True

        If LastItemDesc = itemnum Then Exit Sub ' exit out after setting x + y so we don't reset values

        ' set the name
        Select Case Item(itemnum).Rarity
            Case 0 ' white
                .lblItemName.ForeColor = RGB(255, 255, 255)
            Case 1 ' green
                .lblItemName.ForeColor = RGB(117, 198, 92)
            Case 2 ' blue
                .lblItemName.ForeColor = RGB(103, 140, 224)
            Case 3 ' maroon
                .lblItemName.ForeColor = RGB(205, 34, 0)
            Case 4 ' purple
                .lblItemName.ForeColor = RGB(193, 104, 204)
            Case 5 ' orange
                .lblItemName.ForeColor = RGB(217, 150, 64)
        End Select
        
        ' set captions
        .lblItemName.Caption = Name
        .lblItemDesc.Caption = Trim$(Item(itemnum).Desc)
        
        ' render the item
        BltItemDesc itemnum
    End With

End Sub

Public Sub CacheResources()
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY

            If Map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).x = x
                MapResource(Resource_Count).y = y
            End If

        Next
    Next

    Resource_Index = Resource_Count
End Sub

Public Sub CreateActionMsg(ByVal message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal x As Long, ByVal y As Long)
Dim i As Long

    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .message = message
        .color = color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .x = x
        .y = y
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).y = ActionMsg(ActionMsgIndex).y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).x = ActionMsg(ActionMsgIndex).x + Rand(-8, 8)
    End If
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim i As Long

    ActionMsg(Index).message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).x = 0
    ActionMsg(Index).y = 0
    
    ' find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Created > 0 Then
            Action_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
    Dim looptime As Long
    Dim Layer As Long
    Dim FrameCount As Long
    Dim lockindex As Long
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).FrameIndex(Layer) = 0 Then AnimInstance(Index).FrameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(Layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).FrameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).FrameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).FrameIndex(Layer) = AnimInstance(Index).FrameIndex(Layer) + 1
                End If
                AnimInstance(Index).Timer(Layer) = GetTickCount
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    InShop = shopnum
    ShopAction = 0
    frmMain.picCover.Visible = True
    frmMain.picShop.Visible = True
    BltShop
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
If bankslot = 0 Then: GetBankItemNum = 0: Exit Function
If bankslot > MAX_BANK Then: GetBankItemNum = 0: Exit Function
    GetBankItemNum = Bank.Item(bankslot).num
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemnum As Long)
    Bank.Item(bankslot).num = itemnum
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    GetBankItemValue = Bank.Item(bankslot).Value
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    Bank.Item(bankslot).Value = ItemValue
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockvar As Byte, ByRef Dir As Byte, ByVal block As Boolean)
    If block Then
        blockvar = blockvar Or (2 ^ Dir)
    Else
        blockvar = blockvar And Not (2 ^ Dir)
    End If
End Sub

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function IsHotbarSlot(ByVal x As Single, ByVal y As Single) As Long
    Dim top As Long, Left As Long
    Dim i As Long
    IsHotbarSlot = 0

    For i = 1 To MAX_HOTBAR
        top = HotbarTop
        Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        If x >= Left And x <= Left + PIC_X Then
            If y >= top And y <= top + PIC_Y Then
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next
End Function

Public Sub PlayMapSound(ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).Sound)
            
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).Sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(Npc(entityNum).Sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).Sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).Sound)
        ' other
        Case Else
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "None." Then Exit Sub

    ' play the sound
    PlaySound soundName
End Sub
