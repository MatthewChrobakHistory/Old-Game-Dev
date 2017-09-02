Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public DX7 As New DirectX7  ' Master Object, early binding

Public Sub Main()

    ' set loading screen
    frmLoad.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\loading.jpg")
    frmLoad.Visible = True

    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions

    ' load main menu
    Call SetStatus("Loading Menu...")
    Load frmMenu
    
    ' load gui
    Call SetStatus("Loading interface...")
    loadGUI
    
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "spellicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "gui"
    ChkDir App.Path & "\data files\graphics\gui\", "menu"
    ChkDir App.Path & "\data files\graphics\gui\", "main"
    ChkDir App.Path & "\data files\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "bars"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    
    ' Update the form with the game's name before it's loaded
    frmMain.Caption = GAME_NAME
    
    ' initialize DirectX
    If Not InitDirectDraw Then
        MsgBox "Error Initializing DirectX7 - DirectDraw."
        DestroyGame
    End If
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages
    Call SetStatus("Initializing DirectX...")
    
    ' DX7 Master Object is already created, early binding
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    
    ' populate lists
    Call SetStatus("Populating lists...")
    PopulateLists
    
    ' temp set music/sound vars
    Music_On = True
    Sound_On = True
    
    ' load music/sound engine
    InitSound
    InitMusic
    
    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMidi Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    'Load frmMainMenu
    Load frmMenu
    
    ' cache the buttons then reset & render them
    Call SetStatus("Loading buttons...")
    cacheButtons
    resetButtons_Menu
    
    ' we can now see it
    frmMenu.Visible = True
    
    ' hide all pics
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' hide the load form
    frmLoad.Visible = False
End Sub

Public Sub loadGUI()
    ' menu
    frmMenu.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\background.jpg")
    frmMenu.picMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\main.jpg")
    frmMenu.picLogin.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\login.jpg")
    frmMenu.picRegister.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\register.jpg")
    frmMenu.picCredits.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\credits.jpg")
    frmMenu.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\character.jpg")
    ' main
    frmMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\main.jpg")
    frmMain.picInventory.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\inventory.jpg")
    frmMain.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\character.jpg")
    frmMain.picSpells.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\skills.jpg")
    frmMain.picOptions.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\options.jpg")
    frmMain.picItemDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_item.jpg")
    frmMain.picSpellDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_spell.jpg")
    frmMain.picTempInv.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempSpell.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picShop.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\shop.jpg")
    frmMain.picBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bank.jpg")
    frmMain.picTrade.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\trade.jpg")
    frmMain.picHotbar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\hotbar.jpg")
    ' main - bars
    frmMain.imgHPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\health.jpg")
    frmMain.imgMPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\spirit.jpg")
    frmMain.imgEXPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\experience.jpg")
    ' store the bar widths for calculations
    HPBar_Width = frmMain.imgHPBar.width
    SPRBar_Width = frmMain.imgMPBar.width
    EXPBar_Width = frmMain.imgEXPBar.width
End Sub

Public Sub MenuState(ByVal state As Long)
    frmLoad.Visible = True

    Select Case state
        Case MENU_STATE_ADDCHAR
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending character addition data...")

                If frmMenu.optMale.Value Then
                    Call SendAddChar(frmMenu.txtCName, SEX_MALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                Else
                    Call SendAddChar(frmMenu.txtCName, SEX_FEMALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite)
                End If
            End If
            
        Case MENU_STATE_NEWACCOUNT
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMenu.txtRUser.text, frmMenu.txtRPass.text)
            End If

        Case MENU_STATE_LOGIN
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMenu.txtLUser.text, frmMenu.txtLPass.text)
                Exit Sub
            End If
    End Select

    If frmLoad.Visible Then
        If Not IsConnected Then
            frmMenu.Visible = True
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picCharacter.Visible = False
            frmMenu.picRegister.Visible = False
            frmLoad.Visible = False
            Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, GAME_NAME)
        End If
    End If

End Sub

Sub GameInit()
    EnteringGame = True
    frmMenu.Visible = False
    EnteringGame = False
    
    ' bring all the main gui components to the front
    frmMain.picShop.ZOrder (0)
    frmMain.picBank.ZOrder (0)
    frmMain.picTrade.ZOrder (0)
    
    ' hide gui
    frmMain.picCover.Visible = False
    InBank = False
    InShop = False
    InTrade = False
    
    ' Set font
    Call SetFont(FONT_NAME, FONT_SIZE)
    frmLoad.Visible = False
    frmMain.Show
    
    ' Set the focus
    Call SetFocusOnChat
    frmMain.picScreen.Visible = True
    
    ' Blt inv
    BltInventory
    
    ' blt hotbar
    BltHotbar
    
    ' get ping
    GetPing
    DrawPing
    
    ' set values for amdin panel
    frmMain.scrlAItem.Max = MAX_ITEMS
    frmMain.scrlAItem.Value = 1
    
    'stop the song playing
    StopMidi
End Sub

Public Sub DestroyGame()
    ' break out of GameLoop
    InGame = False
    Call DestroyTCP
    
    'destroy objects in reverse order
    Call DestroyDirectDraw

    ' destory DirectX7 master object
    If Not DX7 Is Nothing Then
        Set DX7 = Nothing
    End If

    Call UnloadAllForms
    End
End Sub

Public Sub UnloadAllForms()
    Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next

End Sub

Public Sub SetStatus(ByVal Caption As String)
    frmLoad.lblStatus.Caption = Caption
    DoEvents
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)

    If NewLine Then
        Txt.text = Txt.text + Msg + vbCrLf
    Else
        Txt.text = Txt.text + Msg
    End If

    Txt.SelStart = Len(Txt.text) - 1
End Sub

Public Sub SetFocusOnChat()

    On Error Resume Next 'prevent RTE5, no way to handle error

    frmMain.txtMyChat.SetFocus
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim GlobalX As Long
    Dim GlobalY As Long
    GlobalX = PB.Left
    GlobalY = PB.top

    If Button = 1 Then
        PB.Left = GlobalX + x - SOffsetX
        PB.top = GlobalY + y - SOffsetY
    End If

End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean

    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
    Dim i As Long

    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, GAME_NAME)
            Exit Function
        End If

    Next

    isStringLegal = True
End Function

' ####################
' ## Buttons - Menu ##
' ####################
Public Sub cacheButtons()
    ' menu - login
    With MenuButton(1)
        .fileName = "login"
        .state = 0 ' normal
    End With
    
    ' menu - register
    With MenuButton(2)
        .fileName = "register"
        .state = 0 ' normal
    End With
    
    ' menu - credits
    With MenuButton(3)
        .fileName = "credits"
        .state = 0 ' normal
    End With
    
    ' menu - exit
    With MenuButton(4)
        .fileName = "exit"
        .state = 0 ' normal
    End With
    
    ' main - inv
    With MainButton(1)
        .fileName = "inv"
        .state = 0 ' normal
    End With
    
    ' main - skills
    With MainButton(2)
        .fileName = "skills"
        .state = 0 ' normal
    End With
    
    ' main - char
    With MainButton(3)
        .fileName = "char"
        .state = 0 ' normal
    End With
    
    ' main - opt
    With MainButton(4)
        .fileName = "opt"
        .state = 0 ' normal
    End With
    
    ' main - trade
    With MainButton(5)
        .fileName = "trade"
        .state = 0 ' normal
    End With
    
    ' main - exit
    With MainButton(6)
        .fileName = "exit"
        .state = 0 ' normal
    End With
End Sub

' menu specific buttons
Public Sub resetButtons_Menu(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' loop through entire array
    For i = 1 To MAX_MENUBUTTONS
        ' only change if different and not exception
        If Not MenuButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            MenuButton(i).state = 0 'normal
            renderButton_Menu i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Menu = 0
End Sub

Public Sub renderButton_Menu(ByVal buttonNum As Long)
Dim bSuffix As String
    ' get the suffix
    Select Case MenuButton(buttonNum).state
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMenu.imgButton(buttonNum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(buttonNum).fileName & bSuffix & ".jpg")
End Sub

Public Sub changeButtonState_Menu(ByVal buttonNum As Long, ByVal bState As Byte)
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MenuButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        MenuButton(buttonNum).state = bState
        renderButton_Menu buttonNum
    End If
End Sub

' main specific buttons
Public Sub resetButtons_Main(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' loop through entire array
    For i = 1 To MAX_MAINBUTTONS
        ' only change if different and not exception
        If Not MainButton(i).state = 0 And Not i = exceptionNum Then
            ' reset state and render
            MainButton(i).state = 0 'normal
            renderButton_Main i
        End If
    Next
    
    If exceptionNum = 0 Then LastButtonSound_Main = 0
End Sub

Public Sub renderButton_Main(ByVal buttonNum As Long)
Dim bSuffix As String
    ' get the suffix
    Select Case MainButton(buttonNum).state
        Case 0 ' normal
            bSuffix = "_norm"
        Case 1 ' hover
            bSuffix = "_hover"
        Case 2 ' click
            bSuffix = "_click"
    End Select
    
    ' render the button
    frmMain.imgButton(buttonNum).Picture = LoadPicture(App.Path & MAINBUTTON_PATH & MainButton(buttonNum).fileName & bSuffix & ".jpg")
End Sub

Public Sub changeButtonState_Main(ByVal buttonNum As Long, ByVal bState As Byte)
    ' valid state?
    If bState >= 0 And bState <= 2 Then
        ' exit out early if state already is same
        If MainButton(buttonNum).state = bState Then Exit Sub
        ' change and render
        MainButton(buttonNum).state = bState
        renderButton_Main buttonNum
    End If
End Sub

Public Sub PopulateLists()
Dim strLoad As String

    ' music in map properties
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    strLoad = Dir(App.Path & MUSIC_PATH & "*.mid")
    Do While strLoad > vbNullString
       frmEditor_MapProperties.lstMusic.AddItem strLoad
       strLoad = Dir
    Loop
    
    ' sounds in editors
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    
    strLoad = Dir(App.Path & SOUND_PATH & "*.wav")
    Do While strLoad > vbNullString
       frmEditor_Animation.cmbSound.AddItem strLoad
       frmEditor_Item.cmbSound.AddItem strLoad
       frmEditor_NPC.cmbSound.AddItem strLoad
       frmEditor_Resource.cmbSound.AddItem strLoad
       frmEditor_Spell.cmbSound.AddItem strLoad
       strLoad = Dir
    Loop
End Sub
