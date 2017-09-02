VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Music"
      Height          =   3015
      Left            =   120
      TabIndex        =   56
      Top             =   4800
      Width           =   2055
      Begin VB.ListBox lstMusic 
         Height          =   2010
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   23
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   720
         TabIndex        =   22
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1320
         TabIndex        =   21
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   855
      Left            =   2280
      TabIndex        =   16
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   960
         List            =   "frmMapProperties.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boot Map:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Boot X:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Boot Y:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   630
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   5895
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   4215
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   30
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   5400
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   29
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   5400
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   28
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   5040
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   27
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   5040
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   26
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   25
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   24
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   23
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   22
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   3960
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   21
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   3960
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   20
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3600
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   19
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   3600
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   18
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   17
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   16
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   15
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   14
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   13
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   12
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   11
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   10
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   9
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   8
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   7
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   6
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   4
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   2
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPlay_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopMidi
    PlayMidi lstMusic.List(lstMusic.ListIndex)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdPlay_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdStop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    StopMidi
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdStop_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOk_Click()
    Dim i As Long
    Dim sTemp As Long
    Dim x As Long, x2 As Long
    Dim y As Long, y2 As Long
    Dim tempArr() As TileRec
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsNumeric(txtMaxX.text) Then txtMaxX.text = Map.MaxX
    If Val(txtMaxX.text) < MAX_MAPX Then txtMaxX.text = MAX_MAPX
    If Val(txtMaxX.text) > MAX_BYTE Then txtMaxX.text = MAX_BYTE
    If Not IsNumeric(txtMaxY.text) Then txtMaxY.text = Map.MaxY
    If Val(txtMaxY.text) < MAX_MAPY Then txtMaxY.text = MAX_MAPY
    If Val(txtMaxY.text) > MAX_BYTE Then txtMaxY.text = MAX_BYTE

    With Map
        .Name = Trim$(txtName.text)
        If lstMusic.ListIndex >= 0 Then
            .Music = lstMusic.List(lstMusic.ListIndex)
        Else
            .Music = vbNullString
        End If
        .Up = Val(txtUp.text)
        .Down = Val(txtDown.text)
        .Left = Val(txtLeft.text)
        .Right = Val(txtRight.text)
        .Moral = cmbMoral.ListIndex
        .BootMap = Val(txtBootMap.text)
        .BootX = Val(txtBootX.text)
        .BootY = Val(txtBootY.text)

        For i = 1 To MAX_MAP_NPCS
            If cmbNpc(i).ListIndex > 0 Then
                sTemp = InStr(1, Trim$(cmbNpc(i).text), ":", vbTextCompare)

                If Len(Trim$(cmbNpc(i).text)) = sTemp Then
                    cmbNpc(i).ListIndex = 0
                End If
            End If
        Next

        For i = 1 To MAX_MAP_NPCS
            .Npc(i) = cmbNpc(i).ListIndex
        Next

        ' set the data before changing it
        tempArr = Map.Tile
        x2 = Map.MaxX
        y2 = Map.MaxY
        ' change the data
        .MaxX = Val(txtMaxX.text)
        .MaxY = Val(txtMaxY.text)
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)

        If x2 > .MaxX Then x2 = .MaxX
        If y2 > .MaxY Then y2 = .MaxY

        For x = 0 To x2
            For y = 0 To y2
                .Tile(x, y) = tempArr(x, y)
            Next
        Next

        ClearTempTile
    End With

    Call UpdateDrawMapName
    Me.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdOk_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Me.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

