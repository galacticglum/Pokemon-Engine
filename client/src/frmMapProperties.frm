VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Data"
      TabPicture(0)   =   "frmMapProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblBack"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblField"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "scrlBack"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlField"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Pokemon"
      TabPicture(1)   =   "frmMapProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbPokemon"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstPokemon"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Music"
      TabPicture(2)   =   "frmMapProperties.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdStop"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdPlay"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lstMusic"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "NPC"
      TabPicture(3)   =   "frmMapProperties.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstNPC"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmbNPC"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.ListBox lstNPC 
         Height          =   3180
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   4215
      End
      Begin VB.ComboBox cmbNPC 
         Height          =   315
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3720
         Width           =   4215
      End
      Begin VB.HScrollBar scrlField 
         Height          =   255
         Left            =   1920
         Max             =   0
         TabIndex        =   30
         Top             =   3720
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBack 
         Height          =   255
         Left            =   1920
         Max             =   0
         TabIndex        =   28
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   -72720
         TabIndex        =   25
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ListBox lstMusic 
         Height          =   2400
         Left            =   -74760
         TabIndex        =   23
         Top             =   480
         Width           =   3975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pokemon Level Range"
         Height          =   855
         Left            =   1920
         TabIndex        =   19
         Top             =   1920
         Width           =   2415
         Begin VB.TextBox txtLvlMax 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtLvlMin 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Text            =   "0"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "-"
            Height          =   255
            Left            =   960
            TabIndex        =   22
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.ComboBox cmbPokemon 
         Height          =   315
         Left            =   -74760
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ListBox lstPokemon 
         Height          =   2205
         Left            =   -74760
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Caption         =   "Map Link"
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
         Begin VB.TextBox txtLink 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   3
            Left            =   480
            TabIndex        =   16
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtLink 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   15
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtLink 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtLink 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   480
            TabIndex        =   13
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Properties"
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4215
         Begin VB.TextBox txtMaxY 
            Height          =   285
            Left            =   3000
            TabIndex        =   7
            Text            =   "0"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtMaxX 
            Height          =   285
            Left            =   960
            TabIndex        =   6
            Text            =   "0"
            Top             =   585
            Width           =   1095
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox cmbMoral 
            Height          =   315
            ItemData        =   "frmMapProperties.frx":0070
            Left            =   960
            List            =   "frmMapProperties.frx":008C
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "Max Y:"
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Max X:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   585
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Moral:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Label lblField 
         Caption         =   "Battle Fields: 0"
         Height          =   255
         Left            =   1920
         TabIndex        =   29
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label lblBack 
         Caption         =   "Battle Background: 0"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Grass Area"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbPokemon_Click()
Dim tmpString() As String
Dim pIndex As Long, tIndex As Long, X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not cmbPokemon.ListCount > 0 Then Exit Sub
    If Not lstPokemon.ListCount > 0 Then Exit Sub
    
    tmpString = Split(cmbPokemon.List(cmbPokemon.ListIndex))
    If Not cmbPokemon.List(cmbPokemon.ListIndex) = "None" Then
        pIndex = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        Map.Pokemon(lstPokemon.ListIndex + 1) = pIndex
    Else
        Map.Pokemon(lstPokemon.ListIndex + 1) = 0
    End If
    
    tIndex = lstPokemon.ListIndex
    lstPokemon.Clear
    For X = 1 To MAX_MAP_POKEMON
        If Map.Pokemon(X) > 0 Then
            lstPokemon.AddItem X & ": " & Trim$(Pokemon(Map.Pokemon(X)).Name)
        Else
            lstPokemon.AddItem X & ": None"
        End If
    Next
    lstPokemon.ListIndex = tIndex

    Exit Sub
errHandler:
    HandleError "cmbPokemon_Click", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbNPC_Click()
Dim tmpString() As String
Dim pIndex As Long, tIndex As Long, X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not cmbNPC.ListCount > 0 Then Exit Sub
    If Not lstNPC.ListCount > 0 Then Exit Sub
    
    tmpString = Split(cmbNPC.List(cmbNPC.ListIndex))
    If Not cmbNPC.List(cmbNPC.ListIndex) = "None" Then
        pIndex = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        Map.NPC(lstNPC.ListIndex + 1) = pIndex
    Else
        Map.NPC(lstNPC.ListIndex + 1) = 0
    End If
    
    tIndex = lstNPC.ListIndex
    lstNPC.Clear
    For X = 1 To MAX_MAP_NPC
        If Map.NPC(X) > 0 Then
            lstNPC.AddItem X & ": " & Trim$(NPC(Map.NPC(X)).Name)
        Else
            lstNPC.AddItem X & ": None"
        End If
    Next
    lstNPC.ListIndex = tIndex

    Exit Sub
errHandler:
    HandleError "cmbNPC_Click", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Unload Me
    
    Exit Sub
errHandler:
    HandleError "cmdCancel_Click", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOk_Click()
Dim X As Long, Y As Long
Dim x2 As Long, y2 As Long
Dim TmpTile() As TileRec
Dim i As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Map
        If Not IsNumeric(txtMaxX.Text) Then txtMaxX.Text = .MaxX
        If Val(txtMaxX.Text) < Max_MapX Then txtMaxX.Text = Max_MapX
        If Val(txtMaxX.Text) > MAX_BYTE Then txtMaxX.Text = MAX_BYTE
        If Not IsNumeric(txtMaxY.Text) Then txtMaxY.Text = .MaxY
        If Val(txtMaxY.Text) < Max_MapY Then txtMaxY.Text = Max_MapY
        If Val(txtMaxY.Text) > MAX_BYTE Then txtMaxY.Text = MAX_BYTE
    
        .Name = Trim$(txtName.Text)
        
        If lstMusic.ListIndex >= 0 Then
            .Music = lstMusic.List(lstMusic.ListIndex)
        Else
            .Music = vbNullString
        End If
        
        .Moral = cmbMoral.ListIndex
        
        TmpTile = .Tile
        X = .MaxX
        Y = .MaxY
        .MaxX = Val(txtMaxX.Text)
        .MaxY = Val(txtMaxY.Text)
        
        If X > .MaxX Then X = .MaxX
        If Y > .MaxY Then Y = .MaxY
        
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)
        
        For x2 = 0 To X
            For y2 = 0 To Y
                Map.Tile(x2, y2) = TmpTile(x2, y2)
            Next
        Next
        
        For i = 0 To 3
            If IsNumeric(txtLink(i).Text) Then
                Map.Link(i) = Val(txtLink(i).Text)
            End If
        Next
        
        If IsNumeric(txtLvlMin.Text) Then
            Map.MinLvl = Val(txtLvlMin.Text)
        End If
        If IsNumeric(txtLvlMax.Text) Then
            Map.MaxLvl = Val(txtLvlMax.Text)
        End If
        
        Map.CurField = scrlField.value
        Map.CurBack = scrlBack.value
    End With
    Unload Me
    
    Exit Sub
errHandler:
    HandleError "cmdOk_Click", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdPlay_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    StopMusic
    PlayMusic lstMusic.List(lstMusic.ListIndex)
    
    Exit Sub
errHandler:
    HandleError "cmdPlay_Click", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdStop_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    StopMusic
    
    Exit Sub
errHandler:
    HandleError "cmdStop_Click", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    scrlField.max = Count_FieldFront
    scrlBack.max = Count_Background
    
    Exit Sub
errHandler:
    HandleError "Form_Load", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstPokemon_Click()
Dim tmpString() As String
Dim pIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not cmbPokemon.ListCount > 0 Then Exit Sub
    If Not lstPokemon.ListCount > 0 Then Exit Sub
    
    tmpString = Split(lstPokemon.List(lstPokemon.ListIndex))
    pIndex = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    cmbPokemon.ListIndex = Map.Pokemon(pIndex)
    
    Exit Sub
errHandler:
    HandleError "lstPokemon_Click", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstNPC_Click()
Dim tmpString() As String
Dim pIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not cmbNPC.ListCount > 0 Then Exit Sub
    If Not lstNPC.ListCount > 0 Then Exit Sub
    
    tmpString = Split(lstNPC.List(lstNPC.ListIndex))
    pIndex = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    cmbNPC.ListIndex = Map.NPC(pIndex)
    
    Exit Sub
errHandler:
    HandleError "lstNPC_Click", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBack_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    lblBack.Caption = "Battle Background: " & scrlBack.value
    
    Exit Sub
errHandler:
    HandleError "scrlBack_Change", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlField_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    lblField.Caption = "Battle Field: " & scrlField.value
    
    Exit Sub
errHandler:
    HandleError "scrlField_Change", "frmMapProperties", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
