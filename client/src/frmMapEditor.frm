VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMapEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame fraAttribute 
      Caption         =   "Attributes"
      Height          =   6375
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Frame fraShop 
         Caption         =   "Shop"
         Height          =   1215
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         Begin VB.CommandButton cmdEditShop 
            Caption         =   "Edit"
            Height          =   255
            Left            =   1440
            TabIndex        =   34
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdShopOk 
            Caption         =   "Confirm"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1335
         End
         Begin VB.HScrollBar scrlShopNum 
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lblShopNum 
            Caption         =   "Shop: None"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame fraWarp 
         Caption         =   "Warp"
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         Begin VB.CommandButton cmdEditWarp 
            Caption         =   "Edit"
            Height          =   255
            Left            =   1440
            TabIndex        =   29
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdWarpOk 
            Caption         =   "Confirm"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtMapY 
            Height          =   285
            Left            =   1200
            TabIndex        =   26
            Text            =   "0"
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtMapX 
            Height          =   285
            Left            =   1200
            TabIndex        =   24
            Text            =   "0"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtMap 
            Height          =   285
            Left            =   1200
            TabIndex        =   22
            Text            =   "0"
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Y:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "X:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Map:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   6120
      Width           =   2175
   End
   Begin VB.PictureBox picTileset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   120
      ScaleHeight     =   382
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   190
      TabIndex        =   3
      Top             =   480
      Width           =   2880
   End
   Begin VB.ComboBox cmbTileset 
      Height          =   315
      ItemData        =   "frmMapEditor.frx":0000
      Left            =   120
      List            =   "frmMapEditor.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      Max             =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   2895
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   5775
      Left            =   3000
      Max             =   0
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Layers"
      TabPicture(0)   =   "frmMapEditor.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdClear"
      Tab(0).Control(1)=   "optLayers(0)"
      Tab(0).Control(2)=   "optLayers(1)"
      Tab(0).Control(3)=   "optLayers(2)"
      Tab(0).Control(4)=   "optLayers(3)"
      Tab(0).Control(5)=   "optLayers(4)"
      Tab(0).Control(6)=   "cmdFill"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Attributes"
      TabPicture(1)   =   "frmMapEditor.frx":0020
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "optAttributes(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "optAttributes(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "optAttributes(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "optAttributes(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "optAttributes(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "optAttributes(6)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "optAttributes(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.OptionButton optAttributes 
         Caption         =   "Shop"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   36
         Top             =   1920
         Width           =   2175
      End
      Begin VB.OptionButton optAttributes 
         Caption         =   "Warp"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   2175
      End
      Begin VB.OptionButton optAttributes 
         Caption         =   "Storage"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optAttributes 
         Caption         =   "Checkpoint"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton optAttributes 
         Caption         =   "Heal"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optAttributes 
         Caption         =   "Grass"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   -74640
         TabIndex        =   14
         Top             =   4200
         Width           =   1935
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Ground"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Mask"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Mask2"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Fringe"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   7
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton optLayers 
         Caption         =   "Fringe2"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optAttributes 
         Caption         =   "Blocked"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   255
         Left            =   -74640
         TabIndex        =   15
         Top             =   3960
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    MapEditorCancel
    
    Exit Sub
errHandler:
    HandleError "cmdCancel_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    MapEditorClearLayer
    
    Exit Sub
errHandler:
    HandleError "cmdClear_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdConfirm_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    MapEditorSend
    
    Exit Sub
errHandler:
    HandleError "cmdConfirm_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEditShop_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    scrlShopNum.Enabled = True
    cmdShopOk.Enabled = True
    cmdEditShop.Enabled = False
    
    Exit Sub
errHandler:
    HandleError "cmdEditShop_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEditWarp_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    txtMap.Enabled = True
    txtMapX.Enabled = True
    txtMapY.Enabled = True
    cmdWarpOk.Enabled = True
    cmdEditWarp.Enabled = False
    
    Exit Sub
errHandler:
    HandleError "cmdEditWarp_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdFill_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    MapEditorFillLayer
    
    Exit Sub
errHandler:
    HandleError "cmdFill_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdProperties_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    InitProperties
    
    Exit Sub
errHandler:
    HandleError "cmdProperties_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTileset_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CurTileset = cmbTileset.ListIndex + 1
    
    If Not gTexture(Tex_Tileset(CurTileset)).loaded Then Call LoadTexture(Tex_Tileset(CurTileset))
    scrlPictureY.max = (GetTextureHeight(Tex_Tileset(CurTileset)) \ Pic_Size) - (picTileset.Height \ Pic_Size)
    scrlPictureX.max = (GetTextureWidth(Tex_Tileset(CurTileset)) \ Pic_Size) - (picTileset.Width \ Pic_Size)
    MapEditorTileScroll
    
    EditorTileX = 0
    EditorTileY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1
    
    Exit Sub
errHandler:
    HandleError "cmbTileset_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdShopOk_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    EditorData1 = scrlShopNum.value
    
    scrlShopNum.Enabled = False
    cmdShopOk.Enabled = False
    cmdEditShop.Enabled = True
    
    Exit Sub
errHandler:
    HandleError "cmdShopOk_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdWarpOk_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not IsNumeric(txtMap.Text) Then txtMap.Text = "0"
    If Not IsNumeric(txtMapX.Text) Then txtMapX.Text = "0"
    If Not IsNumeric(txtMapY.Text) Then txtMapY.Text = "0"
    
    If Val(txtMap.Text) > Count_Map Or Val(txtMap.Text) < 0 Then txtMap.Text = "0"
    If Val(txtMapX.Text) < 0 Then txtMapX.Text = "0"
    If Val(txtMapY.Text) < 0 Then txtMapY.Text = "0"
    
    EditorData1 = Val(txtMap.Text)
    EditorData2 = Val(txtMapX.Text)
    EditorData3 = Val(txtMapY.Text)
    
    txtMap.Enabled = False
    txtMapX.Enabled = False
    txtMapY.Enabled = False
    cmdWarpOk.Enabled = False
    cmdEditWarp.Enabled = True
    
    Exit Sub
errHandler:
    HandleError "cmdWarpOk_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    scrlShopNum.max = Count_Shop
    
    Exit Sub
errHandler:
    HandleError "optAttributes_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub optAttributes_Click(Index As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CurAttribute = Index
    ClearAttributeDialogue
    
    fraAttribute.Visible = True
    Select Case CurAttribute
        Case Attributes.Warp
            fraWarp.Visible = True
        Case Attributes.mShop
            fraShop.Visible = True
    End Select
    
    Exit Sub
errHandler:
    HandleError "optAttributes_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub optLayers_Click(Index As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CurLayer = Index
    
    Exit Sub
errHandler:
    HandleError "optLayers_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub picTileset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    x = x + (scrlPictureX.value * Pic_Size)
    y = y + (scrlPictureY.value * Pic_Size)
    Call MapEditorDrag(Button, x, y)
    
    Exit Sub
errHandler:
    HandleError "picTileset_MouseMove", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    x = x + (scrlPictureX.value * Pic_Size)
    y = y + (scrlPictureY.value * Pic_Size)
    Call MapEditorChooseTile(Button, x, y)
    
    Exit Sub
errHandler:
    HandleError "picTileset_MouseDown", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlShopNum_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If scrlShopNum.value > 0 Then
        lblShopNum.Caption = "Shop: " & Trim$(Shop(scrlShopNum.value).Name)
    Else
        lblShopNum.Caption = "Shop: None"
    End If
    
    Exit Sub
errHandler:
    HandleError "scrlShopNum_Change", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If SSTab1.Tab = 0 Then
        CurLayer = Layers.Ground
        optLayers(CurLayer).value = True
        ClearAttributeDialogue
    ElseIf SSTab1.Tab = 1 Then
        CurAttribute = Attributes.Blocked
        optAttributes(CurAttribute).value = True
        ClearAttributeDialogue
        fraAttribute.Visible = True
    End If
    
    Exit Sub
errHandler:
    HandleError "SSTab1_Click", "frmMapEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
