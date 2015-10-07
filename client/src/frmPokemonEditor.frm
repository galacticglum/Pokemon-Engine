VERSION 5.00
Begin VB.Form frmPokemonEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pokemon Editor"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   5175
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtCatchRate 
         Height          =   285
         Left            =   3720
         TabIndex        =   39
         Text            =   "0"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtEvolveLvl 
         Height          =   285
         Left            =   3480
         TabIndex        =   38
         Text            =   "0"
         Top             =   4680
         Width           =   1335
      End
      Begin VB.ComboBox cmbEvolve 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   4680
         Width           =   1455
      End
      Begin VB.ComboBox cmbSType 
         Height          =   315
         ItemData        =   "frmPokemonEditor.frx":0000
         Left            =   2760
         List            =   "frmPokemonEditor.frx":0037
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3720
         Width           =   2055
      End
      Begin VB.ComboBox cmbPType 
         Height          =   315
         ItemData        =   "frmPokemonEditor.frx":00B6
         Left            =   2760
         List            =   "frmPokemonEditor.frx":00ED
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtBaseExp 
         Height          =   285
         Left            =   4080
         TabIndex        =   25
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtFemaleRate 
         Height          =   285
         Left            =   4080
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "Base Stats"
         Height          =   1455
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   4695
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   6
            Left            =   3600
            TabIndex        =   22
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   5
            Left            =   3600
            TabIndex        =   20
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   4
            Left            =   3600
            TabIndex        =   18
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   16
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   14
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtBaseStat 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Text            =   "0"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Spd"
            Height          =   255
            Left            =   2520
            TabIndex        =   21
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "SpDef"
            Height          =   255
            Left            =   2520
            TabIndex        =   19
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "SpAtk"
            Height          =   255
            Left            =   2520
            TabIndex        =   17
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Def"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Atk"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "HP"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.HScrollBar scrlPokeIcon 
         Height          =   255
         Left            =   840
         Max             =   0
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   7
         Top             =   600
         Width           =   480
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Caption         =   "Move List"
         Height          =   1815
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   2535
         Begin VB.TextBox txtMoveLevel 
            Height          =   285
            Left            =   1680
            TabIndex        =   30
            Text            =   "0"
            Top             =   1320
            Width           =   735
         End
         Begin VB.ComboBox cmbMoveNum 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1320
            Width           =   1455
         End
         Begin VB.ListBox lstMove 
            Height          =   1035
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "%"
         Height          =   255
         Left            =   4440
         TabIndex        =   41
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Catch Rate:"
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "At Lvl:"
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Evolve To:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Secondary Type:"
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Primary Type:"
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Base Exp:"
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Female Rate:"
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblPokeIcon 
         Caption         =   "Icon: 0"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox lstIndex 
      Height          =   5520
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPokemonEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbEvolve_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Pokemon(EditorIndex).EvolveNum = cmbEvolve.ListIndex
    If Pokemon(EditorIndex).EvolveNum = 0 Then
        Pokemon(EditorIndex).EvolveLvl = 0
        txtEvolveLvl.Text = "0"
    End If
    
    Exit Sub
errHandler:
    HandleError "cmbEvolve_Change", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbMoveNum_Click()
Dim TmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Data1Index <= 0 Then Exit Sub
    If cmbMoveNum.ListIndex < 0 Then Exit Sub
    
    Pokemon(EditorIndex).MoveNum(Data1Index) = cmbMoveNum.ListIndex
    If Pokemon(EditorIndex).MoveNum(Data1Index) = 0 Then
        Pokemon(EditorIndex).MoveLevel(Data1Index) = 0
        txtMoveLevel.Text = "0"
    End If
    
    TmpIndex = lstMove.ListIndex
    lstMove.RemoveItem Data1Index - 1
    If Pokemon(EditorIndex).MoveNum(Data1Index) > 0 Then
        lstMove.AddItem Data1Index & ": " & Trim$(Moves(Pokemon(EditorIndex).MoveNum(Data1Index)).Name) & " - Lv" & Pokemon(EditorIndex).MoveLevel(Data1Index), Data1Index - 1
    Else
        lstMove.AddItem Data1Index & ": None.", Data1Index - 1
    End If
    lstMove.ListIndex = TmpIndex
    
    Exit Sub
errHandler:
    HandleError "cmbMoveNum_Change", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbPType_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Pokemon(EditorIndex).pType = cmbPType.ListIndex
    
    Exit Sub
errHandler:
    HandleError "cmbPType_Click", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSType_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Pokemon(EditorIndex).sType = cmbSType.ListIndex
    
    Exit Sub
errHandler:
    HandleError "cmbSType_Click", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtCatchRate_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
        
    If Not IsNumeric(txtCatchRate.Text) Then Exit Sub
    Pokemon(EditorIndex).CatchRate = Val(txtCatchRate.Text)
    
    Exit Sub
errHandler:
    HandleError "txtCatchRate_Validate", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEvolveLvl_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not IsNumeric(txtEvolveLvl.Text) Then Exit Sub
    
    Pokemon(EditorIndex).EvolveLvl = Val(txtEvolveLvl.Text)
    
    Exit Sub
errHandler:
    HandleError "txtEvolveLvl_Change", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMoveLevel_Validate(Cancel As Boolean)
Dim TmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Data1Index <= 0 Then Exit Sub
    If Not IsNumeric(txtMoveLevel.Text) Then Exit Sub
    If Pokemon(EditorIndex).MoveNum(Data1Index) <= 0 Then Exit Sub
    
    Pokemon(EditorIndex).MoveLevel(Data1Index) = Val(txtMoveLevel.Text)
    
    TmpIndex = lstMove.ListIndex
    lstMove.RemoveItem Data1Index - 1
    lstMove.AddItem Data1Index & ": " & Trim$(Moves(Pokemon(EditorIndex).MoveNum(Data1Index)).Name) & " - Lv" & Pokemon(EditorIndex).MoveLevel(Data1Index), Data1Index - 1
    lstMove.ListIndex = TmpIndex
    
    Exit Sub
errHandler:
    HandleError "txtMoveLevel_Validate", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    PokemonEditorCancel
    
    Exit Sub
errHandler:
    HandleError "cmdCancel_Click", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim TmpIndex As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ClearPokemon EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pokemon(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    PokemonEditorInit
    
    Exit Sub
errHandler:
    HandleError "cmdDelete_Click", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOkay_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    PokemonEditorOk
    
    Exit Sub
errHandler:
    HandleError "cmdOkay_Click", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    scrlPokeIcon.max = Count_PokeIcon
    
    Exit Sub
errHandler:
    HandleError "Form_Load", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    PokemonEditorInit
    
    Exit Sub
errHandler:
    HandleError "lstIndex_Click", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstMove_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Data1Index = lstMove.ListIndex + 1
    cmbMoveNum.ListIndex = Pokemon(EditorIndex).MoveNum(Data1Index)
    txtMoveLevel.Text = Pokemon(EditorIndex).MoveLevel(Data1Index)
    
    Exit Sub
errHandler:
    HandleError "lstMove_Click", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPokeIcon_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Pokemon(EditorIndex).Pic = scrlPokeIcon.value
    lblPokeIcon.Caption = "Icon: " & scrlPokeIcon.value
    
    Exit Sub
errHandler:
    HandleError "scrlPokeIcon_Change", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtBaseExp_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsNumeric(txtBaseExp.Text) Then
        Pokemon(EditorIndex).BaseExp = Val(txtBaseExp.Text)
    End If
    
    Exit Sub
errHandler:
    HandleError "txtBaseExp_Change", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtBaseStat_Change(Index As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsNumeric(txtBaseStat(Index).Text) Then
        Pokemon(EditorIndex).BaseStat(Index) = Val(txtBaseStat(Index).Text)
    End If
    
    Exit Sub
errHandler:
    HandleError "txtBaseStat_Change", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtFemaleRate_Validate(Cancel As Boolean)
    On Error GoTo chanceErr
        
    If Not IsNumeric(txtFemaleRate.Text) And Not Right$(txtFemaleRate.Text, 1) = "%" And Not InStr(1, txtFemaleRate.Text, "/") > 0 And Not InStr(1, txtFemaleRate.Text, ".") Then
        txtFemaleRate.Text = "0"
        Pokemon(EditorIndex).FemaleRate = 0
        Exit Sub
    End If
    
    If Right$(txtFemaleRate.Text, 1) = "%" Then
        txtFemaleRate.Text = Left(txtFemaleRate.Text, Len(txtFemaleRate.Text) - 1) / 100
    ElseIf InStr(1, txtFemaleRate.Text, "/") > 0 Then
        Dim i() As String
        i = Split(txtFemaleRate.Text, "/")
        txtFemaleRate.Text = Int(i(0) / i(1) * 1000) / 1000
    End If
    
    If txtFemaleRate.Text > 1 Or txtFemaleRate.Text < 0 Then
        Err.Description = "Value must be between 0 and 1!"
        GoTo chanceErr
    End If
    
    Pokemon(EditorIndex).FemaleRate = txtFemaleRate.Text
    
    Exit Sub
chanceErr:
    txtFemaleRate.Text = "0"
    Pokemon(EditorIndex).FemaleRate = 0
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim TmpIndex As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If EditorIndex = 0 Then Exit Sub
    TmpIndex = lstIndex.ListIndex
    Pokemon(EditorIndex).Name = (txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Pokemon(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    cmbEvolve.Clear
    cmbEvolve.AddItem "None."
    For i = 1 To Count_Pokemon
        cmbEvolve.AddItem i & ": " & Trim$(Pokemon(i).Name)
    Next
    cmbEvolve.ListIndex = Pokemon(EditorIndex).EvolveNum
    
    Exit Sub
errHandler:
    HandleError "txtName_Validate", "frmPokemonEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
