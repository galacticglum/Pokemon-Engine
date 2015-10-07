VERSION 5.00
Begin VB.Form frmNPCEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Editor"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Okay"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   4695
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1800
         Max             =   0
         TabIndex        =   7
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblSprite 
         Caption         =   "Sprite: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox lstIndex 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmNPCEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    NPCEditorCancel
    
    Exit Sub
errHandler:
    HandleError "cmdCancel_Click", "frmNPCEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NPCEditorInit
    
    Exit Sub
errHandler:
    HandleError "cmdDelete_Click", "frmNPCEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    NPCEditorOk
    
    Exit Sub
errHandler:
    HandleError "cmdSave_Click", "frmNPCEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    scrlSprite.max = Count_Sprite
End Sub

Private Sub lstIndex_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    NPCEditorInit
    
    Exit Sub
errHandler:
    HandleError "lstIndex_Click", "frmNPCEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    NPC(EditorIndex).Sprite = scrlSprite.value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).Name = (txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    Exit Sub
errHandler:
    HandleError "txtName_Validate", "frmNPCEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub


