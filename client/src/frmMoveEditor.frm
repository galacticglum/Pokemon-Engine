VERSION 5.00
Begin VB.Form frmMoveEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Move Editor"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox lstIndex 
      Height          =   5130
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   4695
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cmbAtkType 
         Height          =   315
         ItemData        =   "frmMoveEditor.frx":0000
         Left            =   2520
         List            =   "frmMoveEditor.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox cmbPType 
         Height          =   315
         ItemData        =   "frmMoveEditor.frx":0042
         Left            =   120
         List            =   "frmMoveEditor.frx":0079
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtPP 
         Height          =   285
         Left            =   3360
         TabIndex        =   9
         Text            =   "0"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtPower 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Text            =   "0"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Attack Type:"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Primary Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "PP:"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Power:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4920
      Width           =   1335
   End
End
Attribute VB_Name = "frmMoveEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAtkType_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Moves(EditorIndex).AtkType = cmbAtkType.ListIndex
    
    Exit Sub
errHandler:
    HandleError "cmbAtkType_Click", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbPType_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Moves(EditorIndex).Type = cmbPType.ListIndex
    
    Exit Sub
errHandler:
    HandleError "cmbPType_Click", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    MoveEditorCancel
    
    Exit Sub
errHandler:
    HandleError "cmdCancel_Click", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ClearMove EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Moves(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    MoveEditorInit
    
    Exit Sub
errHandler:
    HandleError "cmdDelete_Click", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOkay_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    MoveEditorOk
    
    Exit Sub
errHandler:
    HandleError "cmdOkay_Click", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    MoveEditorInit
    
    Exit Sub
errHandler:
    HandleError "lstIndex_Click", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Moves(EditorIndex).Name = (txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Moves(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    Exit Sub
errHandler:
    HandleError "txtName_Validate", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtPower_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsNumeric(txtPower.Text) Then
        Moves(EditorIndex).Power = Val(txtPower.Text)
    End If
    
    Exit Sub
errHandler:
    HandleError "txtPower_Change", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtPP_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsNumeric(txtPP.Text) Then
        Moves(EditorIndex).PP = Val(txtPP.Text)
    End If
    
    Exit Sub
errHandler:
    HandleError "txtPP_Change", "frmMoveEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
