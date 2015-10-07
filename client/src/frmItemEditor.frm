VERSION 5.00
Begin VB.Form frmItemEditor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
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
      Begin VB.TextBox txtSellPrice 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Text            =   "0"
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Frame fraPokeBall 
         Caption         =   "Pokeball Properties"
         Height          =   1095
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
         Begin VB.TextBox txtCatchRate 
            Height          =   285
            Left            =   1200
            TabIndex        =   17
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "0, 0.3 to 3"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Catch Rate:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   3495
      End
      Begin VB.HScrollBar scrlItemPic 
         Height          =   255
         Left            =   3120
         Max             =   0
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.PictureBox picItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4320
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   10
         Top             =   360
         Width           =   480
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         ItemData        =   "frmItemEditor.frx":0000
         Left            =   960
         List            =   "frmItemEditor.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame fraItem 
         Caption         =   "Item Properties"
         Height          =   2055
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame fraValue 
            Caption         =   "Value"
            Height          =   615
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Visible         =   0   'False
            Width           =   3615
            Begin VB.TextBox txtValue 
               Height          =   285
               Left            =   120
               TabIndex        =   22
               Text            =   "0"
               Top             =   240
               Width           =   3375
            End
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "frmItemEditor.frx":0044
            Left            =   840
            List            =   "frmItemEditor.frx":0051
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label6 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label7 
         Caption         =   "Sell Price:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblItemPic 
         Caption         =   "Item Pic: 0"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
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
      ItemData        =   "frmItemEditor.frx":0073
      Left            =   120
      List            =   "frmItemEditor.frx":0075
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbItemType_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If cmbItemType.ListIndex = ItemType.Items Then
        fraItem.Visible = True
    Else
        fraItem.Visible = False
    End If
    
    If cmbItemType.ListIndex = ItemType.Pokeballs Then
        fraPokeBall.Visible = True
        txtCatchRate.Text = "0"
    Else
        fraPokeBall.Visible = False
    End If
    
    If Not InEditorInit Then Item(EditorIndex).Type = cmbItemType.ListIndex
    
    Exit Sub
errHandler:
    HandleError "cmbItemType_Click", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If cmbType.ListIndex = ItemProperties.RestoreHP Or cmbType.ListIndex = ItemProperties.RestorePP Then
        fraValue.Visible = True
    Else
        fraValue.Visible = False
    End If
    
    If Not InEditorInit Then Item(EditorIndex).IType = cmbType.ListIndex
    
    Exit Sub
errHandler:
    HandleError "cmbType_Click", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ItemEditorCancel
    
    Exit Sub
errHandler:
    HandleError "cmdCancel_Click", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    Exit Sub
errHandler:
    HandleError "cmdDelete_Click", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ItemEditorOk
    
    Exit Sub
errHandler:
    HandleError "cmdSave_Click", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    scrlItemPic.max = Count_ItemPic
    
    Exit Sub
errHandler:
    HandleError "Form_Load", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ItemEditorInit
    
    Exit Sub
errHandler:
    HandleError "lstIndex_Click", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItemPic_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    lblItemPic.Caption = "Item Pic: " & scrlItemPic.value
    Item(EditorIndex).Pic = scrlItemPic.value
    
    Exit Sub
errHandler:
    HandleError "scrlItemPic_Change", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtCatchRate_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Item(EditorIndex).Data3 = txtCatchRate.Text
    
    Exit Sub
errHandler:
    HandleError "txtCatchRate_Change", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Item(EditorIndex).Desc = txtDesc.Text
    
    Exit Sub
errHandler:
    HandleError "txtDesc_Change", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = (txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    Exit Sub
errHandler:
    HandleError "txtName_Validate", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSellPrice_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not IsNumeric(txtSellPrice.Text) Then Exit Sub
    Item(EditorIndex).Sell = Val(txtSellPrice.Text)
    
    Exit Sub
errHandler:
    HandleError "txtSellPrice_Change", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtValue_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not IsNumeric(txtValue.Text) Then Exit Sub
    Item(EditorIndex).Data2 = Val(txtValue.Text)
    
    Exit Sub
errHandler:
    HandleError "txtValue_Change", "frmItemEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
