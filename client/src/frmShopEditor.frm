VERSION 5.00
Begin VB.Form frmShopEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
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
      Begin VB.Frame fraItem 
         Height          =   1335
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   4455
         Begin VB.ComboBox cmbItemNum 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtPrice 
            Height          =   285
            Left            =   1440
            TabIndex        =   9
            Text            =   "0"
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "Item Num:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Price:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.ListBox lstItem 
         Height          =   840
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   3855
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
Attribute VB_Name = "frmShopEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbItemNum_Click()
Dim tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Data1Index = 0 Then Exit Sub
    
    If cmbItemNum.ListIndex <> Shop(EditorIndex).sItem(Data1Index).Num Then
        tmpIndex = lstItem.ListIndex
        Shop(EditorIndex).sItem(Data1Index).Num = (cmbItemNum.ListIndex)
        lstItem.RemoveItem Data1Index - 1
        If Shop(EditorIndex).sItem(Data1Index).Num > 0 Then
            lstItem.AddItem Data1Index & ": " & Trim$(Item(Shop(EditorIndex).sItem(Data1Index).Num).Name), Data1Index - 1
        Else
            lstItem.AddItem Data1Index & ": None"
        End If
        lstItem.ListIndex = tmpIndex
    End If
    
    Exit Sub
errHandler:
    HandleError "cmbItemNum_Click", "frmShopEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ShopEditorCancel
    
    Exit Sub
errHandler:
    HandleError "cmdCancel_Click", "frmShopEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ClearShop EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ShopEditorInit
    
    Exit Sub
errHandler:
    HandleError "cmdDelete_Click", "frmShopEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ShopEditorOk
    
    Exit Sub
errHandler:
    HandleError "cmdSave_Click", "frmShopEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ShopEditorInit
    
    Exit Sub
errHandler:
    HandleError "lstIndex_Click", "frmShopEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstItem_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Data1Index = lstItem.ListIndex + 1
    
    If Data1Index > 0 Then
        fraItem.Visible = True
        cmbItemNum.ListIndex = Shop(EditorIndex).sItem(Data1Index).Num
        txtPrice.Text = Shop(EditorIndex).sItem(Data1Index).Price
    Else
        fraItem.Visible = False
    End If
    
    Exit Sub
errHandler:
    HandleError "lstItem_Click", "frmShopEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).Name = (txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    Exit Sub
errHandler:
    HandleError "txtName_Validate", "frmShopEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtPrice_Change()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Data1Index = 0 Then Exit Sub
    If Not IsNumeric(txtPrice.Text) Then Exit Sub
    Shop(EditorIndex).sItem(Data1Index).Price = Val(txtPrice.Text)
    
    Exit Sub
errHandler:
    HandleError "txtPrice_Change", "frmShopEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
