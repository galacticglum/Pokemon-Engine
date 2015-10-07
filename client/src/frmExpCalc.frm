VERSION 5.00
Begin VB.Form frmExpCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exp Calc Editor"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtExpCalc 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   3720
      Width           =   2655
   End
   Begin VB.ListBox lstExpCalc 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmExpCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private xIndex As Long

Private Sub cmdSave_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    SendExpCalc
    AddText "Saving ExpCalc Complete!", Green
    
    Exit Sub
errHandler:
    HandleError "cmdSave_Click", "frmExpCalc", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    lstExpCalc.Clear
    For i = 1 To MAX_LEVEL
        lstExpCalc.AddItem "Lv" & i & ": " & ExpCalc(i)
    Next
    lstExpCalc.ListIndex = 0
    xIndex = lstExpCalc.ListIndex + 1
    txtExpCalc.Text = ExpCalc(xIndex)
    
    Exit Sub
errHandler:
    HandleError "Form_Load", "frmExpCalc", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstExpCalc_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    xIndex = lstExpCalc.ListIndex + 1
    txtExpCalc.Text = ExpCalc(xIndex)
    
    Exit Sub
errHandler:
    HandleError "lstExpCalc_Click", "frmExpCalc", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lstExpCalc_GotFocus()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    txtExpCalc.SetFocus
    txtExpCalc.SelStart = Len(txtExpCalc.Text)
    
    Exit Sub
errHandler:
    HandleError "lstExpCalc_GotFocus", "frmExpCalc", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub txtExpCalc_Validate(Cancel As Boolean)
Dim i As Byte, tmpIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If IsNumeric(txtExpCalc.Text) Then
        ExpCalc(xIndex) = Val(txtExpCalc.Text)
        tmpIndex = lstExpCalc.ListIndex
        lstExpCalc.RemoveItem xIndex - 1
        lstExpCalc.AddItem "Lv" & xIndex & ": " & ExpCalc(xIndex), xIndex - 1
        lstExpCalc.ListIndex = tmpIndex
    End If
    
    Exit Sub
errHandler:
    HandleError "txtExpCalc_Validate", "frmExpCalc", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
