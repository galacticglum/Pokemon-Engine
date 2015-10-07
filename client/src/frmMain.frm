VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            If ConnectToServer Then
                SendLogin "kirito", "test"
            End If
        Case vbKeyF2
            If ConnectToServer Then
                SendLogin "test", "test"
            End If
        Case vbKeyF3
            If InTradeIndex > 0 Then
                SendTradeAccept
            End If
    End Select
End Sub

Private Sub Form_Load()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Me.Caption = GameTitle
    Me.Width = FormMainWidth
    Me.Height = FormMainHeight
    
    Exit Sub
errHandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CloseApp ' Temporary Only (Change if the main initiation added)
    
    Exit Sub
errHandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
    
    Exit Sub
errHandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    HandleMouseDown Button, Shift, X, Y
    IsClicked = 1
    
    Exit Sub
errHandler:
    HandleError "Form_MouseDown", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    HandleMouseUp Button, Shift, X, Y
    IsClicked = 0
    
    Exit Sub
errHandler:
    HandleError "Form_MouseUp", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    HandleMouseMove Button, Shift, X, Y
    
    Exit Sub
errHandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
Private Sub Form_Click()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    HandleClick
    
    Exit Sub
errHandler:
    HandleError "Form_Click", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If InMenu Then
        HandleMenuKeyPress KeyAscii
        
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
            KeyAscii = 0
        End If
    End If
    If InGame Then
        HandleMainKeyPress KeyAscii
        
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
            KeyAscii = 0
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
