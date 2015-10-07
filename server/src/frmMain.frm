VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Log"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtLog"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdReloadMoves"
      Tab(1).Control(1)=   "cmdReloadPokemons"
      Tab(1).Control(2)=   "cmdReloadMaps"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Players"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwInfo"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdReloadMoves 
         Caption         =   "Reload Moves"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton cmdReloadPokemons 
         Caption         =   "Reload Pokemons"
         Height          =   375
         Left            =   -74760
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdReloadMaps 
         Caption         =   "Reload Maps"
         Height          =   375
         Left            =   -74760
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtLog 
         Height          =   2295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   480
         Width           =   5655
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4048
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
   End
   Begin VB.Menu mnuPlayer 
      Caption         =   "&Player"
      Visible         =   0   'False
      Begin VB.Menu mnuMakeAdmin 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub cmdReloadMaps_Click()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    LoadMaps
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next
    AddLog "All maps reloaded..."
    
    Exit Sub
errHandler:
    HandleError "cmdReloadMaps_Click", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdReloadPokemons_Click()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Call LoadPokemons
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            SendPokemons i
        End If
    Next
    AddLog "All pokemons reloaded..."
    
    Exit Sub
errHandler:
    HandleError "cmdReloadPokemons_Click", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdReloadMoves_Click()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Call LoadMoves
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            SendMoves i
        End If
    Next
    AddLog "All moves reloaded..."
    
    Exit Sub
errHandler:
    HandleError "cmdReloadMoves_Click", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    frmMain.Caption = AppTitle
    UsersOnline_Start
    
    Exit Sub
errHandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CloseApp
    
    Exit Sub
errHandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call AcceptConnection(index, requestID)
    
    Exit Sub
errHandler:
    HandleError "Socket_ConnectionRequest", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call AcceptConnection(index, SocketId)
    
    Exit Sub
errHandler:
    HandleError "Socket_Accept", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If
    
    Exit Sub
errHandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Socket_Close(index As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call CloseSocket(index)
    
    Exit Sub
errHandler:
    HandleError "Socket_Close", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If
    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
    
    Exit Sub
errHandler:
    HandleError "lvwInfo_ColumnClick", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub UsersOnline_Start()
Dim i As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To MAX_PLAYER
        lvwInfo.ListItems.Add (i)
        If i < 10 Then
            lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            lvwInfo.ListItems(i).Text = "0" & i
        Else
            lvwInfo.ListItems(i).Text = i
        End If
        lvwInfo.ListItems(i).SubItems(1) = vbNullString
        lvwInfo.ListItems(i).SubItems(2) = vbNullString
        lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next

    Exit Sub
errHandler:
    HandleError "UsersOnline_Start", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Button = vbRightButton Then
        PopupMenu mnuPlayer
    End If

    Exit Sub
errHandler:
    HandleError "lvwInfo_MouseDown", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuMakeAdmin_Click()
Dim Name As String
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Name = Trim$(lvwInfo.SelectedItem.SubItems(3))
    i = FindPlayer(Name)
    If i > 0 Then
        Player(i).PlayerData(TempPlayer(i).CurSlot).Access = 4
        Call SendPlayerData(i)
    End If

    Exit Sub
errHandler:
    HandleError "mnuMakeAdmin_Click", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuRemoveAdmin_Click()
Dim Name As String
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Name = Trim$(lvwInfo.SelectedItem.SubItems(3))
    i = FindPlayer(Name)
    If i > 0 Then
        Player(i).PlayerData(TempPlayer(i).CurSlot).Access = 0
        Call SendPlayerData(i)
    End If

    Exit Sub
errHandler:
    HandleError "mnuRemoveAdmin_Click", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub mnuDisconnect_Click()
Dim Name As String
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Name = Trim$(lvwInfo.SelectedItem.SubItems(3))
    i = FindPlayer(Name)
    If i > 0 Then
        CloseSocket i
    End If

    Exit Sub
errHandler:
    HandleError "mnuDisconnect_Click", "frmMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
