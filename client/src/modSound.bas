Attribute VB_Name = "modSound"
Option Explicit

Public Const MUSIC_PATH As String = "\data\sfx\music\"
Public Const SOUND_PATH As String = "\data\sfx\sound\"

Public Const Menu_Music As String = "01-opening.mp3"
Public Const Battle_Wild_Music As String = "02-battlewild.mp3"
Public Const Victory_Wild_Music As String = "03-victorywild.mp3"
Public Const Evolve_Music As String = "06-evolution.mp3"

Public musicCache() As String
Public soundCache() As String
Public hasPopulated As Boolean

Public SoundIndex As Long
Public MusicIndex As Long

Public CurMusic As String
Public CurSound As String

Public SoundMusicOn As Boolean

Public Sub InitSound()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ChDrive(App.Path)
    Call ChDir(App.Path)

    If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
        Call MsgBox("An incorrect version of bass.dll was loaded.", vbCritical)
        End
    End If

    If (BASS_Init(-1, 44100, 0, frmMain.hWnd, 0) = 0) Then
        Call MsgBox("Failed to initialise the device.")
        End
    End If

    SoundMusicOn = True
    PopulateLists
    
    Exit Sub
errHandler:
    HandleError "InitSound", "modSound", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseSound()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If SoundMusicOn = False Then Exit Sub
    StopMusic
    StopSound
    Call BASS_Free
    
    Exit Sub
errHandler:
    HandleError "CloseSound", "modSound", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub StopSound()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If SoundMusicOn = False Then Exit Sub
    BASS_ChannelStop (SoundIndex)
    CurSound = vbNullString
    
    Exit Sub
errHandler:
    HandleError "StopSound", "modSound", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub StopMusic()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If SoundMusicOn = False Then Exit Sub
    BASS_ChannelStop (MusicIndex)
    CurMusic = vbNullString
    
    Exit Sub
errHandler:
    HandleError "StopMusic", "modSound", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayMusic(ByVal FileName As String)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Options.Music = 0 Then Exit Sub
    If SoundMusicOn = False Then Exit Sub
    If Not FileExist(App.Path & MUSIC_PATH & FileName) Then Exit Sub
    If CurMusic = FileName Then Exit Sub
    
    StopMusic

    MusicIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & MUSIC_PATH & FileName), 0, 0, BASS_SAMPLE_LOOP)
    If MusicIndex = 0 Then Exit Sub

    Call SetVolume(MusicIndex, 0.7)
    Call BASS_ChannelPlay(MusicIndex, False)
    
    CurMusic = FileName
    
    Exit Sub
errHandler:
    HandleError "PlayMusic", "modSound", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PlaySound(ByVal FileName As String)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Options.Sound = 0 Then Exit Sub
    If SoundMusicOn = False Then Exit Sub
    If Not FileExist(App.Path & SOUND_PATH & FileName) Then Exit Sub
    
    SoundIndex = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & SOUND_PATH & FileName), 0, 0, 0)
    If SoundIndex = 0 Then Exit Sub
    
    Call SetVolume(SoundIndex, 0.7)
    Call BASS_ChannelPlay(SoundIndex, False)

    CurSound = FileName
    
    Exit Sub
errHandler:
    HandleError "PlaySound", "modSound", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SetVolume(ByVal channel As Long, ByVal Volume As Double)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call BASS_ChannelSetAttribute(channel, BASS_ATTRIB_VOL, Volume)
    
    Exit Sub
errHandler:
    HandleError "SetVolume", "modSound", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateLists()
Dim strLoad As String, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    strLoad = dir(App.Path & MUSIC_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    
    strLoad = dir(App.Path & SOUND_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    
    Exit Sub
errHandler:
    HandleError "PopulateLists", "modSound", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
