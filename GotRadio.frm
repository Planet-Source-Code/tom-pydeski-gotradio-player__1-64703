VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmGotRadio 
   Caption         =   "Tom Pydeski's GotRadio Player.  Double Click a format from the list."
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   Icon            =   "GotRadio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6900
      IntegralHeight  =   0   'False
      Left            =   4785
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5040
      Top             =   3480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2535
      Left            =   45
      ScaleHeight     =   2475
      ScaleWidth      =   3195
      TabIndex        =   3
      Top             =   5010
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   7710
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "GotRadio.frx":030A
      Top             =   50
      Width           =   3975
   End
   Begin VB.Label lblFormat 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Double Click to Select a format..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   435
      Left            =   45
      TabIndex        =   6
      Top             =   45
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4755
      TabIndex        =   5
      Top             =   0
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblinfo 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   45
      TabIndex        =   4
      Top             =   3885
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin WMPLibCtl.WindowsMediaPlayer MPlayer 
      Height          =   3375
      Left            =   45
      TabIndex        =   1
      Top             =   480
      Width           =   4695
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8281
      _cy             =   5953
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmGotRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'submitted by Tom Pydeski
'I like to listen to the GotRadio stations at http://www.gotradio.com/,
'but I don't like the ads constantly cluttering up the screen.
'This application utilizes the windows media player control to connect
'to the desired station.  It also retrieves all of the info about the
'song being played and downloads the album cover picture and displays it.
'I also implemented some code to hook the keyboard and utilize the multimedia
'keys (Play, FF, Stop, etc.).  There are 46 channels to choose from and the only
'ads are the audio/video type, not the banners from the website.
'thanks to Christopher Lord for the tray class
'
Dim AID As Long
Dim albumID As Integer
Dim Artist$
Dim Author$
Dim Label$
Dim Title$
Dim adType$
Dim oldTitle$
Dim Album$
Dim SmallCover$
Dim MedCover$
Dim LargeCover$
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
Dim Capt$
Dim sCapt$
Dim LeftChar
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Const LB_ITEMFROMPOINT = &H1A9
Dim i As Integer
Dim eTitle$
Dim EMess$
Dim mError As Long
Private WithEvents Tray As clsTray
Attribute Tray.VB_VarHelpID = -1
Private SaveState As Integer

Private Sub Form_Load()
Me.MousePointer = 11
' Setup how we want the task tray to work and display
Set Tray = New clsTray
' Initialize settings here
Tray.Initialize Me
Tray.AutoRefresh = True
Tray.Tooltip = "Tom Pydeski's GotRadio Player" & vbCrLf & "Double Click a format" & vbCrLf & "to start the Player."
Tray.AddIcon
Tray.Refresh
Tray.ShowBalloonTip Tray.Tooltip, "Tom Pydeski's GotRadio Player", NIIF_INFO + NIIF_NOSOUND, 1000
SaveState = Me.WindowState
'hook the keyboard for usage of the media keys
Hook (Me.hwnd)
'set our normal window size
Me.Width = 4800
Me.Height = 10000
Me.Show
WindowState = vbMaximized
'Fill the list with the channels
List1.Clear
List1.AddItem "Adult Alternative"
List1.ItemData(List1.NewIndex) = 12
List1.AddItem "Adult Contemporary"
List1.ItemData(List1.NewIndex) = 14
List1.AddItem "Alternative Rock"
List1.ItemData(List1.NewIndex) = 47
List1.AddItem "Big Band and Swing"
List1.ItemData(List1.NewIndex) = 54
List1.AddItem "Bluegrass"
List1.ItemData(List1.NewIndex) = 15
List1.AddItem "Blues"
List1.ItemData(List1.NewIndex) = 16
List1.AddItem "Celtic"
List1.ItemData(List1.NewIndex) = 46
List1.AddItem "Christian Contemporary"
List1.ItemData(List1.NewIndex) = 17
List1.AddItem "Christmas Celebration"
List1.ItemData(List1.NewIndex) = 61
List1.AddItem "Classic 60s"
List1.ItemData(List1.NewIndex) = 71
List1.AddItem "Classic Country"
List1.ItemData(List1.NewIndex) = 70
List1.AddItem "Classic Hits"
List1.ItemData(List1.NewIndex) = 19
List1.AddItem "Classic Rock"
List1.ItemData(List1.NewIndex) = 22
List1.AddItem "Classical"
List1.ItemData(List1.NewIndex) = 21
List1.AddItem "Country"
List1.ItemData(List1.NewIndex) = 23
List1.AddItem "Dance"
List1.ItemData(List1.NewIndex) = 24
List1.AddItem "Disco"
List1.ItemData(List1.NewIndex) = 25
List1.AddItem "Electronica"
List1.ItemData(List1.NewIndex) = 26
List1.AddItem "Folk"
List1.ItemData(List1.NewIndex) = 27
List1.AddItem "Forever Fifties"
List1.ItemData(List1.NewIndex) = 53
List1.AddItem "Halloween Rock"
List1.ItemData(List1.NewIndex) = 59
List1.AddItem "Hip Hop"
List1.ItemData(List1.NewIndex) = 28
List1.AddItem "Hot Hits"
List1.ItemData(List1.NewIndex) = 73
List1.AddItem "Indie Rock"
List1.ItemData(List1.NewIndex) = 30
List1.AddItem "Jazz"
List1.ItemData(List1.NewIndex) = 31
List1.AddItem "Mash-Ups"
List1.ItemData(List1.NewIndex) = 63
List1.AddItem "Metal Rock"
List1.ItemData(List1.NewIndex) = 34
List1.AddItem "Musical Magic"
List1.ItemData(List1.NewIndex) = 55
List1.AddItem "Native American"
List1.ItemData(List1.NewIndex) = 36
List1.AddItem "New Age"
List1.ItemData(List1.NewIndex) = 35
List1.AddItem "R&B Classics"
List1.ItemData(List1.NewIndex) = 37
List1.AddItem "Reggae"
List1.ItemData(List1.NewIndex) = 38
List1.AddItem "Retro Radio"
List1.ItemData(List1.NewIndex) = 29
List1.AddItem "Rock"
List1.ItemData(List1.NewIndex) = 39
List1.AddItem "Rockin 80's"
List1.ItemData(List1.NewIndex) = 40
List1.AddItem "Smooth Jazz"
List1.ItemData(List1.NewIndex) = 41
List1.AddItem "Soundtracks"
List1.ItemData(List1.NewIndex) = 64
List1.AddItem "Top 40"
List1.ItemData(List1.NewIndex) = 20
List1.AddItem "Top Alternative 2003"
List1.ItemData(List1.NewIndex) = 52
List1.AddItem "Top Hits 2003"
List1.ItemData(List1.NewIndex) = 51
List1.AddItem "Top Hits 2004"
List1.ItemData(List1.NewIndex) = 62
List1.AddItem "Top Hits 2005"
List1.ItemData(List1.NewIndex) = 72
List1.AddItem "Urban"
List1.ItemData(List1.NewIndex) = 42
List1.AddItem "Vintage Vault"
List1.ItemData(List1.NewIndex) = 57
List1.AddItem "Women's Alternative"
List1.ItemData(List1.NewIndex) = 43
List1.AddItem "World"
List1.ItemData(List1.NewIndex) = 44
'set the listindex to the last channel selected
List1.ListIndex = Val(GetSetting(App.EXEName, "Settings", "Channel"))
Me.MousePointer = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
' This will place an icon in the task tray when the user minimizes this form
If Me.WindowState = 1 Then
    ' Form was minimized
    Me.Hide
    Exit Sub
End If
SaveState = Me.WindowState
Text1.Width = (Me.Width / Screen.TwipsPerPixelX) - Text1.Left - 15
Text1.Height = ((Me.Height / Screen.TwipsPerPixelY) - Text1.Top) - 40
List1.Height = ((Me.Height / Screen.TwipsPerPixelY) - List1.Top) - 40
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Hook(hwnd, False) 'Unhook
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = vbKeyRight
        'get the next song
        MPlayer.Controls.Next
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case UCase$(Chr$(KeyAscii))
    'use the keyboard to control the player
    Case Is = "P"
        'Pause/Play
        If MPlayer.playState = wmppsPlaying Then
            frmGotRadio.MPlayer.Controls.pause
            Debug.Print "Pausing..."
        Else
            MPlayer.Controls.play
            Debug.Print "Resuming..."
        End If
        KeyAscii = 0
    Case Is = "S"
        'Stop
        MPlayer.Controls.stop
    Case Is = "B"
        'Back
        MPlayer.Controls.previous
    Case Is = "N"
        'Next
        MPlayer.Controls.Next
    Case Is = "F"
        'FastForward
        MPlayer.Controls.fastForward
    Case Is = "R"
        'Rewind
        MPlayer.Controls.fastReverse
    Case Is = "Q"
        'let's quit...
        MPlayer.Controls.stop
        Unload Me
End Select
End Sub

Private Sub List1_DblClick()
Dim lIndex As Integer
Me.MousePointer = 11
'double clicking selects the channel to be played
lIndex = List1.ListIndex
'display the format
lblFormat = List1.List(lIndex)
'save the channel
SaveSetting App.EXEName, "Settings", "Channel", lIndex
'launch the channel in the media player
MPlayer.URL = "http://www.gotradio.com/player/launch.asp?refer=web&id=" & List1.ItemData(lIndex)
MPlayer.SetFocus
Tray.Tooltip = "Connecting to " & lblFormat
Tray.Refresh
Tray.ShowBalloonTip Tray.Tooltip, "Tom Pydeski's GotRadio Player", NIIF_INFO + NIIF_NOSOUND, 1000
Me.MousePointer = 0
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
'the return key launches the channel
If KeyAscii = 13 Then
    List1_DblClick
End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
List1.ZOrder 0
' present related tip message
Dim lXPoint As Long
Dim lYPoint As Long
Dim lIndex As Long
If Button = 0 Then ' if no button was pressed
    lXPoint = CLng(X / Screen.TwipsPerPixelX)
    lYPoint = CLng(Y / Screen.TwipsPerPixelY)
    With List1
        ' get selected item from list
        lIndex = SendMessage(.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
        ' show tip or clear last one
        If (lIndex >= 0) And (lIndex <= .ListCount) Then
            .ToolTipText = "Select " & .List(lIndex)
        Else
            .ToolTipText = ""
        End If
    End With '(List1)
End If '(button=0)
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub MPlayer_Buffering(ByVal Start As Boolean)
lblStatus = "Buffering"
Debug.Print lblStatus
End Sub

Private Sub MPlayer_CurrentItemChange(ByVal pdispMedia As Object)
lblStatus = "CurrentItemChange"
Debug.Print lblStatus
End Sub

Private Sub MPlayer_CurrentMediaItemAvailable(ByVal bstrItemName As String)
lblStatus = "CurrentMediaItemAvailable"
'Debug.Print lblStatus
GetInfo
End Sub

Private Sub MPlayer_DeviceConnect(ByVal pDevice As WMPLibCtl.IWMPSyncDevice)
lblStatus = "DeviceConnect"
Debug.Print lblStatus
End Sub

Private Sub MPlayer_DeviceStatusChange(ByVal pDevice As WMPLibCtl.IWMPSyncDevice, ByVal NewStatus As WMPLibCtl.WMPDeviceStatus)
lblStatus = "DeviceStatusChange"
Debug.Print lblStatus
End Sub

Private Sub MPlayer_MediaChange(ByVal Item As Object)
lblStatus = "MediaChange " '; Item
Debug.Print lblStatus
End Sub

Private Sub MPlayer_NewStream()
lblStatus = "NewStream"
Debug.Print lblStatus
End Sub

Private Sub MPlayer_PlayStateChange(ByVal NewState As Long)
If NewState = wmppsPlaying Then '3
    'if the player is playing, fill in the song info
    GetInfo
ElseIf NewState = wmppsTransitioning Then '9
    'clear the song info if the song is done
    ClearInfo
End If
lblStatus = "PlayStateChange = " & NewState
Debug.Print lblStatus
End Sub

Private Sub MPlayer_StatusChange()
lblStatus = "StatusChange"
'Debug.Print lblStatus
End Sub

Sub ClearInfo()
lblinfo = ""
Picture1.Visible = False
Picture1.Picture = LoadPicture("")
End Sub

Sub GetInfo()
On Error GoTo Oops
Me.MousePointer = 11
With MPlayer.currentMedia
    'retrieve the info on the currently playing song
    AID = Val(.getItemInfo("AID"))
    Author$ = .getItemInfo("AUTHOR")
    Artist$ = .getItemInfo("Artist")
    Label$ = .getItemInfo("COPYRIGHT")
    Title$ = .getItemInfo("TITLE")
    adType$ = .getItemInfo("ADTYPE")
    'check if the current media being played is an ad
    If LCase(adType$) <> "none" Then
        'we have an ad, so let's skip it
        Me.Caption = "Skipping Ad..."
        MPlayer.Controls.Next
    End If
    If MPlayer.Controls.isAvailable("FastForward") = False Then
        'we can't skip this one and I don't think there's any way around it
    End If
    If oldTitle$ <> Title$ Then
        'this is a new song... change info
        albumID = Val(.getItemInfo("ALBUMID"))
        'clear the album cover picture
        Picture1.Picture = LoadPicture("")
        'clear our data
        Album$ = ""
        SmallCover$ = ""
        MedCover$ = ""
        LargeCover$ = ""
        'check if we have a valid albumID
        If albumID <> 0 Then
            'it is a valid song, so now let's get the rest of the info
            Album$ = .getItemInfo("ALBUM")
            'get the 3 pictures that are supplied for the album art
            SmallCover$ = .getItemInfo("SCOVER")
            MedCover$ = Replace(.getItemInfo("MCOVER"), "LZ", "MZ", , , vbTextCompare)
            LargeCover$ = .getItemInfo("LCOVER")
            'hide the picture so it loads nicer
            Picture1.Visible = False
            'get the large album cover picture
            Picture1.Picture = OLELoadPicture(LargeCover$)
            'check picture size and if the picture is too big, get a smaller one
            If Picture1.Top + Picture1.Height > Me.Height / Screen.TwipsPerPixelY Then
                'try getting the next size down...
                Picture1.Picture = OLELoadPicture(MedCover$)
            End If
            'show the picture
            Picture1.Visible = True
            'put the picture at the top of the z order in case it overlaps the list box
            Picture1.ZOrder 0
        End If
        'put together the info into the label
        lblinfo = "Title: " & Title$ & vbCrLf
        lblinfo = lblinfo & "Artist: " & Artist$ & vbCrLf
        lblinfo = lblinfo & "Album: " & Album$ & vbCrLf
        lblinfo = lblinfo & "Author: " & Author$ & vbCrLf
        Tray.Tooltip = lblinfo
        'i don't think the timeout works...
        Tray.Refresh
        'instead of showing the balloon, which does not go away,  just set the tip
        'Tray.ShowBalloonTip Tray.Tooltip, "Tom Pydeski's GotRadio Player", NIIF_INFO + NIIF_NOSOUND, 1000
        'Tray.Refresh
        lblinfo = Replace(lblinfo, "&", "&&")
        'now put together the info we want on the window caption
        Capt$ = Title$ & " ; Artist = " & Artist & "; Album = " & Album
        'now display that info in the text box
        Text1.Text = lblinfo & "----------------------------------------" & vbCrLf
        'find out how many items there are to read and set up a loop to get them all
        For i = 0 To .attributeCount - 1
            'add each attribute's name to the text box
            Text1.Text = Text1.Text & i & " " & .getAttributeName(i) & " = "
            'now get the data for that attribute and put it in the text box also
            Text1.Text = Text1.Text & .getItemInfo(.getAttributeName(i)) & vbCrLf
        Next i
        oldTitle$ = Title$
    End If
End With
GoTo Exit_GetInfo
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GetInfo "
EMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in GetInfo"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GetInfo:
Me.MousePointer = 0
End Sub

Public Function OLELoadPicture(ByVal strFilename As String) As Picture
On Error GoTo Oops
'This function gets a picture from a url path
Dim myTGUID As TGUID
myTGUID.Data1 = &H7BF80980
myTGUID.Data2 = &HBF32
myTGUID.Data3 = &H101A
myTGUID.Data4(0) = &H8B
myTGUID.Data4(1) = &HBB
myTGUID.Data4(2) = &H0
myTGUID.Data4(3) = &HAA
myTGUID.Data4(4) = &H0
myTGUID.Data4(5) = &H30
myTGUID.Data4(6) = &HC
myTGUID.Data4(7) = &HAB
OleLoadPicturePath StrPtr(strFilename), 0, 0, 0, myTGUID, OLELoadPicture
GoTo Exit_LoadPicture
LblError:
Set OLELoadPicture = VB.LoadPicture(strFilename)
GoTo Exit_LoadPicture
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine LoadPicture "
EMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in LoadPicture"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_LoadPicture:
End Function

Private Sub Picture1_DblClick()
Picture1.Picture = OLELoadPicture(MedCover$)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.ZOrder 0
End Sub

Private Sub Timer1_Timer()
If Capt$ = "" Then Exit Sub
'scroll the info on the minimized window caption
If Me.WindowState = vbMinimized Then
    LeftChar = LeftChar + 1
    sCapt$ = Mid$(Capt$, LeftChar)
    If Len(sCapt$) = 0 Then
        sCapt$ = Capt$
        LeftChar = 0
    End If
    Me.Caption = sCapt$
Else
    Me.Caption = Capt$
End If
Me.Refresh
DoEvents
End Sub

Private Sub mnuShow_Click()
' This is a command from the menu called show,
'all it does is restore the form to normal
'Let's leave it in the tray
'Tray.RemoveIcon
' Reshow the form here
Me.WindowState = SaveState
Me.Show
End Sub

Private Sub Tray_DoubleClick(Button As Integer)
' If they double click with the left mouse
' button we will simply show the form
If Button = 0 Then
    ' Return to normal
    Call mnuShow_Click
End If
End Sub

Private Sub Tray_MouseDown(Button As Integer)
' If they right click on the task tray then
' we will simply show them a popup menu
If Button = 1 Then
    ' And popup the menu
    PopupMenu mnuPopup, , , , mnuShow
End If
End Sub

