Attribute VB_Name = "MMKeys"
'Modified by Tom Pydeski for use with GotRadio Player
'**************************************
' Name: HID MultiMedia Keyboard
' Description:Let's your code use the multi media keys from intelitype/hid
'compliant multimedia keyboards the way Windows media player does
' By: Techni Rei Myoko
'
' Inputs:First you must hook your form
'
' Side Effects:Responds to Multi media keys (play/pause, stop, prev & Next item, prev & next track)
'
'This code is copyrighted and has limited warranties.Please see
'http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.42033/lngWId.1/qx/vb/scripts/ShowCode.htm
'for details.
'**************************************
Option Explicit 'Prevents human Error
'Code to create HID Multimedia keyboard compliant applications
Public OldProc As Long 'Stores the location of the window handling procedure to be used, and restored upon closing
'Used to get, then set the window handling procedure
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_WNDPROC = (-4)
'Used to call the old window handling procedure
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_APPCOMMAND As Integer = 793 'Monitor Multimedia events
Const WM_SYSCOMMAND = &H112 'Monitor For close/kill
'Multimedia key constants
Public Const MMkey_Play As Long = 917504
Public Const MMkey_Stop As Long = 851968
Public Const MMkey_Prev_Item As Long = 65536
Public Const MMkey_Next_Item As Long = 131072
Public Const MMkey_Prev_Track As Long = 786432
Public Const MMkey_Next_Track As Long = 720896
Public Enum AppCommandConstants
    APPCOMMAND_BROWSER_BACKWARD = 1
    APPCOMMAND_BROWSER_FORWARD = 2
    APPCOMMAND_BROWSER_REFRESH = 3
    APPCOMMAND_BROWSER_STOP = 4
    APPCOMMAND_BROWSER_SEARCH = 5
    APPCOMMAND_BROWSER_FAVORITES = 6
    APPCOMMAND_BROWSER_HOME = 7
    APPCOMMAND_VOLUME_MUTE = 8
    APPCOMMAND_VOLUME_DOWN = 9
    APPCOMMAND_VOLUME_UP = 10
    APPCOMMAND_MEDIA_NEXTTRACK = 11
    APPCOMMAND_MEDIA_PREVIOUSTRACK = 12
    APPCOMMAND_MEDIA_STOP = 13
    APPCOMMAND_MEDIA_PLAY_PAUSE = 14
    APPCOMMAND_LAUNCH_MAIL = 15
    APPCOMMAND_LAUNCH_MEDIA_SELECT = 16
    APPCOMMAND_LAUNCH_APP1 = 17
    APPCOMMAND_LAUNCH_APP2 = 18
    APPCOMMAND_BASS_DOWN = 19
    APPCOMMAND_BASS_BOOST = 20
    APPCOMMAND_BASS_UP = 21
    APPCOMMAND_TREBLE_DOWN = 22
    APPCOMMAND_TREBLE_UP = 23
    APPCOMMAND_MICROPHONE_VOLUME_MUTE = 24
    APPCOMMAND_MICROPHONE_VOLUME_DOWN = 25
    APPCOMMAND_MICROPHONE_VOLUME_UP = 26
    APPCOMMAND_HELP = 27
    APPCOMMAND_FIND = 28
    APPCOMMAND_NEW = 29
    APPCOMMAND_OPEN = 30
    APPCOMMAND_CLOSE = 31
    APPCOMMAND_SAVE = 32
    APPCOMMAND_PRINT = 33
    APPCOMMAND_UNDO = 34
    APPCOMMAND_REDO = 35
    APPCOMMAND_COPY = 36
    APPCOMMAND_CUT = 37
    APPCOMMAND_PASTE = 38
    APPCOMMAND_REPLY_TO_MAIL = 39
    APPCOMMAND_FORWARD_MAIL = 40
    APPCOMMAND_SEND_MAIL = 41
    APPCOMMAND_SPELL_CHECK = 42
    APPCOMMAND_DICTATE_OR_COMMAND_CONTROL_TOGGLE = 43
    APPCOMMAND_MIC_ON_OFF_TOGGLE = 44
    APPCOMMAND_CORRECTION_LIST = 45
End Enum
Public Enum AppCommandDeviceConstants
    FAPPCOMMAND_MOUSE = &H8000&
    FAPPCOMMAND_KEY = 0
    FAPPCOMMAND_OEM = &H1000&
End Enum
Public Enum AppCommandKeyStateConstants
    MK_LBUTTON = &H1
    MK_RBUTTON = &H2
    MK_SHIFT = &H4
    MK_CONTROL = &H8
    MK_MBUTTON = &H10
    MK_XBUTTON1 = &H20
    MK_XBUTTON2 = &H40
End Enum
Private Const FAPPCOMMAND_MASK As Long = &HF000&
'Used to get the parameters needed to monitor for a close program event
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'Usage: On FormLoad, run hook(me.hwnd), and on formunload run hook(me.hwnd, false)
Public Declare Sub CopyMemoryH Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)
Dim eM$
Global AppCommand$(256)
Dim AppKeyName$
Dim LastAppKey As Long

Sub LoadAppKeys()
'this routine was from Tom Pydeski
Dim i As Integer
'set appcommands
For i = 0 To 255
    AppCommand$(i) = "Unk " & i
Next i
'
AppCommand$(1) = "Browser_Backward"
AppCommand$(2) = "Browser_Forward"
AppCommand$(3) = "Browser_Refresh"
AppCommand$(4) = "Browser_Stop"
AppCommand$(5) = "Browser_Search"
AppCommand$(6) = "Browser_Favorites"
AppCommand$(7) = "Browser_Home"
AppCommand$(8) = "Volume_Mute"
AppCommand$(9) = "Volume_Down"
AppCommand$(10) = "Volume_Up"
AppCommand$(11) = "Media_Nexttrack"
AppCommand$(12) = "Media_Previoustrack"
AppCommand$(13) = "Media_Stop"
AppCommand$(14) = "Media_Play_Pause"
AppCommand$(15) = "Launch_Mail"
AppCommand$(16) = "Launch_Media_Select"
AppCommand$(17) = "Launch_App1"
AppCommand$(18) = "Launch_App2"
AppCommand$(19) = "Bass_Down"
AppCommand$(20) = "Bass_Boost"
AppCommand$(21) = "Bass_Up"
AppCommand$(22) = "Treble_Down"
AppCommand$(23) = "Treble_Up"
AppCommand$(24) = "Microphone_Volume_Mute"
AppCommand$(25) = "Microphone_Volume_Down"
AppCommand$(26) = "Microphone_Volume_Up"
AppCommand$(27) = "Help"
AppCommand$(28) = "Find"
AppCommand$(29) = "New"
AppCommand$(30) = "Open"
AppCommand$(31) = "Close"
AppCommand$(32) = "Save"
AppCommand$(33) = "Print"
AppCommand$(34) = "Undo"
AppCommand$(35) = "Redo"
AppCommand$(36) = "Copy"
AppCommand$(37) = "Cut"
AppCommand$(38) = "Paste"
AppCommand$(39) = "Reply_To_Mail"
AppCommand$(40) = "Forward_Mail"
AppCommand$(41) = "Send_Mail"
AppCommand$(42) = "Spell_Check"
AppCommand$(43) = "Dictate_Or_Command_Control_Toggle"
AppCommand$(44) = "Mic_On_Off_Toggle"
AppCommand$(45) = "Correction_List"
End Sub

Public Sub Hook(ByVal hwnd As Long, Optional state As Boolean = True)
If state = True Then 'Stores the old Wndproc, and sets the new to our own
    OldProc = GetWindowLong(hwnd, GWL_WNDPROC)
    SetWindowLong hwnd, GWL_WNDPROC, AddressOf WndProc
Else 'Sets the old WndProc To defualt again
    SetWindowLong hwnd, GWL_WNDPROC, OldProc
End If
End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim MenuHandle As Long, MenuCloseID As Long
'Debug.Print wMsg, wParam, lParam
eM$ = ""
If wMsg = WM_APPCOMMAND Then
    eM$ = "Received app command - "
    If OnMMKeypress(lParam) = True Then
        'If the Function returns true then the key pressed was recognized
        WndProc = 0
        wMsg = 0
        eM$ = eM$ & AppKeyName$
        'Exit Function
    Else
        'Un comment these lines to check for other keys pressed
        'to get their MMKey constants
        eM$ = "An unknown key was pressed!" & vbNewLine & lParam & "(" & Hex$(lParam) & ")"
        eM$ = eM$ & vbNewLine & HiWord(lParam) & "(" & Hex$(HiWord(lParam)) & ")"
        eM$ = eM$ & vbNewLine & LoWord(lParam) & "(" & Hex$(LoWord(lParam)) & ")"
        eM$ = eM$ & vbNewLine & AppCommand$(HiWord(lParam))
        MsgBox eM$, vbInformation, "Unknown Media Key"
        'Clipboard.Clear
        'Clipboard.SetText lParam
    End If
End If
If Len(eM$) > 0 Then
    Debug.Print eM$
    'vKB.lblAppCommand.Visible = True
    'vKB.lblAppCommand.Caption = eM$
    'vKB.Timer1.Enabled = True
    'MsgBox eM$, vbInformation, "AppCommand Keyboard Hook"
End If
'MenuHandle = GetSystemMenu(hwnd, False) 'Get system menu
'MenuCloseID = GetMenuItemID(MenuHandle, 6) 'Get item 6 (close) of system menu
'If wMsg = WM_SYSCOMMAND And wParam = MenuCloseID Then
'    Call Hook(hwnd, False) 'Unhook If close is detected
'End If
WndProc = CallWindowProc(OldProc, hwnd, wMsg, wParam, lParam) 'Pass To old If we have done nothing
End Function

Public Function OnMMKeypress(MMKeyCode As Long) As Boolean
'Place code For Multimedia functions With the Case statements
OnMMKeypress = True
AppKeyName$ = ""
Debug.Print "OnMMKeypress:"; "MMKeyCode= "; MMKeyCode; " HiWord= "; HiWord(MMKeyCode); " LoWord= "; LoWord(MMKeyCode); " AppCommand$= "; AppCommand$(HiWord(MMKeyCode)); ""
Select Case MMKeyCode 'Monitor which key was pressed
    Case MMkey_Play
        'my logitech wireless keyboard uses the same key to toggle
        'between pause and play.  for some reason, it would cause a pause
        'immediately followed by a play
        'we need to capture it once and ignore it the second pass
        AppKeyName$ = "Play"
        If LastAppKey = MMkey_Play Then
            LastAppKey = 0
            Debug.Print "got it again..."
            Exit Function
        End If
        LastAppKey = MMKeyCode
        If frmGotRadio.MPlayer.playState = wmppsPlaying Then
            frmGotRadio.MPlayer.Controls.pause
            Debug.Print "Pausing..."
            Exit Function
        Else
            frmGotRadio.MPlayer.Controls.play
            Debug.Print "Resuming..."
            Exit Function
        End If
        MMKeyCode = 0
    Case MMkey_Stop
        AppKeyName$ = "Stop"
        frmGotRadio.MPlayer.Controls.stop
    Case MMkey_Prev_Item
        AppKeyName$ = "Prev_Item"
    Case MMkey_Next_Item
        AppKeyName$ = "Next_Item"
    Case MMkey_Prev_Track
        AppKeyName$ = "Prev_Track"
        frmGotRadio.MPlayer.Controls.previous
    Case MMkey_Next_Track
        AppKeyName$ = "Next_Track"
        frmGotRadio.MPlayer.Controls.Next
    Case Else
        'OnMMKeypress = False
        AppKeyName$ = AppCommand$(HiWord(MMKeyCode))
End Select
ExitKey:
'vKB.Text1 = "App Key " & AppCommand$(HiWord(MMKeyCode)) & vbCrLf & vKB.Text1
Debug.Print "**************App Key " & AppCommand$(HiWord(MMKeyCode)) & AppKeyName$
End Function

Private Function LoWord(ByVal dw As Long) As Integer
On Error GoTo Err_LoWord
CopyMemoryH LoWord, ByVal VarPtr(dw), 2
Exit_LoWord:
Exit Function
Err_LoWord:
GoTo Exit_LoWord
End Function

Private Function HiWord(ByVal dw As Long) As Integer
On Error GoTo Err_HiWord
CopyMemoryH HiWord, ByVal VarPtr(dw) + 2, 2
Exit_HiWord:
Exit Function
Err_HiWord:
GoTo Exit_HiWord
End Function
