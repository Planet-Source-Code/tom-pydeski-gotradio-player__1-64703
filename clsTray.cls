VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Christopher Lord
' heenix@project-angel.net
' 17th March 2006
'
' You are free to reuse this code, in segments or whole, in
' any projects that you may want. As long as the name and
' date of the original creator is shown in one way or another.
'
' * Added Balloon Tips
' * Changed Icon Structure
' * Added New Events
' The data type for the icon in
' side the task bar, very simple
'
'usage...declare it in the form as follows:
'Private WithEvents Tray As clsTray
'Private SaveState As Integer
'
'Private Sub Form_Load()
'' Setup how we want the
'' task tray to work and display
'Set Tray = New clsTray
'' Initialize settings here
'Tray.Initialize Me
'Tray.AutoRefresh = True
'Tray.Tooltip = "Downloading 3 Files" & vbNewLine & "Speed: 17.9kb/s"
'SaveState = Me.WindowState
'End Sub
'
'Private Sub Form_Resize()
'' This will place an icon in the task
'' tray when the user minimizes this form
'If Me.WindowState = 1 Then
'    ' Form was minimised
'    Tray.AddIcon
'    Me.Hide
'    Exit Sub
'End If
'SaveState = Me.WindowState
'End Sub
'
Private Type NOTIFYICONDATAW
    icoSize As Long
    icoHwnd As Long
    icoId As Long
    icoFlags As Long
    icoCallbackMessage As Long
    icoSource As Long
    icoToolTip(0 To 255) As Byte
    icoState As Long
    icoStateMask As Long
    szInfo(0 To 511) As Byte
    uTimeOutOrVersion As Long
    szInfoTitle(0 To 127) As Byte
    dwInfoFlags As Long
End Type
' The structure that contains all
' the possible types of balloons
Public Enum EBalloonIconTypes
    NIIF_NONE = 0
    NIIF_INFO = 1
    NIIF_WARNING = 2
    NIIF_ERROR = 3
    NIIF_GUID = &H5
    NIIF_ICON_MASK = &HF
    NIIF_NOSOUND = &H10
End Enum
' Events that this class can raise here
Event DoubleClick(Button As Integer)
Event MouseDown(Button As Integer)
Event MouseUp(Button As Integer)
Event BalloonTimeout()
Event BalloonClicked()
Event BalloonShow()
Event BalloonHide()
' These are modular level variables that allow
' us to determine certain aspects of the icon
' and share control of the forms events
Private WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1
Private m_AutoRefresh As Boolean
Private m_IconLoaded As Boolean
Private m_IconData As NOTIFYICONDATAW
' Constants related to events that we can
' extract from the system tray are here
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MOUSEMOVE = &H200
' The events we can extract from the balloons
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
' Constants releated to adding and removing
' the icon from the task tray and response level
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
' These inform windows what action
' we are about to perform with the icon
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5
'shell_notify styles
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
' The windows API call that allows us to
' manage the icon in the task tray
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long

Public Sub Initialize(Form As Form)
' Initialize the icon handler and any variables
' that may be required by the api call
Set m_Form = Form
With m_IconData
    ' Setup the flags and other settings of the icon
    ' like we normally would using the forms settings
    .icoHwnd = m_Form.hwnd
    .icoId = m_Form.Icon
    .icoCallbackMessage = WM_MOUSEMOVE
    .icoSource = m_Form.Icon.Handle
    ' Setup new variables to suit the balloon message
    .uTimeOutOrVersion = 1000 '1000 millisec
    .icoSize = 936
End With
' Ensure we are on a pixel
' scalemode otherwise messages
' will not get interpreted properly
m_Form.ScaleMode = 3
End Sub

Public Sub ShowBalloonTip(Message As String, Optional Title As String, Optional Style As EBalloonIconTypes, Optional Timeout = 1000)
' Attempt to show a balloon tip if the
' icon is loaded in the task tray already
'
'usage: Tray.ShowBalloonTip "Message line 1" & vbNewLine & "Message line 2", "Balloon Title", NIIF_INFO, 1000
If m_IconLoaded Then
    ' Convert the title and message into an array
    stringToArray Message, m_IconData.szInfo, 512
    stringToArray Title, m_IconData.szInfoTitle, 128
    ' Store the timeout value here and the icon
    m_IconData.uTimeOutOrVersion = Timeout
    m_IconData.dwInfoFlags = Style
    m_IconData.icoFlags = NIF_INFO
    ' Update the icon with the new information
    Shell_NotifyIconW NIM_MODIFY, m_IconData
End If
End Sub

Public Sub AddIcon()
' Adds the icon to the task tray so
' the form can handle the events
If m_IconLoaded = False Then
    ' Change the flags to the ones we should use when adding an icon
    m_IconData.icoFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    ' Dont bother adding an icon if its already been done
    Call Shell_NotifyIconW(NIM_ADD, m_IconData)
    Call Shell_NotifyIconW(NIM_SETVERSION, m_IconData)
    ' Only allowed one icon per form
    m_IconLoaded = True
End If
End Sub

Public Sub Refresh()
' Refresh the icon in the task
' tray if it exists at all
If m_IconLoaded Then
    ' Only bother to refresh if it actually exists
    'first set the flag back to the normal tooltip
    m_IconData.icoFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Call Shell_NotifyIconW(NIM_MODIFY, m_IconData)
    DoEvents
    'below shows the tip right away.
    Call Shell_NotifyIconW(NIM_SETFOCUS, m_IconData)
End If
End Sub

Public Sub RemoveIcon()
' Removes the icon from the task tray
' but only if we already have one to delete
If m_IconLoaded = True Then
    ' Remove the icon from the task tray
    Call Shell_NotifyIconW(NIM_DELETE, m_IconData)
    m_IconLoaded = False
End If
End Sub

Private Sub Class_Terminate()
' If the icon is still in the tray, remove it
If m_IconLoaded = True Then RemoveIcon
End Sub

Property Get AutoRefresh() As Boolean
' Return the value of m_AutoRefresh
AutoRefresh = m_AutoRefresh
End Property

Property Let AutoRefresh(Value As Boolean)
' Save the auto update value
m_AutoRefresh = Value
End Property

Property Get Tooltip() As String
' Simply return the tool tip
Tooltip = m_IconData.icoToolTip
End Property

Property Let Tooltip(Message As String)
' Ensure the delimiter of null is kept here
stringToArray Message & Chr(0), m_IconData.icoToolTip, 256
' m_IconData.icoTooltip = Message & Chr(0)
If m_AutoRefresh Then Call Refresh
End Property

Property Get Icon() As Long
' Simply return the icon
Icon = m_IconData.icoSource
End Property

Property Let Icon(Icon As Long)
' Store the data into our data
m_IconData.icoSource = Icon
m_IconData.icoSize = Len(Icon)
If m_AutoRefresh Then Call Refresh
End Property

Private Sub m_Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' We will handle the MouseMove event
' for the form so we can trap events
'
'event usage
'Private Sub Tray_DoubleClick(Button As Integer)
'' If they double click with the left mouse
'' button we will simply show the form
'If Button = 0 Then
'    ' Return to normal
'    Call mnuShow_Click
'End If
'End Sub
'
'Private Sub Tray_MouseDown(Button As Integer)
'' If they right click on the task tray then
'' we will simply show them a popup menu
'If Button = 1 Then
'    ' And popup the menu
'    PopupMenu mnuPopup, , , , mnuShow
'End If
'End Sub
'
Select Case X
    ' Events related to double clicking the icon
    Case WM_LBUTTONDBLCLK
        RaiseEvent DoubleClick(0)
    Case WM_RBUTTONDBLCLK
        RaiseEvent DoubleClick(1)
        ' Events related to button down on the icon
    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(0)
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(1)
        ' Events related to button up on the icon
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(0)
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(1)
        ' Events related to the balloon are here xD
    Case NIN_BALLOONSHOW
        RaiseEvent BalloonShow
    Case NIN_BALLOONHIDE
        RaiseEvent BalloonHide
    Case NIN_BALLOONTIMEOUT
        RaiseEvent BalloonTimeout
    Case NIN_BALLOONUSERCLICK
        RaiseEvent BalloonClicked
End Select
End Sub

Private Sub stringToArray(ByVal sString As String, bArray() As Byte, ByVal lMaxSize As Long)
' Converts a string into a byte array then
' transfers it to the main array and obeying
' any limits that have been set
'
Dim Bytes() As Byte
Dim Pointer As Long
Dim PointerEmpty As Long
If Len(sString) > 0 Then
    ' Get the string into an array
    ' of bytes so we can use it
    Bytes = sString
    For Pointer = 0 To UBound(Bytes)
        ' Store it into the next array and exit
        ' when we have reached the limit
        bArray(Pointer) = Bytes(Pointer)
        If (Pointer = (lMaxSize - 2)) Then Exit For
    Next Pointer
    For PointerEmpty = Pointer To lMaxSize - 1
        ' Fill the rest of the array with an
        ' empty character (in this case 0)
        bArray(PointerEmpty) = 0
    Next PointerEmpty
End If
End Sub
