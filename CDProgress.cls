VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


'Window Messages
Const WM_USER = &H400

'Window Styles
Const WS_ACTIVECAPTION = &H1
Const WS_BORDER = &H800000
Const WS_CAPTION = &HC00000         ' WS_BORDER Or WS_DLGFRAME
Const WS_CHILD = &H40000000
Const WS_CHILDWINDOW = (WS_CHILD)
Const WS_CLIPCHILDREN = &H2000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_DISABLED = &H8000000
Const WS_DLGFRAME = &H400000
Const WS_GROUP = &H20000
Const WS_HSCROLL = &H100000
Const WS_MAXIMIZE = &H1000000
Const WS_MAXIMIZEBOX = &H10000
Const WS_MINIMIZE = &H20000000
Const WS_MINIMIZEBOX = &H20000
Const WS_ICONIC = WS_MINIMIZE
Const WS_OVERLAPPED = &H0&
Const WS_POPUP = &H80000000
Const WS_SYSMENU = &H80000
Const WS_TABSTOP = &H10000
Const WS_THICKFRAME = &H40000
Const WS_SIZEBOX = WS_THICKFRAME
Const WS_TILED = WS_OVERLAPPED
Const WS_VISIBLE = &H10000000
Const WS_VSCROLL = &H200000
Const WS_GT = (WS_GROUP Or WS_TABSTOP)
Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Const WS_EX_DLGMODALFRAME = &H1&
Const WS_EX_WINDOWEDGE = &H100&

'Common Controls Messages
Const CCM_FIRST = &H2000
Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)
Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
Const CCM_GETVERSION = (CCM_FIRST + &H8)
Const CCM_LAST = (CCM_FIRST + &H200)
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)
Const CCM_SETNOTIFYWINDOW = (CCM_FIRST + &H9)
Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Const CCM_SETVERSION = (CCM_FIRST + &H7)

'Common Controls Styles
Const CCS_ADJUSTABLE = &H20
Const CCS_BOTTOM = &H3
Const CCS_NODIVIDER = &H40
Const CCS_NOMOVEY = &H2
Const CCS_NOPARENTALIGN = &H8
Const CCS_NORESIZE = &H4
Const CCS_TOP = &H1
Const CCS_VERT = &H80
Const CCS_RIGHT = (CCS_VERT Or CCS_BOTTOM)
Const CCS_NOMOVEX = (CCS_VERT Or CCS_NOMOVEY)
Const CCS_LEFT = (CCS_VERT Or CCS_TOP)

'Progress Bar Class Name
Const PROGRESS_CLASSA = "msctls_progress32"

'Progress Bar Messages
Const PBM_DELTAPOS = (WM_USER + 3)
Const PBM_GETPOS = (WM_USER + 8)
Const PBM_GETRANGE = (WM_USER + 7)
Const PBM_SETBARCOLOR = (WM_USER + 9)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Const PBM_SETPOS = (WM_USER + 2)
Const PBM_SETRANGE = (WM_USER + 1)
Const PBM_SETRANGE32 = (WM_USER + 6)
Const PBM_SETSTEP = (WM_USER + 4)
Const PBM_STEPIT = (WM_USER + 5)

'Progress Bar Styles
Const PBS_SMOOTH = &H1
Const PBS_VERTICAL = &H4

Public Enum ProgressStyles
PBS_Default = 0
PBS_VerticalBar = 1
PBS_SmoothBar = 2
PBS_SmoothVertical = 3
End Enum

Public Enum BorderStyle
PBB_Raised = WS_EX_DLGMODALFRAME 'Or WS_EX_WINDOWEDGE
PBB_ThickFrame = WS_THICKFRAME 'Or WS_DLGFRAME
PBB_Default = WS_BORDER
PBB_None = 0&
'PBB_Popup = WS_POPUP
End Enum
Dim lhWnd As Long, lType As Long, RetVal As Long

'Public Event OnError()
'Public Event Change()
'Public Event Click()
'Public Event DblClick()
'Public Event MouseDown(Button As Integer)
'Public Event MouseUp(Button As Integer)
'Public Event MouseMove(X As Single, Y As Single, Button As Integer)


' member variable for ProgressStyle property
Private m_ProgressStyle As ProgressStyles
' member variable for Left property
Private m_Left As Integer
' member variable for Top property
Private m_Top As Integer
' member variable for Width property
Private m_Width As Integer
' member variable for Height property
Private m_Height As Integer
' member variable for OwnerhWnd property
Private m_OwnerhWnd As Long
' member variable for Min property
Private m_Min As Long
' member variable for Max property
Private m_Max As Long
' member variable for BarColor property
Private m_BarColor As Long
' member variable for BackColor property
Private m_BackColor As Long
' member variable for Value property
Private m_Value As Long
' member variable for Percent property
Private m_Percent As Long
' member variable for BorderType property
Private m_BorderType As BorderStyle





Private Sub Class_Initialize()
    m_ProgressStyle = PBS_Default
    m_Width = 0
    m_Left = 0
    m_Top = 0
    m_Height = 0
    m_OwnerhWnd = 0
    m_Min = 0
    m_Max = 100
    m_BarColor = &H800000
    m_BackColor = &H8000000F
    m_Value = 0
    m_Percent = 0
    m_BorderType = WS_BORDER
End Sub

' Sets the Border type for the Progress Bar.

Property Let BorderType(ByVal newValue As BorderStyle)
Attribute BorderType.VB_Description = "Sets the Border type for the Progress Bar."
    m_BorderType = newValue
    If Create = 1 Then redraw
End Property

' Sets/Returns the current value of the Progress Bar

Property Get Value() As Long
Attribute Value.VB_Description = "Sets/Returns the current value of the Progress Bar"
Dim temp As Long
    temp& = SendMessage(lhWnd&, PBM_GETPOS, ByVal 0&, ByVal 0&)
    Value = m_Value
    If temp& <> m_Value Then m_Value = temp&: Value = m_Value
End Property

Property Let Value(ByVal newValue As Long)
    m_Value = newValue
    RetVal = SendMessage(lhWnd&, PBM_SETPOS, ByVal m_Value, ByVal 0&)
    If Create = 1 Then redraw
End Property

' Returns the percent of the progress completed

Property Get Percent() As Long
Attribute Percent.VB_Description = "Returns the percent of the progress completed"
    m_Percent = lPercent&
    Percent = m_Percent
End Property

' Sets the Owner Window Handle for the Progress Bar

Property Let OwnerhWnd(ByVal newValue As Long)
Attribute OwnerhWnd.VB_Description = "Sets the Owner Window Handle for the Progress Bar"
    m_OwnerhWnd = newValue
    If Create = 1 Then redraw
End Property

' Sets/Returns the Background Color for the Progress Bar.

Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Sets/Returns the Background Color for the Progress Bar."
    BackColor = m_BackColor
End Property

Property Let BackColor(ByVal newValue As Long)
    m_BackColor = newValue
    RetVal& = SendMessage(lhWnd&, PBM_SETBKCOLOR, ByVal 0&, ByVal m_BackColor&)
    If Create = 1 Then redraw
End Property

' Sets/Returns the Color for the Progress Bar.

Property Get BarColor() As Long
Attribute BarColor.VB_Description = "Sets/Returns the Color for the Progress Bar."
    BarColor = m_BarColor
End Property

Property Let BarColor(ByVal newValue As Long)
    m_BarColor = newValue
    RetVal& = SendMessage(lhWnd&, PBM_SETBARCOLOR, ByVal 0&, ByVal m_BarColor&)
    If Create = 1 Then redraw
End Property

' Sets/Returns the Maximum value for the Progress Bar.

Property Get Max() As Long
Attribute Max.VB_Description = "Sets/Returns the Maximum value for the Progress Bar."
    Max = m_Max
End Property

Property Let Max(ByVal newValue As Long)
    m_Max = newValue
    RetVal& = SendMessage(lhWnd&, PBM_SETRANGE32, ByVal m_Min, ByVal m_Max)
    If Create = 1 Then redraw
End Property

' Sets/Returns the Minimum value for the Progress Bar.

Property Get Min() As Long
Attribute Min.VB_Description = "Sets/Returns the Minimum value for the Progress Bar."
    Min = m_Min
End Property

Property Let Min(ByVal newValue As Long)
    m_Min = newValue
    RetVal& = SendMessage(lhWnd&, PBM_SETRANGE32, ByVal m_Min, ByVal m_Max)
    If Create = 1 Then redraw
End Property

' Sets/Returns the Height of the Progress Bar

Property Get Height() As Integer
Attribute Height.VB_Description = "Sets/Returns the Height of the Progress Bar"
    Height = m_Height
End Property

Property Let Height(ByVal newValue As Integer)
    m_Height = newValue
    If Create = 1 Then redraw
End Property

' Sets/Returns the Width of the Progress Bar

Property Get Width() As Integer
Attribute Width.VB_Description = "Sets/Returns the Width of the Progress Bar"
    Width = m_Width
End Property

Property Let Width(ByVal newValue As Integer)
    m_Width = newValue
    If Create = 1 Then redraw
End Property

' Sets/Returns The Top Position of the Progress Bar

Property Get Top() As Integer
Attribute Top.VB_Description = "Sets/Returns The Top Position of the Progress Bar"
    Top = m_Top
End Property

Property Let Top(ByVal newValue As Integer)
    m_Top = newValue
    redraw
End Property

' Sets/Returns The Left Position of the Progress Bar

Property Get Left() As Integer
Attribute Left.VB_Description = "Sets/Returns The Left Position of the Progress Bar"
    Left = m_Left
End Property

Property Let Left(ByVal newValue As Integer)
    m_Left = newValue
    If Create = 1 Then redraw
End Property

' Set the type of Progress Bar.

Property Let ProgressStyle(ByVal newValue As ProgressStyles)
Attribute ProgressStyle.VB_Description = "Set the type of Progress Bar."
    m_ProgressStyle = newValue
    Select Case m_ProgressStyle
        Case PBS_Default
            lType& = 0&
        Case PBS_VerticalBar
            lType& = PBS_VERTICAL
        Case PBS_SmoothBar
            lType& = PBS_SMOOTH
        Case PBS_SmoothVertical
            lType& = PBS_VERTICAL Or PBS_SMOOTH
    End Select
    If Create = 1 Then redraw
End Property








Public Function Create() As Long
On Error GoTo oops
    lhWnd& = CreateWindowEX(0, PROGRESS_CLASSA, "", WS_VISIBLE Or WS_CHILD Or m_BorderType Or lType, 0, 0, 0, 0, m_OwnerhWnd, 0&, App.hInstance, 0&)
    Call SetParent(lhWnd, m_OwnerhWnd)
    MoveWindow lhWnd, m_Left, m_Top, m_Width, m_Height, True
    Create = 1
    Exit Function
oops:
    Create = 0
End Function

Public Function Destroy() As Long
    Destroy& = DestroyWindow(lhWnd&)
End Function
Private Sub redraw()
Destroy
Create
End Sub
Private Function lPercent() As Long
lPercent& = Int(m_Value * 100 / m_Max)
End Function
Public Sub About()
    frmAbout.Show
End Sub
