Attribute VB_Name = "modShowList"
'This Code was written by Dave Andrews
'and modified by Tom Pydeski
'modifications include:
'made list 3d
'automatically size list based on the number
'of entries and the width of the longest entry
'added the keystroke capability
'added the option of changing the font to that of the calling form
'or any of its container controls that support .textheight.
'added double click capability to the list
'
'
'Feel free to use or modify this module freely
'Special thanks to Joseph Huntley for the skeleton of API forms.
'
'USAGE
'Dim inList() As Variant
'Dim outList() As Variant
'Dim i As Integer
'Dim j As Integer
'Dim listmax As Integer
'listmax = Text1.Text
'ReDim inList(listmax)
'Create a list of  "words"
'For i = 0 To listmax
'    inList(i) = "List Item #" & i & " = "
'    For j = 1 To CInt(Rnd * 25) + 1
'        inList(i) = inList(i) & Chr(CInt(Rnd * 26) + 65)
'    Next j
'Next i
'Get our selection
'If ShowList(inList(), outList(), True, True, "List Test", 0, 0, 300, 350, chkKeys, Form1) Then
'    'output our selection
'    For i = 0 To UBound(outList)
'        MsgBox outList(i)
'    Next i
'End If
'
Option Explicit
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function defWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Listbox Styles
Const LBS_NOTIFY = &H1
Const LBS_SORT = &H2
Const LBS_NOREDRAW = &H4
Const LBS_MULTIPLESEL = &H8
Const LBS_OWNERDRAWFIXED = &H10
Const LBS_OWNERDRAWVARIABLE = &H20
Const LBS_HASSTRINGS = &H40
Const LBS_USETABSTOPS = &H80
Const LBS_NOINTEGRALHEIGHT = &H100
Const LBS_MULTICOLUMN = &H200
Const LBS_WANTKEYBOARDINPUT = &H400
Const LBS_EXTENDEDSEL = &H800
Const LBS_DISABLENOSCROLL = &H1000
Const LBS_NODATA = &H2000
'
'Listbox Constants
Const LB_ADDSTRING = &H180
Const LB_INSERTSTRING = &H181
Const LB_DELETESTRING = &H182
Const LB_SELITEMRANGEEX = &H183
Const LB_RESETCONTENT = &H184
Const LB_SETSEL = &H185
Const LB_SETCURSEL = &H186
Const LB_GETSEL = &H187
Const LB_GETCURSEL = &H188
Const LB_GETTEXT = &H189
Const LB_GETTEXTLEN = &H18A
Const LB_GETCOUNT = &H18B
Const LB_SELECTSTRING = &H18C
Const LB_DIR = &H18D
Const LB_GETTOPINDEX = &H18E
Const LB_FINDSTRING = &H18F
Const LB_GETSELCOUNT = &H190
Const LB_GETSELITEMS = &H191
Const LB_SETTABSTOPS = &H192
Const LB_GETHORIZONTALEXTENT = &H193
Const LB_SETHORIZONTALEXTENT = &H194
Const LB_SETCOLUMNWIDTH = &H195
Const LB_ADDFILE = &H196
Const LB_SETTOPINDEX = &H197
Const LB_GETITEMRECT = &H198
Const LB_GETITEMDATA = &H199
Const LB_SETITEMDATA = &H19A
Const LB_SELITEMRANGE = &H19B
Const LB_SETANCHORINDEX = &H19C
Const LB_GETANCHORINDEX = &H19D
Const LB_SETCARETINDEX = &H19E ' multi-selection lbs only
Const LB_GETCARETINDEX = &H19F
Const LB_SETITEMHEIGHT = &H1A0
Const LB_GETITEMHEIGHT = &H1A1
Const LB_FINDSTRINGEXACT = &H1A2
Const LB_SETLOCALE = &H1A5
Const LB_GETLOCALE = &H1A6
Const LB_SETCOUNT = &H1A7
Const LB_INITSTORAGE = &H1A8
Const LB_ITEMFROMPOINT = &H1A9
Const LB_MSGMAX = &H1B0
'Listbox Notification Codes
Const LBN_ERRSPACE = (-2)
Const LBN_SELCHANGE = 1
Const LBN_DBLCLK = 2
Const LBN_SELCANCEL = 3
Const LBN_SETFOCUS = 4
Const LBN_KILLFOCUS = 5
'------Button Constants
Const BS_USERBUTTON = &H8&
Const BS_CENTER = 768
Const BS_PUSHBUTTON = &H0&
Const BS_AUTORADIOBUTTON = &H9&
Const BS_PUSHLIKE = &H1000&
Const BS_LEFTTEXT = &H20&
Const BM_SETSTATE = &HF3
Const BM_GETSTATE = &HF2
Const BM_SETCHECK = &HF1
Const BM_GETCHECK = &HF0
'-----------Window Style Constants
Const WS_BORDER = &H800000
Const WS_CHILD = &H40000000
Const WS_OVERLAPPED = &H0&
Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Const WS_SYSMENU = &H80000
Const WS_THICKFRAME = &H40000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_VISIBLE = &H10000000
Const WS_POPUP = &H80000000
Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Const WS_VSCROLL = &H200000
Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_TOPMOST = &H8&
Const WS_EX_WINDOWEDGE = &H100
Const WS_EX_CLIENTEDGE = &H200&
Const WS_SIZEBOX = &H40000
Const WS_EX_DLGMODALFRAME = &H1&
'-----------Window Messaging Constants
Const WM_DESTROY = &H2
Const WM_MOVE = &H3
Const WM_SIZE = &H5
Const WM_ENABLE = &HA
Const WM_SETTEXT = &HC
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Const WM_CLOSE = &H10
Const WM_SETCURSOR = &H20
Const WM_SETFONT = &H30
Const WM_GETFONT = &H31
Const WM_NCPAINT = &H85
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const WM_COMMAND = &H111
Const WM_VSCROLL = &H115
Const WM_CTLCOLOREDIT = &H133
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
'--------Window Heiarchy Constants
Const GWL_WNDPROC = (-4)
Const GW_CHILD = 5
Const GW_OWNER = 4
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const SW_SHOWNORMAL = 1
'----------Misc Constants
Const CS_VREDRAW = &H1
Const CS_HREDRAW = &H2
Const CW_USEDEFAULT = &H80000000
Const COLOR_WINDOW = 5
Const SET_BACKGROUND_COLOR = 4103
Const IDC_ARROW = 32512&
Const IDI_APPLICATION = 32512&
Const MB_OK = &H0&
Const MB_ICONEXCLAMATION = &H30&
Dim MyMousePos As POINTAPI 'for getting the mouse positioning
Const gClassName = "Listbox API"
Dim gAppTitle As String
'handle variables
Global FormWindowHwnd As Long
Global ListBoxHwnd As Long
Global OKButtonHwnd As Long
Global CancelButtonHwnd As Long
'Will hold address of the old window proc for the button
Dim ListBoxOldProc As Long
Dim OKButtonOldProc As Long
Dim CancelButtonOldProc As Long
'
Dim ListStyle As Long
Global CurSel() As Variant
Global inList() As Variant
Global isSelected As Boolean
Dim wTop As Long
Dim wLeft As Long
Dim wHeight As Long
Dim wWidth As Long
Dim Created As Boolean
Dim ListCount As Integer
Dim ListHeight As Integer
Global SelectedListItem As Integer
Dim EnableKey As Byte
Dim WaitforEnter As Byte
Dim KeyString$
Dim FontHwnd As Long
Dim FontHeight As Long
Dim ListWidth As Long
Dim eTitle$
Dim EMess$
Dim mError As Long

Public Sub CopyArray(Source() As Variant, ByRef Dest() As Variant)
On Error GoTo Oops
Dim I As Integer
ReDim Preserve Dest(UBound(Source))
For I = 0 To UBound(Source)
    Dest(I) = Source(I)
Next I
eTrap:
GoTo Exit_CopyArray
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine CopyArray "
EMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in CopyArray"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_CopyArray:
End Sub

Function EditClass() As WNDCLASS
EditClass.hbrBackground = vbRed
End Function

Sub MakeSelection()
Dim I As Integer
Dim tLen As Long
Dim tItem As String
Dim sCount As Integer
Dim ret As Long
For I = 0 To UBound(inList)
    'check if each entry is selected
    ret = SendMessage(ListBoxHwnd, LB_GETSEL, I, 0&)
    If ret <> 0 Then
        'if an entry is selected, then put it into the cursel array
        ReDim Preserve CurSel(sCount)
        'retrieve the length of the list item
        tLen = SendMessage(ListBoxHwnd&, LB_GETTEXTLEN, I, 0&)
        'create a string of that length
        tItem = Space(tLen)
        'retrieve the text of the list item
        Call SendMessage(ListBoxHwnd&, LB_GETTEXT, I, ByVal tItem)
        CurSel(sCount) = tItem
        SelectedListItem = I
        'increment the selected count variable
        sCount = sCount + 1
        isSelected = True
    End If
Next I
End Sub

Function ShowList(InputList() As Variant, ByRef SelectionList() As Variant, Optional MultiSelect As Boolean, Optional Sorted As Boolean, Optional Title As String, Optional Left As Long, Optional Top As Long, Optional Width As Long = 150, Optional Height As Long = 200, Optional AllowKey As Byte, Optional FontControl As Object) As Boolean
'option to allow keypress to select item from list
EnableKey = AllowKey
'optional handle of an object whose font is to used for the list font
'must be a form or a picturebox with a textheight property
'set a default font height
FontHeight = 16
If Not IsMissing(FontControl) Then
    FontHwnd = FontControl.hwnd
    FontHeight = FontControl.TextHeight("Test") / Screen.TwipsPerPixelY
End If
Call GetCursorPos(MyMousePos)
If IsMissing(wLeft) Then wLeft = MyMousePos.X
If IsMissing(wTop) Then wTop = MyMousePos.Y
'copy the incoming list
CopyArray InputList(), inList()
'initialize the width of the list
ListWidth = 0
Dim linewidth As Long
Dim I As Integer
Dim tStr As String
For I = 0 To UBound(inList)
    tStr$ = CStr(inList(I))
    'find the longest line in the listbox
    linewidth = FontControl.TextWidth(tStr$ & " ") / Screen.TwipsPerPixelX
    If linewidth > ListWidth Then
        ListWidth = linewidth + 4
        'Debug.Print tStr$; "="; linewidth
    End If
Next I
'Debug.Print ListWidth
'
'set the width of the list to a little more than the longest entry
wWidth = ListWidth + 20 ' Width
If Title <> "" Then gAppTitle$ = Title Else gAppTitle$ = "Make A Selection"
Dim wMsg As msg
Dim tSec As String
If Sorted And EnableKey = 0 Then ListStyle = ListStyle Or LBS_SORT
If MultiSelect And EnableKey = 0 Then ListStyle = ListStyle Or LBS_EXTENDEDSEL
'Call procedure to register window classname. If false, then exit.
If RegisterWindowClass = False Then
    Exit Function
End If
'Create window
If CreateWindows() Then
    'Loop will exit when WM_QUIT is sent to the window.
    Do While GetMessage(wMsg, 0&, 0&, 0&)
        'TranslateMessage takes keyboard messages and converts
        'them to WM_CHAR for easier processing.
        Call TranslateMessage(wMsg)
        'Dispatchmessage calls the default window procedure
        'to process the window message. (WndProc)
        Call DispatchMessage(wMsg)
        DoEvents
    Loop
End If
Call UnregisterClass(gClassName$, App.hInstance)
If isSelected Then
    CopyArray CurSel(), SelectionList()
End If
ShowList = isSelected
End Function

Function CreateWindows() As Boolean
Dim I As Integer
Dim tStr As String
Dim ButtonStyle As Long
ListCount = UBound(inList)
If ListCount < 10 Then
    WaitforEnter = 0
Else
    WaitforEnter = 1
End If
KeyString$ = ""
'set the height of the list based on the number of entries
ListHeight = ((ListCount + 1) * FontHeight) + 10
wHeight = ListHeight + 50
'if the list is larger than the screen height, set it to the screen height
With Screen
    If wHeight > (.Height / .TwipsPerPixelY) Then
        wHeight = (.Height / .TwipsPerPixelY) - 10
        ListHeight = wHeight - 50
    End If
    'center the list on the screen
    wLeft = ((.Width / .TwipsPerPixelX) - wWidth) / 2
    wTop = ((.Height / .TwipsPerPixelY) - wHeight) / 2
    'make sure the top is visible
    If wTop < 0 Then wTop = 0
End With
ButtonStyle = WS_CHILD Or WS_VISIBLE Or WS_BORDER
ListStyle = ListStyle Or WS_CHILD Or WS_VISIBLE Or WS_BORDER Or WS_VSCROLL Or LBS_NOINTEGRALHEIGHT
'Create form window.
'WS_EX_CLIENTEDGE gives a border on the window      Or WS_EX_CLIENTEDGE
FormWindowHwnd& = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, gClassName$, gAppTitle$, WS_POPUPWINDOW Or WS_CAPTION Or WS_VISIBLE Or WS_SIZEBOX, wLeft, wTop, wWidth, wHeight, 0&, 0&, App.hInstance, ByVal 0&)
'Create List Box 'first parameter makes it 3d
ListBoxHwnd& = CreateWindowEx(WS_EX_CLIENTEDGE, "LISTBOX", "", ListStyle, 1, 1, ListWidth, ListHeight, FormWindowHwnd&, 0&, App.hInstance, 0&)
'Create OK and Cancel Buttons
OKButtonHwnd = CreateWindowEx(0&, "BUTTON", "OK", ButtonStyle, 1, wHeight - 50, (wWidth - 11) / 2, 20, FormWindowHwnd&, 0&, App.hInstance, 0&)
CancelButtonHwnd = CreateWindowEx(0&, "BUTTON", "CANCEL", ButtonStyle, (wWidth - 4) / 2, wHeight - 50, (wWidth - 11) / 2, 20, FormWindowHwnd&, 0&, App.hInstance, 0&)
'add the list items to the listbox
For I = 0 To UBound(inList)
    tStr$ = CStr(inList(I))
    SendMessage ListBoxHwnd&, LB_ADDSTRING, 0&, ByVal tStr$
Next I
'below trick from vbthunder.com
'Get rid of that ugly default font
'set the font to the font of the incoming optional control (or container)
If FontHwnd <> 0 Then
    Dim hFont As Long
    hFont = SendMessage(FontHwnd, WM_GETFONT, 0&, ByVal 0&)
    SendMessage ListBoxHwnd&, WM_SETFONT, hFont, ByVal 1&
    SendMessage OKButtonHwnd, WM_SETFONT, hFont, ByVal 1&
    SendMessage CancelButtonHwnd, WM_SETFONT, hFont, ByVal 1&
End If
'
'-------Hook OK CANCEL-----------
'also hook list
ListBoxOldProc& = GetWindowLong(ListBoxHwnd&, GWL_WNDPROC)
Call SetWindowLong(ListBoxHwnd&, GWL_WNDPROC, GetAddress(AddressOf ListWndProc))
OKButtonOldProc& = GetWindowLong(OKButtonHwnd&, GWL_WNDPROC)
Call SetWindowLong(OKButtonHwnd&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
CancelButtonOldProc& = GetWindowLong(CancelButtonHwnd&, GWL_WNDPROC)
Call SetWindowLong(CancelButtonHwnd&, GWL_WNDPROC, GetAddress(AddressOf CancelWndProc))
CreateWindows = (FormWindowHwnd& <> 0)
Created = True
Call SendMessage(FormWindowHwnd, WM_SIZE, 0&, 0&)
End Function

Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'This our default window procedure for the window. It will handle all
'of our incoming window messages and we will write code based on the
'window message what the program should do.
Dim I As Integer
Select Case uMsg&
    Case WM_DESTROY:
        'Since DefWindowProc doesn't automatically call
        'PostQuitMessage (WM_QUIT). We need to do it ourselves.
        'You can use DestroyWindow to get rid of the window manually.
        'SetDate
        Call PostQuitMessage(0&)
        Created = False
        'maybe we need to unhook the process?
        'this was not in the original example, but I think I may put it in
        Call SetWindowLong(ListBoxHwnd&, GWL_WNDPROC, ListBoxOldProc&)
        Call SetWindowLong(OKButtonHwnd&, GWL_WNDPROC, OKButtonOldProc&)
        Call SetWindowLong(CancelButtonHwnd&, GWL_WNDPROC, CancelButtonOldProc&)
    Case WM_SIZE
        If Not Created Then
            Exit Function
        End If
        Dim wSize As RECT
        'get the size of the main window
        GetWindowRect FormWindowHwnd&, wSize
        wLeft = wSize.Left
        wTop = wSize.Top
        wWidth = wSize.Right - wSize.Left
        wHeight = wSize.Bottom - wSize.Top
        Dim butwidth As Integer
        'set the button width
        butwidth = (wWidth - 10) / 2
        MoveWindow ListBoxHwnd&, 1, 0, wWidth - 10, wHeight - 50, True
        MoveWindow OKButtonHwnd&, 2, wHeight - 47, butwidth, 20, True
        MoveWindow CancelButtonHwnd&, butwidth, wHeight - 47, butwidth, 20, True
    Case WM_KEYDOWN
        'Debug.Print wParam
        'Debug.Print lParam
        Dim KeyIn As Integer
        Dim LIndex As Integer
        Dim CharIn$
        KeyIn = CInt(wParam)
        CharIn$ = Chr$(KeyIn)
        'if we get the enter key then we no longer have to wait for it
        If KeyIn = 13 Then WaitforEnter = 0
        'if the 1st key cannot be the first of 2 keys ie 1st key = 8 and listcount can't be 80ish
        If KeyString$ = "" And 10 * (Val(CharIn$)) > ListCount Then WaitforEnter = 0
        'if our key is a number then build a string
        If IsNumeric(CharIn$) Then KeyString$ = KeyString$ & CharIn$
        If EnableKey = 1 Then
            If IsNumeric(KeyString$) Then
                'conver the string of numbers to an index value
                LIndex = Val(KeyString$)
                'if we have entered enough numbers, then why wait for an enter key
                If Len(KeyString$) = Len(Trim(Str$(ListCount))) Then
                    WaitforEnter = 0
                End If
                If WaitforEnter = 0 Then
                    'retieve our list item
                    GetListbyIndex LIndex
                    Call SendMessage(FormWindowHwnd, WM_CLOSE, 0&, 0&)
                End If
            End If
        End If
End Select
'Let windows call the default window procedure since we're done.
WndProc = defWindowProc(hwnd&, uMsg&, wParam&, lParam&)
End Function

Function ListWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg&
    Case WM_LBUTTONDBLCLK
        MakeSelection
        Call SendMessage(FormWindowHwnd, WM_CLOSE, 0&, 0&)
End Select
ListWndProc = CallWindowProc(ListBoxOldProc&, hwnd&, uMsg&, wParam&, lParam&)
End Function

Function OKWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg&
    Case WM_LBUTTONDOWN
        MakeSelection
        Call SendMessage(FormWindowHwnd, WM_CLOSE, 0&, 0&)
End Select
OKWndProc = CallWindowProc(OKButtonOldProc&, hwnd&, uMsg&, wParam&, lParam&)
End Function

Function CancelWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg&
    Case WM_LBUTTONDOWN
        isSelected = False
        Call SendMessage(FormWindowHwnd, WM_CLOSE, 0&, 0&)
End Select
CancelWndProc = CallWindowProc(CancelButtonOldProc&, hwnd&, uMsg&, wParam&, lParam&)
End Function

Function GetAddress(ByVal lngAddr As Long) As Long
'Used with AddressOf to return the address in memory of a procedure.
GetAddress = lngAddr&
End Function

Function RegisterWindowClass() As Boolean
Dim wc As WNDCLASS
'Registers our new window with windows so we can use our classname.
wc.style = CS_HREDRAW Or CS_VREDRAW
wc.lpfnwndproc = GetAddress(AddressOf WndProc) 'Address in memory of default window procedure.
wc.hInstance = App.hInstance
wc.hIcon = LoadIcon(0&, IDI_APPLICATION) 'Default application icon
wc.hCursor = LoadCursor(0&, IDC_ARROW) 'Default arrow
wc.hbrBackground = COLOR_WINDOW 'Default a color for window.
wc.lpszClassName = gClassName$
RegisterWindowClass = RegisterClass(wc) <> 0
End Function

Sub GetListbyIndex(LIndex As Integer)
Dim tLen As Long
Dim tItem As String
Dim ret As Long
On Error GoTo Oops
'set the selection of the list box to our selected index
Call SendMessage(ListBoxHwnd&, LB_SETCARETINDEX, LIndex, ByVal 0&)
Call SendMessage(ListBoxHwnd&, LB_SETCURSEL, LIndex, 0)
Call SendMessage(ListBoxHwnd&, LB_SETSEL, vbTrue, ByVal LIndex)
MakeSelection
Exit Sub
'below is another way to retrieve the selected item
tLen = SendMessage(ListBoxHwnd&, LB_GETTEXTLEN, LIndex, 0&)
tItem = Space(tLen)
Call SendMessage(ListBoxHwnd&, LB_GETTEXT, LIndex, ByVal tItem)
ReDim CurSel(1)
CurSel(0) = tItem
SelectedListItem = LIndex
isSelected = True
GoTo Exit_GetListbyIndex
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine GetListbyIndex "
EMess$ = "Error # " & Err.Number & " - " & Err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in GetListbyIndex"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_GetListbyIndex:
End Sub
