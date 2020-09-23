VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test API List"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Inputbox Replacement Example"
      Height          =   735
      Left            =   3120
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   495
      Left            =   3360
      Max             =   0
      Min             =   100
      TabIndex        =   8
      Top             =   240
      Value           =   100
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   5
      Text            =   "8"
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   480
      Width           =   2655
   End
   Begin VB.CheckBox chkKeys 
      Caption         =   "Enable KeyPress"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Build Random List"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      TabIndex        =   0
      Text            =   "20"
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Font Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Font Size"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   50
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "# of List Entries"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim defUrl$()
Dim UrlName$()
Const TeamUrl$ = "http://yankees.mlb.com/NASApp/mlb/nyy/schedule/nyy_schedule_calendar.jsp?GXHC_gx_session_id_=3b13712f3048b9ee"
Const NJOnlineUrl$ = "http://www.nj.com/yankees/spnet.ssf?/default.asp?c=advance&page=mlb/teams/sched/036.htm"
Const YESUrl$ = "http://www.yesnetworktv.com/mlb/team.cfm?type=sched&team_id=917"
Const CBSUrl$ = "http://cbs.sportsline.com/u/baseball/mlb/teams/NYY/schedule.htm"
Const ESPNUrl$ = "http://sports.espn.go.com/mlb/schedule?team=nyy"
Const MSGUrl$ = "http://www.msgnetwork.com/mlb/home/mlb_team.cfm?type=sched&team_id=917&subnav_key=mlb_nyy"
Const NYPostUrl$ = "http://www.nypostonline.com/sports/yankees/yankeesschedule.htm"
Const SSUrl$ = "http://archive.sportserver.com/newsroom/sports/bbo/1995/mlb/nyy/stat/99log_lo.html"
Const FoxUrl$ = "http://foxsports.lycos.com/named/Story/FullText/MLB/Team_Schedule_Results?statsId=10&statsName=Yankees"
Const SportsNetUrl$ = "http://www.sportsnetwork.com/default.asp?c=sportsnetwork&page=mlb/teams/sched/036.htm"
Dim inList() As Variant
Dim outList() As Variant
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long

Private Sub Form_Load()
Screen.MousePointer = 11
ReDim defUrl$(10)
ReDim UrlName$(10)
'
defUrl$(0) = TeamUrl$
defUrl$(1) = NJOnlineUrl$
defUrl$(2) = YESUrl$
defUrl$(3) = CBSUrl$
defUrl$(4) = ESPNUrl$
defUrl$(5) = MSGUrl$
defUrl$(6) = NYPostUrl$
defUrl$(7) = SSUrl$
defUrl$(8) = FoxUrl$
defUrl$(9) = SportsNetUrl$
'
UrlName$(0) = "N.Y. Yankee Schedule"
UrlName$(1) = "NJ Online"
UrlName$(2) = "YES Network"
UrlName$(3) = "CBS Sportsline"
UrlName$(4) = "ESPN Network"
UrlName$(5) = "MSG Network"
UrlName$(6) = "NY Post"
UrlName$(7) = "NY Daily News Sports Server (Fastest...)"
UrlName$(8) = "Fox Sports"
UrlName$(9) = "SportsNet (same as NJ Online)"
Show
Combo1.Text = "Loading Fonts..."
Refresh
DoEvents
For I = 0 To Screen.FontCount - 1
    Combo1.AddItem Screen.Fonts(I)
Next I
Combo1.Text = Me.FontName
Text2.Text = Me.FontSize
VScroll1.Value = Me.FontSize
'Command1_Click
Screen.MousePointer = 0
End Sub

Private Sub Combo1_Click()
Me.FontName = Combo1.Text
Command1.FontName = Me.FontName
End Sub

Private Sub Text2_Change()
Me.FontSize = Text2.Text
Command1.FontSize = Me.FontSize
End Sub

Sub Command1_Click()
Randomize
Dim inList() As Variant
Dim outList() As Variant
Dim I As Integer
Dim j As Integer
Dim listmax As Integer
listmax = Text1.Text
ReDim inList(listmax)
'Create a list of  "words"
For I = 0 To listmax
    inList(I) = "List Item #" & I & " = "
    For j = 1 To CInt(Rnd * 25) + 1
        inList(I) = inList(I) & Chr(CInt(Rnd * 26) + 65)
    Next j
Next I
'Get our selection
If ShowList(inList(), outList(), True, True, "List Test", 0, 0, 300, 350, chkKeys, Form1) Then
    'output our selection
    For I = 0 To UBound(outList)
        MsgBox outList(I)
    Next I
End If
End Sub

Private Sub Command2_Click()
Dim I As Integer
Dim j As Integer
Dim listmax As Integer
listmax = UBound(defUrl$)
ReDim inList(listmax)
'Create a list of urls to choose from
For I = 0 To listmax - 1
    inList(I) = I & "=" & UrlName$(I)
Next I
'Get our selection
If ShowList(inList(), outList(), False, False, "Select the Yankee Schedule page to launch...", 0, 0, 300, 350, chkKeys, Form1) Then
    'output our selection
    Dim ret As Long
    ret = MsgBox("Do you want to Load the " & UrlName$(SelectedListItem) & " Page?", vbQuestion + vbYesNo, "Browser Launch")
    If ret = vbYes Then
        ShellExecute Me.hwnd, "open", defUrl$(SelectedListItem), vbNullString, vbNullString, vbNormal
    End If
End If
End Sub

Private Sub VScroll1_Change()
Text2.Text = VScroll1.Value
End Sub

