VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DC Clean"
   ClientHeight    =   12990
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   21015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12990
   ScaleWidth      =   21015
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFilter3 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "frmdc.frx":0000
      Left            =   12480
      List            =   "frmdc.frx":0002
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "내용 필터링(&F)"
      Height          =   255
      Left            =   10800
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox lstgall 
      Appearance      =   0  'Flat
      Height          =   225
      ItemData        =   "frmdc.frx":0004
      Left            =   3000
      List            =   "frmdc.frx":0020
      TabIndex        =   11
      Top             =   330
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "이전"
      Height          =   375
      Left            =   10200
      TabIndex        =   10
      Top             =   12240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10920
      TabIndex        =   9
      Text            =   "1"
      Top             =   12240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "다음"
      Height          =   375
      Left            =   12240
      TabIndex        =   8
      Top             =   12240
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "닉네임 필터링(&F)"
      Height          =   255
      Left            =   14160
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox lstFilter2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "frmdc.frx":0066
      Left            =   15840
      List            =   "frmdc.frx":0068
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox lstFroxy 
      Height          =   255
      ItemData        =   "frmdc.frx":006A
      Left            =   14040
      List            =   "frmdc.frx":0071
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   240
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdc.frx":0080
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdc.frx":041A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdc.frx":07B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmdc.frx":0B4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstMobile 
      Height          =   255
      ItemData        =   "frmdc.frx":10E8
      Left            =   14040
      List            =   "frmdc.frx":10EF
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "아이피 필터링(&F)"
      Height          =   255
      Left            =   17520
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.ListBox lstFilter 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "frmdc.frx":10FE
      Left            =   19200
      List            =   "frmdc.frx":1100
      TabIndex        =   2
      Top             =   270
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "댓글 없는 글 삭제"
      Height          =   180
      Left            =   17040
      TabIndex        =   1
      Top             =   12480
      Width           =   1935
   End
   Begin MSComctlLib.ListView Lv_Dc 
      Height          =   11055
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   19500
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList"
      SmallIcons      =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   11.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "아이피"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "제목"
         Object.Width           =   21238
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "닉네임"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "날짜"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "조회"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "필터"
         Object.Width           =   2716
      EndProperty
   End
   Begin VB.Image Image10 
      Height          =   450
      Left            =   19080
      Picture         =   "frmdc.frx":1102
      Top             =   12360
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      Height          =   12975
      Left            =   0
      Top             =   0
      Width           =   21015
   End
   Begin VB.Image Image9 
      Height          =   450
      Left            =   3480
      Picture         =   "frmdc.frx":38A4
      Top             =   12360
      Width           =   1380
   End
   Begin VB.Image Image8 
      Height          =   450
      Left            =   2040
      Picture         =   "frmdc.frx":593E
      Top             =   12360
      Width           =   1380
   End
   Begin VB.Image Image7 
      Height          =   450
      Left            =   240
      Picture         =   "frmdc.frx":79D8
      Top             =   12360
      Width           =   1680
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   18720
      Picture         =   "frmdc.frx":A17A
      Top             =   840
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   16920
      Picture         =   "frmdc.frx":A714
      Top             =   840
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   14880
      Picture         =   "frmdc.frx":ACAE
      Top             =   840
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   7080
      Picture         =   "frmdc.frx":B45C
      Top             =   840
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   240
      Picture         =   "frmdc.frx":B9F6
      Stretch         =   -1  'True
      Top             =   720
      Width           =   20535
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   240
      Picture         =   "frmdc.frx":BAB0
      Top             =   240
      Width           =   2670
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Const Head As String = "<body style='text-align:center;font-size:9pt;background:#FFFFFF;color:#000000;font-family:""malgun gothic"",""굴림"";margin:5px;font' contenteditable='true'>"
Private Const Foot As String = "</body>"

Private Type POINTAPI
x As Long
y As Long
End Type

Dim Gall_Name, Gall_Href, Gall_IP, Gall_Title, Gall_Writer, Gall_Date, Gall_Page, Gall_Ex, Gall_Filter As String
Dim i, r, l, z, q, c, t, o, Gall_Reply As Long
Dim DC As New WinHttp.WinHttpRequest
Dim ContFilter() As String
Dim strContent() As String
Dim strFunction() As String
Dim strFunction2() As String
Dim strFunction3() As String
Dim gallContent() As String
Dim strComment() As String
Dim strHref() As String
Dim viewPicture As Boolean

Private Function Utf82String(ByRef data() As Byte) As String
Dim objStream
Dim strTmp As String
Set objStream = CreateObject("ADODB.Stream")
objStream.Charset = "utf-8"
objStream.Mode = 3
objStream.Type = 1
objStream.Open
objStream.Write data
objStream.Flush
objStream.Position = 0
objStream.Type = 2
strTmp = objStream.ReadText
objStream.Close
Set objStream = Nothing
Utf82String = strTmp
End Function

Private Function Refreshs()
On Error GoTo Passing
Image1.Enabled = False
Lv_Dc.ListItems.Clear
Lv_Dc.Refresh
ReDim strContent(0)
ReDim strFunction2(0)
ReDim strFunction3(0)
ReDim gallContent(0)
ReDim strComment(0)
DC.Open "GET", "http://m.dcinside.com/list.php?id=" & lstgall.List(lstgall.ListIndex) & "&page=" & Gall_Page & Gall_Ex, True
DC.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
DC.SetRequestHeader "User-Agent", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_3_2 like Mac OS X; en-us) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8H7 Safari/6533.18.5"
DC.Send
DC.WaitForResponse

r = InStr(Utf82String(DC.ResponseBody), "<ul class=""list_picture"">") + 25
l = InStr(Utf82String(DC.ResponseBody), "<div id=""list_more"">")
strFunction() = Split(Mid(Utf82String(DC.ResponseBody), r, l - r), "<li>")

For i = 0 To UBound(strFunction)

If InStr(strFunction(i), "list_picture_a") > 0 Then
    r = InStr(strFunction(i), "href=") + 6
    l = InStr(strFunction(i), """>")
    Gall_Href = Mid(strFunction(i), r, l - r)
    ReDim Preserve strHref(Lv_Dc.ListItems.Count)
    strHref(Lv_Dc.ListItems.Count) = Gall_Href
End If

If InStr(strFunction(i), "list_right") > 0 Then
    r = InStr(strFunction(i), "</span>") + 7
    l = InStr(r, strFunction(i), "<")
    Gall_Title = Mid(strFunction(i), r, l - r)
    viewPicture = False
    If InStr(strFunction(i), "list_pic_y") Then viewPicture = True
End If

If InStr(strFunction(i), "list_right") > 0 Then
    r = InStr(strFunction(i), "list_pic_galler") + 24
    l = InStr(r, strFunction(i), "<")
    Gall_Writer = Mid(strFunction(i), r, l - r)
End If

If InStr(strFunction(i), "list_right") > 0 Then
    r = InStr(r, strFunction(i), "<span>") + 6
    l = InStr(r, strFunction(i), "<")
    Gall_Date = Mid(strFunction(i), r, l - r)
End If

If Gall_Href <> "" Then
    DC.Open "GET", Gall_Href, True
    DC.SetRequestHeader "User-Agent", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_3_2 like Mac OS X; en-us) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8H7 Safari/6533.18.5"
    DC.Send
    DC.WaitForResponse
    
    r = InStr(Utf82String(DC.ResponseBody), "contents_top") + 12
    l = InStr(Utf82String(DC.ResponseBody), "gallery_hr_left")
    strFunction2() = Split(Mid(Utf82String(DC.ResponseBody), r, l - r), "<")
    
    If InStr(Utf82String(DC.ResponseBody), "m_reply_title") > 0 Then
        ReDim strFunction3(0)
        r = InStr(Utf82String(DC.ResponseBody), "m_reply_title") + 13
        l = InStr(r, Utf82String(DC.ResponseBody), "comment_more")
        strFunction3() = Split((Mid(Utf82String(DC.ResponseBody), r, l - r)), "<")
    Else
        ReDim strFunction3(999)
    End If
    
    For z = 0 To UBound(strFunction2)
        If InStr(strFunction2(z), "gallery_hr_right2") > 0 Then
            r = InStr(strFunction2(z), "gallery_hr_right2") + 19
            Gall_IP = Mid(strFunction2(z), r, l - r)
            
            If Gall_IP = "" Then Gall_IP = " "
            
            For q = 0 To lstFilter.ListCount
                If lstFilter.List(q) = Gall_IP Then GoTo Pass
            Next q
            
            For q = 0 To lstFilter2.ListCount
                If lstFilter2.List(q) = Gall_Writer Then GoTo Pass
            Next q
        
            If InStr(Utf82String(DC.ResponseBody), "mobile_icon_1.gif") > 0 Then '모바일
                Lv_Dc.ListItems.Add , , Gall_IP
                Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ForeColor = vbBlue
                Lv_Dc.ListItems(Lv_Dc.ListItems.Count).SmallIcon = 1
                GoTo Pass1
            End If
        
            For q = 0 To lstFroxy.ListCount
                If Gall_IP = lstFroxy.List(q) Then '프록시 차단
                    Lv_Dc.ListItems.Add , , Gall_IP
                    Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ForeColor = vbRed
                    Lv_Dc.ListItems(Lv_Dc.ListItems.Count).SmallIcon = 2
                    GoTo Pass1
                End If
            Next q
            
            If Gall_IP <> "" Then
                Lv_Dc.ListItems.Add , , Gall_IP
            Else
                Lv_Dc.ListItems.Add , , "고정닉"
            End If
Pass1:
            r = InStr(Utf82String(DC.ResponseBody), "gallery_reply") + 18 '리플수
            l = InStr(r, Utf82String(DC.ResponseBody), "</li>")
            Gall_Reply = Val(Mid(Utf82String(DC.ResponseBody), r, l - r))
        
                If UBound(strFunction3) < 999 Then
                For c = 0 To UBound(strFunction3)
                    
                    If InStr(strFunction3(c), "m_list_text_bt2") > 0 Then
                        r = InStr(strFunction3(c), "m_list_text_bt2") + 17
                        For q = 0 To lstFilter.ListCount
                            If lstFilter.List(q) = Mid(strFunction3(c), r) Then
                                Gall_Reply = Gall_Reply - 1
                            End If
                        Next q
                    End If
                    
                Next c
                End If
                
                If Check1.Value = 1 Then
                    If Gall_Reply = 0 Then
                        Lv_Dc.ListItems.Remove (Lv_Dc.ListItems.Count)
                        GoTo Pass
                    End If
                End If
                
                If viewPicture = True Then
                    Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ListSubItems.Add , , Gall_Title & " - [" & Gall_Reply & "]", 4
                Else
                    Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ListSubItems.Add , , Gall_Title & " - [" & Gall_Reply & "]"
                End If
                Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ListSubItems.Add , , Gall_Writer
                Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ListSubItems.Add , , Gall_Date
                
                r = InStr(Utf82String(DC.ResponseBody), "gallery_hits") + 17 '조회수
                l = InStr(r, Utf82String(DC.ResponseBody), "</li>")
                Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ListSubItems.Add , , Mid(Utf82String(DC.ResponseBody), r, l - r)
                
                r = InStr(Utf82String(DC.ResponseBody), "<div class=""m_contents"">") + 24
                l = InStr(r, Utf82String(DC.ResponseBody), "<div class=""btn_recommend_love")
                'ReDim Preserve gallContent(Lv_Dc.ListItems.Count)
                Gall_Filter = ""
                For q = 0 To lstFilter3.ListCount
                    If InStr(Mid(Utf82String(DC.ResponseBody), r, l - r), lstFilter3.List(q)) > 0 And Not lstFilter3.List(q) = "" Then
                        Gall_Filter = Gall_Filter & lstFilter3.List(q) & ","
                        Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ListSubItems(1).ForeColor = vbRed
                    End If
                Next q
                    Lv_Dc.ListItems(Lv_Dc.ListItems.Count).ListSubItems.Add , , Gall_Filter
Pass:
            End If
        Next z
    End If
Next

Passing:
Image1.Enabled = True
End Function

Private Sub Command1_Click()
Dim Filter As String
Filter = InputBox("필터링할 아이피를 입력하세요 예=(123.23.**.**)")
lstFilter.AddItem Filter
End Sub

Private Sub Command2_Click()
Dim Filter As String
Filter = InputBox("필터링할 닉네임을 입력하세요 예=(ㅂㅈㄷㄱ)")
lstFilter2.AddItem Filter
End Sub

Private Sub Command3_Click()
Gall_Page = Gall_Page + 1
Refreshs
Command4.Visible = True
Text1 = Gall_Page
End Sub

Private Sub Command4_Click()
Gall_Page = Gall_Page - 1
If Gall_Page = 1 Then Command4.Visible = False
Text1 = Gall_Page
Refreshs
End Sub

Private Sub Command5_Click()
Dim Filter As String
Filter = InputBox("필터링할 내용을 입력하세요 예=(ㅂㅈㄷㄱ)")
lstFilter3.AddItem Filter
End Sub

Private Sub Form_Load()
Gall_Ex = ""
Gall_Page = 1
OpenFile
lstFroxy.AddItem "216.177.**.**"
lstgall.ListIndex = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Image1_Click()
Refreshs
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Image10_Click()
ShellExecute Me.hWnd, vbNullString, "http://gall.dcinside.com/board/write/?id=compose", vbNullString, 1, 1
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Image7_Click()
Gall_Ex = ""
Refreshs
End Sub

Private Sub Image8_Click()
Gall_Ex = "&exception_mode=recommend"
Refreshs
End Sub

Private Sub Image9_Click()
Gall_Ex = "&exception_mode=best"
Refreshs
End Sub

Private Sub lstFilter_DblClick()
lstFilter.RemoveItem lstFilter.ListIndex
End Sub

Private Sub lstFilter2_DblClick()
lstFilter2.RemoveItem lstFilter2.ListIndex
End Sub

Private Sub lstFilter3_DblClick()
lstFilter3.RemoveItem lstFilter3.ListIndex
End Sub

Private Sub Lv_Dc_DblClick()
ShellExecute Me.hWnd, vbNullString, strHref(Lv_Dc.SelectedItem.Index - 1), vbNullString, 1, 1
End Sub

Private Sub Lv_Dc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Call ReleaseCapture
SendMessage Me.hWnd, &HA1, 2, 0&

End Sub

Private Sub txtContent_DblClick()
txtContent.Visible = False
WebBrowser.Visible = False
End Sub

Private Sub WebBrowser_LostFocus()
WebBrowser.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

If Dir(App.Path & "\Filter1.log", vbDirectory) <> "" Then
    Kill App.Path & "\Filter1.log"
End If

Open App.Path & "\Filter1.log" For Output As #1
For o = 0 To lstFilter.ListCount
    Print #1, lstFilter.List(o) & vbCrLf
Next
Close #1

If Dir(App.Path & "\Filter2.log", vbDirectory) <> "" Then
    Kill App.Path & "\Filter2.log"
End If

Open App.Path & "\Filter2.log" For Output As #1
For o = 0 To lstFilter2.ListCount
    Print #1, lstFilter2.List(o) & vbCrLf
Next
Close #1

If Dir(App.Path & "\Filter3.log", vbDirectory) <> "" Then
    Kill App.Path & "\Filter3.log"
End If

Open App.Path & "\Filter3.log" For Output As #1
For o = 0 To lstFilter3.ListCount
    Print #1, lstFilter3.List(o) & vbCrLf
Next
Close #1
End Sub
Private Function OpenFile()
If Dir(App.Path & "\Filter1.log", vbDirectory) <> "" Then
    Open App.Path & "\Filter1.log" For Input As #1
        Do While Not EOF(1)
            Line Input #1, i
            If i <> "" Then lstFilter.AddItem i
        Loop
    Close #1
End If

If Dir(App.Path & "\Filter2.log", vbDirectory) <> "" Then
Open App.Path & "\Filter2.log" For Input As #1
    Do While Not EOF(1)
        Line Input #1, i
        If i <> "" Then lstFilter2.AddItem i
    Loop
Close #1
End If


If Dir(App.Path & "\Filter3.log", vbDirectory) <> "" Then
Open App.Path & "\Filter3.log" For Input As #1
    Do While Not EOF(1)
        Line Input #1, i
        If i <> "" Then lstFilter3.AddItem i
    Loop
Close #1
End If
End Function
