VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  '없음
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "종료"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   960
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   3240
   End
   Begin SHDocVwCtl.WebBrowser w 
      Height          =   8415
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   6015
      ExtentX         =   10610
      ExtentY         =   14843
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label status 
      BackStyle       =   0  '투명
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "잠시만 기다려 주십시오."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   27.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   6090
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ws As String
Dim stp As Long
Dim MAINID As String
Dim MAINPWD As String
Dim REEET As Long
Dim mode As Long

Function Request(ByVal ck As String, ByRef sck As String)
On Error Resume Next
    Dim winhttp As New winhttp.WinHttpRequest
        
    winhttp.Open "POST", "https://mabinogi.nexon.com/page/common/gamestart.asp"
    winhttp.SetRequestHeader "Referer", "http://mabinogi.nexon.com/page/main/index.asp"
    winhttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36"
    winhttp.SetRequestHeader "Origin", "http://mabinogi.nexon.com"
    winhttp.SetRequestHeader "Accept-Language", "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7,lo;q=0.6"
    winhttp.SetRequestHeader "Accept", "text/plain, */*; q=0.01"
    winhttp.SetRequestHeader "Cookie", ck
    winhttp.Send
    
    sck = winhttp.GetAllResponseHeaders
    Request = winhttp.ResponseText
End Function

Function LoginOptions(ByVal idx As Long, ByRef ID As String, ByRef PASWD As String)
On Error Resume Next
    Dim contents As String
    Dim ap() As String
    
    Dim moda As Long
    Dim indexx As Long
    
    moda = 0
    
    If Dir(App.Path & "\login.ini") = "" Then
        Me.Top = Me.Top - 2000
        MsgBox "설정파일이 없습니다.", vbCritical, "Error"
        End
    End If
    
    Open App.Path & "\login.ini" For Input As #1
    
    Do While Not EOF(1) And indexx <= idx
        Line Input #1, contents
        
        If Left(contents, 2) = "G:" Then
            moda = 1
            contents = Mid(contents, 3)
        End If
        
        ap = Split(contents, ";")
        
        ID = RTrim(LTrim(ap(0)))
        PASWD = RTrim(LTrim(ap(1)))
        
        indexx = indexx + 1
    Loop
    
    Close #1
  
    LoginOptions = moda
End Function

Sub execute()
On Error Resume Next
    Dim s As String
    Dim sck As String
    Dim rety As Long
    
    
    setCookie "Cookie: " & w.Document.cookie & vbCrLf

    ws = "location.replace"

    Do While InStr(ws, "location.replace") > 0
        
        status.Caption = "세션 취득중...(" & rety + 1 & ")"
        
        If rety > 10 Then
            Form_Load
            Exit Sub
        End If
        
        ws = Request(getCookieAll(), sck)
        setCookie sck
        
        Sleep 300

        rety = rety + 1
    Loop

    If Len(ws) = 1 Then
        Me.Top = Me.Top - 2000
        MsgBox "로그인 할 수 없습니다. 계정을 확인해 주세요", vbCritical, "Error"
        End
    End If
    
    s = s + "<script src='https://platform.nexon.com/ngm/js/npf_ngm.js' type='text/javascript' charset='euc-kr'></script>"
    s = s + "<script src='https://platform.nexon.com/ngm/js/NGMModuleInfo.js' type='text/javascript' charset='euc-kr'></script>"
    
    w.Document.Clear
    w.Document.write s
    w.Refresh
    
    status.Caption = "세션 취득완료...."
    
    Timer1 = True
End Sub

Function ModeNavigate()
On Error Resume Next
    If mode = 0 Then
        w.Navigate "https://nxlogin.nexon.com/common/login.aspx"
        status.Caption = "넥슨 로그인 입니다..."
    ElseIf mode = 1 Then
        w.Navigate "https://accounts.google.com/signin/oauth/identifier?scope=profile%20email&response_type=code&client_id=919331056041-46d8sbblkb1ek3o02iva7vaiqth50clq.apps.googleusercontent.com&redirect_uri=https%3A%2F%2Flogin.nexon.com%2Flogin%2Fgoogle%2FAccessToken&state=O8zGof2fYUAgoGtHXrquizxmvYH0q4A_Zh8T~IyGsXAkl4eeyPjT9NOwHFQci3nUeUsgatptGxCojHe4UVuJBHOjFQ3iSstG8ctQSPQYybbXA7ZvIBQF5fEXka7VNUESX3ZRhmPfJ0Th~RIGXiXovI1BAHd5Cnty&include_granted_scopes=true&access_type=offline&o2v=2&as=a1HZb7ApMX5XtKjgTOIRxQ&flowName=GeneralOAuthFlow" ', , , , "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.163 Safari/537.36"
        status.Caption = "구글 로그인 입니다...(ie11 인증 필요)"
    End If
End Function

Private Sub Command1_Click()
    w.Navigate "about:blank"
    End
End Sub

Private Sub Form_DblClick()
    Me.Height = 11610
    Me.Top = Me.Top - ((11610 - 1965) / 2)
End Sub

Private Sub Form_Load()
    w.Silent = True
    stp = 0
    
    Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
           
    mode = LoginOptions(Val(LTrim(RTrim(Command))), MAINID, MAINPWD)
    Call ModeNavigate
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Timer1_Timer()
    If stp = 0 Then
        status.Caption = "마비노기를 실행합니다...."
        w.Navigate2 "javascript:NGM.LaunchGame('74245', '" & ws & "')"
        Timer1.Interval = 1000
    ElseIf stp = 1 Then
        End
    End If
    
    stp = stp + 1
End Sub

Private Sub Timer2_Timer()
    status.Caption = "재접속 중...."
    Call ModeNavigate
End Sub

Private Sub w_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
    Dim spstr As String
    spstr = Left(URL, 30)
    
    If InStr(w.Document.cookie, "NPP=;") > 0 Or InStr(w.Document.cookie, "NPP=") = 0 Then
        If InStr(spstr, "https://nxlogin") > 0 Then
            If REEET = 0 Then
                w.Document.getElementById("txtNexonID").value = MAINID
                status.Caption = MAINID & " 으로 로그인 중..."
                w.Document.getElementById("txtPWD").value = MAINPWD
                w.Document.getElementById("btnLogin").Click
                REEET = REEET + 1
            ElseIf REEET > 0 Then
                Me.Top = Me.Top - 2000
                MsgBox "로그인 할 수 없습니다. 계정을 확인해 주세요", vbCritical, "Error"
                End
            End If
        End If
    ElseIf InStr(spstr, "nexon") > 0 Then
        Call execute
        status.Caption = "로그인 완료"
    End If

End Sub
