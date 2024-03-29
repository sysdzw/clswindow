VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Windows程序自动化操作框架V2.2 演示"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   8550
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   5520
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   3240
      Width           =   4575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "微信发送消息"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton Command10 
      Caption         =   "鼠标画圆"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton Command9 
      Caption         =   "发消息到qq群"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "测 试"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "调用notepad进行一系列操作"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "控制“记事本”抖动、淡出、淡入"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "关闭所有记事本程序的窗口"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "设置计算器窗口特定排列"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "测试窗口激活"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "调用计算器进行计算"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton Command1x 
      Caption         =   "关闭纸牌等游戏窗口"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "根据句柄直接点击按钮"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3360
      TabIndex        =   7
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "会激活Chrome、IE、扫雷等窗口"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "会自动安排“计算器”花样出场，表演结束自行退出"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   4140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请先打开任意一个系统自带的游戏"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   2700
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim window As New clsWindow
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'关闭纸牌等游戏窗口
Private Sub Command1x_Click()
    Command1x.Enabled = False
'    If window.GetWindowByTitle("纸牌", 10).hWnd > 0 Then                                            '方法1：完全匹配方式，不支持正则
'    If window.GetWindowByTitleEx("蛛", 10).hWnd > 0 Then                                           '方法2：模糊匹配方式，含有此文字的就获取
'    使用正则匹配，对系统自带的几种游戏通杀
    If window.GetWindowByTitleEx("蜘蛛|纸牌|空当接龙|扫雷|红心大战|三维弹球", 10, , True).hWnd > 0 Then   '方法3：使用正则匹配，功能强大
        window.CloseWindow  '关闭窗口
'        window.CloseApp     '关闭进程，注意和上面方法的区别
    End If
    Command1x.Enabled = True
End Sub
'
Private Sub Command10_Click()
    Dim i%, w As New clsWindow
    Dim k As Single
    Const pi As Single = 3.14159
    Do While k < 2 * pi
        w.ClickPoint Cos(k) * 300 / 4 + 500, Sin(k) * 500 / 4 + 300, , , 10
        k = k + pi / 180
        DoEvents
    Loop
End Sub

Private Sub Command11_Click()
    sendWeixinMsg "朱磊", "hello"
End Sub
Private Sub sendWeixinMsg(ByVal strName$, ByVal strMsg$)
    Dim w As New clsWindow
    If w.GetWindowByClassName("WeChatMainWndForPC").hWnd <> 0 Then
        w.Focus
        w.ClickPoint 30, 100, , , 1200 '延时200ms后点击聊天板块
        w.ClickPoint 130, 40, , , 1200, 1500 '延时200ms后点击搜索框，点击后再延时500ms
        SendKeys strName
        w.ClickPoint 150, 130, , , 1500, 1500 '点击搜索结果
        SendKeys strMsg & "{ENTER}"
    Else
        MsgBox "未发现微信桌面版，请打开微信桌面版再测试！", vbExclamation
    End If
    Set w = Nothing
End Sub

'调用计算器进行计算
Private Sub Command2_Click()
    Dim strPID$
    Command2.Enabled = False
    strPID = Shell("calc.exe", 1)
'    window.GetWindowByTitleEx "计算器", 2        '方法1：不知道pid的情况下可使用，有同样标题的窗口时容易误取
    window.GetWindowByTitleEx "计算器", 2, strPID '方法2：如果知道pid那么进行验证获取窗口，防止误取
'    MsgBox window.AppPath '得到应用程序路径
'    MsgBox window.AppName '得到应用程序文件名
'    MsgBox window.AppCommandLine '得到应用程序启动参数
    
    '逐个点击计算器按键
'    window.SetElementTextByClassName "SciCalc", "计算2+3的结果："'方法1：通过类名定位设置窗口文字
'    window.setElementTextByText "计算器", "计算2+3的结果："      '方法2：通过窗口文字定位设置窗口文字
    window.Caption = "计算2+3的结果："                            '方法3：直接使用当前存储的句柄设置窗口文字
    window.Wait 500
    window.ClickElementByText "2"
    window.ClickElementByText "+"
    window.ClickElementByText "3"
    window.ClickElementByText "="
    
    '循环处理计算器按键
    Dim s$, i%
    window.Caption = "计算5*3*4/2+1+69=的结果："
    s = "5*3*4/2+1+69="
    For i = 1 To Len(s)
        window.ClickElementByText Mid(s, i, 1)
    Next
    window.Wait 2000
    
    window.Caption = "两秒后关闭计算器"
    window.Wait 2000
    window.CloseWindow
    Command2.Enabled = True
    window.Wait 2000
End Sub
'激活窗口获得焦点
Private Sub Command3_Click()
    Command3.Enabled = False
    
    Dim w As New clsWindow
    
    w.Wait 500
    w.GetWindowByTitleEx("Google Chrome", 1).Focus   '如果谷歌浏览器窗口获取到就激活
    w.Wait 500
    If w.GetWindowByTitleEx("扫雷", 1).hWnd > 0 Then w.Focus '如果扫雷游戏窗口获取到就激活，默认等待1秒
    w.Wait 500
    If w.GetWindowByTitleEx("Internet Explorer", 1).hWnd > 0 Then  '如果IE浏览器窗口获取到就激活
        w.Focus
        Randomize
        w.Move 800 * Rnd + 100, 300 * Rnd, 800 * Rnd + 500, 800 * Rnd + 100  '设置窗口位置的测试
        w.Caption = "窗口左顶点坐标：(" & w.Left & "," & w.Top & ") Internet Explorer"   '得到窗口左顶点坐标的测试
    End If
    
    Command3.Enabled = True
End Sub
'设置窗口位置
Private Sub Command4_Click()
    Dim i&, j&, lngLeft&, lngTop&, strHwnd$, v
    Command4.Enabled = False
    Dim w As New clsWindow

    '横排依次打开4个计算器
    For i = 1 To 4
        w.GetWindowByTitleEx "计算器", 2, Shell("calc.exe", 1)
        w.Caption = "句柄为:" & w.hWnd
        w.Move lngLeft, 300
        lngLeft = lngLeft + w.Width
        strHwnd = strHwnd & w.hWnd & ","
    Next
    
    w.GetWindowByTitleEx "计算器", 2, Shell("calc.exe", 1)
    strHwnd = strHwnd & w.hWnd & ","
    w.Wait 20
    '向下移动
    For i = 0 To Screen.Height / 15 - w.Height Step 10
        w.Wait 5
        w.Move lngLeft, i
    Next
    '再回头
    For i = Screen.Height / 15 - w.Height To 300 Step -10
        w.Wait 5
        w.Move lngLeft, i
    Next
    w.Move lngLeft, 300
    
    '竖排依次打开3个计算器，并且处理为渐渐移动到右边的效果
    For i = 1 To 3
        w.GetWindowByTitleEx "计算器", 2, Shell("calc.exe", 1)
        w.Caption = "句柄为:" & w.hWnd
        For j = 0 To Screen.Width / 15 - w.Width Step 10
            w.Wait 10
            w.Caption = "坐标(" & j & "." & lngTop & ")"
            w.Move j, lngTop
        Next
        lngTop = lngTop + w.Height
        strHwnd = strHwnd & w.hWnd & ","
    Next

    v = Split(strHwnd, ",")
    
    '激活测试，窗口全部激活一遍，注意标题会闪动一下
    w.Wait 1000
    For i = 0 To UBound(v) - 1
        w.hWnd = v(i)
        w.Focus
        w.FadeIn 20
    Next
    
    '1秒后逐步关闭所有刚刚打开的窗口
    w.Wait 1000
    For i = UBound(v) - 1 To 0 Step -1
        w.hWnd = v(i)
        w.Wait 100 '延时是为了看到逐步关闭的进程，不然一下子关了看不到效果
        w.CloseWindow
    Next
    Command4.Enabled = True
End Sub
'得到所有为资源管理器类的窗口句柄
Private Sub Command5_Click()
Command5.Enabled = False
'方法1 通过类名一次性获取Notepad所有窗口句柄：
'Dim w As New clsWindow
'Dim sHwndAll$, v, i%
'If w.GetWindowByClassNameEx("Notepad", 3, sHwndAll).hWnd > 0 Then    'xp下资源浏览器类名是CabinetWClass
'    v = Split(sHwndAll, " ") '所有符合条件的窗口句柄的集合
'    For i = 0 To UBound(v)
'        w.hWnd = v(i)
'        w.CloseWindow
'    Next
'End If

'方法2 循环关闭所有类名为Notepad的窗口，直到没有：
Dim w As New clsWindow
Do While w.GetWindowByClassName("Notepad", 1).hWnd > 0
    w.CloseWindow
Loop

'方法3 循环关闭所有标题栏含关键字“ - 记事本”窗口，直到没有：
'Dim w As New clsWindow
'Do While w.GetWindowByTitleEx(" - 记事本", 1).hWnd > 0
'    w.CloseWindow
'Loop


'方法4 根据进程名关闭Notepad，为防止错关闭，可以加if结合窗口标题控制：（win7下好像无效，问题没仔细研究）
'Dim w As New clsWindow
'Dim sHwndAll$, v, i%
'If w.GetWindowByAppNameEx("notepad.exe", 3, sHwndAll).hWnd > 0 Then    'xp下资源浏览器进程名是explorer.exe
'    v = Split(sHwndAll, " ") '所有符合条件的窗口句柄的集合
'    For i = 0 To UBound(v)
'        w.hWnd = v(i)
'            If w.Caption <> "" And w.Caption <> "Program Manager" And w.Caption <> "M" And w.Caption <> "Default IME" And w.Visible Then '必须可见并且内容不为空不是系统特殊的窗口
'            w.CloseWindow
'        End If
'    Next
'End If

Command5.Enabled = True
End Sub

Private Sub Command6_Click()
'Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl", 0
Dim w As New clsWindow
'MsgBox w.GetWindowByAppName("taskmgr.exe").Caption
'MsgBox w.GetWindowByTitle("Windows 任务管理器").Caption
'MsgBox w.GetWindowByAppName("notepad++.exe").Caption
'w.GetWindowByTitleEx("- 记事本").CloseWindow '关闭记事本窗口
'MsgBox w.GetWindowByTitleEx("记事本").GetElementTextByClassName("Edit") '取得记事本中的内容

'w.GetWindowByTitleEx("记事本").Hide
'w.GetWindowByTitleEx("记事本").Show

'w.GetWindowByTitleEx("记事本").SetTop

If w.GetWindowByTitleEx("记事本", 2).hWnd > 0 Then
    w.Focus '设置为焦点
    w.Shake  '抖动
    w.FadeOut '淡出
    w.FadeIn '淡入
Else
    MsgBox "找了2秒钟也没有发现打开着的notepad，逗我玩呢吧？", vbInformation
End If

'w.GetWindowByPID Shell("C:\Program Files (x86)\Microsoft Office\Office14\powerpnt.exe c:\实战销售技巧超级销售.ppt", 1)
'w.WAIT 2000
'w.Focus
'SendKeys "{F5}"
'w.WAIT 1000
'w.ClickElementByClassName "paneClassDC"
End Sub

Private Sub Command1_Click()
    Dim w As New clsWindow
''    If w.GetWindowByClassName("Notepad").hwnd <> 0 Then
''        w.Focus
''        w.Shake '默认抖动
''    End If
''
'Dim s$
''w.GetWindowByClassNameEx "DirectUIHWND", 1, s
''MsgBox s
'
'If w.GetWindowByTitle("Internet Explorer", 0).hwnd > 0 Then
'    w.GetWindowByHwnd(w.GetElementHwndByClassName("DirectUIHWND")).CloseWindow
'End If
'
''    w.GetWindowByClassName("DirectUIHWND", 0).CloseWindow
'
'
'GetWindowByTitle
'Dim w As New clsWindow
'If w.GetWindowByTitle("计算器", 0).hwnd <> 0 Then
'    Do
'        If Not w.CheckWindow Then
'            MsgBox "计算器已关闭"
'            Exit Do
'        End If
'        DoEvents
'    Loop
'Else
'    MsgBox "计算器未打开"
'End If

'    Dim s$, strTmp$
''    w.hwnd = 65962
''    w.Visible = True
''    Exit Sub
'
'    w.GetWindowByTitleEx "", 0, s, , , HiddenWindow
''    MsgBox s
'    Dim i%, v
'    v = Split(s, " ")
'    s = ""
'    For i = 0 To UBound(v)
'        strTmp = w.GetWindowTextByHwnd(v(i))
'        If strTmp = "WeNote" Then w.GetWindowByHwnd(v(i)).Visible = True
''        If strTmp <> "" Then s = s & v(i) & " " & strTmp & vbCrLf
'    Next
''    w.writeToFile "结果.txt", s

'w.DebugMe = True
'Dim w As New clsWindow
'If w.GetWindowByTitle("转到指定行", 0).hWnd <> 0 Then
'        w.SetElementTextByClassName "Edit", 999 '将行号改成999
'        w.ClickElementByClassName "Button" '点击“转到”按钮，如果要点击“取消”按钮可以用w.ClickElementByClassName "Button", 1
'    End If

'w.hWnd = Me.hWnd
'w.WindowState = vbMaximized

'If w.GetWindowByTitleEx("计算器").hWnd <> 0 Then
'    Dim hwndButton As Long
'    hwndButton = w.GetElementHwndByText("8")
'    w.GetWindowByHwnd(hwndButton).Focus
'End If

'MsgBox w.GetWindowByPID(12132).CommandLine
'MsgBox w.CommandLine(5504)
'MsgBox w.GetWindowByTitleEx("记事本").CommandLine
w.GetWindowByTitleEx "记事本"
w.Focus
w.Wait 1000
w.Focus
w.test
Me.Caption = Now
End Sub
'调用记事本然后写入一些内容后保存到c:\test.txt
Private Sub Command8_Click()
    Dim w As New clsWindow
'    w.DebugMe = True
    w.GetWindowByPID(Shell("notepad", 1)).Focus '根据pid直接获取窗口并设置为当前活动窗口
    w.Caption = "看到记事本打开了吗？" & Now '设置应用程序标题内容
    w.Shake '抖动窗口，可以通过参数调节抖动方向、速度、幅度、次数
    
    w.SetElementTextByClassName "Edit", "clsWindow" & w.Version & "发布 " & Now()    '设置输入框文字
    w.Wait 1000
    w.FadeOut '淡出
    w.AppendElementTextByClassName "Edit", vbCrLf & "功能全面" '向输入框追加内容
    w.FadeIn '淡入
    w.AppendElementTextByClassName "Edit", vbCrLf & "欢迎测试反馈" '向输入框追加内容
    w.Wait 1000
    
    w.Focus
    w.Wait 500
    w.SendKeys "^{s}" '设置焦点后按快捷键保存
    Dim w2 As New clsWindow
    w2.GetWindowByTitleEx("另存为").SetElementTextByClassName "Edit", "c:\test" & Format(Now, "yyyymmddhhnnss") & ".txt"
    w2.ClickElementByText "保存(&S)"
    
    MsgBox w.GetElementTextByClassName("Edit")
End Sub

Private Sub Command9_Click()
'    Dim i%
'    Dim w As New clsWindow
'    For i = 1 To 40
'        w.GetWindowByTitle("VB高级语言学习交流群").Focus
'        w.ClickPoint w.Left + 35, w.Top + w.Height - 100, absolute
'        w.Wait 20
'        Clipboard.Clear
'        Clipboard.SetText "[" & i & "]发送 本消息由程序clswindow" & w.Version & "类发出 QQ ：788028734" & Now
'        SendKeys "^{v}"
'        SendKeys "%{s}"
'        w.Wait 2
'    Next
    
    sendQQMsg "clswindow交流群", "大家好，框架交流群QQ：788028734"
    sendQQMsg "石头", "hello!"
End Sub
'sendQQMsg "QQ窗口标题", "测试"
Private Sub sendQQMsg(ByVal strName$, ByVal strMsg$)
    Dim w As New clsWindow
    If w.GetWindowByTitle(strName).hWnd <> 0 Then
        w.SetPosNormal
        w.Focus
        w.ClickPoint w.Left + 35, w.Top + w.Height - 100, ABSOLUTE
        w.Wait 20
        Clipboard.Clear
        Clipboard.SetText strMsg
        SendKeys "^{v}"
        SendKeys "%{s}"
        w.Wait 2
    Else
        MsgBox "未发现包含“strName”的QQ聊天窗口，请打开对应的窗口再测试，注意请在面板上取消勾选“合并勾选窗口”", vbExclamation
    End If
    Set w = Nothing
End Sub

Private Sub Form_Load()
    Dim w As New clsWindow
    Me.Show
    w.Load(Me.hWnd).FadeIn '淡入
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then '始终将本窗口设置为活动窗口
        Dim w As New clsWindow
        w.hWnd = Me.hWnd
        w.Focus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub HScroll1_Change()
    Dim w As New clsWindow
    If w.GetWindowByClassName("Notepad", 0).hWnd <> 0 Then
        w.MakeTransparent HScroll1.Value / HScroll1.Max * 100 & "%"
    Else
        MsgBox "请先打开一个记事本文件，然后再拉动滚动条，注意观察记事本窗口变化", vbInformation
    End If
End Sub

Private Sub HScroll1_Scroll()
    Call HScroll1_Change
End Sub

