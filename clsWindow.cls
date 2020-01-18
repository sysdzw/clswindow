VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================================
'名    称：windows窗体控制类v2.2
'描    述：一个操作windows窗口的类，可对窗口进行很多常用的操作(类名为clsWindow)
'使用范例：Dim window As New clsWindow
'          window.GetWindowByTitle("计算器").closeWindow
'编    程：sysdzw 原创开发，如果有需要对模块扩充或更新的话请邮箱发我一份，共同维护
'发布日期：2013/06/01
'博    客：http://blog.csdn.net/sysdzw
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                                           2012/12/03
'          V1.1 修正了几个正则相关的函数，调整了部分类结构                       2013/05/28
'          V1.2 增加属性Caption，可以获取或设置当前标题栏                        2013/05/29
'          V1.3 增加了方法Focus，可以激活当前窗口                               2013/06/01
'               增加了方法Left,Top,Width,Height,Move，处理窗口位置等
'          V1.4 增加了窗口位置调整的几个函数                                    2013/06/04
'               增加了得到应用程序路径的函数AppName
'               增加了得到应用程序启动参数的函数AppCommandLine
'          V1.5 增加了窗口最大最小化，隐藏显示正常的几个函数                     2013/06/06
'               增加了获取控件相关函数是否使用正则的参数UseRegExp默认F
'          V1.6 将Left，Top函数改为属性，可获得可设置                           2013/06/10
'          V1.7 增加函数：CloseApp 结束进程                                    2013/06/13
'               修正了部分跟正则匹配相关的函数
'               增加函数：GetElementTextByText
'               增加函数：GetElementHwndByText
'          V1.8 增加函数：GetWindowByClassName                                 2013/06/26
'               增加函数：GetWindowByClassNameEx
'               增加函数：GetWindowByAppName
'               增加私有变量hWnd_
'               增加属性hWnd，可设置，单设置时候会检查，非法则设置为0
'               更新GetWindowByTitleEx函数，使之可以选择性支持正则
'               删除GetWindowByTitleRegExp函数，合并到上面函数
'               增加SetFocus函数，调用Focus实现，为了是兼容VB习惯
'               扩了ProcessID、AppPath、AppName、AppCommandLine三个函数，可带参数
'               网友wwb(wwbing@gmail.com)提供了一些函数和方法属性：
'                 CheckWindow, Load, WindowState, Visible, hDC, ZOrder
'                 AlphaBlend, Enabled, Refresh, TransparentColor
'               采纳wwb网友的部分意见，将句柄变量改为hWnd_，但是hWnd作为公共属性
'          V1.9 修正函数：GetMatchHwndFromWindow 正则表达式的错误               2013/08/07
'               修正函数：GetMatchHwndFromWindow 函数中的一些错误               2014/09/23
'               增加函数：GetWindowByClassNameEx
'               增加函数：GetWindowByPID 根据PID取窗口句柄
'               增加函数：GetCaptionByHwnd 根据句柄取得标题
'               增加函数：SetTop设置窗体置顶，传入参数false则取消                2014/09/24
'               增加函数：Shake、FadeIn、FadeOut 抖动、淡入、淡出特效
'          V2.0 修正函数：GetWindowByPID 遍历窗体Win7下有一些问题               2015/09/29
'               修正函数：GetWindowByAppName 遍历窗体Win7下有一些问题
'               修正函数：GetWindowByAppNameEx 遍历窗体Win7下有一些问题
'          V2.1 修正函数：ClickPoint 增加位置模式参数相对和绝对，默认相对        2018/06/05
'               增加函数：SelectComboBoxIndex 根据指定的index选择下拉框中的项
'                         上述方法得到网友Chen8013的不少帮助，特此感谢
'               增加函数：GetWindowByHwnd 根据指定的句柄确定窗口                2018/07/22
'               增加函数：GetWindowByCursorPos 根据当前光标获取窗口（控件）
'               增加函数：GetWindowByPoint 根据指定的位置获取窗口（控件）
'               升级ClickPoint函数，支持点击前后分别延时，默认延时为0            2018/07/23
'          V2.2 根据网友小凡的bug反馈（句柄和id负数的情况），所以修正了相关正则   2020/01/08
'               优化属性：Caption(Get)，根据网友小凡的建议改成可获得文本框内容
'               增加方法：Wait 此方法原为clsWaitableTimer模块中，现集成进来      2020/01/09
'               增加方法：ClickCurrentPoint 点击当前点                          2020/01/10
'               增加方法：SetCursor(别名:SetPoint MoveCursor MoveCursorTo)
'               更新函数：将所有默认等待超时60秒的函数中默认等待时间都改为10秒
'               增加属性：Text、Value、Title（均为Caption别名）                 2020/01/12
'               优化代码：GetCaptionByHwnd采用原Caption(Get)代码，后者也做了调整
'               增加函数：GetCursorPosCurrent(别名：GetCursorPoint)得到当前坐标
'               优化函数：所有窗口获取的函数增加了是否过滤可见的参数              2020/01/16
'               增加函数：GetTextByHwnd（同GetCaptionByHwnd）
'==============================================================================================
Option Explicit

Private Const DelayOneTime = 500
Private hWnd_ As Long '窗口句柄
Private wReturn As New clsWindow '作为返回对象以便连用
Private Const strVersion = "V2.2"
Public DebugMe As Boolean '表示是否启用调试模式

Public Function Move(Optional ByVal x, Optional ByVal y, Optional ByVal nWidth, Optional ByVal nHeight) As Long
    If IsMissing(x) Then x = Left
    If IsMissing(y) Then y = Top
    If IsMissing(nWidth) Then nWidth = Width
    If IsMissing(nHeight) Then nHeight = Height
    Move = MoveWindow(hWnd_, x, y, nWidth, nHeight, True)
End Function
'得到当前窗口的句柄
Public Property Get hWnd() As Long
    hWnd = hWnd_
End Property
'设置类的句柄
Public Property Let hWnd(ByVal hWnd As Long)
    If IsWindow(hWnd) Then
        hWnd_ = hWnd
    Else
        hWnd_ = 0
    End If
End Property
'得到窗口的Left位置
Public Property Get Left() As Long
    Dim rect As rect
    If hWnd_ <> 0 Then
        GetWindowRect hWnd_, rect
        Left = rect.Left
    Else
        Left = -1
    End If
End Property
'设置窗口的Left位置
Public Property Let Left(ByVal lngLeft As Long)
    Move lngLeft
End Property
'得到窗口的Top位置
Public Property Get Top() As Long
    Dim rect As rect
    If hWnd_ <> 0 Then
        GetWindowRect hWnd_, rect
        Top = rect.Top
    Else
        Top = -1
    End If
End Property
'设置窗口的Top位置
Public Property Let Top(ByVal lngTop As Long)
    Move , lngTop
End Property
'得到窗口的Width大小
Public Property Get Width() As Long
    Dim rect As rect
    If hWnd_ <> 0 Then
        GetWindowRect hWnd_, rect
        Width = rect.Right - rect.Left
    Else
        Width = -1
    End If
End Property
'设置窗口的Width
Public Property Let Width(ByVal lngWidth As Long)
    Move , , lngWidth
End Property
'得到窗口的Height大小
Public Property Get Height() As Long
    Dim rect As rect
    If hWnd_ <> 0 Then
        GetWindowRect hWnd_, rect
        Height = rect.Bottom - rect.Top
    Else
        Height = -1
    End If
End Property
'设置窗口的Width
Public Property Let Height(ByVal lngHeight As Long)
    Move , , , lngHeight
End Property

'设置当前窗口的标题栏文字
Public Property Let Caption(ByVal strNewText As String)
    If hWnd_ <> 0 Then
        SendMessage hWnd_, WM_SETTEXT, 0&, ByVal strNewText
    End If
End Property
'得到当前窗口的标题栏文字
Public Property Get Caption() As String
    Caption = GetTextByHwnd(hWnd_)
End Property
'根据句柄获得窗口内容
Public Function GetCaptionByHwnd(ByVal hWnd As Long) As String
    GetCaptionByHwnd = GetTextByHwnd(hWnd_)
End Function
'根据句柄获得窗口内容
Public Function GetTextByHwnd(ByVal hWnd As Long) As String
    Dim Txt2(64000) As Byte
    SendMessage hWnd, WM_GETTEXT, 64000, Txt2(0)
    GetTextByHwnd = Split(StrConv(Split(Txt2, Chr$(0), 2)(0), vbUnicode) & Chr$(0), Chr$(0), 2)(0)
End Function
'设置当前窗口的标题栏文字
Public Property Let Text(ByVal strNewText As String)
    Caption = strNewText
End Property
'得到当前窗口的标题栏文字
Public Property Get Text() As String
    Text = Caption
End Property
'设置当前窗口的标题栏文字
Public Property Let Value(ByVal strNewText As String)
    Caption = strNewText
End Property
'得到当前窗口的标题栏文字
Public Property Get Value() As String
    Value = Caption
End Property
'设置当前窗口的标题栏文字
Public Property Let Title(ByVal strNewText As String)
    Caption = strNewText
End Property
'得到当前窗口的标题栏文字
Public Property Get Title() As String
    Title = Caption
End Property

'窗口置前
Public Function SetPosFront() As Long
    If hWnd_ <> 0 Then
        SetPosFront = SetWindowPos(hWnd_, -1, 0, 0, 0, 0, 3)
    End If
End Function
'窗口置后
Public Function SetPosBottom() As Long
    If hWnd_ <> 0 Then
        SetPosBottom = SetWindowPos(hWnd_, 1, 0, 0, 0, 0, 3)
    End If
End Function
'设置窗口位置正常
Public Function SetPosNormal() As Long
    If hWnd_ <> 0 Then
        SetPosNormal = SetWindowPos(hWnd_, -2, 0, 0, 0, 0, 3)
    End If
End Function
'设置窗口最小化
Public Function Minimized() As Long
    If hWnd_ <> 0 Then
        Minimized = ShowWindow(hWnd_, SW_SHOWMINIMIZED)
    End If
End Function
'设置窗口最大化
Public Function Maximized() As Long
    If hWnd_ <> 0 Then
        Maximized = ShowWindow(hWnd_, SW_SHOWMAXIMIZED)
    End If
End Function
'设置窗口位置隐藏
Public Function Hide() As Long
    If hWnd_ <> 0 Then
        Hide = ShowWindow(hWnd_, SW_HIDE)
    End If
End Function
'设置窗口显示
'Public Const SW_MINIMIZE = 6
'Public Const SW_SHOW = 5
'Public Const SW_SHOWMAXIMIZED = 3
'Public Const SW_SHOWMINIMIZED = 2
'Public Const SW_HIDE = 0
'Public Const SW_NORMAL = 1
'Public Const SW_SHOWNORMAL = 1
Public Function Show(Optional ByVal ShowStatus As VbAppWinStyle = 5) As Long '默认是5，直接显示
    If hWnd_ <> 0 Then
        Show = ShowWindow(hWnd_, ShowStatus)
    End If
End Function
'设置窗口正常显示
Public Function Normal() As Long
    If hWnd_ <> 0 Then
        Normal = ShowWindow(hWnd_, SW_NORMAL)
    End If
End Function
'设置当前窗口为活动窗口
Public Sub Focus()
    Dim hWnd&, hForeWnd&, dwForeID&, dwCurID&
    hWnd = hWnd_
    hForeWnd = GetForegroundWindow()
    dwForeID = GetWindowThreadProcessId(hForeWnd, 0)
    dwCurID = GetCurrentThreadId()
    AttachThreadInput dwCurID, dwForeID, True
'    ShowWindow hWnd, SW_SHOWNORMAL
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    SetForegroundWindow hWnd
    AttachThreadInput dwCurID, dwForeID, False
End Sub
'支持和VB中通常名称的方法
Public Sub SetFocus()
    Call Focus
End Sub
'得到进程标识
Public Function ProcessID(Optional ByVal my_hWnd) As Long
    Dim l&
    If IsMissing(my_hWnd) Then
        If hWnd_ <> 0 Then GetWindowThreadProcessId hWnd_, l
    Else
        GetWindowThreadProcessId my_hWnd, l
    End If
    ProcessID = l
End Function
'得到进程标识
Public Function PID(Optional ByVal my_hWnd) As Long
    If IsMissing(my_hWnd) Then
        PID = ProcessID
    Else
        PID = ProcessID(my_hWnd)
    End If
End Function
'根据窗口句柄得到应用程序的路径
Public Function AppPath(Optional ByVal hwn1) As String
    Dim dwProcessID As Long, hProcess As Long, hModule As Long
    Dim nSize As Long
    If IsMissing(hwn1) Then '如果没写则表示处理当前存储的句柄
        If hWnd_ <> 0 Then dwProcessID = ProcessID() 'GetWindowThreadProcessId hWnd_, dwProcessID
    Else
        dwProcessID = ProcessID(hwn1)
    End If
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, dwProcessID)
    AppPath = Space$(255)
    nSize = GetModuleFileNameEx(hProcess, 0, StrPtr(AppPath), 255)
    AppPath = Mid$(AppPath, 1, nSize)
    Call CloseHandle(hProcess)
End Function
'得到exe名字,如果缺省就获取自己的exe名称
Public Function AppName(Optional ByVal hwnd1) As String
    Dim v
    Static i%
    If IsMissing(hwnd1) Then
        If hWnd_ <> 0 Then v = Split(AppPath, "\")
    Else
        v = Split(AppPath(hwnd1), "\")
    End If
    i = i + 1
    If UBound(v) >= 0 Then AppName = v(UBound(v))
End Function
'得到exe文件的启动参数
Public Function AppCommandLine(Optional ByVal strAppName) As String
    If hWnd_ <> 0 Then
        Dim objWMIService As Object
        Dim colProcessList As Object
        Dim objProcess As Object
        Dim objProType As Object
        Dim strResult As String
        Set objWMIService = GetObject("winmgmts=" & "{impersonationlevel=impersonate}!//./root/cimv2")
        If IsMissing(strAppName) Then
            Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & AppName() & "'")
        Else
            Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & strAppName & "'")
        End If
         
        If colProcessList.Count <> 0 Then
            For Each objProcess In colProcessList
                For Each objProType In objProcess.Properties_
                    If objProType.Name = "CommandLine" Then
                        strResult = strResult & objProType.Value & vbCrLf
                        Exit For
                    End If
                Next
            Next
        End If
        If strResult <> "" Then
            AppCommandLine = strResult
        End If
    End If
End Function
'设置根据窗体标题得到窗口句柄，可以指定等待几秒检测指定标题的窗体是否出现，默认是等待10秒
Public Function GetWindowByTitle(ByVal strTitle$, Optional ByVal intWaitSeconds& = 10, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    Dim lngDelayCount&
    hWnd_ = 0
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetWindowByTitle被调用" & vbCrLf & "参数为 strTitle=" & strTitle & vbTab & " intWaitSeconds=" & intWaitSeconds, False
    End If
            
    Do
        hWnd_ = FindWindow(vbNullString, strTitle)
        If CheckWindowVisible(hWnd_, windowVisible) Then Exit Do '窗口可见度检查，符合条件表示找到了即退出

        lngDelayCount = lngDelayCount + DelayOneTime
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '未发现关键字为 strTitle
        Wait DelayOneTime
        
        hWnd_ = 0
    Loop
    
    wReturn.hWnd = hWnd_
    Set GetWindowByTitle = wReturn
End Function
'检查窗口可见是否符合指定要求
Private Function CheckWindowVisible(ByVal lngHwnd&, ByVal windowVisible As enumWindowVisible) As Boolean
    If lngHwnd <> 0 Then
        If windowVisible = AllWindow Then
            CheckWindowVisible = True
        ElseIf windowVisible = DisplayedWindow Then
            If IsWindowVisible(lngHwnd) Then CheckWindowVisible = True
        ElseIf windowVisible = HiddenWindow Then
            If Not IsWindowVisible(lngHwnd) Then CheckWindowVisible = True
        End If
    End If
End Function
'和GetWindowByTitle函数功能类似，只是这个是模糊匹配
Public Function GetWindowByTitleEx(ByVal strTitle$, Optional ByVal intWaitSeconds& = 10, Optional ByRef hWndAll, Optional isUseRegExp As Boolean = False, Optional ByVal checkPid = "", Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    Dim lngHwnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strHwndAllTmp$
    Dim isMatch As Boolean
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetWindowByTitleEx被调用" & vbCrLf & "参数为 strTitle=" & strTitle & vbTab & " intWaitSeconds=" & intWaitSeconds, False
    End If
    Do
        hWnd_ = 0
        Do
            lngHwnd = FindWindowEx(0, lngHwnd, vbNullString, vbNullString)
            If CheckWindowVisible(lngHwnd, windowVisible) Then  '窗口可见度检查，符合条件则继续下面的检查
                strCaption = GetTextByHwnd(lngHwnd)
                
                If Not isUseRegExp Then
                    isMatch = InStr(strCaption, strTitle) > 0
                Else
                    isMatch = regTest(strCaption, strTitle)
                End If
                If isMatch Then
                    strHwndAllTmp = strHwndAllTmp & CStr(lngHwnd) & " "
                    If checkPid <> "" Then '如果pid相等，优先放到第一个
                        If checkPid = ProcessID() Then strHwndAllTmp = lngHwnd & " " & strHwndAllTmp
                    End If
                End If
            End If
        Loop Until lngHwnd = 0 '这层loop是遍历所有窗口的
        If strHwndAllTmp <> "" Then Exit Do '如果一遍循环得到结果了就退出

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    strHwndAllTmp = Trim$(strHwndAllTmp)
    If Not IsMissing(hWndAll) Then hWndAll = strHwndAllTmp '如果需要搜集所有句柄，那么则使用这个参数返回

    If strHwndAllTmp <> "" Then
        wReturn.hWnd = Split(strHwndAllTmp, " ")(0)
    Else
        wReturn.hWnd = 0
    End If
    Set GetWindowByTitleEx = wReturn
    hWnd_ = wReturn.hWnd
End Function
'设置根据窗体类名返回窗体句柄，可以指定等待几秒检测指定标题的窗体是否出现，默认是等待10秒
Public Function GetWindowByClassName(ByVal strClassName$, Optional ByVal intWaitSeconds& = 10, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    Dim lngDelayCount&
    hWnd_ = 0
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetWindowByClassName被调用" & vbCrLf & "参数为 strClassName=" & strClassName & vbTab & " intWaitSeconds=" & intWaitSeconds, False
    End If
            
    Do
        hWnd_ = FindWindow(strClassName, vbNullString)
        If CheckWindowVisible(hWnd_, windowVisible) Then Exit Do '窗口可见度检查，符合条件表示找到了即退出

        lngDelayCount = lngDelayCount + DelayOneTime
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strClassName
        Wait DelayOneTime
        
        hWnd_ = 0
    Loop

    wReturn.hWnd = hWnd_
    Set GetWindowByClassName = wReturn
End Function
'设置根据窗体类名得到窗口句柄，可以指定等待几秒检测指定类名的窗体是否出现，默认是等待10秒，允许使用正则表达式
Public Function GetWindowByClassNameEx(ByVal strClassName$, Optional ByVal intWaitSeconds& = 10, Optional ByRef hWndAll, Optional isUseRegExp = False, Optional ByVal checkPid = "", Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    Dim lngHwnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strWindowClass As String * 255
    Dim strHwndAllTmp$
    Dim isMatch As Boolean
    
    hWnd_ = 0
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetWindowByClassNameEx被调用" & vbCrLf & "参数为 strClassName=" & strClassName & vbTab & " intWaitSeconds=" & intWaitSeconds, False
    End If
    Do
        Do
            lngHwnd = FindWindowEx(0, lngHwnd, vbNullString, vbNullString)
            If CheckWindowVisible(lngHwnd, windowVisible) Then
                GetClassName lngHwnd, strWindowClass, 255  '获得窗口类
                strWindowClass = Replace(strWindowClass, Chr$(0), "")
                If Not isUseRegExp Then
                    isMatch = InStr(strWindowClass, strClassName) > 0
                Else
                    isMatch = regTest(strWindowClass, strClassName)
                End If
                If isMatch Then
                    strHwndAllTmp = strHwndAllTmp & CStr(lngHwnd) & " "
                    If checkPid <> "" Then '如果pid相等，优先放到第一个
                        If checkPid = ProcessID() Then strHwndAllTmp = lngHwnd & " " & strHwndAllTmp
                    End If
                End If
            End If
        Loop Until lngHwnd = 0
        If strHwndAllTmp <> "" Then Exit Do

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    strHwndAllTmp = Trim$(strHwndAllTmp)
    If Not IsMissing(hWndAll) Then hWndAll = strHwndAllTmp '如果需要搜集所有句柄，那么则使用这个参数返回
    
    If strHwndAllTmp <> "" Then
        wReturn.hWnd = Split(strHwndAllTmp, " ")(0)
    Else
        wReturn.hWnd = 0
    End If
    Set GetWindowByClassNameEx = wReturn
    hWnd_ = wReturn.hWnd
End Function
'设置根据进程名称返回窗体句柄，可以指定等待几秒检测指定进程名称的窗体是否出现，默认是等待10秒
Public Function GetWindowByAppName(ByVal strAppName$, Optional ByVal intWaitSeconds& = 10, Optional ByVal isWholeMatch As Boolean = False, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    Dim lngHwnd&, currWnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strWindowClass As String * 255
    Dim isMatch As Boolean
    Dim strGetHwndAppName$
    Dim strHwnd
    
    strAppName = LCase$(strAppName)
    writeToFile "DEBUG_exe.txt", ""
    hWnd_ = 0
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetWindowByAppName被调用" & vbCrLf & "参数为 strAppName=" & strAppName & vbTab & " intWaitSeconds=" & intWaitSeconds, False
    End If
    Do
        currWnd = GetDesktopWindow()
        currWnd = GetWindow(currWnd, GW_CHILD)
        Do While currWnd <> 0
            If CheckWindowVisible(currWnd, windowVisible) Then
                strGetHwndAppName = LCase$(AppName(currWnd)) '根据句柄得到进程名字，都转换成小写
                If isWholeMatch Then '判断是否完全相等
                    isMatch = (LCase$(strGetHwndAppName) = strAppName)
                Else
                    isMatch = (InStr(LCase$(strGetHwndAppName), strAppName) > 0)
                End If
                strCaption = GetTextByHwnd(currWnd)
                If InStr("|MSCTFIME UI|Program Manager|M|Default IME|", strCaption) = 0 And isMatch Then '过滤掉输入法相关的窗口
                    hWnd_ = currWnd
                    Exit Do
                End If
            End If
            currWnd = GetWindow(currWnd, GW_HWNDNEXT)
            
            DoEvents
        Loop
        If hWnd_ <> 0 Then Exit Do

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    wReturn.hWnd = hWnd_
    Set GetWindowByAppName = wReturn
End Function
'设置根据进程名称返回窗体句柄，可以指定等待几秒检测指定进程名称的窗体是否出现，默认是等待10秒，原理还是枚举所有窗体检查它们的所属进程，只不过是根据目标的进程名跟strAppName比对
Public Function GetWindowByAppNameEx(ByVal strAppName$, Optional ByVal intWaitSeconds& = 10, Optional ByRef hWndAll, Optional isUseRegExp = False, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    Dim lngHwnd&, currWnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strWindowClass As String * 255
    Dim strHwndAllTmp$
    Dim isMatch As Boolean
    Dim strGetHwndAppName$
    Dim strHwnd
    
    strAppName = LCase$(strAppName)
    
    hWnd_ = 0
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetWindowByAppNameEx被调用" & vbCrLf & "参数为 strAppName=" & strAppName & vbTab & " intWaitSeconds=" & intWaitSeconds, False
    End If
    strHwndAllTmp = ""
    Do
        currWnd = GetDesktopWindow()
        currWnd = GetWindow(currWnd, GW_CHILD)
        Do While currWnd <> 0
            If CheckWindowVisible(currWnd, windowVisible) Then
                strGetHwndAppName = LCase$(AppName(currWnd))  '根据句柄得到进程名字，都转换成小写
                If Not isUseRegExp Then '如果不用正则就包含关系匹配
                    isMatch = (InStr(strGetHwndAppName, strAppName) > 0)
                Else
                    isMatch = regTest(strGetHwndAppName, strAppName)
                End If
                strCaption = GetTextByHwnd(currWnd)
                If InStr("|MSCTFIME UI|Program Manager|M|Default IME|", strCaption) = 0 And isMatch Then '过滤掉输入法相关的窗口
                    strHwndAllTmp = strHwndAllTmp & CStr(currWnd) & " "
                End If
            End If
            currWnd = GetWindow(currWnd, GW_HWNDNEXT)
            DoEvents
        Loop
        If strHwndAllTmp <> "" Then Exit Do '表示匹配到了一部分数据

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
    Loop
    strHwndAllTmp = Trim$(strHwndAllTmp)
    If Not IsMissing(hWndAll) Then hWndAll = strHwndAllTmp '如果需要搜集所有句柄，那么则使用这个参数返回
    
    If strHwndAllTmp <> "" Then
        wReturn.hWnd = Split(strHwndAllTmp, " ")(0)
    Else
        wReturn.hWnd = 0
    End If
    Set GetWindowByAppNameEx = wReturn
    hWnd_ = wReturn.hWnd
End Function
'根据pid获得窗体
Public Function GetWindowByPID(ByVal PID As Long, Optional ByVal intWaitSeconds& = 10, Optional ByRef hWndAll, Optional ByVal windowVisible As enumWindowVisible = AllWindow) As clsWindow
    Dim lngHwnd&, currWnd&, l&, lngDelayCount&
    Dim strCaption As String
    Dim strWindowClass As String * 255
    Dim strHwndAllTmp$
    Dim isMatch As Boolean
    Dim currentPID$
    Dim strHwnd
    
    hWnd_ = 0
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetWindowByPID被调用" & vbCrLf & "参数为 PID=" & PID & vbTab & " hWndAll=" & hWndAll, False
    End If
    strHwndAllTmp = ""
    Dim aaa&, bbb&
    Do '循环找几遍
        currWnd = GetDesktopWindow()
        currWnd = GetWindow(currWnd, GW_CHILD)
        Do While currWnd <> 0
            If CheckWindowVisible(currWnd, windowVisible) Then
                currentPID = ProcessID(currWnd) '得到pid
                isMatch = (currentPID = PID)
                If isMatch Then
                    strCaption = GetTextByHwnd(currWnd)
    
                    If InStr("|MSCTFIME UI|Program Manager|M|Default IME|", strCaption) = 0 Then '过滤掉输入法相关的窗口
                        strHwndAllTmp = strHwndAllTmp & CStr(currWnd) & " "
                    End If
                End If
            End If
            currWnd = GetWindow(currWnd, GW_HWNDNEXT)
            DoEvents
        Loop
        If strHwndAllTmp <> "" Then Exit Do '表示匹配到了一部分数据

        Wait DelayOneTime
        lngDelayCount = lngDelayCount + DelayOneTime '计算累计延时了多少，然后和用户设置的延时时间对比
        If lngDelayCount >= intWaitSeconds * 1000 Then Exit Do '超时未发现关键字strTitle
        Exit Do
    Loop
    strHwndAllTmp = Trim$(strHwndAllTmp)
    If Not IsMissing(hWndAll) Then hWndAll = strHwndAllTmp '如果需要搜集所有句柄，那么则使用这个参数返回
    
    If strHwndAllTmp <> "" Then
        Dim vHwnd
        vHwnd = Split(strHwndAllTmp, " ")
        wReturn.hWnd = vHwnd(UBound(vHwnd))
    Else
        wReturn.hWnd = 0
    End If
    Set GetWindowByPID = wReturn
    hWnd_ = wReturn.hWnd
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'以下几个方法为网友wwb（邮箱：wwbing@gmail.com）所增加
'日期：2013-6-26
'具体有：CheckWindow, Load, WindowState, Visible, hDC,
'       AlphaBlend, Enabled, Refresh, TransparentColor, ZOrder
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'检查当前窗口句柄是否还有效
Private Function CheckWindow() As Long
    CheckWindow = IsWindow(hWnd_)
    If CheckWindow = 0 Then hWnd_ = 0 '如果窗口不存在，我们要将当前对象的句柄修改为0，标准是不去修改，但是不修改，一个无效的句柄也没用，而且可能会被其他窗口随机分配上导致出错
End Function
'Load 窗口载入的入口
Public Function Load(WindowHwnd As Variant) As clsWindow
    If IsNumeric(WindowHwnd) Then
        hWnd_ = CLng(WindowHwnd)
        If CheckWindow Then
            wReturn.hWnd = hWnd_
            Set Load = wReturn
        Else
            MsgBox "类初始化错误！句柄“" & WindowHwnd & "”对应的窗口不存在！", vbExclamation
        End If
    Else
        MsgBox "类初始化错误！设置的句柄“" & WindowHwnd & "”应当为一个数字。", vbExclamation
    End If
End Function
'直接由句柄指定窗口，与直接设置w.hwnd或者load方法同等效果
Public Function GetWindowByHwnd(WindowHwnd As Variant) As clsWindow
    If IsNumeric(WindowHwnd) Then
        hWnd_ = CLng(WindowHwnd)
        If CheckWindow Then
            wReturn.hWnd = hWnd_
            Set GetWindowByHwnd = wReturn
        Else
            MsgBox "类初始化错误！句柄“" & WindowHwnd & "”对应的窗口不存在！", vbExclamation
        End If
    Else
        MsgBox "类初始化错误！设置的句柄“" & WindowHwnd & "”应当为一个数字。", vbExclamation
    End If
End Function
'直接获得当前鼠标下的窗口或控件的句柄
Public Function GetWindowByCursorPos() As clsWindow
    Dim tPoint As POINTAPI
    GetCursorPos tPoint '获得当前鼠标位置
    hWnd_ = WindowFromPoint(tPoint.x, tPoint.y)
    wReturn.hWnd = hWnd_
    Set GetWindowByCursorPos = wReturn
End Function
'根据指定的点获取窗口或控件的句柄
Public Function GetWindowByPoint(ByVal x As Long, ByVal y As Long) As clsWindow
    hWnd_ = WindowFromPoint(x, y)
    wReturn.hWnd = hWnd_
    Set GetWindowByPoint = wReturn
End Function
'根据控件的类名设置控件的显示文字
Public Function SetElementTextByClassName(ByVal strClassName$, ByVal strNewText$, Optional ByVal intIndex% = 1, Optional ByVal UseRegExp As Boolean = False) As Boolean
    Dim lngControlsHwnd As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数SetElementTextByClassName被调用" & vbCrLf & "参数为 strClassName=" & strClassName & vbTab & " strNewText=" & strNewText & vbTab & " intIndex=" & intIndex & vbTab & " UseRegExp=" & UseRegExp, False
        writeToFile "DEBUG.txt", Now & " 函数SetElementTextByClassName中调用函数getMatchHwndFromWindow", False
    End If
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If
    
    If lngControlsHwnd <> 0 Then
        SendMessage lngControlsHwnd, WM_SETTEXT, 0&, ByVal strNewText
        SetElementTextByClassName = True
    Else
        SetElementTextByClassName = False
    End If
End Function
'根据控件的类名追加显示控件的文字
Public Function AppendElementTextByClassName(ByVal strClassName$, ByVal strNewText$, Optional ByVal intIndex% = 1, Optional ByVal UseRegExp As Boolean = False) As Long
    Dim lngControlsHwnd As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数AppendElementTextByClassName被调用" & vbCrLf & "参数为 strClassName=" & strClassName & vbTab & " strNewText=" & strNewText & vbTab & " UseRegExp=" & UseRegExp, False
        writeToFile "DEBUG.txt", Now & " 函数AppendElementTextByClassName中调用函数getMatchHwndFromWindow", False
    End If
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If
    
    If lngControlsHwnd <> 0 Then
        strNewText = RTrim$(GetElementTextByClassName(strClassName, intIndex, UseRegExp)) & strNewText
        AppendElementTextByClassName = SendMessage(lngControlsHwnd, WM_SETTEXT, 0&, ByVal strNewText)
    Else
        AppendElementTextByClassName = -1
    End If
End Function
'根据控件的类名得到控件的显示文字
Public Function GetElementTextByClassName(ByVal strClassName$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As String
    Dim lngControlsHwnd As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetElementTextByClassName被调用" & vbCrLf & "参数为 strClassName=" & strClassName & vbTab & " intIndex=" & intIndex, False
        writeToFile "DEBUG.txt", Now & " 函数GetElementTextByClassName中调用函数getMatchHwndFromWindow", False
    End If
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If
    
    If lngControlsHwnd <> 0 Then
        GetElementTextByClassName = GetTextByHwnd(lngControlsHwnd)
    Else
        GetElementTextByClassName = ""
    End If
End Function
'根据控件的类名得到控件的显示文字
Public Function GetElementTextByText(ByVal strText$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As String
    Dim lngControlsHwnd As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetElementTextByText被调用" & vbCrLf & "参数为 strText=" & strText & vbTab & " intIndex=" & intIndex, False
        writeToFile "DEBUG.txt", Now & " 函数GetElementTextByText中调用函数GetMatchHwndFromWindow", False
    End If
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strText = replaceReg(strText, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s.*?\s" & strText & "$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strText, intIndex, UseRegExp, False) 'False表示使用文本，默认是true表示类名
    End If
    
    If lngControlsHwnd <> 0 Then
        GetElementTextByText = GetTextByHwnd(lngControlsHwnd)
    Else
        GetElementTextByText = ""
    End If
End Function
'根据控件的类名得到句柄
Public Function GetElementHwndByClassName(ByVal strClassName$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As Long
    Dim lngControlsHwnd As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetElementHwndByClassName被调用" & vbCrLf & "参数为 strClassName=" & strClassName & vbTab & " intIndex=" & intIndex, False
        writeToFile "DEBUG.txt", Now & " 函数GetElementHwndByClassName中调用函数getMatchHwndFromWindow", False
    End If
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        GetElementHwndByClassName = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        GetElementHwndByClassName = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If
End Function
'根据控件的文本内容得到控件的显示文字
Public Function GetElementHwndByText(ByVal strText$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As Long
    Dim lngControlsHwnd As Long
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数GetElementHwndByClassName被调用" & vbCrLf & "参数为 strText=" & strText & vbTab & " intIndex=" & intIndex, False
        writeToFile "DEBUG.txt", Now & " 函数GetElementHwndByClassName中调用函数GetElementHwndByText", False
    End If
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strText = replaceReg(strText, "([\\+-\.()\[\]{}?*\|])", "\$1")
        GetElementHwndByText = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s.*?\s" & strText & "$", intIndex)
    Else
        GetElementHwndByText = GetMatchHwndFromWindow(strText, intIndex, UseRegExp, False) 'False表示使用文本，默认是true表示类名
    End If
End Function
'根据控件的显示文字设置新的显示文字
'函数有问题，调用记事本，发现不能根据文本框内容定位 20130613
Public Function SetElementTextByText(ByVal strText$, ByVal strNewText$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp As Boolean = False) As Boolean
    Dim lngControlsHwnd As Long
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strText = replaceReg(strText, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s.*?\s" & strText & "$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strText, intIndex, UseRegExp, False) 'False表示使用文本，默认是true表示类名
    End If
    If lngControlsHwnd <> 0 Then
        SendMessage lngControlsHwnd, WM_SETTEXT, 0&, ByVal strNewText
        SetElementTextByText = True
    Else
        SetElementTextByText = False
    End If
End Function
'根据控件的类名点击控件
Public Function ClickElementByClassName(ByVal strClassName$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp = False) As Long
    Dim lngControlsHwnd As Long
    '如果使用非正则模式需要替换特殊字符，因为函数内部始终用正则查询的，防止干扰
    If Not UseRegExp Then
        strClassName = replaceReg(strClassName, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s" & strClassName & "\s.*?$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strClassName, intIndex, UseRegExp, True) '最后的true可以省略，默认为类名，否则就是文本
    End If
    If lngControlsHwnd <> 0 Then
        ClickElementByClassName = PostMessage(lngControlsHwnd, BM_CLICK, 0&, ByVal 0)
    End If
End Function
'根据控件的显示文字点击控件
Public Function ClickElementByText(ByVal strText$, Optional ByVal intIndex = 1, Optional ByVal UseRegExp = False) As Long
    Dim lngControlsHwnd As Long
    If Not UseRegExp Then
        strText = replaceReg(strText, "([\\+-\.()\[\]{}?*\|])", "\$1")
        lngControlsHwnd = GetMatchHwndFromWindow("^([-\d]+)\s+[-\d]+\s.*?\s" & strText & "$", intIndex)
    Else
        lngControlsHwnd = GetMatchHwndFromWindow(strText, intIndex, UseRegExp, False) 'False表示使用文本，默认是true表示类名
    End If
    If lngControlsHwnd <> 0 Then
        ClickElementByText = PostMessage(lngControlsHwnd, BM_CLICK, 0&, ByVal 0)
    Else
        ClickElementByText = 0
    End If
End Function
'从匹配结果中得到句柄
Private Function GetMatchHwndFromWindow(ByVal Pattern As String, Optional ByVal intIndex = 1, Optional ByVal UseRegExp = False, Optional ByVal isUseClass = True) As Long
    Dim reg As Object
    Dim matchs As Object, match As Object
    Dim strControlsInfo As String
    Dim intGet As Integer
    
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数getMatchHwndFromWindow被调用" & vbCrLf & "参数为 Pattern=" & Pattern & vbTab & " intIndex=" & intIndex, False
        writeToFile "DEBUG.txt", Now & " 函数getMatchHwndFromWindow中调用函数ControlsInfo，参数为 hWnd_=" & hWnd_, False
    End If
    strControlsInfo = ControlsInfo(hWnd_, DebugMe)
    Set reg = CreateObject("vbscript.regexp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    
    If Not UseRegExp Then
        reg.Pattern = Pattern
        Set matchs = reg.Execute(strControlsInfo)
        If matchs.Count >= intIndex Then
            GetMatchHwndFromWindow = CLng(matchs(intIndex - 1).SubMatches(0))
        End If
    Else '每行结构为="43452836    0   ToolbarWindow32 缩放级别”
        reg.Pattern = "^([-\d]+)\t[-\d]+\t(.*?)\t(.*?)$" '默认检索获得“ToolbarWindow32”
        If Not isUseClass Then reg.Pattern = "^([-\d]+)\t[-\d]+\t.*?\t(.*?)$" '如果不是类名那么就用文本内容进行正则检索“缩放级别”
        Set matchs = reg.Execute(strControlsInfo)
        For Each match In matchs
            If regTest(match.SubMatches(1), Pattern) Then '匹配的
                intGet = intGet + 1
                GetMatchHwndFromWindow = CLng(match.SubMatches(0))
                If intGet = intIndex Then Exit For '如果已经确定当前索引次序就是要处理的，那么退出检索循环
            End If
        Next
    End If
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数getMatchHwndFromWindow执行完毕" & vbCrLf & "匹配数量为=" & matchs.Count, False
    End If
End Function
Public Function CloseWindow() As Long
    CloseWindow = PostMessage(hWnd_, WM_CLOSE, 0, 0)
    If DebugMe Then
        writeToFile "DEBUG.txt", Now & " 函数closeWindow执行完毕" & vbCrLf & "关闭窗口，句柄为=" & hWnd_, False
    End If
End Function
Public Function CloseApp() As Long
    Dim hProcess&
    hProcess = OpenProcess(&H1F0FFF, False, ProcessID())
    CloseApp = TerminateProcess(hProcess, 0)
    CloseHandle hProcess
End Function
Public Sub ClickPoint(x1 As Long, y1 As Long, Optional ByVal PositionMod As enumPositionMode = relative, Optional ByVal clickTimes As Integer = 1, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    If PositionMod = absolute Then
        SetCursorPos x1, y1
    ElseIf PositionMod = relative Then
        SetCursorPos Left + x1, Top + y1
    Else
        MsgBox "函数ClickPoint执行一次。未能识别点击模式" & PositionMod & "，要么是relative(相对位置)要么是absolute(绝对位置)，默认是相对位置。", vbExclamation
        Exit Sub
    End If
    Do While clickTimes > 0
        Wait delayMilliSecondsBefore
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        Wait delayMilliSecondsAfter
        clickTimes = clickTimes - 1
    Loop
End Sub
'点击当前坐标
Public Sub ClickCurrentPoint()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
'移动鼠标到指定点
Public Sub SetCursor(x1 As Long, y1 As Long, Optional ByVal PositionMod As enumPositionMode = absolute, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    If PositionMod = absolute Then
        Wait delayMilliSecondsBefore
        SetCursorPos x1, y1
        Wait delayMilliSecondsAfter
    ElseIf PositionMod = relative Then
        Wait delayMilliSecondsBefore
        SetCursorPos Left + x1, Top + y1
        Wait delayMilliSecondsAfter
    Else
        MsgBox "函数ClickPoint执行一次。未能识别点击模式" & PositionMod & "，要么是relative(相对位置)要么是absolute(绝对位置)，默认是相对位置。", vbExclamation
        Exit Sub
    End If
End Sub
'同SetCursor
Public Sub SetPoint(x1 As Long, y1 As Long, Optional ByVal PositionMod As enumPositionMode = absolute, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    SetCursor x1, y1, PositionMod, delayMilliSecondsBefore, delayMilliSecondsAfter
End Sub
'同SetCursor
Public Sub MoveCursor(x1 As Long, y1 As Long, Optional ByVal PositionMod As enumPositionMode = absolute, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    SetCursor x1, y1, PositionMod, delayMilliSecondsBefore, delayMilliSecondsAfter
End Sub
'同SetCursor
Public Sub MoveCursorTo(x1 As Long, y1 As Long, Optional ByVal PositionMod As enumPositionMode = absolute, Optional ByVal delayMilliSecondsBefore As Long = 0, Optional ByVal delayMilliSecondsAfter As Long = 0)
    SetCursor x1, y1, PositionMod, delayMilliSecondsBefore, delayMilliSecondsAfter
End Sub
'直接获得当前鼠标下的窗口或控件的句柄
Public Function GetCursorPosCurrent() As String
    Dim tPoint As POINTAPI
    GetCursorPos tPoint  '获得当前鼠标位置
    GetCursorPosCurrent = tPoint.x & "," & tPoint.y
End Function
'直接获得当前鼠标下的窗口或控件的句柄
Public Function GetCursorPoint() As String
    GetCursorPoint = GetCursorPosCurrent()
End Function
'以下是几个特效的方法
'特效1：窗体抖动
Public Sub Shake(Optional ByVal shakeShift As enumShift = Both, Optional ByVal shakeRepeats = 30, Optional ByVal shakePads = 10, Optional ByVal shakeInterval = 25)
    Dim i%
    
    If CheckWindow() = 0 Then Exit Sub
    Select Case shakeShift
        Case Horizontal
            For i = 1 To shakeRepeats
                If i Mod 2 = 0 Then
                    Left = Left + shakePads
                Else
                    Left = Left - shakePads
                End If
                Wait shakeInterval
            Next
        Case Vertical
            For i = 1 To shakeRepeats
                If i Mod 2 = 0 Then
                    Top = Top + shakePads
                Else
                    Top = Top - shakePads
                End If
                Wait shakeInterval
            Next
        Case Both
            For i = 1 To shakeRepeats
                If i Mod 2 = 0 Then
                    Left = Left + shakePads
                    Wait 10
                    Top = Top + shakePads
                Else
                    Left = Left - shakePads
                    Wait 10
                    Top = Top - shakePads
                End If
                Wait shakeInterval
            Next
    End Select
End Sub
'特效2：淡入
Public Sub FadeIn(Optional ByVal speed = 10)
    Dim sty As Long, i As Long
    
    If CheckWindow() = 0 Then Exit Sub
    sty = GetWindowLong(hWnd_, GWL_EXSTYLE)
    sty = sty Or WS_EX_LAYERED
    SetWindowLong hWnd_, GWL_EXSTYLE, sty
    
    SetLayeredWindowAttributes hWnd_, 0, 100, LWA_ALPHA
    If speed > 255 Then speed = 255
    If speed < 1 Then speed = 1
    For i = 1 To 255 Step speed
        SetLayeredWindowAttributes hWnd_, 0, i, LWA_ALPHA   '设置窗体透明度
        Wait 50
    Next
    SetLayeredWindowAttributes hWnd_, 0, 255, LWA_ALPHA
End Sub
'特效3：淡出
'可以选择是否关闭窗口或者经常
Public Sub FadeOut(Optional ByVal speed = 10, Optional ByVal isCloseWindow = False, Optional ByVal isCloseApp = False)
    Dim sty As Long, i As Long
    
    If CheckWindow() = 0 Then Exit Sub
    sty = GetWindowLong(hWnd_, GWL_EXSTYLE)
    sty = sty Or WS_EX_LAYERED
    SetWindowLong hWnd_, GWL_EXSTYLE, sty
    
    SetLayeredWindowAttributes hWnd_, 0, 100, LWA_ALPHA
    If speed > 255 Then speed = 255
    If speed < 1 Then speed = 1
    For i = 255 To 1 Step -speed
        SetLayeredWindowAttributes hWnd_, 0, i, LWA_ALPHA   '设置窗体透明度
        Wait 50
    Next
    SetLayeredWindowAttributes hWnd_, 0, 0, LWA_ALPHA
    If isCloseWindow Then CloseWindow
    If isCloseApp Then CloseApp
End Sub
Public Property Get WindowState() As FormWindowStateConstants
    If CheckWindow() = 0 Then Exit Property
    If IsZoomed(hWnd_) Then
        WindowState = vbMaximized
    ElseIf IsIconic(hWnd_) Then
        WindowState = vbMinimized
    Else
        WindowState = vbNormal
    End If
End Property
Public Property Let WindowState(ByVal vNewValue As FormWindowStateConstants)
    If CheckWindow() = 0 Then Exit Property
    Select Case vNewValue
        Case 0
            ShowWindow hWnd_, SW_NORMAL
        Case 1
            ShowWindow hWnd_, SW_SHOWMINIMIZED
        Case 2
            ShowWindow hWnd_, SW_SHOWMAXIMIZED
    End Select
End Property

Public Property Get Visible() As Boolean
    If CheckWindow() = 0 Then Exit Property
    Visible = CBool(IsWindowVisible(hWnd_) = 1)
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
    If vNewValue = True Then
        Call Show
    Else
        Call Hide
    End If
End Property

Public Property Get hDC() As Long
    If CheckWindow() = 0 Then Exit Property
    hDC = GetDC(hWnd_)
End Property

Public Property Let hDC(ByVal vNewValue As Long)
'
End Property

Public Property Get AlphaBlend() As Byte
    If CheckWindow() = 0 Then Exit Property
    GetLayeredWindowAttributes hWnd_, vbNull, AlphaBlend, vbNull
End Property

Public Property Let AlphaBlend(ByVal vNewValue As Byte)
    Dim R As Long, v As Long
    If CheckWindow() = 0 Then Exit Property
    R = GetWindowLong(hWnd_, GWL_EXSTYLE)
    If (R And WS_EX_LAYERED) = 0 Then
        R = R Or WS_EX_LAYERED
        SetWindowLong hWnd_, GWL_EXSTYLE, R
    End If
    R = 0
    GetLayeredWindowAttributes hWnd_, v, vbNull, R
    SetLayeredWindowAttributes hWnd_, v, vNewValue, R Or LWA_ALPHA
End Property

Public Property Get Enabled() As Boolean
    If CheckWindow() = 0 Then Exit Property
    Enabled = CBool(IsWindowEnabled(hWnd_) = 1)
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    If CheckWindow() = 0 Then Exit Property
    EnableWindow hWnd_, vNewValue
End Property

Public Sub Refresh()
    If CheckWindow() = 0 Then Exit Sub
    UpdateWindow hWnd_
End Sub

Public Property Get TransparentColor() As Long
    If CheckWindow() = 0 Then Exit Property
    GetLayeredWindowAttributes hWnd_, TransparentColor, vbNull, vbNull
End Property

Public Property Let TransparentColor(ByVal vNewValue As Long)
    Dim R As Long, v As Byte
    If CheckWindow() = 0 Then Exit Property
    R = GetWindowLong(hWnd_, GWL_EXSTYLE)
    If (R And WS_EX_LAYERED) = 0 Then
        R = R Or WS_EX_LAYERED
        SetWindowLong hWnd_, GWL_EXSTYLE, R
    End If
    R = 0
    GetLayeredWindowAttributes hWnd_, vbNull, v, R
    SetLayeredWindowAttributes hWnd_, vNewValue, v, R Or LWA_COLORKEY
End Property
'设置窗体置顶
Public Sub SetTop(Optional ByVal isSetTop As Boolean = True)
    SetWindowPos hWnd, IIf(isSetTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
End Sub
'此方法未经测试
Public Sub ZOrder()
    If CheckWindow() = 0 Then Exit Sub
    BringWindowToTop hWnd_
End Sub
'得到版本信息
Public Property Get Version() As String
    Version = strVersion
End Property
'根据指定的index选择下拉框中内容
Public Sub SelectComboBoxIndex(ByVal hWnd As Long, ByVal intIndex As Integer)
    SendMessage hWnd, CB_SETCURSEL, intIndex, 0
End Sub
'延时
Public Sub Wait(ByVal MilliSeconds As Long)
    On Error GoTo ErrHandler
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    
    If MilliSeconds = 0 Then Exit Sub
    
    Randomize
    mlTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer" & Format$(Now(), "NNSS") & Int(Rnd * 10000))
    
    If Err.LastDllError <> ERROR_ALREADY_EXISTS Then
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, 0)
    End If
    
    dblDelay = CDbl(MilliSeconds) * 10000#
    
    ft.dwHighDateTime = -CLng(dblDelay / UNITS) - 1
    dblDelayLow = -UNITS * (dblDelay / UNITS - Fix(CStr(dblDelay / UNITS)))
    
    If dblDelayLow < MAX_LONG Then dblDelayLow = UNITS + dblDelayLow
    
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, False)
    
    Do
        lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
        
        DoEvents
'        Exit Sub '**DEBUG**调试时可以加上
    Loop Until lBusy = WAIT_OBJECT_0
    
    CloseHandle mlTimer
    mlTimer = 0
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsWaitableTimer.Wait]" & Err.Description
End Sub
