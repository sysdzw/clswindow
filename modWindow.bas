'===========================================================================
'描    述：是clsWindow.cls类的调用模块，一些常用函数和api在此声明 (modWindow)
'编    程：sysdzw 原创开发，如果有需要对模块进行更新请发我一份，共同维护
'发布日期：2013/05/28
'博    客：http://blog.163.com/sysdzw
'          http://blog.csdn.net/sysdzw
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                        2012/12/3
'          V1.1 将类中的api函数以及部分变量挪到此模块       2013/05/28
'          V1.2 将EnumChildProc中获取控件文字函数修改了     2013/06/13
'===========================================================================
Option Explicit

'常量定义
Public Const SW_MINIMIZE = 6
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOWNORMAL = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Const WM_SETTEXT = &HC
Public Const BM_CLICK = &HF5
Public Const WM_SETFOCUS = &H7
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD


Private Const SC_MOVE = &HF010&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&

Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up

Public Const GWL_EXSTYLE = -20

Public Const WS_VISIBLE = &H10000000

Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const WS_EX_LAYERED = &H80000

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const CB_SETCURSEL = &H14E
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

'结构体定义
Public Type rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Enum enumShift
    Horizontal
    Vertical
    Both
End Enum

Public Enum enumPositionMode
    absolute
    relative
End Enum

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As Any, ByVal nSize As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As rect) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByRef crKey As Long, ByRef bAlpha As Byte, ByRef dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function QueryFullProcessImageName Lib "Kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As Any, ByVal nSize As Long) As Long
Public Declare Function GetProcessImageFileName Lib "Kernel32.dll" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lpFilename As Any, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'自定义变量声明
Public strControlInfo$ '保存容器内所有控件的信息
Public strWindowInfo$ '保存所有窗口的信息，格式为 句柄 文本内容
Private strWindowKeyWord$ '要参与的过滤的窗口的关键字，如果不需要过滤就留空
Dim strTmp$, isWholeEx As Boolean

Public objTimer As New clsWaitableTimer

'得到所有窗口的信息
Private Function GetWindowInfo(Optional ByVal strKeyWord = "", Optional ByVal isWhole As Boolean = False) As String
    strWindowInfo = ""
    isWholeEx = isWhole
    strWindowKeyWord = strKeyWord
    Call EnumWindows(AddressOf EnumWindowProc, 0)
    If Right(strWindowInfo, 2) = vbCrLf Then strWindowInfo = Left(strWindowInfo, Len(strWindowInfo) - 2)
    GetWindowInfo = strWindowInfo
End Function

Private Function EnumWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    If (GetWindowLong(hWnd, GWL_STYLE) And &HCF0000) = &HCF0000 And (IsWindowVisible(hWnd) = 1) Then
        strTmp = GetWinText(hWnd)
        If InStr(strTmp, strWindowKeyWord) > 0 Then '如果在关键字内就显示
            strWindowInfo = strWindowInfo & CStr(hWnd) & " " & strTmp & vbCrLf
        End If
    End If
    EnumWindowProc = 1
End Function

Public Function GetWinText(ByVal hWnd As Long) As String
    GetWinText = String(1024, Chr(0))
    GetWindowText hWnd, GetWinText, Len(GetWinText)
    GetWinText = Left$(GetWinText, InStr(GetWinText, Chr(0)) - 1)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：得到所有控件的信息，是按次序获得的，可用于编写脚本的参考和程序设置值时使用。此函数需要和EnumChildProc一起使用
'函数名：ControlsInfo
'入口参数：hWnd   long型  容器句柄，一般指窗体句柄。可缺省，缺省为最新获取到的窗体的句柄，也可以指定一个句柄
'返回值：string   保存了容器内所有控件的信息，包含“句柄、ID、类名、显示文字”
'备注：sysdzw 于 2010-11-13 提供
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ControlsInfo(ByVal lngMainHwnd As Long, Optional isDebug = False) As String
    Dim strHwnd$, strCtlId$, strClass$
    Dim Rtn&, hWin&
    Dim Txt(64000) As Byte
    Dim strWindowClass As String * 255
    Dim strWindowTitle$
    
    GetClassName lngMainHwnd, strWindowClass, 255  '获得窗口类
    SendMessage lngMainHwnd, &HD, 64000, Txt(0) '获得窗口标题(也可使用 API 函数:GetWindowText,但效果不佳)
    strWindowTitle = StrConv(Txt, vbUnicode)
    strWindowTitle = Replace(strWindowTitle, Chr(0), "")
    strWindowClass = Replace(strWindowClass, Chr(0), "")
    strControlInfo = lngMainHwnd & vbTab & "0" & vbTab & Replace(strWindowClass, " ", "") & vbTab & strWindowTitle & vbCrLf
    
    Rtn = EnumChildWindows(lngMainHwnd, AddressOf EnumChildProc, 0&)
    ControlsInfo = strControlInfo
    If isDebug Then writeToFile "controls.txt", strControlInfo
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：和api函数EnumChildWindows结合使用得到一个窗体容器内的所有child控件
'函数名：EnumChildProc
'入口参数：hWnd   long型  容器句柄，一般指窗体句柄
'返回值：long   这里直接返回的true，如果是true则继续调用
'备注：sysdzw 于 2010-11-13 提供
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim strClassName As String * 255
    Dim strCaption As String
    Dim lngCtlId As Long
    Dim strHwnd$, strCtlId$, strClass$
    Dim Txt2(64000) As Byte
    
    EnumChildProc = True
    
    lngCtlId = GetWindowLong(hWnd, GWL_ID)
    Call GetClassName(hWnd, strClassName, 255)
    SendMessage hWnd, &HD, 64000, Txt2(0)
    strCaption = StrConv(Txt2, vbUnicode)
    strCaption = Left$(strCaption, InStr(strCaption, Chr$(0)) - 1)
    strCaption = Replace(strCaption, vbCrLf, " ") '强制将文本框中内容回车替换成空格，以防止影响正则获取
    
    strHwnd$ = CStr(hWnd)
    strHwnd$ = strHwnd$ & vbTab
    strCtlId$ = CStr(lngCtlId)
    strCtlId$ = strCtlId$ & vbTab
    strClass$ = Left$(strClassName, InStr(strClassName, Chr$(0)) - 1)
    strClass$ = strClass$ & vbTab
    
    strControlInfo = strControlInfo & strHwnd$ & _
                    strCtlId$ & _
                    strClass$ & _
                    strCaption & vbCrLf
End Function

Public Function replaceReg(ByVal str1$, ByVal strPattern$, ByVal strNew$) As String
    Dim strData As String
    Dim reg As Object

    strData = str1 ' "1)2{}?*3.7(7"

    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = strPattern ' "([\\+-\.()\[\]{}?*])"
    replaceReg = reg.Replace(strData, strNew) '"\$1"
End Function

Public Function regTest(ByVal str1$, ByVal strPattern$) As Boolean
    Dim strData As String
    Dim reg As Object

    strData = str1 ' "1)2{}?*3.7(7"

    Set reg = CreateObject("vbscript.regExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.MultiLine = True
    reg.Pattern = strPattern ' "([\\+-\.()\[\]{}?*])"
    regTest = reg.test(strData)  '"\$1"
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：根据所给文件名和内容直接写文件
'函数名：writeToFile
'入口参数(如下)：
'  strFileName 所给的文件名；
'  strContent 要输入到上述文件的字符串
'  isCover 是否覆盖该文件，默认为覆盖
'返回值：True或False，成功则返回前者，否则返回后者
'备注：sysdzw 于 2007-5-2 提供
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function writeToFile(ByVal strFileName$, ByVal strContent$, Optional isCover As Boolean = True) As Boolean
    On Error GoTo Err1
    Dim fileHandl%
    fileHandl = FreeFile
    If isCover Then
        Open strFileName For Output As #fileHandl
    Else
        Open strFileName For Append As #fileHandl
    End If
    Print #fileHandl, strContent
    Close #fileHandl
    writeToFile = True
    Exit Function
Err1:
    writeToFile = False
End Function
