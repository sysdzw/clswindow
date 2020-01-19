Attribute VB_Name = "modWindow"
'===========================================================================
'��    ������clsWindow.cls��ĵ���ģ�飬һЩ���ú�����api�ڴ����� (modWindow)
'��    �̣�sysdzw ԭ���������������Ҫ��ģ����и����뷢��һ�ݣ���ͬά��
'�������ڣ�2013/05/28
'��    �ͣ�http://blog.163.com/sysdzw
'          http://blog.csdn.net/sysdzw
'Email   ��sysdzw@163.com
'QQ      ��171977759
'��    ����V1.0 ����                                        2012/12/3
'          V1.1 �����е�api�����Լ����ֱ���Ų����ģ��       2013/05/28
'          V1.2 ��EnumChildProc�л�ȡ�ؼ����ֺ����޸���     2013/06/13
'===========================================================================
Option Explicit

'��������
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

Public Const WAIT_ABANDONED& = &H80&
Public Const WAIT_ABANDONED_0& = &H80&
Public Const WAIT_FAILED& = -1&
Public Const WAIT_IO_COMPLETION& = &HC0&
Public Const WAIT_OBJECT_0& = 0
Public Const WAIT_OBJECT_1& = 1
Public Const WAIT_TIMEOUT& = &H102&
Public Const INFINITE = &HFFFF
Public Const QS_HOTKEY& = &H80
Public Const QS_KEY& = &H1
Public Const QS_MOUSEBUTTON& = &H4
Public Const QS_MOUSEMOVE& = &H2
Public Const QS_PAINT& = &H20
Public Const QS_POSTMESSAGE& = &H8
Public Const QS_SENDMESSAGE& = &H40
Public Const QS_TIMER& = &H10
Public Const ERROR_ALREADY_EXISTS = 183&
Public Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Public Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Public Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Public Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Public Const UNITS = 4294967296#
Public Const MAX_LONG = -2147483648#

'�ṹ�嶨��
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

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
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
'���˴����Ƿ�ɼ�
Public Enum enumWindowVisible
    HiddenWindow
    DisplayedWindow
    AllWindow
End Enum


Public Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Public Declare Function OpenWaitableTimer Lib "kernel32" Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Public Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long

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

'�Զ����������
Public strControlInfo$ '�������������пؼ�����Ϣ
Public strWindowInfo$ '�������д��ڵ���Ϣ����ʽΪ ��� �ı�����
Private strWindowKeyWord$ 'Ҫ����Ĺ��˵Ĵ��ڵĹؼ��֣��������Ҫ���˾�����
Dim strTmp$, isWholeEx As Boolean

Public mlTimer As Long

'�õ����д��ڵ���Ϣ
Private Function GetWindowInfo(Optional ByVal strKeyWord = "", Optional ByVal isWhole As Boolean = False) As String
    strWindowInfo = ""
    isWholeEx = isWhole
    strWindowKeyWord = strKeyWord
    Call EnumWindows(AddressOf EnumWindowProc, 0)
    If Right$(strWindowInfo, 2) = vbCrLf Then strWindowInfo = Left$(strWindowInfo, Len(strWindowInfo) - 2)
    GetWindowInfo = strWindowInfo
End Function

Private Function EnumWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    If (GetWindowLong(hWnd, GWL_STYLE) And &HCF0000) = &HCF0000 And (IsWindowVisible(hWnd) = 1) Then
        strTmp = GetTextByHwnd(hWnd)
        If InStr(strTmp, strWindowKeyWord) > 0 Then '����ڹؼ����ھ���ʾ
            strWindowInfo = strWindowInfo & CStr(hWnd) & " " & strTmp & vbCrLf
        End If
    End If
    EnumWindowProc = 1
End Function

'���ݾ����ô�������
Public Function GetTextByHwnd(ByVal hWnd As Long) As String
    Dim Txt2(64000) As Byte
    SendMessage hWnd, WM_GETTEXT, 64000, Txt2(0)
    GetTextByHwnd = Split(StrConv(Split(Txt2, Chr$(0), 2)(0), vbUnicode) & Chr$(0), Chr$(0), 2)(0)
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ��õ����пؼ�����Ϣ���ǰ������õģ������ڱ�д�ű��Ĳο��ͳ�������ֵʱʹ�á��˺�����Ҫ��EnumChildProcһ��ʹ��
'��������ControlsInfo
'��ڲ�����hWnd   long��  ���������һ��ָ����������ȱʡ��ȱʡΪ���»�ȡ���Ĵ���ľ����Ҳ����ָ��һ�����
'����ֵ��string   ���������������пؼ�����Ϣ�������������ID����������ʾ���֡�
'��ע��sysdzw �� 2010-11-13 �ṩ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ControlsInfo(ByVal lngMainHwnd As Long, Optional isDebug As Boolean = False) As String
    Dim Rtn&
    Dim strWindowClass As String * 255
    Dim strWindowTitle$
    
    GetClassName lngMainHwnd, strWindowClass, 255  '��ô�����
    strWindowTitle = GetTextByHwnd(lngMainHwnd)
    strWindowClass = Replace(strWindowClass, Chr$(0), "")
    strControlInfo = lngMainHwnd & vbTab & "0" & vbTab & Replace(strWindowClass, " ", "") & vbTab & strWindowTitle & vbCrLf
    
    Rtn = EnumChildWindows(lngMainHwnd, AddressOf EnumChildProc, 0&)
    ControlsInfo = strControlInfo
    If isDebug Then writeToFile "controls.txt", strControlInfo
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ���api����EnumChildWindows���ʹ�õõ�һ�����������ڵ�����child�ؼ�
'��������EnumChildProc
'��ڲ�����hWnd   long��  ���������һ��ָ������
'����ֵ��long   ����ֱ�ӷ��ص�true�������true���������
'��ע��sysdzw �� 2010-11-13 �ṩ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim strClassName As String * 255
    Dim strCaption As String
    Dim lngCtlId As Long
    Dim strHwnd$, strCtlId$, strClass$
    
    EnumChildProc = True
    
    lngCtlId = GetWindowLong(hWnd, GWL_ID)
    Call GetClassName(hWnd, strClassName, 255)
    
    strCaption = GetTextByHwnd(hWnd)
    strCaption = Replace(strCaption, vbCrLf, " ") 'ǿ�ƽ��ı��������ݻس��滻�ɿո��Է�ֹӰ�������ȡ
    
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
'���ܣ����������ļ���������ֱ��д�ļ�
'��������writeToFile
'��ڲ���(����)��
'  strFileName �������ļ�����
'  strContent Ҫ���뵽�����ļ����ַ���
'  isCover �Ƿ񸲸Ǹ��ļ���Ĭ��Ϊ����
'����ֵ��True��False���ɹ��򷵻�ǰ�ߣ����򷵻غ���
'��ע��sysdzw �� 2007-5-2 �ṩ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function writeToFile(ByVal strFileName$, ByVal strContent$, Optional isCover As Boolean = True) As Boolean
    On Error GoTo err1
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
err1:
    writeToFile = False
End Function

