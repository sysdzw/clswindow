Attribute VB_Name = "modWindow"
'=====================================================================================
'��    ������clsWindow.cls�������ģ�飬һЩ�޷��ŵ���ģ���еĴ���������� (modWindow)
'��    �̣�sysdzw ԭ���������������Ҫ��ģ����и����뷢��һ�ݣ���ͬά��
'�������ڣ�2013/05/28
'��    �ͣ�http://blog.csdn.net/sysdzw
'�û��ֲ᣺https://www.kancloud.cn/sysdzw/clswindow/
'Email   ��sysdzw@163.com
'QQ      ��171977759
'��    ����V1.0 ����                                        2012/12/3
'          V1.1 �����е�api�����Լ����ֱ���Ų����ģ��         2013/05/28
'          V1.2 ��EnumChildProc�л�ȡ�ؼ����ֺ����޸���      2013/06/13
'          V1.3 ����ģ�������Ƶ���ģ���еĶ��ƹ�ȥ��          2020/01/19
'=====================================================================================
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public strControlInfo$ '�������������пؼ�����Ϣ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ���api����EnumChildWindows���ʹ�õõ�һ�����������ڵ�����child�ؼ�
'��������EnumChildProc
'��ڲ�����hWnd   long��  ���������һ��ָ������
'����ֵ��long   ����ֱ�ӷ��ص�true�������true���������
'��ע��sysdzw �� 2010-11-13 �ṩ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim Txt2(64000) As Byte
    Dim strClassName As String * 255
    Dim strText As String
    Dim lngCtlId As Long
    Dim strHwnd$, strCtlId$, strClass$, lRet&
    
    EnumChildProc = True
    
    lngCtlId = GetWindowLong(hWnd, (-12))
    lRet = GetClassName(hWnd, strClassName, 255)
    
    SendMessage hWnd, &HD, 64000, Txt2(0)
    strText = Split(StrConv(Split(Txt2, Chr$(0), 2)(0), vbUnicode) & Chr$(0), Chr$(0), 2)(0)
    strText = Replace(strText, vbCrLf, " ") 'ǿ�ƽ��ı��������ݻس��滻�ɿո��Է�ֹӰ�������ȡ
    
    strHwnd$ = CStr(hWnd) & vbTab
    strCtlId$ = CStr(lngCtlId) & vbTab
    strClass$ = Left$(strClassName, lRet) & vbTab
    
    strControlInfo = strControlInfo & strHwnd$ & _
                    strCtlId$ & _
                    strClass$ & _
                    strText & vbCrLf
End Function
