VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Windows�����Զ����������V2.2 ��ʾ"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8550
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "����Ϣ��qqȺ"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   3015
   End
   Begin VB.CommandButton Command7 
      Caption         =   "�� ��"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "����notepad����һϵ�в���"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "���ơ����±�������������������"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�ر����м��±�����Ĵ���"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "���ü����������ض�����"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���Դ��ڼ���"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���ü��������м���"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ر�ֽ�Ƶ���Ϸ����"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "���ݾ��ֱ�ӵ����ť"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3360
      TabIndex        =   7
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�ἤ��Chrome��IE��ɨ�׵ȴ���"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���Զ����š����������������������ݽ��������˳�"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3360
      TabIndex        =   5
      Top             =   2400
      Width           =   4140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���ȴ�����һ��ϵͳ�Դ�����Ϸ"
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

'�ر�ֽ�Ƶ���Ϸ����
Private Sub Command1_Click()
    Command1.Enabled = False
'    If window.GetWindowByTitle("ֽ��", 10).hWnd > 0 Then                                            '����1����ȫƥ�䷽ʽ����֧������
'    If window.GetWindowByTitleEx("��", 10).hWnd > 0 Then                                           '����2��ģ��ƥ�䷽ʽ�����д����ֵľͻ�ȡ
'    ʹ������ƥ�䣬��ϵͳ�Դ��ļ�����Ϸͨɱ
    If window.GetWindowByTitleEx("֩��|ֽ��|�յ�����|ɨ��|���Ĵ�ս|��ά����", 10, , True).hWnd > 0 Then   '����3��ʹ������ƥ�䣬����ǿ��
        window.CloseWindow  '�رմ���
'        window.CloseApp     '�رս��̣�ע������淽��������
    End If
    Command1.Enabled = True
End Sub

Private Sub Command10_Click()
Dim w As New clsWindow
w.GetWindowByPID(Shell("notepad", 1)).Caption = "�ı���� CMB666"
w.SetElementTextByClassName "Edit", "����ı����� CBM666"
End Sub

'���ü��������м���
Private Sub Command2_Click()
    Dim strPID$
    Command2.Enabled = False
    strPID = Shell("calc.exe", 1)
'    window.GetWindowByTitleEx "������", 2        '����1����֪��pid������¿�ʹ�ã���ͬ������Ĵ���ʱ������ȡ
    window.GetWindowByTitleEx "������", 2, strPID '����2�����֪��pid��ô������֤��ȡ���ڣ���ֹ��ȡ
'    MsgBox window.AppPath '�õ�Ӧ�ó���·��
'    MsgBox window.AppName '�õ�Ӧ�ó����ļ���
'    MsgBox window.AppCommandLine '�õ�Ӧ�ó�����������
    
    '����������������
'    window.SetElementTextByClassName "SciCalc", "����2+3�Ľ����"'����1��ͨ��������λ���ô�������
'    window.setElementTextByText "������", "����2+3�Ľ����"      '����2��ͨ���������ֶ�λ���ô�������
    window.Caption = "����2+3�Ľ����"                            '����3��ֱ��ʹ�õ�ǰ�洢�ľ�����ô�������
    window.Wait 500
    window.ClickElementByText "2"
    window.ClickElementByText "+"
    window.ClickElementByText "3"
    window.ClickElementByText "="
    
    'ѭ���������������
    Dim s$, i%
    window.Caption = "����5*3*4/2+1+69=�Ľ����"
    s = "5*3*4/2+1+69="
    For i = 1 To Len(s)
        window.ClickElementByText Mid(s, i, 1)
    Next
    window.Wait 2000
    
    window.Caption = "�����رռ�����"
    window.Wait 2000
    window.CloseWindow
    Command2.Enabled = True
    window.Wait 2000
End Sub
'����ڻ�ý���
Private Sub Command3_Click()
    Command3.Enabled = False
    
    Dim w As New clsWindow
    
    w.Wait 500
    w.GetWindowByTitleEx("Google Chrome", 1).Focus   '����ȸ���������ڻ�ȡ���ͼ���
    w.Wait 500
    If w.GetWindowByTitleEx("ɨ��", 1).hWnd > 0 Then w.Focus '���ɨ����Ϸ���ڻ�ȡ���ͼ��Ĭ�ϵȴ�1��
    w.Wait 500
    If w.GetWindowByTitleEx("Internet Explorer", 1).hWnd > 0 Then  '���IE��������ڻ�ȡ���ͼ���
        w.Focus
        Randomize
        w.Move 800 * Rnd + 100, 300 * Rnd, 800 * Rnd + 500, 800 * Rnd + 100  '���ô���λ�õĲ���
        w.Caption = "�����󶥵����꣺(" & w.Left & "," & w.Top & ") Internet Explorer"   '�õ������󶥵�����Ĳ���
    End If
    
    Command3.Enabled = True
End Sub
'���ô���λ��
Private Sub Command4_Click()
    Dim i&, j&, lngLeft&, lngTop&, strHwnd$, v
    Command4.Enabled = False
    Dim w As New clsWindow

    '�������δ�4��������
    For i = 1 To 4
        w.GetWindowByTitleEx "������", 2, Shell("calc.exe", 1)
        w.Caption = "���Ϊ:" & w.hWnd
        w.Move lngLeft, 300
        lngLeft = lngLeft + w.Width
        strHwnd = strHwnd & w.hWnd & ","
    Next
    
    w.GetWindowByTitleEx "������", 2, Shell("calc.exe", 1)
    strHwnd = strHwnd & w.hWnd & ","
    w.Wait 20
    '�����ƶ�
    For i = 0 To Screen.Height / 15 - w.Height Step 10
        w.Wait 5
        w.Move lngLeft, i
    Next
    '�ٻ�ͷ
    For i = Screen.Height / 15 - w.Height To 300 Step -10
        w.Wait 5
        w.Move lngLeft, i
    Next
    w.Move lngLeft, 300
    
    '�������δ�3�������������Ҵ���Ϊ�����ƶ����ұߵ�Ч��
    For i = 1 To 3
        w.GetWindowByTitleEx "������", 2, Shell("calc.exe", 1)
        w.Caption = "���Ϊ:" & w.hWnd
        For j = 0 To Screen.Width / 15 - w.Width Step 10
            w.Wait 10
            w.Caption = "����(" & j & "." & lngTop & ")"
            w.Move j, lngTop
        Next
        lngTop = lngTop + w.Height
        strHwnd = strHwnd & w.hWnd & ","
    Next

    v = Split(strHwnd, ",")
    
    '������ԣ�����ȫ������һ�飬ע����������һ��
    w.Wait 1000
    For i = 0 To UBound(v) - 1
        w.hWnd = v(i)
        w.Wait 100
        w.Hide
        w.Focus
    Next
    
    '1����𲽹ر����иոմ򿪵Ĵ���
    w.Wait 1000
    For i = UBound(v) - 1 To 0 Step -1
        w.hWnd = v(i)
        w.Wait 100 '��ʱ��Ϊ�˿����𲽹رյĽ��̣���Ȼһ���ӹ��˿�����Ч��
        w.CloseWindow
    Next
    Command4.Enabled = True
End Sub
'�õ�����Ϊ��Դ��������Ĵ��ھ��
Private Sub Command5_Click()
Command5.Enabled = False
'����1 ͨ������һ���Ի�ȡNotepad���д��ھ����
'Dim w As New clsWindow
'Dim sHwndAll$, v, i%
'If w.GetWindowByClassNameEx("Notepad", 3, sHwndAll).hWnd > 0 Then    'xp����Դ�����������CabinetWClass
'    v = Split(sHwndAll, " ") '���з��������Ĵ��ھ���ļ���
'    For i = 0 To UBound(v)
'        w.hWnd = v(i)
'        w.CloseWindow
'    Next
'End If

'����2 ѭ���ر���������ΪNotepad�Ĵ��ڣ�ֱ��û�У�
Dim w As New clsWindow
Do While w.GetWindowByClassName("Notepad", 1).hWnd > 0
    w.CloseWindow
Loop

'����3 ѭ���ر����б��������ؼ��֡� - ���±������ڣ�ֱ��û�У�
'Dim w As New clsWindow
'Do While w.GetWindowByTitleEx(" - ���±�", 1).hWnd > 0
'    w.CloseWindow
'Loop


'����4 ���ݽ������ر�Notepad��Ϊ��ֹ��رգ����Լ�if��ϴ��ڱ�����ƣ���win7�º�����Ч������û��ϸ�о���
'Dim w As New clsWindow
'Dim sHwndAll$, v, i%
'If w.GetWindowByAppNameEx("notepad.exe", 3, sHwndAll).hWnd > 0 Then    'xp����Դ�������������explorer.exe
'    v = Split(sHwndAll, " ") '���з��������Ĵ��ھ���ļ���
'    For i = 0 To UBound(v)
'        w.hWnd = v(i)
'            If w.Caption <> "" And w.Caption <> "Program Manager" And w.Caption <> "M" And w.Caption <> "Default IME" And w.Visible Then '����ɼ��������ݲ�Ϊ�ղ���ϵͳ����Ĵ���
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
'MsgBox w.GetWindowByTitle("Windows ���������").Caption
'MsgBox w.GetWindowByAppName("notepad++.exe").Caption
'w.GetWindowByTitleEx("- ���±�").CloseWindow '�رռ��±�����
'MsgBox w.GetWindowByTitleEx("���±�").GetElementTextByClassName("Edit") 'ȡ�ü��±��е�����

'w.GetWindowByTitleEx("���±�").Hide
'w.GetWindowByTitleEx("���±�").Show

'w.GetWindowByTitleEx("���±�").SetTop

If w.GetWindowByTitleEx("���±�", 2).hWnd > 0 Then
    w.Focus '����Ϊ����
    w.Shake  '����
    w.FadeOut '����
    w.FadeIn '����
Else
    MsgBox "����2����Ҳû�з��ִ��ŵ�notepad���������ذɣ�", vbInformation
End If

'w.GetWindowByPID Shell("C:\Program Files (x86)\Microsoft Office\Office14\powerpnt.exe c:\ʵս���ۼ��ɳ�������.ppt", 1)
'w.WAIT 2000
'w.Focus
'SendKeys "{F5}"
'w.WAIT 1000
'w.ClickElementByClassName "paneClassDC"
End Sub


Private Sub Command7_Click()
'Dim i%
Dim w As New clsWindow
'w.GetWindowByClassName("Notepad").SetElementTextByClassName "Edit", "csdn��ӭ�㣡"

'Do While w.GetWindowByTitleEx("ͳ����").SetElementTextByClassName("ThunderRT6TextBox", "����" & i + 1, i + 1)
'    i = i + 1
'Loop

'w.GetWindowByTitle("������").Shake Vertical, 10, 50, 100
'w.FadeOut , True
'w.FadeIn

'w.GetWindowByTitle("��½").SetElementTextByClassName "WindowsForms10.EDIT.app.0.378734a", "admin", 2
'w.SetElementTextByClassName "WindowsForms10.EDIT.app.0.378734a", "123456", 1
'w.ClickElementByClassName "WindowsForms10.BUTTON.app.0.378734a", 2
'w.ClickElementByText "��½"

'������±����ڴ���
'w.GetWindowByTitleEx "���±�"
'w.Focus
'w.ClickPoint w.Left + 242, w.Top + 36
'w.ClickPoint w.Left + 242, w.Top + 97
'w.GetWindowByTitleEx("���±�").ClickPoint w.Left + 242, w.Top + 36

'�Զ��������±����ļ�
'w.GetWindowByTitleEx("���±�").Focus
'SendKeys "^{o}" 'ʹ�ÿ�ݼ�Ctrl+O��
'w.GetWindowByTitle("��").SetElementTextByClassName "Edit", "c:\1.txt" '�����ļ�·��
'w.ClickElementByText "��(&O)" '����򿪰�ť

'w.GetWindowByTitleEx("���±�").SetElementTextByClassName "Edit", "hello world!" & vbCrLf & Now

'�رռ��±����� ͨ��ģ�ⰴ���İ취
'w.GetWindowByTitleEx("���±�").Focus
'SendKeys "% "
'SendKeys "c"

'Dim hwnds$
'w.GetWindowByClassNameEx "Notepad", , hwnds
'MsgBox hwnds

'w.GetWindowByClassNameEx("Notepad").Focus
'w.ClickPoint w.Left + 40, w.Top + 40

'2018-05-12 19:19:17
'w.GetWindowByTitleEx "��µ����շ���ͳ��ϵͳ", 1
'w.hWnd = w.GetElementHwndByClassName("ThunderRT6FormDC")
'i = 0
'Do
'    i = i + 1
'    If Not w.SetElementTextByClassName("ThunderRT6TextBox", "����" & i, i) Then Exit Do
'Loop
'MsgBox "��д����ˣ�"
'
''���濪ʼ��ȡ
'MsgBox w.GetElementTextByClassName("ThunderRT6TextBox", 2)

'2018-05-14 21:03:11
'w.GetWindowByTitle("�е��Ʊʼ�").ClickPoint w.Left + 205, w.Top + w.Height - 130
'Dim w As New clsWindow
'w.GetWindowByTitle "�е��Ʊʼ�"
'w.ClickPoint w.Left + 205, w.Top + w.Height - 130
'w.GetWindowByClassName("GeneralHtmlDialog").CloseWindow

'w.GetWindowByTitleEx("�Զ����������").ClickPoint 144, 185
'SendMessage Combo1.hWnd, &H14E, "4324", 0
'Combo1.Text = "4324"
'MsgBox Combo1.ListIndex

'w.SelectComboBoxIndex Combo1.hWnd, 0
'Dim w As New clsWindow
'w.GetWindowByTitleEx("��ʱ��").SetElementTextByClassName "Edit", "12345678"
'w.SetElementTextByClassName "Edit", "8888888", 2
'MsgBox w.GetWindowByTitleEx("imm").AppName

'w.WAIT 5000
'Dim i%
'For i = 1 To 3000
'    w.WAIT 50
'    w.ClickPoint2
'Next
'MsgBox "�������"
'SetParent w.GetWindowByPID(Shell("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", 1)).hWnd, Me.hWnd
'MsgBox w.GetWindowByClassName("ConsoleWindowClass").Caption
'MsgBox w.GetCursorPoint
'w.SetPoint 200, 200
'MsgBox w.GetCursorPoint
'SetParent
'
'w.GetWindowByPID(Shell("G:\Program Files\Microsoft Office\Office14\winword.exe", 1), , , DisplayedWindow).Focus
'w.Wait 2000
'w.Focus
'Me.Caption = w.hWnd
'SendKeys "love"
'Shell "G:\Program Files\Microsoft Office\Office14\winword.exe", 1
'w.GetWindowByClassName "NetUIHWND"
'MsgBox w.Caption



'Dim w As New clsWindow
'w.GetWindowByPID(Shell("notepad", 1)).Caption = "�ı���� CMB666"
'w.SetElementTextByClassName "Edit", "����ı����� CBM666"

'�õ����м��±��ľ��
'Dim w As New clsWindow
'Dim s As String
'w.GetWindowByTitleEx "���±�", , s
'MsgBox s
End Sub

'���ü��±�Ȼ��д��һЩ���ݺ󱣴浽c:\test.txt
Private Sub Command8_Click()
    Dim w As New clsWindow
    w.GetWindowByPID(Shell("notepad", 1)).Focus '����pidֱ�ӻ�ȡ���ڲ�����Ϊ��ǰ�����
    w.Caption = "�������±�������" & Now '����Ӧ�ó����������
    w.Shake '�������ڣ�����ͨ���������ڶ��������ٶȡ����ȡ�����
    
    w.SetElementTextByClassName "Edit", "clsWindowv" & w.Version & "���� " & Now()    '�������������
    w.Wait 1000
    w.FadeOut '����
    w.AppendElementTextByClassName "Edit", vbCrLf & "����ȫ��" '�������׷������
    w.FadeIn '����
    w.AppendElementTextByClassName "Edit", vbCrLf & "��ӭ���Է���" '�������׷������
    w.Wait 1000
    
    w.Focus
    w.Wait 500
    SendKeys "^{s}" '���ý���󰴿�ݼ�����
    Dim w2 As New clsWindow
    w2.GetWindowByTitleEx("���Ϊ").SetElementTextByClassName "Edit", "c:\test" & Format(Now, "yyyymmddhhnnss") & ".txt"
    w2.ClickElementByText "����(&S)"
    
    MsgBox w.GetElementTextByClassName("Edit")
End Sub

Private Sub Command9_Click()
    Dim i%
    Dim w As New clsWindow
    For i = 1 To 40
        w.GetWindowByTitle("VB�߼�����ѧϰ����Ⱥ").Focus
        w.ClickPoint w.Left + 35, w.Top + w.Height - 100, absolute
        w.Wait 20
        Clipboard.Clear
        Clipboard.SetText "[" & i & "]���� ����Ϣ�ɳ���clswindow2.1�෢�� " & Now
        SendKeys "^{v}"
        SendKeys "%{s}"
        w.Wait 2
    Next
End Sub

Private Sub Form_Load()
    Dim w As New clsWindow
    Me.Show
    w.Load(Me.hWnd).FadeIn '����
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then 'ʼ�ս�����������Ϊ�����
        Dim w As New clsWindow
        w.hWnd = Me.hWnd
        w.Focus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
