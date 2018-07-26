VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWaitableTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsWaitableTimer
'描述:延时类
'以下代码放入类模块中，clsWaitableTimer
'使用范例:
'Private Sub Command1_Click()
'    Dim objTimer As New clsWaitableTimer
'    WAIT 5000
'    Set objTimer = Nothing
'End Sub
Option Explicit

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&
Private Const INFINITE = &HFFFF
Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Const UNITS = 4294967296#
Private Const MAX_LONG = -2147483648#

Private Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Private Declare Function OpenWaitableTimer Lib "kernel32" Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Private Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long

Private mlTimer As Long

Private Sub Class_Terminate()
    On Error Resume Next
    If mlTimer <> 0 Then CloseHandle mlTimer
End Sub

Public Sub WAIT(MilliSeconds As Long)
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
