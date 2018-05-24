VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type Msg
        hWnd As Long
        Message As Long
        wParam As Long
        lParam As Long
        Time As Long
        pt As POINTAPI
End Type
Private Const PM_NOREMOVE = 0
Private Const PM_REMOVE = 1
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'APP属性
Public CommandLine    As String '命令行
Public hInstance        As Long '实例
Public ErrDescription As String '错误描述


'卸载窗体
Public Sub Unload(ByRef CForm As Class_Form)
        Const WM_CLOSE = &H10
        PostMessage CForm.hWnd, WM_CLOSE, 0&, 0&
End Sub



'DoEvents增强版
Public Sub DoEventsEx()
        Dim wMsg As Msg
        While PeekMessage(wMsg, 0, 0, 0, PM_REMOVE)
                Call TranslateMessage(wMsg)
                Call DispatchMessage(wMsg)
        Wend
End Sub

'修正GetCommandLine函数让VB6崩溃的问题
Private Function GetCommLineVB6() As String
        Dim RetStr As Long, SLen As Long
        Dim Buffer As String
        RetStr = GetCommandLine
        SLen = lstrlen(RetStr)
        If SLen > 0 Then
                Dim CommLineVB6 As String
                CommLineVB6 = Space$(SLen)
                CopyMemory ByVal CommLineVB6, ByVal RetStr, SLen
                GetCommLineVB6 = CommLineVB6
        End If
End Function

'构造函数
Private Sub Class_Initialize()
        hInstance = GetModuleHandle(vbNullString)       '获取模块句柄
        CommandLine = GetCommLineVB6                    '获取命令行参数
End Sub
