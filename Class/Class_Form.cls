VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'句柄是否有效
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
'可见
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
'标题
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'坐标
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'样式
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'----------------------------------------------------------------------------------------------------------------
'窗体样式
Public Enum eWindowStyle
        None = &H6000000
        Fixed_Single = &H6C80000
        Sizable = &H6CF0000
        Fixed_Dialog = &H6C80080
        Fixed_ToolWindow = &H6C80000
        Sizable_ToolWindow = &H6CC0000
End Enum

Private mVisible        As Boolean      '可见
Private mCaption        As String       '标题
Private mHeight         As Long         '高度
Private mWidth          As Long         '宽度
Private mTop            As Long         '顶边
Private mLeft           As Long         '左边
Private mMostTop        As Boolean      '层次
Private mCenter         As Boolean      '居中
Private mWindowStyle    As eWindowStyle '样式

Public hWnd             As Long         '窗体句柄
Public hDC              As Long         '设备句柄

'事件
Public Event Create()
Public Event Resize()
Public Event MouseMove(ByVal X As Integer, ByVal Y As Integer)
Public Event Unload(ByRef Cancel As Boolean)
Public Event Command(ByVal wParam As Long, ByVal lParam As Long)



'构造函数
Private Sub Class_Initialize()
        '缺省值
        mWidth = 200
        mHeight = 200
        mWindowStyle = Sizable
        mCaption = "Windows GUI App"
End Sub

'析构函数
Private Sub Class_Terminate()
        '释放类
        Set CScreen = Nothing
End Sub



'=====事件接口===================================================================================
Public Sub ICreate()
        RaiseEvent Create
End Sub

Public Sub IMouseMove(ByVal X As Integer, ByVal Y As Integer)
        RaiseEvent MouseMove(X, Y)
End Sub

Public Function IResize()
        RaiseEvent Resize
End Function

Public Function IUnload(ByRef Cancel As Boolean)
        RaiseEvent Unload(Cancel)
End Function

Public Function ICommand(ByVal wParam As Long, ByVal lParam As Long)
        RaiseEvent Command(wParam, lParam)
End Function


'=====是否可见===================================================================================
Public Property Get Visible() As Boolean
        Visible = mVisible
End Property
Public Property Let Visible(ByVal State As Boolean)
        mVisible = State
        If IsWindow(hWnd) <> 0 Then
                Const SW_NORMAL As Long = 1
                Const SW_HIDE = 0
                
                '显示/隐藏 窗体
                If State = True Then
                        ShowWindow hWnd, SW_NORMAL
                Else
                        ShowWindow hWnd, SW_HIDE
                End If
                '更新窗体
                UpdateWindow hWnd
        End If
End Property

'=====标题文字===================================================================================
Public Property Get Caption() As String
        Caption = mCaption
End Property

Public Property Let Caption(ByVal Text As String)
        mCaption = Text
        If IsWindow(hWnd) <> 0 Then SetWindowText hWnd, Text
End Property


'=====窗体高度===================================================================================
Public Property Get height() As Long
        height = mHeight
End Property

Public Property Let height(ByVal Value As Long)
        mHeight = Value
        Call SetWinPos
End Property

'=====窗体宽度===================================================================================
Public Property Get width() As Long
        width = mWidth
End Property

Public Property Let width(ByVal Value As Long)
        mWidth = Value
        Call SetWinPos
End Property

'=====窗体顶边===================================================================================
Public Property Get Top() As Long
        Top = mTop
End Property

Public Property Let Top(ByVal Value As Long)
        mTop = Value
        Call SetWinPos
End Property

'=====窗体左边===================================================================================
Public Property Get Left() As Long
        Left = mLeft
End Property

Public Property Let Left(ByVal Value As Long)
        mLeft = Value
        Call SetWinPos
End Property

'=====窗体置顶===================================================================================
Public Property Get MostTop() As Boolean
        MostTop = mMostTop
End Property

Public Property Let MostTop(ByVal Value As Boolean)
        mMostTop = Value
        Call SetWinPos
End Property

'=====窗体居中===================================================================================
Public Property Get Center() As Boolean
        Center = mCenter
End Property

Public Property Let Center(ByVal Value As Boolean)
        mCenter = Value
        mLeft = (CScreen.width - mWidth) \ 2
        mTop = (CScreen.height - mHeight) \ 2
        Call SetWinPos
End Property

'=====工具函数===================================================================================
Private Sub SetWinPos()
        Const HWND_TOPMOST As Long = -1
        Const SWP_SHOWWINDOW = &H40
        If IsWindow(hWnd) <> 0 Then
                SetWindowPos hWnd, IIf(mMostTop = True, HWND_TOPMOST, 0), mLeft, mTop, mWidth, mHeight, SWP_SHOWWINDOW
                '更新窗体
                UpdateWindow hWnd
        End If
End Sub

'=====窗体样式===================================================================================
Public Property Get WindowStyle() As eWindowStyle
        WindowStyle = mWindowStyle
End Property

Public Property Let WindowStyle(ByVal Value As eWindowStyle)
        mWindowStyle = Value
        
        Const GWL_STYLE = (-16)
        If IsWindow(hWnd) <> 0 Then
                SetWindowLong hWnd, GWL_STYLE, mWindowStyle
                '更新窗体
                UpdateWindow hWnd
        End If
End Property



