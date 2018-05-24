Attribute VB_Name = "ModuleTrusteeship"
'托管模块
Option Explicit
'结构体
Private Type WNDCLASS   '窗体结构
        Style As Long
        lpfnwndproc As Long
        cbClsextra As Long
        cbWndExtra2 As Long
        hInstance As Long
        hIcon As Long
        hCursor As Long
        hbrBackground As Long
        lpszMenuName As String
        lpszClassName As String
End Type
Private Type POINTAPI   '坐标结构
        X As Long
        Y As Long
End Type
Private Type Msg        '消息结构
        hWnd As Long
        Message As Long
        wParam As Long
        lParam As Long
        Time As Long
        pt As POINTAPI
End Type
'API函数
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)


'App对象
Public CApp As Class_Application
'屏幕对象
Public CScreen As Class_Screen

'事件托管窗体
Private IForm As Class_Form

'初始化系统对象
Public Function sysInitialize()
        '系统对象类实例化
        Set CApp = New Class_Application
        Set CScreen = New Class_Screen
End Function

'高低位
Public Function GetHiWord(ByVal Value As Long) As Integer
        RtlMoveMemory GetHiWord, ByVal VarPtr(Value) + 2, 2
End Function
Public Function GetLoWord(ByVal Value As Long) As Integer
        RtlMoveMemory GetLoWord, Value, 2
End Function

'托管函数
Public Function Trusteeship(ByRef EventForm As Class_Form) As Boolean
        '类实例化
        Set IForm = EventForm
        Const WinClassName = "MyWinClass"               '定义窗口类名
        
        Dim WC As WNDCLASS '设置窗体参数
        With WC
                .hIcon = 0                                      '窗体图标 使用 LoadIcon(hInstance, ID)   加载RES图标
                .hCursor = 0                                    '窗体光标 使用 LoadCursor(hInstance, ID) 加载RES光标
                .lpszMenuName = vbNullString                    '窗体菜单 使用 LoadMenu(hInstance,ID)    加载RES菜单
                .hInstance = CApp.hInstance                     '实例
                .cbClsextra = 0
                .cbWndExtra2 = 0
                .Style = 0
                .hbrBackground = 16
                .lpszClassName = WinClassName                   '类名
                .lpfnwndproc = GetAddress(AddressOf WinProc)    '消息函数地址
        End With
        '注册窗体类
        If RegisterClass(WC) = 0 Then CApp.ErrDescription = "RegisterClass Faild.": Exit Function
        '获取窗体句柄
        With IForm
                .hWnd = CreateWindowEx(0&, WinClassName, .Caption, .WindowStyle, .Left, .Top, .width, .height, 0, 0, CApp.hInstance, ByVal 0&)
                If .hWnd = 0 Then CApp.ErrDescription = "CreateWindowEx Faild.": Exit Function
                .hDC = GetDC(.hWnd)     '获取窗体GDI句柄
                .Visible = True         '显示窗体
                
                '窗体创建
                Call .ICreate
                
                Dim WinMsg As Msg       '消息结构
                '消息循环
                Do While GetMessage(WinMsg, 0, 0, 0) > 0
                        TranslateMessage WinMsg
                        DispatchMessage WinMsg
                Loop
        End With
        
        '返回值
        Trusteeship = True
End Function

'窗体过程
Private Function WinProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Const WM_CREATE = &H1
        Const WM_COMMAND = &H111
        Const WM_CLOSE = &H10
        Const WM_MOUSEMOVE = &H200
        Const WM_SIZE = &H5
        Const WM_DESTROY = &H2


        Dim bRet As Boolean '取返回值
        With IForm
                Select Case wMsg
                Case WM_COMMAND
                        Call .ICommand(wParam, lParam)
                Case WM_CLOSE
                        Call .IUnload(bRet)
                        If bRet = True Then Exit Function
                        DestroyWindow .hWnd '销毁窗体
                Case WM_MOUSEMOVE
                        Call .IMouseMove(LoWord(lParam), HiWord(lParam))
                Case WM_SIZE
                        Call .IResize
                Case WM_DESTROY
                        PostQuitMessage 0
                Case Else
                        WinProc = DefWindowProc(hWnd, wMsg, wParam, lParam)
                End Select
        End With
End Function

'取地址
Private Function GetAddress(Address) As Long
        GetAddress = Address
End Function

'低字
Private Function LoWord(ByVal DWord As Long) As Integer
        If DWord And &H8000& Then
                LoWord = DWord Or &HFFFF0000
        Else
                LoWord = DWord And &HFFFF&
        End If
End Function

'高字
Private Function HiWord(ByVal DWord As Long) As Integer
        HiWord = (DWord And &HFFFF0000) \ 65536
End Function
