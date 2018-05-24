VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_MenuManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'菜单收集器
Private Items As New Class_MenuItems

Private hWnd As Long '对象句柄
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private hFormMenu As Long
Private hSystemMenu As Long
Private hPopupMenu As Long

Private Const MF_BYCOMMAND = &H0&
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long

'Private Type RECT                                       '坐标结构
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type
'Private Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
'Private mRECT As RECT

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Long) As Long

'自定义事件
Public Event Click(ByRef MenuItem As Class_MenuItem)

'子类化
Private WithEvents CSubClass As Class_SubClass
Attribute CSubClass.VB_VarHelpID = -1

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'事件触发方法
Public Sub EventMethod(ByRef Menu As Class_MenuItem)
        With Menu
                Const WM_COMMAND = &H111
                PostMessage hWnd, WM_COMMAND, .hItem, 0&
        End With
End Sub


'构造函数
Private Sub Class_Initialize()
        '类实例化
        Set CSubClass = New Class_SubClass
End Sub

'析构函数
Private Sub Class_Terminate()
        '释放类
        Set CSubClass = Nothing
End Sub

'初始化类
Public Function Initial(ByVal DesthWnd As Long) As Boolean
        hWnd = DesthWnd
        If IsWindow(hWnd) = 0 Then Exit Function
        hFormMenu = CreateMenu
        If hFormMenu = 0 Then Exit Function
        hSystemMenu = GetSystemMenu(hWnd, False)
        If hSystemMenu = 0 Then Exit Function
        hPopupMenu = CreatePopupMenu
        If hPopupMenu = 0 Then Exit Function
        Initial = True
End Function

Public Function AddStd(ByVal ParentTag As Variant, Optional ByVal Caption As String, Optional ByVal Alias As String, Optional BindhItem As Long) As Class_MenuItem
        Set AddStd = Items.AddMenu()
        With AddStd
                If VarType(ParentTag) = vbString Then
                        .hMenu = Items.GetItemByAlias(CStr(ParentTag)).hItem
                Else
                        .hMenu = Items.GetItemByIndex(CLng(ParentTag)).hItem
                End If
                .hItem = IIf(BindhItem = 0, CreateMenu, BindhItem)
                .Alias = Alias
                .Caption = Caption
                AppendMenu .hMenu, .Flags, .hItem, .Caption
'                GetMenuItemRect hWnd, .hMenu, .hItem, mRECT
'                .RectR = mRECT.Right
'                .RectB = mRECT.Bottom
'                .RectL = mRECT.Left
'                .RectT = mRECT.Top
        End With
End Function

Public Function AddForm(Optional ByVal Caption As String, Optional ByVal Alias As String, Optional BindhItem As Long) As Class_MenuItem
        Set AddForm = Items.AddMenu()
        With AddForm
                .hMenu = hFormMenu
                .hItem = IIf(BindhItem = 0, CreateMenu, BindhItem)
                .HaveSub = True
                .Alias = Alias
                .Caption = Caption
                AppendMenu .hMenu, .Flags, .hItem, .Caption
'                GetMenuItemRect hWnd, .hMenu, .hItem, mRECT
'                .RectR = mRECT.Right
'                .RectB = mRECT.Bottom
'                .RectL = mRECT.Left
'                .RectT = mRECT.Top

        End With
End Function

Public Function AddPop(Optional ByVal Caption As String, Optional ByVal Alias As String, Optional BindhItem As Long) As Class_MenuItem
        Set AddPop = Items.AddMenu()
        With AddPop
                .hMenu = hPopupMenu
                .hItem = IIf(BindhItem = 0, CreateMenu, BindhItem)
                .Alias = Alias
                .Caption = Caption
                AppendMenu .hMenu, .Flags, .hItem, .Caption
'                GetMenuItemRect hWnd, .hMenu, .hItem, mRECT
'                .RectR = mRECT.Right
'                .RectB = mRECT.Bottom
'                .RectL = mRECT.Left
'                .RectT = mRECT.Top
        End With
End Function

Public Function AddSys(Optional ByVal Caption As String, Optional ByVal Alias As String, Optional BindhItem As Long) As Class_MenuItem
        Set AddSys = Items.AddMenu()
        With AddSys
                .hMenu = hSystemMenu
                .hItem = IIf(BindhItem = 0, CreateMenu, BindhItem)
                .Alias = Alias
                .Caption = Caption
                AppendMenu .hMenu, .Flags, .hItem, .Caption
'                GetMenuItemRect hWnd, .hMenu, .hItem, mRECT
'                .RectR = mRECT.Right
'                .RectB = mRECT.Bottom
'                .RectL = mRECT.Left
'                .RectT = mRECT.Top
        End With
End Function

Public Property Let Enabled(ByVal State As Boolean)
        If State = True Then
                Call CSubClass.Hook(hWnd)
        Else
                Call CSubClass.UnHook
        End If
End Property


'显示弹出式菜单
Public Sub ShowPopup(Optional ByVal UsePos As Boolean = False, Optional ByVal X As Long, Optional ByVal Y As Long)
        Dim CurPos As POINTAPI
        Call GetCursorPos(CurPos)
        Const TPM_LEFTALIGN = &H0&
        TrackPopupMenu hPopupMenu, TPM_LEFTALIGN, IIf(UsePos = True, X, CurPos.X), IIf(UsePos = True, Y, CurPos.Y), 0, hWnd, ByVal 0&
End Sub

'拦截菜单消息
Private Sub CSubClass_NewMessage(ByVal hWnd As Long, FunctionPtr As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
        Const WM_PAINT = &HF
        
        Const WM_COMMAND = &H111
        Const WM_SYSCOMMAND = &H112
        
        Const WM_DRAWITEM = &H2B                     ' 消息: 绘制菜单项
        Const WM_EXITMENULOOP = &H212                ' 消息: 退出菜单消息循环
        Const WM_MEASUREITEM = &H2C                  ' 消息: 处理菜单高度和宽度
        Const WM_MENUSELECT = &H11F                  ' 消息: 选择菜单项
        

        Select Case uMsg
        Case WM_PAINT
                SetMenu hWnd, hFormMenu
        Case WM_COMMAND, WM_SYSCOMMAND              ' 单击菜单项
                
                Dim eMenu As Class_MenuItem
                Set eMenu = Items.GetItemByhItem(wParam)
                If Not eMenu Is Nothing Then RaiseEvent Click(eMenu)
                Set eMenu = Nothing
                
                'If MyItemInfo(wParam).itemType = MIT_CHECKBOX Then
                '        If MyItemInfo(wParam).itemState = MIS_CHECKED Then
                '            MyItemInfo(wParam).itemState = MIS_UNCHECKED
                '        Else
                '            MyItemInfo(wParam).itemState = MIS_CHECKED
                '        End If
                'End If
                'MenuItemSelected wParam
        Case WM_EXITMENULOOP                                            ' 退出菜单消息循环(保留)
        
        Case WM_MEASUREITEM                                             ' 处理菜单项高度和宽度
                
            'MeasureItem hWnd, lParam
        Case WM_MENUSELECT                                              ' 选择菜单项
            'Dim itemID As Long
            'itemID = GetMenuItemID(lParam, wParam And &HFF)
            'If itemID <> -1 Then
            '    MenuItemSelecting itemID
            'End If
        Case WM_DRAWITEM                                                ' 绘制菜单项
            'DrawItem lParam
        End Select
End Sub


'继承访问接口
Public Function GetItemByIndex(ByVal Index As Long) As Class_MenuItem
        Set GetItemByIndex = Items.GetItemByIndex(Index)
End Function

Public Function GetItemByhItem(ByVal hItem As Long) As Class_MenuItem
        Set GetItemByhItem = Items.GetItemByhItem(hItem)
End Function

Public Function GetItemByAlias(ByVal Alias As String) As Class_MenuItem
        Set GetItemByAlias = Items.GetItemByAlias(Alias)
End Function
