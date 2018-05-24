VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_MenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'菜单条目数组
Private Menus() As Class_MenuItem
'上标
Public Count As Long

'构造函数
Private Sub Class_Initialize()
        '初始化数组
        ReDim Menus(0&)
End Sub

'析构函数
Private Sub Class_Terminate()
        Dim i As Long
        For i = 0& To Count
                '释放类
                Set Menus(i) = Nothing
        Next i
        '释放数组
        Erase Menus
End Sub

'添加一个菜单项
Public Function AddMenu() As Class_MenuItem
        '上标累加
        Count = Count + 1&
        '申请内存
        ReDim Preserve Menus(Count)
        '类实例化
        Set Menus(Count) = New Class_MenuItem
        '保存索引
        Menus(Count).Index = Count
        '传出索引
        Set AddMenu = Menus(Count)
End Function

'访问接口
Public Function GetItemByIndex(ByVal Index As Long) As Class_MenuItem
        Set GetItemByIndex = Menus(Index)
End Function

Public Function GetItemByhItem(ByVal hItem As Long) As Class_MenuItem
        Dim i As Long
        For i = 1& To Count
                Set GetItemByhItem = GetItemByIndex(i)
                If GetItemByhItem.hItem = hItem Then
                        Exit Function
                End If
        Next i
        Set GetItemByhItem = Nothing
End Function

Public Function GetItemByAlias(ByVal Alias As String) As Class_MenuItem
        Dim i As Long
        For i = 1& To Count
                Set GetItemByAlias = GetItemByIndex(i)
                If GetItemByAlias.Alias = Alias Then
                        Exit Function
                End If
        Next i
        Set GetItemByAlias = Nothing
End Function
