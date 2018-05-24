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

'�˵���Ŀ����
Private Menus() As Class_MenuItem
'�ϱ�
Public Count As Long

'���캯��
Private Sub Class_Initialize()
        '��ʼ������
        ReDim Menus(0&)
End Sub

'��������
Private Sub Class_Terminate()
        Dim i As Long
        For i = 0& To Count
                '�ͷ���
                Set Menus(i) = Nothing
        Next i
        '�ͷ�����
        Erase Menus
End Sub

'����һ���˵���
Public Function AddMenu() As Class_MenuItem
        '�ϱ��ۼ�
        Count = Count + 1&
        '�����ڴ�
        ReDim Preserve Menus(Count)
        '��ʵ����
        Set Menus(Count) = New Class_MenuItem
        '��������
        Menus(Count).Index = Count
        '��������
        Set AddMenu = Menus(Count)
End Function

'���ʽӿ�
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