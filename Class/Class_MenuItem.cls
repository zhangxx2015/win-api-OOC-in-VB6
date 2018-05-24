VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Const MF_BYCOMMAND = &H0&
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long

Private Type RectL                                      '坐标结构
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum wFlags
        MF_OWNERDRAW = &H100&                           '自画菜单
        
        MF_HILITE = &H80&                               '高亮显示
        
        MFT_RADIOCHECK = &H200&
        MF_RADIOCHECK = MFT_RADIOCHECK                  '单选菜单
        
        MF_CHECKED = &H8&                               '复选菜单
        MF_UNCHECKED = &H0&
        
        MF_HELP = &H4000&                               '帮助菜单
        
        CMF_DISABLED = &H2&
        CMF_GRAYED = &H1&
        MF_DISABLED = CMF_DISABLED Or CMF_GRAYED        '禁用
        MF_ENABLED = &H0&                               '可用
        
        MF_SEPARATOR = &H800&                           '分隔线
        
        MF_STRING = &H0&                                '字符串
        
        MF_POPUP = &H10&                                '有子菜单
End Enum

Public hMenu As Long            '菜单句柄
Public hItem As Long            '条目句柄
Public hDC As Long              '条目设备
Public Index As Long            '条目索引
Public Alias As String          '别名

Private mFlags As wFlags        '样式
Private RECT As RectL           '坐标

Private mOwnerDraw As Boolean   '自画
Private mHighSelect As Boolean  '高亮
Private mCaption As String      '标题
Private mHaveSub As Boolean     '有子
Private mEnabled As Boolean     '可用
Private mChecked As Boolean     '复选
Private mIsRadio As Boolean     '单选
Private mVisible As Boolean     '可见
Private mIsLine As Boolean      '为线
Private mIsHelp As Boolean      '帮助

'构造函数
Private Sub Class_Initialize()
        '设置初始值
        Flags = MF_STRING
        Enabled = True
        Visible = True
        mHighSelect = True
End Sub

'外部只读
Public Property Get Flags() As wFlags
        Flags = mFlags
End Property
'内部调用
Private Property Let Flags(ByVal Value As wFlags)
        mFlags = Value
        If IsMenu(hMenu) <> 0 And IsMenu(hItem) <> 0 Then
                '修改样式
                If ModifyMenu(hMenu, hItem, MF_BYCOMMAND Or mFlags, hItem, Caption) = 0 Then
                        Exit Property
                End If
        End If
End Property
'-------------------------------------------------------------------
Public Property Get OwnerDraw() As Boolean
        OwnerDraw = mOwnerDraw
End Property
Public Property Let OwnerDraw(ByVal Value As Boolean)
        mOwnerDraw = Value
        If Value = True Then
                Flags = Flags Or MF_OWNERDRAW
        Else
                Flags = Flags And Not MF_OWNERDRAW
        End If
End Property
'-------------------------------------------------------------------
Public Property Get HighSelect() As Boolean
        HighSelect = mHighSelect
End Property
Public Property Let HighSelect(ByVal Value As Boolean)
        mHighSelect = Value
        If Value = True Then
                Flags = Flags Or MF_HILITE
        Else
                Flags = Flags And Not MF_HILITE
        End If
End Property
'-------------------------------------------------------------------
Public Property Get IsRadio() As Boolean
        IsRadio = mIsRadio
End Property
Public Property Let IsRadio(ByVal Value As Boolean)
        mIsRadio = Value
        If Value = True Then
                Flags = Flags Or MF_RADIOCHECK
        Else
                Flags = Flags And Not MF_RADIOCHECK
        End If
End Property
'-------------------------------------------------------------------
Public Property Get IsHelp() As Boolean
        IsHelp = mIsHelp
End Property
Public Property Let IsHelp(ByVal Value As Boolean)
        mIsHelp = Value
        If Value = True Then
                Flags = Flags Or MF_HELP
        Else
                Flags = Flags And Not MF_HELP
        End If
End Property
'-------------------------------------------------------------------
Public Property Get IsLine() As Boolean
        IsLine = mIsLine
End Property
Public Property Let IsLine(ByVal Value As Boolean)
        mIsLine = Value
        If Value = True Then
                Flags = Flags And Not MF_STRING
                Flags = Flags Or MF_SEPARATOR
        Else
                Flags = Flags And Not MF_SEPARATOR
                Flags = Flags Or MF_STRING
        End If
End Property
'-------------------------------------------------------------------
Public Property Get Visible() As Boolean
        Visible = mVisible
End Property
Public Property Let Visible(ByVal Value As Boolean)
        mVisible = Value
        Flags = Flags
End Property
'-------------------------------------------------------------------
Public Property Get Checked() As Boolean
        Checked = mChecked
End Property
Public Property Let Checked(ByVal Value As Boolean)
        mChecked = Value
        If Value = True Then
                Flags = Flags And Not MF_UNCHECKED
                Flags = Flags Or MF_CHECKED
        Else
                Flags = Flags And Not MF_CHECKED
                Flags = Flags Or MF_UNCHECKED
        End If
End Property
'-------------------------------------------------------------------
Public Property Get Enabled() As Boolean
        Enabled = mEnabled
End Property
Public Property Let Enabled(ByVal Value As Boolean)
        mEnabled = Value
        If Value = True Then
                Flags = Flags And Not MF_DISABLED
                Flags = Flags Or MF_ENABLED
        Else
                Flags = Flags And Not MF_ENABLED
                Flags = Flags Or MF_DISABLED
        End If
End Property
'-------------------------------------------------------------------
Public Property Get HaveSub() As Boolean
        HaveSub = mHaveSub
End Property
Public Property Let HaveSub(ByVal Value As Boolean)
        mHaveSub = Value
        If Value = True Then
                Flags = Flags Or MF_POPUP
        Else
                Flags = Flags And Not MF_POPUP
        End If
End Property
'-------------------------------------------------------------------
Public Property Get Caption() As String
        Caption = mCaption
End Property
Public Property Let Caption(ByVal Value As String)
        mCaption = Value
        Flags = Flags And Not MF_SEPARATOR
        Flags = Flags Or MF_STRING
End Property

'-------------------------------------------------------------------
Public Property Get RectR() As Long
        RectR = RECT.Right
End Property
Public Property Let RectR(ByVal Value As Long)
        RECT.Right = Value
End Property
'-------------------------------------------------------------------
Public Property Get RectL() As Long
        RectL = RECT.Left
End Property
Public Property Let RectL(ByVal Value As Long)
        RECT.Left = Value
End Property
'-------------------------------------------------------------------
Public Property Get RectT() As Long
        RectT = RECT.Top
End Property
Public Property Let RectT(ByVal Value As Long)
        RECT.Top = Value
End Property
'-------------------------------------------------------------------
Public Property Get RectB() As Long
        RectB = RECT.Bottom
End Property
Public Property Let RectB(ByVal Value As Long)
        RECT.Bottom = Value
End Property
'-------------------------------------------------------------------
