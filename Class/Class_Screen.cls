VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class_Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'ÆÁÄ»¿í¶È
Public Function width() As Long
        Const SM_CXSCREEN = 0
        width = GetSystemMetrics(SM_CXSCREEN)
End Function

'ÆÁÄ»¸ß¶È
Public Function height() As Long
        Const SM_CYSCREEN = 1
        height = GetSystemMetrics(SM_CYSCREEN)
End Function

