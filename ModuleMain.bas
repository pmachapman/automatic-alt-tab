Attribute VB_Name = "ModuleMain"
' Require variable declaration
Option Explicit

' Win32 API declarations
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_ALT = &H12
Public Const VK_TAB = &H9

