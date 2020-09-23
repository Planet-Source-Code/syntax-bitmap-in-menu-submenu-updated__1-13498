Attribute VB_Name = "modGlobals"
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewitem As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Public Const LR_LOADFROMFILE = &H10
Public Const MF_BITMAP = &H4
Public Const MF_BYPOSITION = &H400

