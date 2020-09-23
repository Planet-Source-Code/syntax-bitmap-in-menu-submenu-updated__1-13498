VERSION 5.00
Begin VB.Form frmMenubmp 
   Caption         =   "BitMap in Menu Updated on Dec 11,2000"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu m1 
         Caption         =   "Item1"
      End
      Begin VB.Menu m2 
         Caption         =   "Item2"
      End
      Begin VB.Menu m3 
         Caption         =   "Item3"
      End
   End
End
Attribute VB_Name = "frmMenubmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code places a bitmap in your Menu / SubMenu
' (c)
Private Sub Form_Load()

Dim hMenu As Long
Dim hSubMenu As Long
Dim lMenuID As Long
Dim hBitmap As Long
Dim hImage As Long

' Get handle of top menu
hMenu = GetMenu(Me.hwnd)

' Check for a valid menu handle
If IsMenu(hMenu) = 0 Then
    MsgBox ("Menu handle invalid"), vbInformation
    Exit Sub
End If

' Get handle of submenu 0 (File Menu)
hSubMenu = GetSubMenu(hMenu, 0)

' Check for a valid submenu handle
If IsMenu(hSubMenu) = 0 Then
    MsgBox ("Submenu handle invalid"), vbInformation
    Exit Sub
End If

' Need menu item ID (item 1 is the second item)
lMenuID = GetMenuItemID(hSubMenu, 1)

' Load bitmap to be used in the menu
hImage = LoadImage(0, App.Path & "\new.bmp", IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)

' Stick it in the menu (the bitmap sticks in place of the second item)
' The number 1 defines which item should be replaced by the bitmap
'
'
' The line below puts the bitmap in the SubMenu - controlled by hSubMenu
' Just activate the line and you'll see what I mean
'
ModifyMenu hSubMenu, 1, MF_BITMAP Or MF_BYPOSITION, lMenuID, hImage

' The line below puts the bitmap in the Menu - controlled by hMenu
' Just activate the line and you'll see what I mean
'
'ModifyMenu hMenu, 0, MF_BITMAP Or MF_BYPOSITION, lMenuID, hImage
End Sub
