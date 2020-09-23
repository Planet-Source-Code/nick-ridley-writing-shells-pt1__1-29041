Attribute VB_Name = "modWindows"


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40


Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Boolean

Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

Public Function WindowPos(frm As Object, setting As Integer)
'Change positions of windows, make top most etc...


Dim i As Integer
Select Case setting
Case 1
i = HWND_TOPMOST
Case 2
i = HWND_TOP
Case 3
i = HWND_NOTOPMOST
Case 4
i = HWND_BOTTOM
End Select

SetWindowPos frm.hwnd, i, frm.Left / 15, _
frm.Top / 15, frm.Width / 15, _
frm.Height / 15, SWP_SHOWWINDOW Or SWP_NOACTIVATE

End Function

Public Sub SetFGWindow(ByVal hwnd As Long, Show As Boolean)
If Show Then
If IsIconic(hwnd) Then
ShowWindow hwnd, SW_RESTORE
Else
BringWindowToTop hwnd
End If
Else
ShowWindow hwnd, SW_MINIMIZE
End If
End Sub

