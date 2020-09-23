VERSION 5.00
Begin VB.Form frmTasks 
   BackColor       =   &H00E0A878&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tasklisting"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTaskUpdate 
      Interval        =   250
      Left            =   120
      Top             =   1140
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0A878&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ListBox lstApps 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lstNames 
      Height          =   255
      Left            =   1380
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstHwnd 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lstHwndNames 
      Height          =   255
      Left            =   1380
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblTask 
      BackColor       =   &H00E0A878&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF4040&
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple tasklisting example

'Sample for tutorial:
'Writing Shells Pt1

Private Sub Form_Load()
Me.Icon = Nothing
WindowPos Me, 1
picIcon(0).Height = 16 * Screen.TwipsPerPixelY
picIcon(0).Width = 16 * Screen.TwipsPerPixelX
fEnumWindows Me.lstApps

picIcon(0).Top = picIcon(0).Top - picIcon(0).Height
lblTask(0).Top = lblTask(0).Top - lblTask(0).Height
End Sub

Private Sub tmrTaskUpdate_Timer()
ListApps
End Sub

Public Function ListApps()

On Error Resume Next

Dim i As Long, c As Long
Dim d As Long
Dim e As Boolean

Me.lstApps.Clear
Me.lstNames.Clear

fEnumWindows Me.lstApps

DoEvents

i = lstApps.ListCount - 1
c = lstApps.ListCount

Do Until i < 0

d = 0
e = False

'check if window allready has an entry
Do Until d = lstHwnd.ListCount
If lstHwnd.List(d) = lstApps.List(i) Then e = True: Exit Do
d = d + 1
Loop

'Add it if its not there

If e = False Then
Load lblTask(lblTask.ubound + 1)
Load picIcon(picIcon.ubound + 1)
lblTask(lblTask.ubound).Caption = lstNames.List(i)
lblTask(lblTask.ubound).Top = lblTask(lblTask.ubound - 1).Top + picIcon(picIcon.ubound).Height + 30
lblTask(lblTask.ubound).ZOrder 0
lblTask(lblTask.ubound).Tag = lstApps.List(i)
lblTask(lblTask.ubound).Visible = True
picIcon(picIcon.ubound).Top = picIcon(picIcon.ubound - 1).Top + picIcon(picIcon.ubound).Height + 30
picIcon(picIcon.ubound).ZOrder 0
picIcon(picIcon.ubound).AutoRedraw = True
picIcon(picIcon.ubound).Visible = True
Call DrawIcon(picIcon(picIcon.ubound).hdc, lstApps.List(i), 0, 0)
lstHwnd.AddItem lstApps.List(i)
lstHwndNames.AddItem lstNames.List(i)
End If

'Change the buttons text if the one on the form has changed

If e = True Then
c = 0
Do Until lblTask(c).Caption = lstHwndNames.List(d)
c = c + 1
Loop

lstHwndNames.List(d) = lstNames.List(i)
lblTask(c).Caption = lstHwndNames.List(d)

End If

i = i - 1

Loop


i = 0
d = lstApps.ListCount

'Now check top see if windows that we have on the list still exits

Do Until i >= lstHwnd.ListCount

c = 0
e = False

Do Until c = lstApps.ListCount

If lstHwnd.List(i) = lstApps.List(c) Then e = True: Exit Do
c = c + 1

Loop

If e = False And c <> 0 Then
c = 0

Do Until lblTask(c).Caption = lstHwndNames.List(i)
c = c + 1
If c > lblTask.ubound Then GoTo kill
Loop

RemTask c
DoEvents

lstHwnd.RemoveItem i
lstHwndNames.RemoveItem i

End If

kill:

i = i + 1

Loop

End Function

Public Function RemTask(i As Long)
Dim c As Long
c = i
Do Until c = lblTask.ubound
lblTask(c).Caption = lblTask(c + 1).Caption
lblTask(c).Tag = lblTask(c + 1).Tag
picIcon(c).Picture = Nothing
Call DrawIcon(picIcon(c).hdc, lblTask(c + 1).Tag, 0, 0)
c = c + 1
Loop
Unload lblTask(lblTask.ubound)
Unload picIcon(picIcon.ubound)
End Function

Public Sub DrawIcon(hdc As Long, hwnd As Long, x As Integer, y As Integer)
ico = GetIcon(hwnd)
DrawIconEx hdc, x, y, ico, 16, 16, 0, 0, DI_NORMAL
End Sub

Public Function GetIcon(hwnd As Long) As Long
Call SendMessageTimeout(hwnd, WM_GETICON, 0, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICONSM)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_GETICON, 1, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICON)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Do Until i = lblTask.ubound + 1
lblTask(i).FontUnderline = False
i = i + 1
Loop
End Sub

Private Sub imgBG_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Do Until i = lblTask.ubound + 1
lblTask(i).FontUnderline = False
i = i + 1
Loop
End Sub

Private Sub lblTask_Click(Index As Integer)
SetFGWindow lblTask(Index).Tag, True
End Sub

Private Sub lblTask_DblClick(Index As Integer)
SetFGWindow lblTask(Index).Tag, False
End Sub

Private Sub lblTask_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Do Until i = lblTask.ubound + 1
lblTask(i).FontUnderline = False
i = i + 1
Loop
lblTask(Index).FontUnderline = True
End Sub

