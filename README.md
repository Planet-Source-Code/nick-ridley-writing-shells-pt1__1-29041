<div align="center">

## Writing Shells \- pt1


</div>

### Description

The first of my tutorials on writting Shells (like explorer) in VB
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-11-20 18:09:10
**By**             |[Nick Ridley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nick-ridley.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Writting\_S3629811202001\.zip](https://github.com/Planet-Source-Code/nick-ridley-writing-shells-pt1__1-29041/archive/master.zip)





### Source Code

<h1>Writing A Shell</h1>
<p>(Part1) By Nick Ridley</p>
<p>Date: 20/11/2001</p>
<p><strong>Contents:</strong></p>
<p>1- Introduction<br>
2- Getting started<br>
3- Taskbar buttons<br>
4- Next Issue</p>
<h3>1 - Introduction</h3>
<p>I have stated to write these tutorials to try and get some more people into writing
shells in VB. I know that I am  not the best shell writer but I do know how to get
started in making one and these tutorials are meant to give newbies that boost of info
they need so they will start.</p>
<p>Nick Ridley</p>
<p> </p>
<h3>2- Getting started</h3>
<p>Before you even start to make your shell decide on some things first:</p>
<p>1- Will it be free or commercial?<br>
2- Will it be open source?<br>
3- What colour scheme will you use?<br>
4- What versions of window will it be compatible with</p>
<p>Decide on all of these things and then write them down on a bit of paper. Below start a
brainstorm of the word SHELL and come up with as much info. Now finalise what you want in
light of this info and decide on a name. Write down all this on a bit of paper and stick
it to your monitor or something. Get some paper and a pen and keep this handy at all times
to write down ideas. You may also need a calculator to do any sums and stuff.</p>
<p>Now you have most the info you will need, now we can start.</p>
<p><strong>You must now:</strong></p>
<p>Create your project<br>
Do your splash screen<br>
Design the place were the task buttons will be</p>
<p> </p>
<h3>3- The task buttons</h3>
<p>Now we will move on to task listing:</p>
<p>I have re written some parts of a .bas file I got of PSC (I think this is made up of
Softshell and RepShell) and you must now add this to your project:</p>
<p>NOTE: I did not fully write this, this is a rewritten version of what was in softshell
and repshell, although I have re-written some of it</p>
<p>[BEGIN TaskListing.bas]</p>
<p><em><font color="#008040">'I hope this bit encourages you newbies to<br>
'start new shells (use this to make a taskbar)</font><br>
<br>
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long,
ByVal lParam As Long) As Long<br>
Public Declare Function GetForegroundWindow Lib "user32" () As Long<br>
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long<br>
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd
As Long) As Long<br>
Public Declare Function GetWindowLong Lib "user32" Alias
"GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long<br>
Public Declare Function GetWindowText Lib "user32" Alias
"GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As
Long) As Long<br>
<br>
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA"
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long<br>
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA"
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long<br>
<br>
Public Const LB_ADDSTRING = &H180<br>
Public Const LB_FINDSTRINGEXACT = &H1A2<br>
Public Const LB_ERR = (-1)<br>
<br>
Public Const GW_OWNER = 4<br>
Public Const GWL_EXSTYLE = (-20)<br>
<br>
Public Const WS_EX_APPWINDOW = &H40000<br>
Public Const WS_EX_TOOLWINDOW = &H80<br>
<br>
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Boolean<br>
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long<br>
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As
Long<br>
<br>
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft
As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As
Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As
Long) As Long<br>
Public Const DI_NORMAL = &H3<br>
<br>
Public Declare Function GetClassLong Lib "user32" Alias
"GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Integer) As Long<br>
<br>
Public Const WM_GETICON = &H7F<br>
Public Const GCL_HICON = (-14)<br>
Public Const GCL_HICONSM = (-34)<br>
Public Const WM_QUERYDRAGICON = &H37<br>
<br>
Public Declare Function SendMessageTimeout Lib "user32" Alias
"SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As
Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As
Long) As Long<br>
<br>
<font color="#008040">'This is used to get icons from windows >>>></font><br>
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As
Long, ByVal y As Long, ByVal hIcon As Long) As Long<br>
<br>
Public Function fEnumWindows(lst As ListBox) As Long<br>
With lst<br>
.Clear<br>
frmTasks.lstNames.Clear<font color="#008040"> ' replace this as neccessary</font><br>
Call EnumWindows(AddressOf fEnumWindowsCallBack, .hwnd)<br>
fEnumWindows = .ListCount<br>
End With<br>
End Function<br>
<br>
Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Long) As Long<br>
<br>
Dim lExStyle As Long, bHasNoOwner As Boolean, sAdd As String, sCaption As String<br>
<br>
If IsWindowVisible(hwnd) Then<br>
bHasNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)<br>
lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)<br>
<br>
If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bHasNoOwner) Or _<br>
((lExStyle And WS_EX_APPWINDOW) And Not bHasNoOwner) Then<br>
sAdd = hwnd: sCaption = GetCaption(hwnd)<br>
Call SendMessage(lParam, LB_ADDSTRING, 0, ByVal sAdd)<br>
Call SendMessage(frmTasks.lstNames.hwnd, LB_ADDSTRING, 0, ByVal sCaption)<font
color="#008040"> ' replace this as neccessary</font><br>
End If<br>
End If<br>
<br>
fEnumWindowsCallBack = True<br>
End Function<br>
<br>
Public Function GetCaption(hwnd As Long) As String<br>
Dim mCaption As String, lReturn As Long<br>
mCaption = Space(255)<br>
lReturn = GetWindowText(hwnd, mCaption, 255)<br>
GetCaption = Left(mCaption, lReturn)<br>
End Function<br>
</em></p>
<p>[END TaskListing.bas]</p>
<p>If you are not going to download the sample project you will need to write your own
function to use this. In my project i have included a function to do this.</p>
<p>Basically the functions do this:</p>
<p><em>fEnumWindows</em></p>
<p><em>lst</em> = the list box were the window hWnd's will be held</p>
<p>You will also need to change a few lines (these are marked) to suit your project, You
do not need to directly call the rest of the functions.</p>
<p>You may also find this useful to set FG windows and make your taskbar stay on top:</p>
<p>[BEGIN modWindows.bas]</p>
<p><em>Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long,
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal
cy As Long, ByVal wFlags As Long) As Long<br>
<br>
Public Const HWND_BOTTOM = 1<br>
Public Const HWND_NOTOPMOST = -2<br>
Public Const HWND_TOP = 0<br>
Public Const HWND_TOPMOST = -1<br>
<br>
Public Const SWP_NOACTIVATE = &H10<br>
Public Const SWP_SHOWWINDOW = &H40<br>
<br>
<br>
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As
Long) As Long<br>
Public Const SW_HIDE = 0<br>
Public Const SW_NORMAL = 1<br>
Public Const SW_SHOWMINIMIZED = 2<br>
Public Const SW_SHOWMAXIMIZED = 3<br>
Public Const SW_SHOWNOACTIVATE = 4<br>
Public Const SW_SHOW = 5<br>
Public Const SW_MINIMIZE = 6<br>
Public Const SW_SHOWMINNOACTIVE = 7<br>
Public Const SW_SHOWNA = 8<br>
Public Const SW_RESTORE = 9<br>
Public Const SW_SHOWDEFAULT = 10<br>
<br>
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As
Boolean<br>
<br>
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long<br>
<br>
Public Function WindowPos(frm As Object, setting As Integer)<br>
<font color="#008040">'Change positions of windows, make top most etc...</font><br>
<br>
<br>
Dim i As Integer<br>
Select Case setting<br>
Case 1<br>
i = HWND_TOPMOST<br>
Case 2<br>
i = HWND_TOP<br>
Case 3<br>
i = HWND_NOTOPMOST<br>
Case 4<br>
i = HWND_BOTTOM<br>
End Select<br>
<br>
SetWindowPos frm.hwnd, i, frm.Left / 15, _<br>
frm.Top / 15, frm.Width / 15, _<br>
frm.Height / 15, SWP_SHOWWINDOW Or SWP_NOACTIVATE<br>
<br>
End Function<br>
<br>
Public Sub SetFGWindow(ByVal hwnd As Long, Show As Boolean)<br>
If Show Then<br>
If IsIconic(hwnd) Then<br>
ShowWindow hwnd, SW_RESTORE<br>
Else<br>
BringWindowToTop hwnd<br>
End If<br>
Else<br>
ShowWindow hwnd, SW_MINIMIZE<br>
End If<br>
End Sub</em></p>
<p>[END modWindows.bas]</p>
<p>Now you can either use this info to build your own project or use mine.</p>
<h1>I HIGHLY RECOMEND YOU DOWNLOAD MY SAMPLE</h1>
<h3>This DOES NOT cover everything</h3>
<h3>4- Next Issue:</h3>
<p>In the next issue I plan to describe how to make a start menu (hopefully in more detail
than this) describing how to get icons from files and how to make menus appear and
disappear. And in further issues i will describe how to make a system tray for example.</p>
<p> </p>
<p>I hope you find this useful and <strong>PLEASE VOTE</strong> and <strong>LEAVE COMMENTS</strong>.
What annoys me is when people read your code and use it but dont vote so please show your
appreciation and even if you vote poor every vote counts.</p>
<h3>Thanx for reading</h3>
<h4>Nick Ridley</h4>
<p><a href="http://www.spyderhackers.co.uk">http://www.spyderhackers.co.uk</a></p>
<p><a href="http://www.spyderhackers.com">http://www.spyderhackers.com</a></p>
<p><a href="mailto:nick@spyderhackers.com">nick@spyderhackers.com</a></p>

