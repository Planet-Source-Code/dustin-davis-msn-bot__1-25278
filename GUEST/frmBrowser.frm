VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmbrowser 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M$N"
   ClientHeight    =   7125
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11550
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enter Chat"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "CHAT URL HERE"
      Top             =   6600
      Width           =   4455
   End
   Begin VB.Timer Timer5 
      Left            =   3240
      Top             =   5520
   End
   Begin VB.Timer Timer4 
      Left            =   2880
      Top             =   5520
   End
   Begin VB.Timer Timer3 
      Left            =   2520
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   5520
   End
   Begin VB.Timer TimerNoWhisper 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1800
      Top             =   5520
   End
   Begin VB.Timer timFlood 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   5520
   End
   Begin VB.Timer timBan 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   1440
      Top             =   5520
   End
   Begin VB.Timer timKick 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   5520
   End
   Begin VB.Timer timScroll 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   5520
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6735
      Left            =   -240
      TabIndex        =   1
      Top             =   -240
      Width           =   12255
      ExtentX         =   21616
      ExtentY         =   11880
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "HIT ALT+D TO STOP KICKING"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   0
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Menu sfdf 
      Caption         =   "Kick'em"
      Begin VB.Menu Kicking 
         Caption         =   "&Kick All"
         Shortcut        =   ^K
      End
      Begin VB.Menu kjam 
         Caption         =   "Kick all with &message"
      End
      Begin VB.Menu KFT 
         Caption         =   "Kick First Two"
      End
   End
   Begin VB.Menu nopt 
      Caption         =   "Ban'em"
      Begin VB.Menu Banning 
         Caption         =   "&Ban All (24 Hour)"
         Shortcut        =   ^B
      End
      Begin VB.Menu rtrt 
         Caption         =   "&Ban All (1 Hour)"
      End
      Begin VB.Menu rttr 
         Caption         =   "&Ban All (15 Min)"
      End
      Begin VB.Menu balm 
         Caption         =   "Ban all with M&essage"
      End
      Begin VB.Menu bfocus 
         Caption         =   "Ban all set fo&cus"
      End
      Begin VB.Menu BFT 
         Caption         =   "Ban First Two"
      End
   End
   Begin VB.Menu fdsfs 
      Caption         =   "Host'em"
      Begin VB.Menu Host 
         Caption         =   "&Host All"
         Shortcut        =   ^H
      End
      Begin VB.Menu HFT 
         Caption         =   "Host First Two"
      End
   End
   Begin VB.Menu dsd 
      Caption         =   "Dehost'em"
      Begin VB.Menu dfd 
         Caption         =   "&Dehost All"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu EP 
      Caption         =   "&Enter Pass"
   End
   Begin VB.Menu Flood 
      Caption         =   "&Flood Message"
   End
   Begin VB.Menu Refresh 
      Caption         =   "&Refresh Chatroom"
   End
   Begin VB.Menu Disable 
      Caption         =   "&Disable "
   End
   Begin VB.Menu dfdf 
      Caption         =   "&Msn Home"
   End
   Begin VB.Menu ferr 
      Caption         =   "&Exit Chat"
   End
End
Attribute VB_Name = "frmbrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5


Dim mLCO As Integer
Dim Window As Long
Dim L As Long
Dim Str As String
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSEEVENTF_LEFTDOWN = &H2      ' left button down
Private Const MOUSEEVENTF_LEFTUP = &H4        ' left button up
Private Const MOUSEEVENTF_ABSOLUTE = &H8000   ' absolute move
Private Const MOUSEEVENTF_MOVE = &H1          ' move
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim pt As POINTAPI
Dim temp2 As POINTAPI
Public Scroll As Single
Public ScrollCause As String
Public temp As Single
Public HostingNum As Single
Public FloodMsg As String
Public ChatPageMsg As String
Public key As String
Public Mess As String
Public MessBan As String
Public NewSet As Integer
Public NewSetB As Integer
Dim xp, yp

Private Sub balm_Click()
MessBan = InputBox("What is the message to ban with?", "Message", "(*)LaMeR(*)")
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timBan.Enabled = False
timScroll.Enabled = True
ScrollCause = "WhileBanningMess"
End Sub

Private Sub Banning_Click()
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timBan.Enabled = False
timScroll.Enabled = True
ScrollCause = "WhileBanning"
End Sub

Private Sub bfocus_Click()
NewSetB = InputBox("Enter the number of people to skip banning kicking at!", "Set Ban", "1")
NewSetB = NewSetB * 15
NewSetB = NewSetB + 132
HostingNum = 300
Scroll = NewSet
timFlood.Enabled = False
timKick.Enabled = False
timScroll.Enabled = True
timBan.Enabled = False
ScrollCause = "WhileSetBan"
End Sub

Private Sub BFT_Click()
timFlood.Enabled = False
timScroll.Enabled = False
timKick.Enabled = False
timBan.Enabled = True
temp = 0
End Sub


Private Sub Command1_Click()
WebBrowser1.Navigate Text1.Text
End Sub

Private Sub dfd_Click()
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timBan.Enabled = False
timScroll.Enabled = True
Timer2.Enabled = False
ScrollCause = "WhileDehosting"
End Sub

Private Sub dfdf_Click()
ChatPageMsg = InputBox("What Chat Page?", "Chat Page")
If ChatPageMsg = 1 Then WebBrowser1.Navigate "http://chat.msn.com/find.msnw?cat=TN&page=1"
If ChatPageMsg = 2 Then WebBrowser1.Navigate "http://chat.msn.com/find.msnw?cat=TN&page=2"
If ChatPageMsg = 3 Then WebBrowser1.Navigate "http://chat.msn.com/find.msnw?cat=TN&page=3"
If ChatPageMsg = 4 Then WebBrowser1.Navigate "http://chat.msn.com/find.msnw?cat=TN&page=4"
If ChatPageMsg = 5 Then WebBrowser1.Navigate "http://chat.msn.com/find.msnw?cat=TN&page=5"
If ChatPageMsg = 6 Then WebBrowser1.Navigate "http://chat.msn.com/find.msnw?cat=TN&page=6"
If ChatPageMsg = 7 Then WebBrowser1.Navigate "http://chat.msn.com/find.msnw?cat=TN&page=7"
If ChatPageMsg = 8 Then WebBrowser1.Navigate "http://chat.msn.com/find.msnw?cat=TN&page=8"
End Sub

Private Sub dfdf34234534_Click()


End Sub

Private Sub Disable_Click()
timKick.Enabled = False
timFlood.Enabled = False
timBan.Enabled = False
timScroll.Enabled = False
Scroll = 118
End Sub



Private Sub EP_Click()
Dim temp As String
Dim a As Long
Open "C:\regloc.dat" For Input As #1
Line Input #1, temp
Close #1
Shell temp + " /e C:\temp.reg HKEY_CURRENT_USER\SOFTWARE\Microsoft\msnchat\4.0"
Open "C:\temp.reg" For Input As #1
Do While a < 4
a = a + 1
Line Input #1, temp
Loop
SendKeys "/pass " + Mid(temp, 14, Len(temp) - 14) + Chr$(13) + Chr$(13)
Close #1
End Sub



Private Sub fdf_Click()
HostingNum = 150
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timScroll.Enabled = True
timBan.Enabled = False
Timer2.Enabled = False
ScrollCause = "WhileDehosting"
End Sub

Private Sub ferr_Click()
SendKeys "/part + {Enter}"
End Sub

Private Sub Flood_Click()
On Error Resume Next
timKick.Enabled = False
timFlood.Enabled = False
timBan.Enabled = False
timScroll.Enabled = False
Scroll = 118
FloodMsg = InputBox("Flood Message", "Flood")
If Len(FloodMsg) < 1 Then Exit Sub
timFlood.Enabled = True
End Sub



Private Sub focus_Click()

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    WebBrowser1.Navigate "C:\Chatroom.html"
    Scroll = 118
    temp = 0
End Sub

Private Sub HFT_Click()
HostingNum = 150
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timScroll.Enabled = True
timBan.Enabled = False
ScrollCause = "WhileHosting"
End Sub

Private Sub Host_Click()
HostingNum = 300
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timScroll.Enabled = True
timBan.Enabled = False
ScrollCause = "WhileHosting"
End Sub




Private Sub KFT_Click()
timFlood.Enabled = False
timScroll.Enabled = False
timBan.Enabled = False
temp = 0
timKick.Enabled = True
End Sub

Private Sub Kicking_Click()
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timBan.Enabled = False
timScroll.Enabled = True
ScrollCause = "WhileKicking"
End Sub

Private Sub kjam_Click()
Mess = InputBox("What is the message to kick with?", "messsage", "Message")
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timBan.Enabled = False
timScroll.Enabled = True
ScrollCause = "WhileKickingMess"
End Sub

Private Sub kkk_Click()
SetCursorPos 618, 118
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KD"
End Sub

Private Sub Refresh_Click()
timKick.Enabled = False
timFlood.Enabled = False
timBan.Enabled = False
timScroll.Enabled = False
Scroll = 118
WebBrowser1.Refresh
End Sub

Private Sub rtrt_Click()
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timBan.Enabled = False
timScroll.Enabled = True
Timer4.Enabled = False
ScrollCause = "WhileBanning1"
End Sub

Private Sub rttr_Click()
Scroll = 118
timFlood.Enabled = False
timKick.Enabled = False
timBan.Enabled = False
timScroll.Enabled = True
Timer3.Enabled = False
ScrollCause = "WhileBanning15"
End Sub

Private Sub timBan_Timer()
If temp = 0 Then
pt.X = 618
pt.Y = 118
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 118
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + Chr$(9) + Chr$(9) + "B" + Chr$(9) + "2" + Chr$(13)
temp = 1
Exit Sub
Else
pt.X = 618
pt.Y = 138
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 138
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + Chr$(9) + Chr$(9) + "B" + Chr$(9) + "2" + Chr$(13)
temp = 0
Exit Sub
End If
End Sub



Private Sub Timer2_Timer()
If temp = 0 Then
pt.X = 618
pt.Y = 118
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 118
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "P"
temp = 1
Exit Sub
Else
pt.X = 618
pt.Y = 138
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 138
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "P"
temp = 0
Exit Sub
End If
End Sub

Private Sub Timer3_Timer()
If temp = 0 Then
pt.X = 618
pt.Y = 118
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 118
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + Chr$(9) + Chr$(9) + "B" + Chr$(9) + "2" + Chr$(13)
temp = 1
Exit Sub
Else
pt.X = 618
pt.Y = 138
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 138
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + Chr$(9) + Chr$(9) + "B" + Chr$(9) + Chr$(13)
temp = 0
Exit Sub
End If
End Sub

Private Sub Timer4_Timer()
If temp = 0 Then
pt.X = 618
pt.Y = 118
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 118
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + Chr$(9) + Chr$(9) + "B" + Chr$(9) + "2" + Chr$(13)
temp = 1
Exit Sub
Else
pt.X = 618
pt.Y = 138
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 138
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + Chr$(9) + Chr$(9) + "B" + Chr$(9) + "1" + "1" + "1" + Chr$(13)
temp = 0
Exit Sub
End If
End Sub

Private Sub Timer5_Timer()
If temp = 0 Then
pt.X = 618
pt.Y = 118
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 118
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KD"
temp = 1
Exit Sub
Else
pt.X = 618
pt.Y = 138
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 138
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KD"
temp = 0

Exit Sub
End If
End Sub

Private Sub timFlood_Timer()
SendKeys FloodMsg + Chr$(13)
End Sub

Private Sub timKick_Timer()
If temp = 0 Then
pt.X = 618
pt.Y = 118
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 118
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KD"
temp = 1
Exit Sub
Else
pt.X = 618
pt.Y = 138
ClientToScreen Me.hwnd, pt
SetCursorPos 618, 138
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KD"
temp = 0
Exit Sub
End If
End Sub
Private Sub timScroll_Timer()
If ScrollCause = "WhileKicking" Then
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KD"
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = 118
Exit Sub
End If

If ScrollCause = "WhileSetKick" Then
Scroll = NewSetB
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KD"
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = NewSetB
Exit Sub
End If

If ScrollCause = "WhileSetBan" Then
Scroll = NewSetB
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + "%{B}" + "2" + Chr$(13)
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = NewSetB
Exit Sub
End If

If ScrollCause = "WhileBanning" Then
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + "%{B}" + "2" + Chr$(13)
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = 118
Exit Sub
End If

If ScrollCause = "WhileBanningMess" Then
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + MessBan + "%{B}" + "2" + Chr$(13)
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = 118
Exit Sub
End If

If ScrollCause = "WhileHosting" Then
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "H"
Scroll = Scroll + 20
If Scroll > HostingNum Then Scroll = 118
Exit Sub
End If

If ScrollCause = "WhileKickingMess" Then
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + Mess + Chr$(13)
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = 118
Exit Sub
End If

If ScrollCause = "WhileBanning15" Then
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + "%{B}" + Chr$(13)
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = 118
Exit Sub
End If

If ScrollCause = "WhileBanning1" Then
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "KC" + Chr$(9) + Chr$(9) + "B" + Chr$(9) + "1" + "1" + "1" + Chr$(13)
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = 118
Exit Sub
End If

If ScrollCause = "WhileDehosting" Then
pt.X = 618
pt.Y = Scroll
ClientToScreen Me.hwnd, pt
SetCursorPos 618, Scroll
mouse_event MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0
SendKeys "P"
Scroll = Scroll + 20
If Scroll > 300 Then Scroll = 118
Exit Sub
End If
End Sub

Private Sub WB_StatusTextChange(ByVal Text As String)

End Sub

