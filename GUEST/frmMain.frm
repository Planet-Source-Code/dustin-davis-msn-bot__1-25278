VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M$N Bot"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   5655
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Leo - - bigbadmotherfukkingx@hotmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Dustin - - Supa_Thug21@hotmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright© 2001 Dustin and Leo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Timer TimerUnban 
      Left            =   5520
      Top             =   1080
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Auto Un-Ban"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   13
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unban"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmMain.frx":08CA
      Left            =   120
      List            =   "frmMain.frx":08DA
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Join "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddNickName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Nick"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddRoomName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Room"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Pass"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Pass"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.ComboBox cboRoomName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin VB.ComboBox cboNickName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pass Tools"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Room Name"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nick Name"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileDeleted As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
TimerUnban.Enabled = True
Label2.Caption = "Auto Un-Ban On"
End If
If Check1.Value = 0 Then
TimerUnban.Enabled = False
Label2.Caption = "Auto Un-Ban Off"
End If
End Sub

Private Sub cmdAddNickName_Click()
cboNickName.AddItem cboNickName.Text
Open "C:\cheatnick.txt" For Append As #1
Print #1, cboNickName.Text
Close #1
End Sub

Private Sub cmdAddRoomName_Click()
If cboRoomName.Text = "" Then Exit Sub
cboRoomName.AddItem cboRoomName.Text
Open "C:\cheatroom.txt" For Append As #1
Print #1, cboRoomName.Text
Close #1
End Sub



Private Sub cmdDestroy_Click()
On Error Resume Next
If FloodNameList.ListCount < 1 Then Exit Sub
FloodNameList.ListIndex = 0
Open "C:\chatroomflood.html" For Output As #1
Do While FloodNameList.ListIndex < FloodNameList.ListCount - 1
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "CLSID:81361155-FAF9-11d3-B0D3-00C04F612FF1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "63%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "C:\msnchat4.cab" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + FloodNameList.List(FloodNameList.ListIndex) + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "C:\WINDOWS\Profiles\KubeNMå§²ér\Desktop\" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=0>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + "TN" + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
FloodNameList.ListIndex = FloodNameList.ListIndex + 1
Loop
continue:
Print #1, "</HTML>"
Close #1
Load frmbrowser2
frmbrowser2.Show
End Sub
Private Sub cmdJoin_Click()
On Error Resume Next
If FileDeleted = 1 Then
Open frmOptions.Text1.Text & "chatroom.html" For Output As #1
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "CLSID:81361155-FAF9-11d3-B0D3-00C04F612FF1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "http://fdl.msn.com/public/chat/msnchat3.cab#Version=1,1,7,058" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboNickName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + "TN" + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Close #1
Unload Me
Load frmbrowser1
frmbrowser1.Show
FileDeleted = 0
Exit Sub
End If

If FileDeleted = 0 Then
Open "C:\chatroom.html" For Output As #1
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "CLSID:81361155-FAF9-11d3-B0D3-00C04F612FF1" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "C:\msnchat3.cab" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboNickName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + "TN" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "MessageOfTheDay\" + Chr$(34) + " VALUE=\" + Chr$(34) + "Welcome To Dust|n's BoT.\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Locale" + Chr$(34) + " VALUE=" + Chr$(34) + "EN-US" + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Close #1
Unload Me
Load frmbrowser1
frmbrowser1.Show
Exit Sub
End If

End Sub
Private Sub cmdAddName_Click()
If Text1.Text = "" Then Exit Sub
FloodNameList.AddItem Text1.Text
Text1.Text = ""
Text1.SetFocus
End Sub
Private Sub cmdDeleteName_Click()
FloodNameList.RemoveItem FloodNameList.ListIndex
End Sub
Private Sub Command2_Click()
On Error Resume Next
If FileDeleted = 1 Then
Open frmOptions.Text1.Text & "chatroom.html" For Output As #1
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "http://fdl.msn.com/public/chat/msnchat4.cab#Version=1,1,7,058" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboNickName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Close #1
Unload Me
Load frmbrowser
frmbrowser.Show
FileDeleted = 0
Exit Sub
End If

If FileDeleted = 0 Then
Open "C:\chatroom.html" For Output As #1
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "CLSID:e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "C:\msnchat4.cab" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboNickName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + "TN" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "MessageOfTheDay\" + Chr$(34) + " VALUE=\" + Chr$(34) + "(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)(*)~Dustin and Leo~.\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Locale" + Chr$(34) + " VALUE=" + Chr$(34) + "EN-US" + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Close #1
Unload Me
Load frmbrowser
frmbrowser.Show
Exit Sub
End If
End Sub
Private Sub Command1_Click()
Open "C:\regloc.dat" For Input As #1
Line Input #1, temp
Close #1
Shell temp + " /e C:\temp.reg HKEY_CURRENT_USER\SOFTWARE\Microsoft\msnchat\4.0"
Open "C:\temp.reg" For Input As #1
Do While a < 4
a = a + 1
Line Input #1, temp
Loop
Text2.Text = Mid(temp, 14, Len(temp) - 14)
Close #1
End Sub
Private Sub Command3_Click()

Open "C:\RegistryCheat.reg" For Output As #1
Print #1, "REGEDIT4"
Print #1, ""
Print #1, "[HKEY_CURRENT_USER\SOFTWARE\Microsoft\msnchat\4.0]"
Print #1, Chr$(34) + "userdata1" + Chr$(34) + "=" + Chr$(34) + Text2.Text + Chr$(34)
Close #1
Open "C:\regloc.dat" For Input As #1
Line Input #1, temp
Close #1
Shell temp + " C:\registrycheat.reg", vbNormalFocus
End Sub
Private Sub Exit_Click()
End
End Sub

Private Sub Command4_Click()
Kill "C:\WINDOWS\Downloaded Program Files\MSNChat4.ini"
Kill "C:\WINDOWS\Downloaded Program Files\MSNChat40.ocx"
End Sub

Private Sub DList_Click()
FloodNameList.Clear
Open "C:\cheatnicks.txt" For Output As #1
Close #1
End Sub

Private Sub DName_Click()
FloodNameList.RemoveItem FloodNameList.ListIndex
End Sub

Private Sub FloodNameList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo StopMenu
If Button = 2 Then
If FloodNameList.ListIndex < 0 Then FloodNameList.ListIndex = 0
PopupMenu EditList, vbPopupMenuRightButton
End If
StopMenu:
End Sub

Private Sub Command5_Click()
On Error Resume Next
If FileDeleted = 1 Then
Open frmOptions.Text1.Text & "chatroom.html" For Output As #1
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "http://fdl.msn.com/public/chat/msnchat4.cab#Version=1,1,7,058" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + "FlOodErMAIN" + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "http://fdl.msn.com/public/chat/msnchat4.cab#Version=1,1,7,058" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + "FlOodEr5624534" + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "http://fdl.msn.com/public/chat/msnchat4.cab#Version=1,1,7,058" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + "FlOodEr5676" + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "http://fdl.msn.com/public/chat/msnchat4.cab#Version=1,1,7,058" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + "FlOodEr3424" + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Close #1
Unload Me
Load frmbrowser
frmbrowser.Show
FileDeleted = 0
Exit Sub
End If

If FileDeleted = 0 Then
Open "C:\chatroom.html" For Output As #1
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "CLSID:e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "C:\msnchat4.cab" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + "FlooDeR2" + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "MessageOfTheDay\" + Chr$(34) + " VALUE=\" + Chr$(34) + "(*)New 4.0 Sux Ass...(*)~~Dustin and Leo.\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Locale" + Chr$(34) + " VALUE=" + Chr$(34) + "EN-US" + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "CLSID:e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "C:\msnchat4.cab" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + "FlooDeR3" + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "MessageOfTheDay\" + Chr$(34) + " VALUE=\" + Chr$(34) + "(*)New 4.0 Sux Ass...(*)~~Dustin and Leo.\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Locale" + Chr$(34) + " VALUE=" + Chr$(34) + "EN-US" + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "CLSID:e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "C:\msnchat4.cab" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + "FlooDeR4" + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "MessageOfTheDay\" + Chr$(34) + " VALUE=\" + Chr$(34) + "(*)New 4.0 Sux Ass...(*)~~Dustin and Leo.\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Locale" + Chr$(34) + " VALUE=" + Chr$(34) + "EN-US" + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Print #1, "<script language=" + Chr$(34) + "JavaScript" + Chr$(34) + ">"
Print #1, "var nMode = 0;"
Print #1, "var temp = '<OBJECT ID=" + Chr$(34) + "ChatFrame" + Chr$(34) + " CLASSID=" + Chr$(34) + "CLSID:e87a6788-1d0f-4444-8898-1d25829b6755" + Chr$(34) + " width=" + Chr$(34) + "100%" + Chr$(34) + " height=" + Chr$(34) + "100%" + Chr$(34) + " CODEBASE=" + Chr$(34) + "C:\msnchat4.cab" + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "RoomName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + cboRoomName.Text + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "NickName\" + Chr$(34) + " VALUE=\" + Chr$(34) + "" + "FlooDeR5" + "\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Server" + Chr$(34) + " VALUE=" + Chr$(34) + "207.46.216.29:6667" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "BaseURL" + Chr$(34) + " VALUE=" + Chr$(34) + "http://msn.chat.com/" + Chr$(34) + ">';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "ChatMode" + Chr$(34) + " VALUE=2>';"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Category" + Chr$(34) + " VALUE=" + Chr$(34) + Combo1.Text + Chr$(34) + ">';"
Print #1, "temp += " + Chr$(34) + "<PARAM NAME=\" + Chr$(34) + "MessageOfTheDay\" + Chr$(34) + " VALUE=\" + Chr$(34) + "(*)New 4.0 Sux Ass...(*)~~Dustin and Leo.\" + Chr$(34) + ">" + Chr$(34) + ";"
Print #1, "temp += '<PARAM NAME=" + Chr$(34) + "Locale" + Chr$(34) + " VALUE=" + Chr$(34) + "EN-US" + Chr$(34) + ">';"
Print #1, "temp += '</OBJECT>';"
Print #1, "document.write(temp);"
Print #1, "</script>"
Close #1
Unload Me
Load frmbrowser
frmbrowser.Show
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim temp As String
Open "C:\regloc.dat" For Input As #1
Line Input #1, temp
Close #1
Shell temp + " /e C:\temp.reg HKEY_CURRENT_USER\SOFTWARE\Microsoft\msnchat\4.0"
On Error GoTo filemake3
Open "C:\temp.reg" For Input As #1
Do While a < 4
a = a + 1
Line Input #1, temp
Loop
Text2.Text = Mid(temp, 14, Len(temp) - 14)
Close #1
Open "C:\cheatnick.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, temp
cboNickName.AddItem temp
Loop
Close #1
Open "C:\cheatnicks.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, temp
FloodNameList.AddItem temp
Loop
Close #1
Open "C:\cheatroom.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, temp
cboRoomName.AddItem temp
Loop
Close #1
Exit Sub
filemake3:
Open "C:\cheatnicks.txt" For Output As #1
Close #1
Open "C:\cheatnick.txt" For Output As #1
Close #1
Open "C:\cheatroom.txt" For Output As #1
Close #1
End Sub

Private Sub Image1_Click()
Frame1.Visible = True
End Sub

Private Sub Label5_Click()
Text1.Text = ""
Text2.Text = ""
cboNickName.Clear
cboRoomName.Clear
cboNickName.Text = ""
cboRoomName.Text = ""
Open "C:\cheatnicks.txt" For Output As #1
Close #1
Open "C:\cheatnick.txt" For Output As #1
Close #1
Open "C:\cheatroom.txt" For Output As #1
Close #1
Open "C:\regloc.dat" For Input As #1
Line Input #1, temp
Close #1
Shell temp + " /e C:\temp.reg HKEY_CURRENT_USER\SOFTWARE\Microsoft\msnchat\3.0"
Open "C:\temp.reg" For Input As #1
Do While a < 4
a = a + 1
Line Input #1, temp
Loop
Text2.Text = Mid(temp, 14, Len(temp) - 14)
Close #1
End Sub

Private Sub Label6_Click()
Shell "Start Mailto:supa_thug21@hotmail.com"
End Sub

Private Sub Label7_Click()
Shell "start mailto:bigbadmotherfukkingx@hotmail.com"
End Sub

Private Sub SList_Click()
Open "C:\cheatnick.txt" For Output As #1
Close #1
If FloodNameList.ListCount < 1 Then
MsgBox "No List", vbOKOnly, "Error"
Exit Sub
End If
Open "C:\cheatnicks.txt" For Output As #1
FloodNameList.ListIndex = 0
Do While FloodNameList.ListCount - 1 > FloodNameList.ListIndex
Print #1, FloodNameList.List(FloodNameList.ListIndex)
FloodNameList.ListIndex = FloodNameList.ListIndex + 1
Loop
Close #1
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub TimerUnban_Timer()
On Error GoTo er
Kill "C:\WINDOWS\Downloaded Program Files\MSNChat4.ini"
Kill "C:\WINDOWS\Downloaded Program Files\MSNChat40.ocx"
FileDeleted = 2
Exit Sub
er: TimerUnban.Enabled = False
End Sub
