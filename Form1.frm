VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Object = "{F27AA381-8600-11D1-AD8F-DB21EA843472}#4.3#0"; "TrayIcn2.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Chat 1.40"
   ClientHeight    =   4095
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   9999
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg 
      Left            =   8400
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DevPowerTrayIcon.TrayIcon tray 
      Left            =   8400
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "Network Chat"
      Icon            =   "Form1.frx":1272
   End
   Begin SysInfoLib.SysInfo SysInf 
      Left            =   8400
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfTemp 
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":24F4
   End
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   7920
      Top             =   1800
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   7920
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtSave 
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form1.frx":256F
   End
   Begin RichTextLib.RichTextBox txtOutput 
      Height          =   630
      Left            =   120
      TabIndex        =   0
      Top             =   3030
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   1111
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   2
      MaxLength       =   1024
      Appearance      =   0
      TextRTF         =   $"Form1.frx":25EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MCI.MMControl mci 
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   873
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComctlLib.ImageList btnImgLst 
      Left            =   7920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   41
      ImageHeight     =   37
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2665
            Key             =   "send_up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3113
            Key             =   "send_dis"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3BC1
            Key             =   "send_over"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":466F
            Key             =   "em_up"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A71
            Key             =   "em_over"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5553
            Key             =   "font_up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5E05
            Key             =   "font_over"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   3795
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   529
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7964
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7858
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   7920
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   1001
   End
   Begin RichTextLib.RichTextBox txtInput 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":66B7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   5880
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Image imgFont 
      Height          =   300
      Left            =   1440
      MouseIcon       =   "Form1.frx":6733
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":6A3D
      Top             =   2595
      Width           =   810
   End
   Begin VB.Image imgEm 
      Height          =   300
      Left            =   120
      MouseIcon       =   "Form1.frx":72DF
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":75E9
      Top             =   2595
      Width           =   1215
   End
   Begin MSForms.ListBox lstUsers 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """£""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   2
      EndProperty
      Height          =   3135
      Left            =   6120
      TabIndex        =   5
      Top             =   480
      Width           =   1575
      BackColor       =   16777215
      BorderStyle     =   1
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "2778;5212"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   9120
      X2              =   -480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image imgBtnSend 
      Enabled         =   0   'False
      Height          =   555
      Left            =   5205
      MouseIcon       =   "Form1.frx":79DB
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":7CE5
      Top             =   3075
      Width           =   615
   End
   Begin VB.Image imgBG 
      Height          =   3795
      Left            =   0
      Picture         =   "Form1.frx":8783
      Top             =   0
      Width           =   7875
   End
   Begin VB.Menu mnuChat 
      Caption         =   "&Chat"
      Begin VB.Menu mnuChatConfig 
         Caption         =   "Network &Configuration..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuMiscOpt 
         Caption         =   "&Options..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuChatSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChatExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMessage 
      Caption         =   "&Message"
      Begin VB.Menu mnuMessagePrivate 
         Caption         =   "Send Private &Message..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuMessageIP 
         Caption         =   "Send Message to IP &Address..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuMessageSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMessagePrivacy 
         Caption         =   "&Privacy..."
         Begin VB.Menu mnuPrivOnSave 
            Caption         =   "On - &Save Incoming Messages"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuPrivOnIgnore 
            Caption         =   "On - &Ignore Incoming Messages"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuPrivSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPrivOff 
            Caption         =   "Privacy &Off"
            Checked         =   -1  'True
            Shortcut        =   ^F
         End
      End
      Begin VB.Menu mnuMessageSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMessageClear 
         Caption         =   "&Clear Messages"
         Enabled         =   0   'False
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpUpgradeChk 
         Caption         =   "Check for &Updates..."
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChatAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpOpen 
         Caption         =   "&Open Network Chat"
      End
      Begin VB.Menu mnuPopUpAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuPopUpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpExit 
         Caption         =   "E&xit Network Chat"
      End
   End
   Begin VB.Menu mnuPopUp1 
      Caption         =   "PopUp1"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUp1Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPopUp1Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Speech As New SpVoice
Sub findPublicSmiley(str As String)
    Dim lp2 As Integer
    
    ' 2 char codes
    For lp2 = 1 To Len(str$)
        ' find occurances of smileys:
        DoEvents
        If Mid$(str$, lp2, 2) = ":)" Then
            Call avatar("happy")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ":)", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 2) = ";)" Then
            Call avatar("wink")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ";)", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 2) = ":(" Then
            Call avatar("sad")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ":(", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 2) = ":$" Then
            Call avatar("blush")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ":$", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 2) = ":D" Then
            Call avatar("grin")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ":D", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 2) = "8)" Then
            Call avatar("shades")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, "8)", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 2) = ":@" Then
            Call avatar("angry")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ":@", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 2) = ":%" Then
            Call avatar("worried")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ":%", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 2) = ":P" Or Mid$(str$, lp2, 2) = ":p" Then
            Call avatar("tongue")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ":P", rtfAvatar$)
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, ":p", rtfAvatar$)
        End If
        DoEvents
    Next lp2
    
    ' 3 char codes
    For lp2 = 1 To Len(str$)
        ' find occurances of smileys:
        DoEvents
        If Mid$(str$, lp2, 3) = "(b)" Or Mid$(str$, lp2, 3) = "(B)" Then
            Call avatar("beer")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, "(b)", rtfAvatar$)
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, "(B)", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 3) = "(l)" Or Mid$(str$, lp2, 3) = "(L)" Then
            Call avatar("heart")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, "(l)", rtfAvatar$)
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, "(L)", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 3) = "(%)" Then
            Call avatar("guns")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, "(%)", rtfAvatar$)
        End If
        DoEvents
        If Mid$(str$, lp2, 3) = "(6)" Then
            Call avatar("devil")
            rtfTemp.TextRTF = Replace(rtfTemp.TextRTF, "(6)", rtfAvatar$)
        End If
        DoEvents
    Next lp2
End Sub

Sub updateCheck()
    Dim retVal As Long
    
    On Error GoTo errTrap
    
    retVal = Shell(App.Path & "\gwebupdate.exe", vbNormalFocus)
    
    Exit Sub
    
errTrap:
    If Err.Number = 53 Then
        MsgBox "It appears that Web Update is not installed, or cannot be run." & vbCrLf & _
                "Please visit www.global-devtech.com to check for updated versions.", vbCritical, "Error"
    End If
End Sub




Private Sub Form_Load()
    On Error GoTo errTrap

    Dim filename As String
    Dim filenum As Integer
    Dim str As String
    Dim optVal As Integer
       
    ' Allow active hyper links in chat window
    EnableURLDetect txtInput.hwnd, Me.hwnd
       
    Me.Caption = "Network Chat " & App.Major & "." & Format(App.Minor, "0#")
      
    ' Check for OS and display Speech Message for NT based OS's
    'If SysInf.OSPlatform = 0 Or frmMain.SysInf.OSPlatform = 1 Then ' Windows 9x or Win32s
        'mnuMessageSpeech.Visible = False
    'ElseIf SysInf.OSPlatform = 2 Then ' Windows NT
        'mnuMessageSpeech.Visible = True
    'End If
         
    ' Set size of form if user list shown
    optVal = GetSetting(App.Title, "Config", "UserList", CLng(1))
    If optVal = 1 Then
        Me.Width = 7965
        userList = True
    Else
        Me.Width = 6090
        userList = False
    End If
    
    Me.Left = CLng(GetSetting(App.Title, "Config", "x", Screen.Width / 2 - Me.Width / 2))
    Me.Top = CLng(GetSetting(App.Title, "Config", "y", Screen.Height / 2 - Me.Height / 2))
    
    ' Setup MCI audio device
    mci.DeviceType = "WaveAudio"
    mci.Notify = False
    mci.Wait = False
    mci.Shareable = False
    
    status.Panels(2).Text = "No Messages"
    
    ' Set WinSock port...
    Winsock(0).RemotePort = CLng(GetSetting(App.Title, "Config", "Port", "1001"))
    
    ' Set options...
    ' --------------
    ' Set Font
    txtInput.Font = GetSetting(App.Title, "Config", "Font", "Verdana")
    txtInput.Font.Size = CLng(GetSetting(App.Title, "Config", "FontSize", "8"))
    txtInput.Font.Bold = GetSetting(App.Title, "Config", "FontBold", "False")
    txtInput.Font.Italic = GetSetting(App.Title, "Config", "FontItalic", "False")
    txtOutput.Font = GetSetting(App.Title, "Config", "Font", "Verdana")
    txtOutput.Font.Size = CLng(GetSetting(App.Title, "Config", "FontSize", "8"))
    txtOutput.Font.Bold = GetSetting(App.Title, "Config", "FontBold", "False")
    txtOutput.Font.Italic = GetSetting(App.Title, "Config", "FontItalic", "False")
    txtSave.Font = GetSetting(App.Title, "Config", "Font", "Verdana")
    txtSave.Font.Size = CLng(GetSetting(App.Title, "Config", "FontSize", "8"))
    txtSave.Font.Bold = GetSetting(App.Title, "Config", "FontBold", "False")
    txtSave.Font.Italic = GetSetting(App.Title, "Config", "FontItalic", "False")
    rtfTemp.Font = GetSetting(App.Title, "Config", "Font", "Verdana")
    rtfTemp.Font.Size = CLng(GetSetting(App.Title, "Config", "FontSize", "8"))
    rtfTemp.Font.Bold = GetSetting(App.Title, "Config", "FontBold", "False")
    rtfTemp.Font.Italic = GetSetting(App.Title, "Config", "FontItalic", "False")
    
    ' Get user name
    optVal = GetSetting(App.Title, "Config", "useDisplayName", CLng("0"))
    If optVal = 1 Then
        useDisplayName = True
    Else
        useDisplayName = False
    End If
    displayName$ = GetSetting(App.Title, "Config", "displayName", "")

    ' Play Sound When...
    optVal = CLng(GetSetting(App.Title, "Config", "notifySnd", "1"))
    If optVal = 1 Then
        playSnd = True
    Else
        playSnd = False
    End If
    
    ' PopUp Windows...
    optVal = CLng(GetSetting(App.Title, "Config", "notifyPopUp", "1"))
    If optVal = 1 Then
        winPopUp = True
    Else
        winPopUp = False
    End If
    
    ' Minimise to SystemTray option...
    optVal = CLng(GetSetting(App.Title, "Config", "minSysTray", "1"))
    If optVal = 1 Then
        minSysTray = True
    Else
        minSysTray = False
    End If
    
    NotifyOff = CBool(GetSetting(App.Title, "Config", "NotifyOff", "0"))
    
    privacy = False
    saveMsg = False
    
    imgBtnSend.Picture = btnImgLst.ListImages.Item("send_dis").Picture
    sendDisabled = True
    
    ' Bind winsock to port
    Dim portNum As Integer
    portNum = CLng(GetSetting(App.Title, "Config", "Port", "1001"))
    Winsock(0).Bind portNum
    
    ' open NETWORK.CFG file and send ENTERED message to each user
    ' Also at the same time, add user names to user list.
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    Open filename$ For Input As #filenum
        If LOF(filenum) = 0 Then
            Close #filenum
        Else
            Do Until EOF(filenum)
                DoEvents
                Line Input #filenum, str$
                
                If str$ = UCase$(Winsock(0).LocalHostName) Then
                    ' ignore own computer name if user added to CFG file! Doh!
                    Close #filenum
                    Exit Sub
                ElseIf Len(str$) > 1 Then
                    Winsock(0).RemoteHost = str$
                    Winsock(0).SendData "*** " & UCase$(Winsock(0).LocalHostName) & " HAS ENTERED NETWORK CHAT ***"
                    ' add to user list
                    lstUsers.AddItem str$
                Else
                    ' ignore blank lines
                End If
            Loop
        End If
    Close #filenum
        
    If ShowAtStartup = 1 Then
        Me.Show
        Me.Refresh
        frmTip.Show vbModal
    End If
    
    Exit Sub
    
errTrap:
    If Err.Number = 10048 Then
        MsgBox "The selected Port number (" & frmMain.Winsock(0).RemotePort & ") is currently in use. You will not be able to send " & _
                "or receive any messages until the port becomes available.", vbOKOnly Or vbCritical, "Port In Use..."
    Else
        Resume Next
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errTrap
    Dim retVal As Integer
    Dim filename As String
    Dim filenum As Integer
    Dim str As String
    
    ' Check to see if the Show Exit Confirmation option is enabled
    retVal = GetSetting(App.Title, "Config", "ConfirmExit", CLng(1))
    
    If UnloadMode = 2 Or UnloadMode = 3 Then
        retVal = 0
    End If
    If retVal = 1 Then
        retVal = MsgBox("Are you sure you want to exit?", vbOKCancel Or vbQuestion, "Confirm Exit")
        If retVal = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    ' open NETWORK.CFG file and send LEFT message to each user
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    Open filename$ For Input As #filenum
    If LOF(filenum) = 0 Then
        Close #filenum
    Else
        Do Until EOF(filenum)
            DoEvents
            Line Input #filenum, str$
                
            If Len(str$) > 1 Then
                Winsock(0).RemoteHost = str$
                Winsock(0).SendData "*** " & UCase$(Winsock(0).LocalHostName) & " HAS LEFT NETWORK CHAT ***"
            Else
                ' ignore blank lines
            End If
        Loop
    End If
    Close #filenum
    Exit Sub

errTrap:
    Exit Sub
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbMinimized And minSysTray = True Then
        tray.Visible = True
        Me.Hide
    Else
        tray.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frmForm As Form

    mci.Command = "Close"
    
    SaveSetting App.Title, "Config", "x", CStr(Me.Left)
    SaveSetting App.Title, "Config", "y", CStr(Me.Top)
    
    DisableURLDetect

    For Each frmForm In Forms
        Unload frmForm
    Next
    
    Unload Me
End Sub

Private Sub imgBG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEm.Picture = btnImgLst.ListImages.Item("em_up").Picture
    imgFont.Picture = btnImgLst.ListImages.Item("font_up").Picture
    If sendDisabled = True Then
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_dis").Picture
    Else
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_up").Picture
    End If
End Sub


Private Sub imgBtnSend_Click()
    On Error Resume Next
    
    Dim filenum As Long
    Dim str As String
    Dim str2 As String
    Dim filename As String
    Dim lp As Long
    
    'Me.MousePointer = vbHourglass
    
    ' Send message to own computer!
    Winsock(0).RemoteHost = Winsock(0).LocalHostName
    
    If useDisplayName = True Then
        str$ = UCase$(Winsock(0).LocalHostName) & " - " & displayName$ & " says:" & vbCrLf & "    " & txtOutput.Text
    Else
        str$ = UCase$(Winsock(0).LocalHostName) & " says:" & vbCrLf & "    " & txtOutput.Text
    End If
       
    str2 = txtOutput.Text
    txtOutput.Text = ""
    imgBtnSend.Enabled = False
       
    Winsock(0).SendData str$
    
    ' open NETWORK.CFG file and send message to each user
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    Open filename$ For Input As #filenum
        If LOF(filenum) = 0 Then
            Close #filenum
            txtOutput.Text = ""
            'Me.MousePointer = vbDefault
        Else
            Do Until EOF(filenum)
                DoEvents
                Line Input #filenum, str$
                
                If Len(str$) > 1 Then
                    Winsock(0).RemoteHost = str$
                    If useDisplayName = True Then
                        Winsock(0).SendData UCase$(Winsock(0).LocalHostName) & " - " & displayName$ & " says:" & vbCrLf & "    " & str2
                    Else
                        Winsock(0).SendData UCase$(Winsock(0).LocalHostName) & " says:" & vbCrLf & "    " & str2
                    End If
                Else
                    ' ignore blank lines
                End If
            Loop
        End If
    Close #filenum
    
    'Me.MousePointer = vbDefault
End Sub

Private Sub imgBtnSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEm.Picture = btnImgLst.ListImages.Item("em_up").Picture
    imgFont.Picture = btnImgLst.ListImages.Item("font_up").Picture
    If sendDisabled = True Then
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_dis").Picture
    Else
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_up").Picture
    End If
End Sub


Private Sub imgEm_Click()
    frmAvatar.Left = Me.Left + 150
    frmAvatar.Top = Me.Top + 3500
    
    frmAvatar.Show
End Sub

Private Sub imgEm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEm.Picture = btnImgLst.ListImages.Item("em_over").Picture
    imgFont.Picture = btnImgLst.ListImages.Item("font_up").Picture
End Sub


Private Sub imgFont_Click()
    On Error GoTo errTrap
    
    dlg.CancelError = True
    dlg.Flags = cdlCFScreenFonts
    
    dlg.FontName = GetSetting(App.Title, "Config", "Font", "Verdana")
    dlg.FontSize = CLng(GetSetting(App.Title, "Config", "FontSize", "8"))
    dlg.FontBold = GetSetting(App.Title, "Config", "FontBold", "False")
    dlg.FontItalic = GetSetting(App.Title, "Config", "FontItalic", "False")

    dlg.ShowFont
    
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.TextRTF)
    txtInput.SelFontName = dlg.FontName
    txtInput.SelBold = dlg.FontBold
    txtInput.SelItalic = dlg.FontItalic
    txtInput.SelFontSize = dlg.FontSize
    txtInput.SelStart = Len(frmMain.txtInput.TextRTF)
    
    rtfTemp.SelStart = 0
    rtfTemp.SelLength = Len(txtInput.TextRTF)
    rtfTemp.SelFontName = dlg.FontName
    rtfTemp.SelBold = dlg.FontBold
    rtfTemp.SelItalic = dlg.FontItalic
    rtfTemp.SelFontSize = dlg.FontSize
    rtfTemp.SelStart = Len(frmMain.txtInput.TextRTF)

    txtSave.SelStart = 0
    txtSave.SelLength = Len(txtInput.TextRTF)
    txtSave.SelFontName = dlg.FontName
    txtSave.SelBold = dlg.FontBold
    txtSave.SelItalic = dlg.FontItalic
    txtSave.SelFontSize = dlg.FontSize
    txtSave.SelStart = Len(frmMain.txtInput.TextRTF)
    
    txtOutput.Font = dlg.FontName
    txtOutput.Font.Bold = dlg.FontBold
    txtOutput.Font.Italic = dlg.FontItalic
    txtOutput.Font.Size = dlg.FontSize
    
    ' Save font settings
    SaveSetting App.Title, "Config", "Font", dlg.FontName
    SaveSetting App.Title, "Config", "FontSize", dlg.FontSize
    SaveSetting App.Title, "Config", "FontBold", dlg.FontBold
    SaveSetting App.Title, "Config", "FontItalic", dlg.FontItalic
    
    Exit Sub
    
errTrap:
    
End Sub

Private Sub imgFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgFont.Picture = btnImgLst.ListImages.Item("font_over").Picture
    imgEm.Picture = btnImgLst.ListImages.Item("em_up").Picture
End Sub


Private Sub lstUsers_DblClick(Cancel As MSForms.ReturnBoolean)
    On Error GoTo errTrap
    
    Dim user As String
    user$ = lstUsers.List(lstUsers.ListIndex)
    
    With frmMsg
        .Caption = "Private Message"
        .cmbNames.Visible = False
        .lblName.Caption = user$
        .lblName.Visible = True
        .txtMsg.Enabled = True
        .Show vbModal
    End With
    Exit Sub
    
errTrap:
    Exit Sub
End Sub


Private Sub lstUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEm.Picture = btnImgLst.ListImages.Item("em_up").Picture
    imgFont.Picture = btnImgLst.ListImages.Item("font_up").Picture
End Sub


Private Sub mnuChatAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuChatConfig_Click()
    frmConfig.Show vbModal
End Sub

Private Sub mnuChatExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpContents_Click()
    SendKeys "{F1}"
End Sub

Private Sub mnuHelpUpgradeChk_Click()
    Call updateCheck
End Sub

Private Sub mnuMessageClear_Click()
    txtInput.Text = ""
    txtSave.Text = ""
    txtOutput.SetFocus
    numPublic = 0
    numPrivate = 0
    status.Panels(1).Text = ""
    status.Panels(2).Text = "No Messages"
    mnuMessageClear.Enabled = False
End Sub

Private Sub mnuMessageIP_Click()
    frmIP.Show vbModal
End Sub

Private Sub mnuMessagePrivate_Click()
    On Error GoTo errTrap
    
    Dim filename As String
    Dim filenum As Integer
    Dim str As String
    
    ' Load NETWORK.CFG file and add to listbox
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    Open filename$ For Input As #filenum
        If LOF(filenum) = 0 Then
            MsgBox "There are no computer names in the configuration file. Please use the Chat -> Network Configuration " & _
                    "menu option to create a configuration file.", vbExclamation, "Error..."
            Close #filenum
            Exit Sub
        Else
            Do Until EOF(filenum)
                Line Input #filenum, str$
                If Len(str$) > 0 Then
                    frmMsg.cmbNames.AddItem str$
                Else
                    '
                End If
            Loop
            frmMsg.Show vbModal
        End If
    Close #filenum
    
    Exit Sub
    
errTrap:
    If Err.Number = 53 Then
        MsgBox "The Network Configuration file was not found. Please use the Chat -> Network Configuration " & _
                "menu option to create a configuration file.", vbCritical, "Error..."
    End If
End Sub

Private Sub mnuMiscOpt_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuPopUp1Copy_Click()
    Clipboard.Clear
    Clipboard.SetText frmMain.ActiveControl.SelText
End Sub

Private Sub mnuPopUp1Paste_Click()
    txtOutput.SelText = Clipboard.GetText(vbCFText)
End Sub

Private Sub mnuPopUpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuPopUpExit_Click()
    Unload Me
End Sub

Private Sub mnuPopUpOpen_Click()
    Me.WindowState = vbNormal
    Me.Show
    tray.Visible = False
End Sub

Private Sub mnuPrivOff_Click()
    privacy = False
    mnuPrivOnIgnore.Checked = False
    mnuPrivOnSave.Checked = False
    mnuPrivOff.Checked = True
    
    If saveMsg = True Then
        txtInput.TextRTF = txtSave.TextRTF
    End If
    
    If txtInput.Text = "" Then
        status.Panels(1).Text = ""
    Else
        status.Panels(1).Text = lastMsgTime$
    End If
    
    txtInput.Visible = True
    txtOutput.Visible = True
    txtInput.SetFocus
    txtInput.SelStart = Len(txtInput.TextRTF)
    txtOutput.SetFocus
    
    imgEm.Enabled = True
    imgFont.Enabled = True
    lstUsers.Enabled = True
    
    saveMsg = False
    
    mnuMessageClear.Enabled = True

    Dim filename As String
    Dim filenum As Integer
    Dim str As String
    
    ' open NETWORK.CFG file and send ENTERED message to each user
    ' Also at the same time, add user names to user list.
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    Open filename$ For Input As #filenum
        If LOF(filenum) = 0 Then
            Close #filenum
        Else
            Do Until EOF(filenum)
                DoEvents
                Line Input #filenum, str$
                
                If str$ = UCase$(Winsock(0).LocalHostName) Then
                    ' ignore own computer name if user added to CFG file! Doh!
                    Close #filenum
                    Exit Sub
                ElseIf Len(str$) > 1 Then
                    Winsock(0).RemoteHost = str$
                    Winsock(0).SendData "*** " & UCase$(Winsock(0).LocalHostName) & " IS NOW AVAILABLE FOR CHAT ***"
                Else
                    ' ignore blank lines
                End If
            Loop
        End If
    Close #filenum
End Sub

Private Sub mnuPrivOnIgnore_Click()
    privacy = True
    saveMsg = False
    mnuPrivOnIgnore.Checked = True
    mnuPrivOnSave.Checked = False
    mnuPrivOff.Checked = False
    
    If saveMsg = True Then
        lastMsgTime$ = status.Panels(1).Text
    End If
    
    imgEm.Enabled = False
    imgFont.Enabled = False
    lstUsers.Enabled = False
    
    status.Panels(1).Text = "Privacy (Ignore)"
    txtInput.Visible = False
    txtOutput.Visible = False
    
    mnuMessageClear.Enabled = False
    Me.WindowState = vbMinimized
    
    Dim filename As String
    Dim filenum As Integer
    Dim str As String
    
    ' open NETWORK.CFG file and send ENTERED message to each user
    ' Also at the same time, add user names to user list.
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    Open filename$ For Input As #filenum
        If LOF(filenum) = 0 Then
            Close #filenum
        Else
            Do Until EOF(filenum)
                DoEvents
                Line Input #filenum, str$
                
                If str$ = UCase$(Winsock(0).LocalHostName) Then
                    ' ignore own computer name if user added to CFG file! Doh!
                    Close #filenum
                    Exit Sub
                ElseIf Len(str$) > 1 Then
                    Winsock(0).RemoteHost = str$
                    Winsock(0).SendData "*** " & UCase$(Winsock(0).LocalHostName) & " IS IN PRIVACY (IGNORE) MODE ***"
                Else
                    ' ignore blank lines
                End If
            Loop
        End If
    Close #filenum
End Sub

Private Sub mnuPrivOnSave_Click()
    privacy = True
    saveMsg = True
    mnuPrivOnIgnore.Checked = False
    mnuPrivOnSave.Checked = True
    mnuPrivOff.Checked = False
    status.Panels(1).Text = "Privacy (Save)"
    txtInput.Visible = False
    txtOutput.Visible = False
    
    imgEm.Enabled = False
    imgFont.Enabled = False
    lstUsers.Enabled = False
    
    mnuMessageClear.Enabled = False
    Me.WindowState = vbMinimized
    
    Dim filename As String
    Dim filenum As Integer
    Dim str As String
    
    ' open NETWORK.CFG file and send ENTERED message to each user
    ' Also at the same time, add user names to user list.
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    Open filename$ For Input As #filenum
        If LOF(filenum) = 0 Then
            Close #filenum
        Else
            Do Until EOF(filenum)
                DoEvents
                Line Input #filenum, str$
                
                If str$ = UCase$(Winsock(0).LocalHostName) Then
                    ' ignore own computer name if user added to CFG file! Doh!
                    Close #filenum
                    Exit Sub
                ElseIf Len(str$) > 1 Then
                    Winsock(0).RemoteHost = str$
                    Winsock(0).SendData "*** " & UCase$(Winsock(0).LocalHostName) & " IS IN PRIVACY (SAVE) MODE ***"
                Else
                    ' ignore blank lines
                End If
            Loop
        End If
    Close #filenum
End Sub


Private Sub status_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEm.Picture = btnImgLst.ListImages.Item("em_up").Picture
    imgFont.Picture = btnImgLst.ListImages.Item("font_up").Picture
    If sendDisabled = True Then
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_dis").Picture
    Else
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_up").Picture
    End If
End Sub

Private Sub tmr_Timer()
    Dim retVal
    
    ' Check for update if option is on
    retVal = GetSetting(App.Title, "Config", "UpdateCheck", CLng(0))
    If retVal = 1 Then
        Call updateCheck
    End If
    
    tmr.Enabled = False
End Sub

Private Sub tray_Click()
    Me.WindowState = vbNormal
    Me.Show
    tray.Visible = False
End Sub

Private Sub tray_RightClick()
    PopupMenu mnuPopUp, , , , mnuPopUpOpen
End Sub


Private Sub txtInput_Change()
    If Len(txtInput.Text) > 0 Then
        mnuMessageClear.Enabled = True
    Else
        mnuMessageClear.Enabled = False
    End If
End Sub

Private Sub txtInput_Click()
    'txtInput.SelStart = Len(txtInput.TextRTF)
    'txtOutput.SetFocus
End Sub


Private Sub txtInput_KeyPress(KeyAscii As Integer)
    txtOutput.Text = txtOutput.Text & Chr$(KeyAscii)
    
    txtOutput.SelStart = Len(txtInput.TextRTF)
    txtOutput.SetFocus
End Sub


Private Sub txtInput_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEm.Picture = btnImgLst.ListImages.Item("em_up").Picture
    imgFont.Picture = btnImgLst.ListImages.Item("font_up").Picture
    If sendDisabled = True Then
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_dis").Picture
    Else
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_up").Picture
    End If
End Sub


Private Sub txtInput_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mnuPopUp1Paste.Visible = False
        If txtInput.SelLength > 0 Then
            mnuPopUp1Copy.Enabled = True
        Else
            mnuPopUp1Copy.Enabled = False
        End If
        PopupMenu mnuPopUp1
    End If
End Sub

Private Sub txtoutput_Change()
    If Len(txtOutput.Text) > 0 Then
        imgBtnSend.Enabled = True
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_up").Picture
        sendDisabled = False
    Else
        imgBtnSend.Enabled = False
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_dis").Picture
        sendDisabled = True
    End If
End Sub

Private Sub txtoutput_Click()
    'lstUsers
End Sub


Private Sub txtoutput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(txtOutput.Text) > 0 Then
        KeyAscii = 0
        imgBtnSend_Click
    End If
End Sub


Private Sub txtoutput_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEm.Picture = btnImgLst.ListImages.Item("em_up").Picture
    imgFont.Picture = btnImgLst.ListImages.Item("font_up").Picture
    If sendDisabled = True Then
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_dis").Picture
    Else
        imgBtnSend.Picture = btnImgLst.ListImages.Item("send_up").Picture
    End If
End Sub

Private Sub txtOutput_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        mnuPopUp1Copy.Visible = True
        mnuPopUp1Paste.Visible = True
        If Clipboard.GetFormat(vbCFText) = True Then
            mnuPopUp1Paste.Enabled = True
        Else
            mnuPopUp1Paste.Enabled = False
        End If
        
        If txtOutput.SelLength > 0 Then
            mnuPopUp1Copy.Enabled = True
        Else
            mnuPopUp1Copy.Enabled = False
        End If
        
        PopupMenu mnuPopUp1
    End If
End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim strInput As String
    Dim lp1 As Integer
    Dim lp2 As Integer
    Dim fromUser As String
    Dim message As String
            
    ' Check OS platform... If Win32s or Win9x then add a Linefeed to
    ' incoming string. This is done due to compensate for the differences
    ' between Win9x and WinNT text formatting
    If SysInf.OSPlatform = 0 Or SysInf.OSPlatform = 1 Then ' Windows 9x or Win32s
        Winsock(0).GetData strInput$
        strInput$ = strInput$ & vbCrLf
    ElseIf SysInf.OSPlatform = 2 Then ' Windows NT
        Winsock(0).GetData strInput$
    End If

    txtInput.SelStart = Len(txtInput.TextRTF)
    
    If Len(strInput$) < 1 Or strInput$ = vbCrLf Then
        Exit Sub
    End If
    
    If Left$(strInput$, 9) = "/cmd:test" Then
        Exit Sub
    End If
       
    If privacy = True And saveMsg = False Then
        Exit Sub
    End If
        
    If Me.WindowState = vbMinimized And winPopUp = True And privacy = False Then
        If Left$(strInput$, 3) = "***" And NotifyOff = True Then
            ' Do nothing
        Else
            Me.WindowState = vbNormal
            Me.Show
        End If
    End If
    
    ' Check to see if sender is local host, if so don't play a sound!
    Dim tmpName As String
    tmpName$ = UCase$(Winsock(0).LocalHostName)

    If Left$(strInput$, Len(tmpName$)) = tmpName$ Then
        'do nothing
    ElseIf playSnd = True And privacy = False Then
        If Left$(strInput$, 3) = "***" And NotifyOff = True Then
            ' Do nothing
        Else
            soundFile$ = GetSetting(App.Title, "Config", "SoundFile", App.Path & "\Sounds\Bleep1.wav")
            mci.filename = soundFile$
            mci.Command = "Open"
            mci.Command = "Prev"
            mci.Command = "Play"
        End If
    End If
    
    If Left$(strInput$, 7) = "/msgsp;" Then
        ' Speech Message (counted as Private)
        message$ = Right$(strInput, Len(strInput$) - 7)
        rtfTemp.TextRTF = message$
        
        If minSysTray = True And winPopUp = False And privacy = False Then
            tray.ShowBalloon "Incoming Message", message
        End If
               
        Call findPublicSmiley(rtfTemp.TextRTF)
        txtInput.SelRTF = rtfTemp.TextRTF
        txtSave.TextRTF = txtInput.TextRTF
            
        ' ????
        If saveMsg = False Then
            txtInput.SetFocus
            txtInput.SelStart = Len(txtInput.TextRTF)
            txtOutput.SetFocus
        End If
        
        numPrivate = numPrivate + 1
        
        ' this makes sure the message is displayed in the window before
        ' the speech engine reads it out.
        Me.Refresh
        
        Speech.Speak message$
    ElseIf Left$(strInput$, 5) = "/msg;" Then
        ' Private message
        message$ = Right$(strInput, Len(strInput$) - 5)
        rtfTemp.TextRTF = message$
        
        If minSysTray = True And winPopUp = False And privacy = False Then
            tray.ShowBalloon "Incoming Message", message
        End If
               
        Call findPublicSmiley(rtfTemp.TextRTF)
        txtInput.SelRTF = rtfTemp.TextRTF
        txtSave.TextRTF = txtInput.TextRTF
            
        ' ????
        If saveMsg = False Then
            txtInput.SetFocus
            txtInput.SelStart = Len(txtInput.TextRTF)
            txtOutput.SetFocus
        End If
               
        numPrivate = numPrivate + 1
    Else
        ' public message
        rtfTemp.TextRTF = strInput
        
        If minSysTray = True And winPopUp = False And privacy = False Then
            If Left$(strInput$, 3) = "***" And NotifyOff = True Then
                ' Do Nothing
            Else
                tray.ShowBalloon "Incoming Message", strInput
            End If
        End If
        
        Call findPublicSmiley(rtfTemp.TextRTF)
        txtInput.SelRTF = rtfTemp.TextRTF
        txtSave.TextRTF = txtInput.TextRTF
            
        txtInput.SetFocus
        txtInput.SelStart = Len(txtInput.TextRTF)
        txtOutput.SetFocus
        
        numPublic = numPublic + 1
    End If
        
    If Left$(strInput$, 3) = "***" Then
        numPublic = numPublic - 1
    ElseIf privacy = False And saveMsg = False Then
        status.Panels(1).Text = "Last Message Received at " & Time$
        lastMsgTime$ = "Last Message Received at " & Time$
        status.Panels(2).Text = "Messages: " & CStr(numPublic) & " Public, " & CStr(numPrivate) & " Private"
    End If
End Sub

