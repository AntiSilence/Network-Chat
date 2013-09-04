VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options...."
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7440
   HelpContextID   =   4
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmOptions 
      BorderStyle     =   0  'None
      Height          =   2655
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   6975
      Begin VB.CheckBox chkTips 
         Caption         =   "Display Tips when Network Chat starts."
         CausesValidation=   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   4095
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         MaxLength       =   32
         TabIndex        =   18
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CheckBox chkUserName 
         Caption         =   "Use custom &Display name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CheckBox chkExitConfirm 
         Caption         =   "Show &Exit confirmation."
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chkUpdates 
         Caption         =   "Check for &Updates when Network Chat is run (Requires Internet connection)."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   6015
      End
      Begin VB.CheckBox chkUserList 
         Caption         =   "Show User &List in Main Window."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   2655
      End
      Begin VB.CheckBox chkSystemTray 
         Caption         =   "Minimise to &System Tray (includes when in Privacy Mode)."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label lblMax 
         Caption         =   "(Max 32 Chars.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   6840
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame frmOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5640
         ScaleHeight     =   375
         ScaleWidth      =   1095
         TabIndex        =   10
         Top             =   840
         Width           =   1095
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "&Do not notify upon users Entering or Exiting Network Chat."
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtSoundFile 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   5295
      End
      Begin VB.CheckBox chkPopUp 
         Caption         =   "&Popup window on incoming message."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   4215
      End
      Begin VB.CheckBox chkPlaySnd 
         Caption         =   "Play &sound on incoming message."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label lblSnd 
         Alignment       =   1  'Right Justify
         Caption         =   "Sound File:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblSndInf 
         Caption         =   "Incoming message sound and window Popup are both disabled while in Privacy Mode (both Ignore and Save Modes)."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   6495
      End
   End
   Begin MSComctlLib.TabStrip tabs 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5953
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Notification"
            ImageVarType    =   2
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
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintCurFrame As Integer

Private Sub chkPlaySnd_Click()
    If chkPlaySnd.Value = 1 Then
        txtSoundFile.Enabled = True
        lblSnd.Enabled = True
        cmdBrowse.Enabled = True
    Else
        txtSoundFile.Enabled = False
        lblSnd.Enabled = False
        cmdBrowse.Enabled = False
    End If
End Sub

Private Sub chkUserName_Click()
    If chkUserName.Value = 1 Then
        txtUserName.Enabled = True
    Else
        txtUserName.Enabled = False
    End If
End Sub


Private Sub cmdBrowse_Click()
    On Error GoTo getErr
    
    dlgDialog.Filter = "WAVE Files|*.wav"
    dlgDialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    dlgDialog.ShowOpen
    
    txtSoundFile.Text = dlgDialog.filename
    soundFile$ = txtSoundFile.Text

    Exit Sub
    
getErr:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox Err.Description, vbExclamation, "Error..."
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    ' Save selected options and set flags
    '
    ' Play a sound when...
    SaveSetting App.Title, "Config", "notifySnd", CStr(chkPlaySnd)
    If chkPlaySnd.Value = 1 Then
        playSnd = True
    Else
        playSnd = False
    End If
    
    ' Popup Window
    SaveSetting App.Title, "Config", "notifyPopUp", CStr(chkPopUp)
    If chkPopUp.Value = 1 Then
        winPopUp = True
    Else
        winPopUp = False
    End If
    
    ' Display Name
    SaveSetting App.Title, "Config", "useDisplayName", CStr(chkUserName)
    SaveSetting App.Title, "Config", "displayName", txtUserName.Text
    displayName$ = txtUserName.Text
    If chkUserName.Value = 1 Then
        useDisplayName = True
    Else
        useDisplayName = False
    End If
    
    ' User List
    SaveSetting App.Title, "Config", "UserList", CStr(chkUserList)
    If chkUserList.Value = 1 Then
        userList = True
        frmMain.Width = 7965
    Else
        userList = False
        frmMain.Width = 6090
    End If
    
    ' Minimise to SystemTray
    SaveSetting App.Title, "Config", "minSysTray", CStr(chkSystemTray)
    If chkSystemTray.Value = 1 Then
        minSysTray = True
    Else
        minSysTray = False
    End If
    
    ' Check for updates on startup
    SaveSetting App.Title, "Config", "UpdateCheck", CStr(chkUpdates)
    
    ' Exit Confirmation
    SaveSetting App.Title, "Config", "ConfirmExit", CStr(chkExitConfirm)
    
    ' Save sound file path
    SaveSetting App.Title, "Config", "SoundFile", soundFile$
    
    ' Save Tips option
    SaveSetting App.Title, "Config", "Tips", chkTips.Value
    
    ' Entry/Exit Notify
    SaveSetting App.Title, "Config", "NotifyOff", CStr(chkDisable)
    If chkDisable.Value = 1 Then
        NotifyOff = True
    Else
        NotifyOff = False
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
    chkExitConfirm.Value = GetSetting(App.Title, "Config", "ConfirmExit", CLng(1))
    chkUpdates.Value = GetSetting(App.Title, "Config", "UpdateCheck", CLng(0))
    chkUserList.Value = GetSetting(App.Title, "Config", "UserList", CLng(1))
    chkPlaySnd.Value = GetSetting(App.Title, "Config", "notifySnd", CLng("1"))
    chkPopUp.Value = GetSetting(App.Title, "Config", "notifyPopUp", CLng("1"))
    chkUserName.Value = GetSetting(App.Title, "Config", "useDisplayName", CLng("0"))
    chkSystemTray.Value = GetSetting(App.Title, "Config", "minSysTray", CLng("0"))
    chkTips.Value = GetSetting(App.Title, "Config", "Tips", 1)
    chkDisable.Value = GetSetting(App.Title, "Config", "NotifyOff", CLng("0"))
    txtUserName.Text = GetSetting(App.Title, "Config", "displayName", "")
    txtSoundFile.Text = GetSetting(App.Title, "Config", "SoundFile", App.Path & "\Sounds\Bleep1.wav")
    
    soundFile$ = txtSoundFile.Text
    mintCurFrame = 1
    
    If chkPlaySnd.Value = 1 Then
        txtSoundFile.Enabled = True
        lblSnd.Enabled = True
        cmdBrowse.Enabled = True
    Else
        txtSoundFile.Enabled = False
        lblSnd.Enabled = False
        cmdBrowse.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


Private Sub TabStrip1_Change()

End Sub


Private Sub tabs_Click()
    If Tabs.SelectedItem.Index = mintCurFrame Then Exit Sub
    
    frmOptions(Tabs.SelectedItem.Index).Visible = True
    frmOptions(mintCurFrame).Visible = False
    ' Set mintCurFrame to new value.
    mintCurFrame = Tabs.SelectedItem.Index
End Sub


