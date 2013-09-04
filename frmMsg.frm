VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Private Message..."
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSpeech 
      Caption         =   "Send as S&peech"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ComboBox cmbNames 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtMsg 
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
      Height          =   1455
      Left            =   120
      MaxLength       =   1024
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4095
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
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Send To:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbNames_Click()
    txtMsg.Enabled = True
    txtMsg.SetFocus
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdSend_Click()
    On Error Resume Next
        
    Dim name As String
    Dim str As String
    
    Me.MousePointer = vbHourglass
    
    If cmbNames.Visible = True Then
        frmMain.Winsock(0).RemoteHost = cmbNames.List(cmbNames.ListIndex)
        name$ = cmbNames.List(cmbNames.ListIndex)
    Else
        frmMain.Winsock(0).RemoteHost = lblName.Caption
        name$ = lblName.Caption
    End If
    
    If useDisplayName = True Then
        If chkSpeech.Value = 0 Then
            str$ = "/msg;PRIVATE MESSAGE from " & UCase$(frmMain.Winsock(0).LocalHostName) & " - " & displayName$ & ":" & vbCrLf & "    " & txtMsg.Text
        ElseIf chkSpeech.Value = 1 Then
            str$ = "/msgsp;SPEECH MESSAGE from " & UCase$(frmMain.Winsock(0).LocalHostName) & " - " & displayName$ & ":" & vbCrLf & "    " & txtMsg.Text
        End If
    Else
        If chkSpeech.Value = 0 Then
            str$ = "/msg;PRIVATE MESSAGE from " & UCase$(frmMain.Winsock(0).LocalHostName) & ":" & vbCrLf & "    " & txtMsg.Text
        ElseIf chkSpeech.Value = 1 Then
            str$ = "/msgsp;SPEECH MESSAGE from " & UCase$(frmMain.Winsock(0).LocalHostName) & ":" & vbCrLf & "    " & txtMsg.Text
        End If
    End If
    
    frmMain.Winsock(0).SendData str$
    
    ' TEMP OUT OF ORDER!
    'frmMain.Winsock(0).RemoteHost = frmMain.Winsock(0).LocalHostName
    'str$ = "/msg;PRIVATE MESSAGE to " & name$ & ":" & vbCrLf & "    " & txtMsg.Text & vbCrLf
    'frmMain.Winsock(0).SendData str$
    
    frmMain.txtOutput.SetFocus
    
    Me.MousePointer = vbDefault

    Unload Me
End Sub

Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
    ' Check for OS and display Get Name for NT based OS's
    If frmMain.SysInf.OSPlatform = 0 Or frmMain.SysInf.OSPlatform = 1 Then ' Windows 9x or Win32s
        chkSpeech.Visible = False
    ElseIf frmMain.SysInf.OSPlatform = 2 Then ' Windows NT
        chkSpeech.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


Private Sub txtMsg_Change()
    If Len(txtMsg.Text) > 0 Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(txtMsg.Text) > 0 Then
        KeyAscii = 0
        cmdSend_Click
    End If
End Sub


Private Sub txtTo_Change()

End Sub


