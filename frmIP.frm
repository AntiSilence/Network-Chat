VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Send Message to IP Address..."
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox txtIP 
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393217
      TextRTF         =   $"frmIP.frx":0000
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
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtMsg 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Send to IP Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdSend_Click()
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    
    frmMain.Winsock(0).RemoteHost = txtIP.Text
    
    If chkSpeech.Value = 0 Then
        frmMain.Winsock(0).SendData "/msg;MESSAGE from " & frmMain.Winsock(0).LocalIP & ":" & vbCrLf & "    " & txtMsg.Text
    ElseIf chkSpeech.Value = 1 Then
        frmMain.Winsock(0).SendData "/msgsp;MESSAGE from " & frmMain.Winsock(0).LocalIP & ":" & vbCrLf & "    " & txtMsg.Text
    End If
    
    frmMain.txtOutput.SetFocus
    
    Me.MousePointer = vbDefault

    Unload Me
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


Private Sub txtIP_Change()
    If Len(txtIP.Text) > 0 And Len(txtMsg.Text) > 0 Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
  If KeyAscii = Asc(".") Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtIP1_Change()

End Sub

Private Sub txtMsg_Change()
    If Len(txtIP.Text) > 0 And Len(txtMsg.Text) > 0 Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub


Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(txtMsg.Text) > 0 And Len(txtIP.Text) > 0 Then
        KeyAscii = 0
        cmdSend_Click
    End If
End Sub


