Attribute VB_Name = "Global"
Option Explicit

' Global variables
Public numPublic As Integer
Public numPrivate As Integer
Public optMin As Boolean
Public sendDisabled As Boolean
Public lastMsgTime As String
Public saveMsgText As String
Public soundFile As String
Public rtfAvatar As String
Public displayName As String
Public sndStr As String
Public copyStr As String
Public ShowAtStartup As Long

' Global Option Flags
Public playSnd As Boolean
Public winPopUp As Boolean
Public privacy As Boolean
Public saveMsg As Boolean
Public userList As Boolean
Public useDisplayName As Boolean
Public minSysTray As Boolean
Public NotifyOff As Boolean

' THIS IS FOR THE XP-STYLE INTERFACE
Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

' THIS IS FOR THE XP-STYLE INTERFACE
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
Sub Main()
    ' Check to see if already running
    If App.PrevInstance Then
        MsgBox "Network Chat is already running.", vbExclamation, "Error..."
        End
    Else
        ' Set path to Help File
        App.HelpFile = App.Path & "\Network Chat.hlp"
        
        On Error Resume Next
        ' this will fail if Comctl not available
        '  - unlikely now though!
        Dim iccex As tagInitCommonControlsEx
        With iccex
            .lngSize = LenB(iccex)
            .lngICC = ICC_USEREX_CLASSES
        End With
        InitCommonControlsEx iccex
   
        ' now start the application
        On Error GoTo 0
        
        Debug.Print Command$
            
        If Command$ = LCase("/systray") Then
            Load frmMain
            'frmMain.Hide
            'frmMain.WindowState = vbMinimized
            ShowAtStartup = 0
            frmMain.tray.Visible = True
        Else
            ShowAtStartup = 1
            frmMain.Show
        End If
    End If
End Sub


Sub updateUserList()
    Dim filename As String
    Dim filenum As Integer
    Dim str As String
    
    ' open NETWORK.CFG file and send ENTERED message to each user
    ' Also at the same time, add user names to user list.
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    frmMain.lstUsers.Clear
    
    Open filename$ For Input As #filenum
        If LOF(filenum) = 0 Then
            Close #filenum
        Else
            Do Until EOF(filenum)
                DoEvents
                Line Input #filenum, str$
                
                If str$ = UCase$(frmMain.Winsock(0).LocalHostName) Then
                    ' ignore own computer name if user added to CFG file! Doh!
                ElseIf Len(str$) > 1 Then
                    ' add to user list
                    frmMain.lstUsers.AddItem str$
                Else
                    ' ignore blank lines
                End If
            Loop
        End If
    Close #filenum
End Sub


