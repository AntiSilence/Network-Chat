VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Network Configuration..."
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Selected Port"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   4575
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblPrt 
         Caption         =   "NOTE: If you change the Port number, you must restart Network Chat before the change will take effect."
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label lblPort 
         Alignment       =   1  'Right Justify
         Caption         =   "Use Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Computers on the Network"
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.PictureBox picContainer1 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   4335
         TabIndex        =   7
         Top             =   240
         Width           =   4335
         Begin VB.CommandButton cmdGetNames 
            Caption         =   "&Get Names"
            Height          =   375
            Left            =   3240
            TabIndex        =   13
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Cl&ear All"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2040
            TabIndex        =   12
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "&Remove >>"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2040
            TabIndex        =   11
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox txtAdd 
            Height          =   285
            Left            =   2040
            TabIndex        =   10
            Top             =   600
            Width           =   2295
         End
         Begin VB.ListBox lstNames 
            Height          =   2400
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   1935
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<< &Add..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   2040
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblYours 
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   2640
            Width           =   4335
         End
         Begin VB.Label lblName 
            Caption         =   "Please type the name of a computer on your network:"
            Height          =   495
            Left            =   2040
            TabIndex        =   14
            Top             =   120
            Width           =   2295
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_PREFERRED_LENGTH As Long = -1
Private Const NERR_SUCCESS As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234&

Private Const SV_TYPE_WORKSTATION As Long = &H1
Private Const SV_TYPE_SERVER As Long = &H2
Private Const SV_TYPE_SQLSERVER As Long = &H4
Private Const SV_TYPE_DOMAIN_CTRL As Long = &H8
Private Const SV_TYPE_DOMAIN_BAKCTRL As Long = &H10
Private Const SV_TYPE_TIME_SOURCE As Long = &H20
Private Const SV_TYPE_AFP As Long = &H40
Private Const SV_TYPE_NOVELL As Long = &H80
Private Const SV_TYPE_DOMAIN_MEMBER As Long = &H100
Private Const SV_TYPE_PRINTQ_SERVER As Long = &H200
Private Const SV_TYPE_DIALIN_SERVER As Long = &H400
Private Const SV_TYPE_XENIX_SERVER As Long = &H800
Private Const SV_TYPE_SERVER_UNIX As Long = SV_TYPE_XENIX_SERVER
Private Const SV_TYPE_NT As Long = &H1000
Private Const SV_TYPE_WFW As Long = &H2000
Private Const SV_TYPE_SERVER_MFPN As Long = &H4000
Private Const SV_TYPE_SERVER_NT As Long = &H8000
Private Const SV_TYPE_POTENTIAL_BROWSER As Long = &H10000
Private Const SV_TYPE_BACKUP_BROWSER As Long = &H20000
Private Const SV_TYPE_MASTER_BROWSER As Long = &H40000
Private Const SV_TYPE_DOMAIN_MASTER As Long = &H80000
Private Const SV_TYPE_SERVER_OSF As Long = &H100000
Private Const SV_TYPE_SERVER_VMS As Long = &H200000
Private Const SV_TYPE_WINDOWS As Long = &H400000 'Windows95 and above
Private Const SV_TYPE_DFS As Long = &H800000 'Root of a DFS tree
Private Const SV_TYPE_CLUSTER_NT As Long = &H1000000 'NT Cluster
Private Const SV_TYPE_TERMINALSERVER As Long = &H2000000 'Terminal Server
Private Const SV_TYPE_DCE As Long = &H10000000 'IBM DSS
Private Const SV_TYPE_ALTERNATE_XPORT As Long = &H20000000 'rtn alternate transport
Private Const SV_TYPE_LOCAL_LIST_ONLY As Long = &H40000000 'rtn local only
Private Const SV_TYPE_DOMAIN_ENUM As Long = &H80000000
Private Const SV_TYPE_ALL As Long = &HFFFFFFFF

Private Const SV_PLATFORM_ID_OS2 As Long = 400
Private Const SV_PLATFORM_ID_NT As Long = 500

Private Const MAJOR_VERSION_MASK As Long = &HF

Private Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_name As Long
End Type

Private Declare Function NetServerEnum Lib "netapi32" _
    (ByVal servername As Long, _
    ByVal level As Long, _
    buf As Any, _
    ByVal prefmaxlen As Long, _
    entriesread As Long, _
    totalentries As Long, _
    ByVal servertype As Long, _
    ByVal domain As Long, _
    resume_handle As Long) As Long

Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetServers(sDomain As String) As Long
    Dim bufptr As Long
    Dim dwEntriesread As Long
    Dim dwTotalentries As Long
    Dim dwResumehandle As Long
    Dim se100 As SERVER_INFO_100
    Dim success As Long
    Dim nStructSize As Long
    Dim cnt As Long
    Dim tmpName As String
    
    nStructSize = LenB(se100)
    success = NetServerEnum(0&, 100, bufptr, MAX_PREFERRED_LENGTH, dwEntriesread, dwTotalentries, SV_TYPE_ALL, 0&, dwResumehandle)
    If success = NERR_SUCCESS And success <> ERROR_MORE_DATA Then
        For cnt = 0 To dwEntriesread - 1
            CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize
            tmpName = GetPointerToByteStringW(se100.sv100_name)
            
            If tmpName = UCase(frmMain.Winsock(0).LocalHostName) Then
                ' Ignore own computer name
            Else
                lstNames.AddItem GetPointerToByteStringW(se100.sv100_name)
            End If
        Next
    End If
    Call NetApiBufferFree(bufptr)
    cmdClear.Enabled = True
    txtAdd.SetFocus
    GetServers = dwEntriesread
End Function
Public Function GetPointerToByteStringW(ByVal dwData As Long) As String
    Dim tmp() As Byte
    Dim tmplen As Long
    If dwData <> 0 Then
        tmplen = lstrlenW(dwData) * 2
        If tmplen <> 0 Then
            ReDim tmp(0 To (tmplen - 1)) As Byte
            CopyMemory tmp(0), ByVal dwData, tmplen
            GetPointerToByteStringW = tmp
        End If
    End If
End Function
Private Sub cmdAdd_Click()
    On Error GoTo errChk
    
    Dim idx As Integer
    Dim lp As Integer
    
    If txtAdd.Text = frmMain.Winsock(0).LocalHostName Then
        MsgBox UCase$(txtAdd.Text) & " is your computer! You cannot add your computer" & _
                " to the list. " & vbCrLf & vbCrLf & "Messages are sent to your computer automatically.", vbExclamation, "Error - Cannot Add Your Computer..."
        txtAdd.Text = ""
        txtAdd.SetFocus
        Exit Sub
    End If
    
    ' Used to set the index of the list box items
    idx = lstNames.ListCount
    
    Me.MousePointer = vbHourglass

    ' Check to make sure name is not in list already
    For lp = 0 To idx - 1
        If UCase$(txtAdd.Text) = lstNames.List(lp) Then
            MsgBox UCase$(txtAdd.Text) & " is already in the list.", vbExclamation, "Error..."
            txtAdd.Text = ""
            txtAdd.SetFocus
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    Next lp
    
    ' Test if computer is on network
    'Dim tst As String
    frmMain.Winsock(0).RemoteHost = UCase$(txtAdd.Text)
    frmMain.Winsock(0).SendData "/cmd:test"
    
    lstNames.AddItem UCase$(txtAdd.Text), idx
    cmdClear.Enabled = True
    
    txtAdd.Text = ""
    txtAdd.SetFocus
    Me.MousePointer = vbDefault
    
    Exit Sub
    
errChk:
    Me.MousePointer = vbDefault
    MsgBox UCase$(txtAdd.Text) & " was not found on the network. The name you entered may be wrong, or " & _
            "the user may have logged off.", vbExclamation, "Error..."
    txtAdd.Text = ""
    txtAdd.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    lstNames.Clear
    txtAdd.SetFocus
    cmdRemove.Enabled = False
    cmdClear.Enabled = False
End Sub

Private Sub cmdGetNames_Click()
    lstNames.Clear
    Me.MousePointer = vbHourglass
    Call GetServers(vbNullString)
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdRemove_Click()
    lstNames.RemoveItem lstNames.ListIndex
    txtAdd.SetFocus

    If lstNames.ListCount = 0 Then
        cmdRemove.Enabled = False
    End If
End Sub

Private Sub cmdSave_Click()
    ' This saves the NETWORK.CFG file to the application path
    
    Dim nameList As String
    Dim idx As Integer
    Dim filenum As Integer
    Dim filename As String
    
    SaveSetting App.Title, "Config", "Port", CLng(txtPort.Text)
    frmMain.Winsock(0).RemotePort = CLng(txtPort.Text)
    
    ' Generate name list
    For idx = 0 To lstNames.ListCount - 1
        nameList$ = nameList$ + lstNames.List(idx) & vbCrLf
    Next idx
    
    ' Set the filename
    filename = App.Path & "\Network.cfg"
    
    ' Get free file number
    filenum = FreeFile
    Open filename$ For Output As #filenum
        Print #filenum, nameList$
    Close #filenum
    'End If
    
    Call updateUserList
    
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo errTrap
    
    Dim filenum As Integer
    Dim str As String
    Dim filename As String
    
    txtPort.Text = frmMain.Winsock(0).RemotePort
        
    ' Load NETWORK.CFG file and add to listbox
    filename$ = App.Path & "\Network.cfg"
    filenum = FreeFile
    
    lblYours.Caption = "Your computer name is " & UCase$(frmMain.Winsock(0).LocalHostName)
    
    Open filename$ For Input As #filenum
        Do Until EOF(filenum)
            Line Input #filenum, str$
            If str$ = UCase$(frmMain.Winsock(0).LocalHostName) Then
                '
            ElseIf Len(str$) > 1 Then
                lstNames.AddItem str$
            End If
        Loop
    Close #filenum
        
    If lstNames.ListCount > 0 Then
        cmdClear.Enabled = True
    End If
    
    ' Check for OS and display Get Name for NT based OS's
    If frmMain.SysInf.OSPlatform = 0 Or frmMain.SysInf.OSPlatform = 1 Then ' Windows 9x or Win32s
        cmdGetNames.Visible = False
    ElseIf frmMain.SysInf.OSPlatform = 2 Then ' Windows NT
        cmdGetNames.Visible = True
    End If
    
    Exit Sub
    
errTrap:
    If Err.Number = 53 Then
        MsgBox "Network configuration file was not found. Please add some computer names " & _
                "to your list and choose Save.", vbCritical, "Error..."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


Private Sub lstNames_Click()
    If lstNames.ListCount > 0 Then
        cmdRemove.Enabled = True
    Else
        cmdRemove.Enabled = False
    End If
End Sub

Private Sub txtAdd_Change()
    If Len(txtAdd.Text) > 0 Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End Sub


Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(txtAdd.Text) > 0 Then
        KeyAscii = 0
        cmdAdd_Click
    End If
End Sub


