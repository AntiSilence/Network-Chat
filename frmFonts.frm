VERSION 5.00
Begin VB.Form frmFonts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFonts 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lstFonts.Width = frmFonts.Width
    lstFonts.Height = frmFonts.Height
    
    Dim i As Integer
    
    Me.Refresh
    
    Me.MousePointer = vbHourglass
    For i = 1 To Screen.FontCount - 1
        lstFonts.AddItem Screen.Fonts(i)
    Next i
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub


Private Sub Label1_Click()

End Sub

Private Sub lstFonts_Click()
    frmMain.txtInput.SelStart = 0
    frmMain.txtInput.SelLength = Len(frmMain.txtInput.TextRTF)
    frmMain.txtInput.SelFontName = lstFonts.List(lstFonts.ListIndex)
    frmMain.txtInput.SelBold = False
    frmMain.txtInput.SelFontSize = 8
    frmMain.txtInput.SelStart = Len(frmMain.txtInput.TextRTF)
    
    frmMain.rtfTemp.SelStart = 0
    frmMain.rtfTemp.SelLength = Len(frmMain.txtInput.TextRTF)
    frmMain.rtfTemp.SelFontName = lstFonts.List(lstFonts.ListIndex)
    frmMain.rtfTemp.SelBold = False
    frmMain.rtfTemp.SelFontSize = 8
    frmMain.rtfTemp.SelStart = Len(frmMain.txtInput.TextRTF)
    
    frmMain.txtSave.SelStart = 0
    frmMain.txtSave.SelLength = Len(frmMain.txtInput.TextRTF)
    frmMain.txtSave.SelFontName = lstFonts.List(lstFonts.ListIndex)
    frmMain.txtSave.SelBold = False
    frmMain.txtSave.SelFontSize = 8
    frmMain.txtSave.SelStart = Len(frmMain.txtInput.TextRTF)
    frmMain.txtOutput.SetFocus
    
    frmMain.txtOutput.Font = lstFonts.List(lstFonts.ListIndex)
    frmMain.txtOutput.Font.Size = 8
    frmMain.txtOutput.Font.Bold = False
     
    SaveSetting App.Title, "Config", "Font", lstFonts.List(lstFonts.ListIndex)
    Unload frmFonts
End Sub

Private Sub lstFonts_LostFocus()
    Unload frmFonts
End Sub


Private Sub lstFonts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.imgEm.Picture = frmMain.btnImgLst.ListImages.Item("em_up").Picture
    frmMain.imgFont.Picture = frmMain.btnImgLst.ListImages.Item("font_up").Picture
End Sub


