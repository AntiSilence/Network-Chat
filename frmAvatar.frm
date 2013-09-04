VERSION 5.00
Begin VB.Form frmAvatar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1455
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1935
   ControlBox      =   0   'False
   HelpContextID   =   6
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgTongue 
      Height          =   225
      Left            =   1560
      MouseIcon       =   "frmAvatar.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":030A
      ToolTipText     =   "Tongue out! :P"
      Top             =   600
      Width           =   225
   End
   Begin VB.Image imgEvil 
      Height          =   225
      Left            =   1200
      MouseIcon       =   "frmAvatar.frx":0404
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":070E
      ToolTipText     =   "Evil (6)"
      Top             =   600
      Width           =   225
   End
   Begin VB.Image imgWorried 
      Height          =   225
      Left            =   840
      MouseIcon       =   "frmAvatar.frx":0808
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":0B12
      ToolTipText     =   "Worried :%"
      Top             =   600
      Width           =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   3240
      X2              =   0
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image imgBlush 
      Height          =   225
      Left            =   120
      MouseIcon       =   "frmAvatar.frx":0C0C
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":0F16
      ToolTipText     =   "Blush :$"
      Top             =   600
      Width           =   225
   End
   Begin VB.Image imgWink 
      Height          =   225
      Left            =   840
      MouseIcon       =   "frmAvatar.frx":1010
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":131A
      ToolTipText     =   "Wink ;)"
      Top             =   120
      Width           =   225
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Image imgHeart 
      Height          =   210
      Left            =   1560
      MouseIcon       =   "frmAvatar.frx":1414
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":171E
      ToolTipText     =   "Love Heart (L)"
      Top             =   1080
      Width           =   225
   End
   Begin VB.Image imgGuns 
      Height          =   240
      Left            =   765
      MouseIcon       =   "frmAvatar.frx":1810
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":1B1A
      ToolTipText     =   "Guns (%)"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image imgBeer 
      Height          =   240
      Left            =   120
      MouseIcon       =   "frmAvatar.frx":1CDC
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":1FE6
      ToolTipText     =   "Hmmm... Beer! (B)"
      Top             =   1080
      Width           =   390
   End
   Begin VB.Image imgAngry 
      Height          =   225
      Left            =   480
      MouseIcon       =   "frmAvatar.frx":2168
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":2472
      ToolTipText     =   "Angry :@"
      Top             =   600
      Width           =   225
   End
   Begin VB.Image imgShades 
      Height          =   225
      Left            =   1200
      MouseIcon       =   "frmAvatar.frx":256C
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":2876
      ToolTipText     =   "Cool 8)"
      Top             =   120
      Width           =   225
   End
   Begin VB.Image imgHappy 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   120
      MouseIcon       =   "frmAvatar.frx":2970
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":2C7A
      ToolTipText     =   "Happy :)"
      Top             =   120
      Width           =   225
   End
   Begin VB.Image imgGrin 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   480
      MouseIcon       =   "frmAvatar.frx":2D74
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":307E
      ToolTipText     =   "Big Grin :D"
      Top             =   120
      Width           =   225
   End
   Begin VB.Image imgSad 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   1560
      MouseIcon       =   "frmAvatar.frx":3178
      MousePointer    =   99  'Custom
      Picture         =   "frmAvatar.frx":3482
      ToolTipText     =   "Sad :("
      Top             =   120
      Width           =   225
   End
End
Attribute VB_Name = "frmAvatar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_LostFocus()
    Unload Me
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.imgEm.Picture = frmMain.btnImgLst.ListImages.Item("em_up").Picture
    frmMain.imgFont.Picture = frmMain.btnImgLst.ListImages.Item("font_up").Picture
End Sub

Private Sub Image1_Click()

End Sub

Private Sub imgAngry_Click()
    frmMain.txtOutput.SelRTF = ":@"
    Unload Me
End Sub

Private Sub imgBeer_Click()
    frmMain.txtOutput.SelRTF = "(B)"
    Unload Me
End Sub

Private Sub imgBlush_Click()
    frmMain.txtOutput.SelRTF = ":$"
    Unload Me
End Sub

Private Sub imgEvil_Click()
    frmMain.txtOutput.SelRTF = "(6)"
    Unload Me
End Sub

Private Sub imgGrin_Click()
    frmMain.txtOutput.SelRTF = ":D"
    Unload Me
End Sub

Private Sub imgGuns_Click()
    frmMain.txtOutput.SelRTF = "(%)"
    Unload Me
End Sub

Private Sub imgHappy_Click()
    frmMain.txtOutput.SelRTF = ":)"
    Unload Me
End Sub


Private Sub imgHeart_Click()
    frmMain.txtOutput.SelRTF = "(L)"
    Unload Me
End Sub

Private Sub imgSad_Click()
    frmMain.txtOutput.SelRTF = ":("
    Unload Me
End Sub

Private Sub imgShades_Click()
    frmMain.txtOutput.SelRTF = "8)"
    Unload Me
End Sub


Private Sub imgTongue_Click()
    frmMain.txtOutput.SelRTF = ":P"
    Unload Me
End Sub

Private Sub imgWink_Click()
    frmMain.txtOutput.SelRTF = ";)"
    Unload Me
End Sub


Private Sub imgWorried_Click()
    frmMain.txtOutput.SelRTF = ":%"
    Unload Me
End Sub


