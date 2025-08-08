VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "LeechSTRO"
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2129.199
   ScaleMode       =   0  'User
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Fainéant 
      Interval        =   2000
      Left            =   3240
      Top             =   1080
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Mettre LeechSTRO dans le systray au démarrage, si l'option était coché
Private Sub Form_Load()
    If ModIniFile.GetIni("Options", "Le placer dans le systray au démarrage", App.path & "\configurations.ini") = "Oui" Then
        Call frmMain.mnuSystray_Click
    End If
End Sub

' Au bout de 2 secondes, la fenêtre frmSplash s'enlève automatiquement
Private Sub Fainéant_Timer()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    frmMain.Show
    Unload Me
End Sub
