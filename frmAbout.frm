VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A propos de LeechSTRO"
   ClientHeight    =   3600
   ClientLeft      =   4965
   ClientTop       =   4680
   ClientWidth     =   4785
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":08CA
   ScaleHeight     =   3600
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin LeechSTRO.btn btnOk 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAbout.frx":3D15
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblCourriel 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail : redacted@example.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblWeb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Site web : http://www.example.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003 Urgo, Inc."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3210
      Width           =   2415
   End
   Begin VB.Label lblProgramme 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Programmé par Urgo en Visual Basic 6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    ' Inscrit la version exacte du programme automatiquement
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

' Lorqu'on appui sur [Entrée] ou [Espace] -> la fenêtre se ferme
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub btnOk_Click()
    Unload Me
End Sub
