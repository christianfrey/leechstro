VERSION 5.00
Begin VB.Form frmQuickConnection 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "...::: LeechSTRO :::... Connection rapide à"
   ClientHeight    =   1095
   ClientLeft      =   5550
   ClientTop       =   5160
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin LeechSTRO.btn btnQuitter 
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      BTYPE           =   14
      TX              =   "Quitter"
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
      MICON           =   "frmQuickConnection.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin LeechSTRO.btn btnConnection 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      BTYPE           =   14
      TX              =   "Connection"
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
      MICON           =   "frmQuickConnection.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Hôte ou URL :"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmQuickConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Si on apppui sur [entrer]
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnConnection_Click
End Sub

Private Sub btnConnection_Click()

    If Len(txtURL.Text) = 0 Then
        MsgBox "Veuillez entrer l'hôte ou l'URL auquelle vous voulez vous connectez.", vbExclamation
        Exit Sub
    Else
    
    End If
    
End Sub

Private Sub btnQuitter_Click()

    mvarAction = comdCancel
    Unload Me
    
End Sub
