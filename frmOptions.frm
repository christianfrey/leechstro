VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "...::: LeechSTRO :::... Options"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Option windows"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4215
      Begin LeechSTRO.btn btnCorbeille 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Vider la corbeille de windows"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   33023
         BCOLO           =   16576
         FCOL            =   12648447
         FCOLO           =   8454143
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmOptions.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin LeechSTRO.btn btnAnnuler 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Annuler"
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
      MICON           =   "frmOptions.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin LeechSTRO.btn btnAccepter 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Accepter"
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
      MICON           =   "frmOptions.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Option de la fenêtre"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
      Begin VB.CheckBox CheckVisible 
         Caption         =   "LeechSTRO toujours visible -> [ Always on top ]"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option de démarrage"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox CheckSystray 
         Caption         =   "Le placer dans le systray au démarrage"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox CheckDemarrage 
         Caption         =   "Lancer LeechSTRO au démarrage de Windows"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                   Pour vider la corbeille
'
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                    Pour [ ALWAYS ON TOP ]
'
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Bouton pour vider la corbeille de windows
Private Sub btnCorbeille_Click()

    ' La vide
    SHEmptyRecycleBin Me.hwnd, vbNullString, 0
    ' La met à jour
    SHUpdateRecycleBinIcon
    
End Sub

Private Sub Form_Load()

    ' Permet de mettre au premier plan frmOptions,
    ' meme quand frmMain est toujours visible
    RendreFormTjsVisible Me
    
    ' Sert à cocher ou décocher la CheckBox [ CheckDemarrage ]
    Dim ret As String
    ret = ModRegedit.getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "LeechSTRO")
    
    If ret = App.Path & "\" & App.EXEName & ".exe /tray" Then
        CheckDemarrage.value = vbChecked
    Else
        CheckDemarrage.value = 0
    End If
    
    ' Sert à cocher ou décocher la CheckBox [ CheckSystray ]
    If ModIniFile.GetIni("Options", "Le placer dans le systray au démarrage", App.Path & "\configurations.ini") = "Non" Then
        CheckSystray.value = 0
    Else
        CheckSystray.value = vbChecked
    End If

    ' Sert à cocher ou décocher la CheckBox [ CheckVisible ]
    If ModIniFile.GetIni("Options", "La fenêtre principale de LeechSTRO est toujours au premier plan", App.Path & "\configurations.ini") = "Non" Then
        CheckVisible.value = 0
    Else
        CheckVisible.value = vbChecked
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call btnAnnuler_Click
End Sub

' LeechSTRO reste au premier plan
Public Sub RendreFormTjsVisible(MonForm As Object)
     SetWindowPos MonForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

' LeechSTRO ne reste pas au premier plan
Public Sub RendreFormPasTjsVisible(MonForm As Object)
     SetWindowPos MonForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

'Si on clique sur [ CheckDemarrage ], on pourra cliquer sur [ CheckSystray ]
Private Sub CheckDemarrage_Click()

    If CheckDemarrage.value = vbChecked Then
        CheckSystray.Enabled = True
    Else
        CheckSystray.Enabled = False
        CheckSystray.value = 0
    End If
    
End Sub

Private Sub btnAccepter_Click()

    ' Option [ Lancer LeechSTRO au démarrage ]
    If CheckDemarrage.value = vbChecked Then
        ' Ecrit dans le fichier "configurations.ini" les paramètres sauvegardés
        ModIniFile.WriteIni "Options", "Lancer LeechSTRO au démarrage de windows", "Oui", App.Path & "\configurations.ini"
        ModRegedit.savestring HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "LeechSTRO", App.Path & "\" & App.EXEName & ".exe /tray"
    Else
        ' Ecrit dans le fichier "configurations.ini" les paramètres sauvegardés
        ModIniFile.WriteIni "Options", "Lancer LeechSTRO au démarrage de windows", "Non", App.Path & "\configurations.ini"
        Call ModRegedit.DeleteValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "LeechSTRO")
    End If
    
    ' Option [ Le placer dans le systray au démarrage ]
    If CheckSystray.value = 1 Then
        ModIniFile.WriteIni "Options", "Le placer dans le systray au démarrage", "Oui", App.Path & "\configurations.ini"
    Else
        ModIniFile.WriteIni "Options", "Le placer dans le systray au démarrage", "Non", App.Path & "\configurations.ini"
    End If
    
    ' Option [ Always on top ]
    If CheckVisible.value = 1 Then
        ' Ecrit dans le fichier "configurations.ini" les paramètres sauvegardés
        ModIniFile.WriteIni "Options", "La fenêtre principale de LeechSTRO est toujours au premier plan", "Oui", App.Path & "\configurations.ini"
        RendreFormTjsVisible frmMain
        Unload Me
    Else
        ' Ecrit dans le fichier "configurations.ini" les paramètres sauvegardés
        ModIniFile.WriteIni "Options", "La fenêtre principale de LeechSTRO est toujours au premier plan", "Non", App.Path & "\configurations.ini"
        RendreFormPasTjsVisible frmMain
        Unload Me
    End If
    
End Sub

' Ne pas appliquer les paramètres
Private Sub btnAnnuler_Click()
    Unload Me
End Sub
