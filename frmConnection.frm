VERSION 5.00
Begin VB.Form frmConnection 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "...::: LeechSTRO :::... Se connecter à"
   ClientHeight    =   3765
   ClientLeft      =   4680
   ClientTop       =   4260
   ClientWidth     =   5310
   Icon            =   "frmConnection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin LeechSTRO.btn btnQuitter 
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   3240
      Width           =   2295
      _ExtentX        =   4048
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
      MICON           =   "frmConnection.frx":0442
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
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   2295
      _ExtentX        =   4048
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
      MICON           =   "frmConnection.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtTimeOut 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "60"
      Top             =   1750
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Informations"
      ForeColor       =   &H00000000&
      Height          =   3550
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   5055
      Begin VB.TextBox txtRootDirectory 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   4815
      End
      Begin VB.CheckBox chkPassMode 
         Caption         =   "PASV Mode"
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
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ComboBox cboTransMode 
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
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   2295
      End
      Begin VB.OptionButton OptAnonyme 
         Appearance      =   0  'Flat
         Caption         =   "Anonyme"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   2760
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptPersonnel 
         Appearance      =   0  'Flat
         Caption         =   "Login personnel"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   4200
         TabIndex        =   2
         Text            =   "21"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2640
         TabIndex        =   4
         Text            =   "user@leechstro.com"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "anonymous"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "Dossier distant :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mode de transfert :"
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
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   15
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label Label6 
         Caption         =   "Idle Timeout (s) :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Port :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Mot de passe :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Nom d'utilisateur :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hôte ou URL :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private WithEvents m_objFtpClientConnection As CFtpClient
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                 Pour la connection au FTP
'
Private m_varFtpServer          As Variant
Private m_lngRemotePort         As Long
Private m_strUserName           As String
Private m_strPassword           As String
Private m_bPassiveMode          As Boolean
Private m_TransferMode          As FtpTransferModes
Private m_objTimeOut            As Integer
Private m_strRootDirectory      As String
Public Enum Command
    comdOK
    comdCancel
End Enum
Private mvarAction              As Command
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''               Pour l'URL sous la forme ftp://
'
Dim URL                         As String
Dim POS1                        As Integer
Dim POS2                        As Integer
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
    
  '  Set m_objFtpClientConnection = New CFtpClient
    
    ' Sert à sélectionner les textes lorsqu'on appui sur [ Tab ]
    txtURL.SelLength = Len(txtURL.Text)
    txtPort.SelLength = Len(txtPort.Text)
    txtUserName.SelLength = Len(txtUserName.Text)
    txtPassword.SelLength = Len(txtPassword.Text)
    txtTimeOut.SelLength = Len(txtTimeOut.Text)
    txtRootDirectory.SelLength = Len(txtRootDirectory.Text)

    ' Place dans la ComboBox les mots [ ASCII ] et [ Image ]
    With cboTransMode
        .AddItem "ASCII"
        .AddItem "Image"
        .ListIndex = 1
    End With

End Sub

' Si on apppui sur [entrer]
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnConnection_Click
End Sub

Private Sub txtUserName_Change()

    If txtUserName <> "anonymous" Then
        OptPersonnel.value = True
        OptAnonyme.value = False
    End If
    
End Sub

Public Property Let Action(ByVal vData As Command)
     mvarAction = vData
End Property

Public Property Get Action() As Command
' Utilisé en recherchant la valeur d'une propriété, du bon côté d'une tâche
' Syntaxe : Debug.Print X.Action
    Action = mvarAction
End Property

Public Property Let FtpServer(ByVal vData As String)
    m_varFtpServer = vData
End Property

Public Property Get FtpServer() As String
    FtpServer = m_varFtpServer
End Property

Public Property Let RemotePort(NewValue As Long)

    If NewValue < 1 Or NewValue > 65535 Then
        'Err.Raise 380, "CFtpClient.RemotePort", "Invalid property value."
        m_lngRemotePort = 21
    Else
        m_lngRemotePort = NewValue
    End If

End Property

Public Property Get RemotePort() As Long
    RemotePort = m_lngRemotePort
End Property

Public Property Let UserName(ByVal vData As String)
    m_strUserName = vData
End Property

Public Property Get UserName() As String
    UserName = m_strUserName
End Property

Public Property Let Password(ByVal vData As String)
    m_strPassword = vData
End Property

Public Property Get Password() As String
    Password = m_strPassword
End Property

Public Property Let TimeOut(ByVal intSeconds As Integer)
    m_objTimeOut = intSeconds
End Property

Public Property Get TimeOut() As Integer
    TimeOut = m_objTimeOut
End Property

Public Property Let TransferMode(NewValue As FtpTransferModes)
    m_TransferMode = NewValue
End Property

Public Property Get TransferMode() As FtpTransferModes
    TransferMode = m_TransferMode
End Property

Public Property Get RootDirectory() As String
    RootDirectory = m_strRootDirectory
End Property

Public Property Let RootDirectory(ByVal strDirectoryPath As String)
    m_strRootDirectory = strDirectoryPath
End Property

Public Property Get PassiveMode() As Boolean
    PassiveMode = m_bPassiveMode
End Property

Public Property Let PassiveMode(NewValue As Boolean)
    m_bPassiveMode = NewValue
End Property

Private Sub OptAnonyme_Click()

    txtUserName = "anonymous"
    txtPassword.PasswordChar = ""
    txtPassword = "leecher@leechtstro.com"
    
End Sub

Private Sub OptPersonnel_Click()
    txtPassword.PasswordChar = "*"
End Sub

Private Sub btnConnection_Click()

  '  Dim uAns As VbMsgBoxResult
    
  '  If m_objFtpClientConnection.FtpSessionState = ftpConnected Then
  '      uAns = MsgBox("Vous êtes encore connecté à un serveur. Se déconnecter ?", vbQuestion + vbOKCancel, "Confirmation LeechSTRO")
  '  End If
    
  '  If uAns = vbCancel Then
  '      Unload frmConnection
  '      Exit Sub
  '  Else
  '      lolmdr.CloseControlConnection
  '  End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''                  Si l'URL est du style ftp://
    '
    If LCase(Left(URL, 6)) = "ftp://" Then
        URL = txtURL.Text
        POS1 = InStr(7, URL, ":")
        txtUserName.Text = Mid$(URL, 7, POS1 - 7)
        POS2 = InStr(POS1, URL, "@")
        txtPassword.Text = Mid$(URL, POS1 + 1, POS2 - POS1 - 1)
        POS1 = InStr(POS2, URL, ":")
        txtPort.Text = Mid$(URL, POS1 + 1)
        txtURL.Text = Mid$(URL, POS2 + 1, POS1 - POS2 - 1)
    End If
    ''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Len(txtURL.Text) = 0 Then
        MsgBox "Veuillez entrer l'hôte ou l'URL auquelle vous voulez vous connectez.", vbExclamation
        Exit Sub
    Else
        m_varFtpServer = LCase(txtURL.Text)
    End If

    If Len(txtPort.Text) > 0 Then
        m_lngRemotePort = CLng(txtPort.Text)
    End If

    If Len(txtUserName.Text) = 0 Then
        m_strUserName = vbNullString
    Else
        m_strUserName = txtUserName.Text
    End If

    If Len(txtPassword.Text) = 0 Then
        m_strPassword = vbNullString
    Else
        m_strPassword = txtPassword.Text
    End If

    If chkPassMode.value = vbChecked Then
        PassiveMode = True
    Else
        PassiveMode = False
    End If

    If Len(txtTimeOut.Text) > 0 Then
        m_objTimeOut = CInt(txtTimeOut.Text)
    End If

    If Len(txtDirectoryPath) > 0 Then
        m_strRootDirectory = txtDirectoryPath.Text
    End If

    mvarAction = comdOK
    Unload Me

End Sub

Private Sub btnQuitter_Click()

    mvarAction = comdCancel
    Unload Me
    
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    mvarAction = comdCancel
'End Sub
