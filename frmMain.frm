VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMain 
   Caption         =   "...::: LeechSTRO :::... Non connecté"
   ClientHeight    =   8715
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   10935
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   10935
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSeparateurV 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   50
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   10845
      TabIndex        =   25
      Top             =   2425
      Width           =   10845
   End
   Begin VB.PictureBox picSeparateurH 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5900
      Left            =   7350
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5895
      ScaleWidth      =   45
      TabIndex        =   4
      Top             =   2475
      Width           =   50
   End
   Begin MSComctlLib.ImageList imlListViews 
      Left            =   7560
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "folder_back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A4
            Key             =   "folder_next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":227E
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F58
            Key             =   "file"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListMain 
      Left            =   3360
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":474C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5266
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":689A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolbarMain 
      Height          =   540
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   953
      ButtonWidth     =   979
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageListMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connection"
            Object.ToolTipText     =   "Connection"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FastConnection"
            Object.ToolTipText     =   "Connection Rapide"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BookMarks"
            Object.ToolTipText     =   "BookMarks"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AbortLoginProcess"
            Object.ToolTipText     =   "Arrêter L'Opération En Cours"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Disconnect"
            Object.ToolTipText     =   "Se Déconnecter"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Quit"
            Object.ToolTipText     =   "Quitter"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RtbFTP 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   3201
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":7ECE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   8445
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6667
            MinWidth        =   6667
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6253
            MinWidth        =   6253
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6253
            MinWidth        =   6253
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvLocal 
      Height          =   5925
      Left            =   3840
      TabIndex        =   2
      Top             =   2475
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   10451
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgLst"
      SmallIcons      =   "imlListViews"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nom"
         Object.Width           =   5327
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Taille"
         Object.Width           =   1976
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView lvFTP 
      Height          =   5925
      Left            =   7395
      TabIndex        =   3
      Top             =   2475
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   10451
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlListViews"
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nom"
         Object.Width           =   5327
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Taille"
         Object.Width           =   1976
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Autorisations"
         Object.Width           =   2540
      EndProperty
   End
   Begin TabDlg.SSTab SSTabTransferts 
      Height          =   5925
      Left            =   0
      TabIndex        =   6
      Top             =   2475
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   10451
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "File d'attente"
      TabPicture(0)   =   "frmMain.frx":7F4B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lNbFilesQueue"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvQueue"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ImgListQueue"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ToolbarQueue"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TimerQueue"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Threads"
      TabPicture(1)   =   "frmMain.frx":7F67
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "UpDown1"
      Tab(1).Control(1)=   "txtThreads"
      Tab(1).Control(2)=   "FrameConnection"
      Tab(1).Control(3)=   "FrameThread1"
      Tab(1).Control(4)=   "lMaxThreads"
      Tab(1).ControlCount=   5
      Begin ComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   -74400
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   327681
         Value           =   2
         BuddyControl    =   "txtThreads"
         BuddyDispid     =   196611
         OrigLeft        =   600
         OrigTop         =   360
         OrigRight       =   855
         OrigBottom      =   675
         Max             =   16
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtThreads 
         Height          =   310
         Left            =   -74925
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "2"
         Top             =   360
         Width           =   525
      End
      Begin VB.Frame FrameConnection 
         Height          =   615
         Left            =   -74950
         TabIndex        =   13
         Top             =   720
         Width           =   3675
         Begin VB.PictureBox PicConnection 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   50
            Picture         =   "frmMain.frx":7F83
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   14
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lIp 
            Caption         =   "[ThreadPrincipal] Hôte :"
            Height          =   400
            Left            =   330
            TabIndex        =   15
            Top             =   120
            Width           =   3165
         End
      End
      Begin VB.Timer TimerQueue 
         Interval        =   100
         Left            =   1560
         Top             =   1320
      End
      Begin MSComctlLib.Toolbar ToolbarQueue 
         Height          =   330
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         ButtonWidth     =   529
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImgListQueue"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Pause"
               Object.ToolTipText     =   "Mettre En Pause La Queue Pour Le Transfert"
               ImageIndex      =   3
               Style           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EffacerList"
               Object.ToolTipText     =   "Effacer La Queue"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImgListQueue 
         Left            =   2400
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":82C5
               Key             =   "download"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8597
               Key             =   "upload"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8869
               Key             =   "pause"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8B8B
               Key             =   "effacer"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvQueue 
         Height          =   5100
         Left            =   75
         TabIndex        =   7
         Top             =   720
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   8996
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImgListQueue"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fichier"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hôte"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Taille (bytes)"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame FrameThread1 
         Height          =   2415
         Left            =   -74950
         TabIndex        =   9
         Top             =   1320
         Width           =   3675
         Begin VB.PictureBox PicThread1DL 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   50
            Picture         =   "frmMain.frx":8E5D
            ScaleHeight     =   375
            ScaleWidth      =   330
            TabIndex        =   17
            Top             =   720
            Width           =   330
         End
         Begin VB.PictureBox PicThread1UL 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   50
            Picture         =   "frmMain.frx":9257
            ScaleHeight     =   375
            ScaleWidth      =   330
            TabIndex        =   18
            Top             =   720
            Width           =   330
         End
         Begin MSComctlLib.ProgressBar ProgressBarTransfert 
            Height          =   255
            Left            =   330
            TabIndex        =   10
            Top             =   200
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lTempsRestant 
            Caption         =   "Temps restant:"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label lPourcentage 
            Caption         =   "Pourcentage:"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label lVitesse 
            Caption         =   "Vitesse;"
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label lIp2 
            Caption         =   "Hôte: "
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   490
            Width           =   3255
         End
         Begin VB.Label lInfosThread1 
            Caption         =   "InfosThread1"
            Height          =   735
            Left            =   360
            TabIndex        =   16
            Top             =   730
            Width           =   3255
         End
      End
      Begin VB.Label lNbFilesQueue 
         Caption         =   "Fichiers Dans La Queue: 0"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label lMaxThreads 
         Caption         =   "threads max (pour l'instant)"
         Height          =   255
         Left            =   -74040
         TabIndex        =   8
         Top             =   420
         Width           =   1935
      End
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connection"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuQuickConnect 
         Caption         =   "C&onnection rapide"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "&Téléchager le fichier"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "&Uploader le fichier"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "&Mettre dans le systray"
         Shortcut        =   {F9}
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCloseWin 
         Caption         =   "&Fermer windows"
         Shortcut        =   {F11}
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Affichage"
      Begin VB.Menu mnuViewLocal 
         Caption         =   "&Local"
         Begin VB.Menu mnuViewlvLocal 
            Caption         =   "Grandes icônes"
            Index           =   0
         End
         Begin VB.Menu mnuViewlvLocal 
            Caption         =   "Petites icônes"
            Index           =   1
         End
         Begin VB.Menu mnuViewlvLocal 
            Caption         =   "Liste"
            Index           =   2
         End
         Begin VB.Menu mnuViewlvLocal 
            Caption         =   "Liste Détaillé"
            Index           =   3
         End
      End
      Begin VB.Menu mnuViewFtp 
         Caption         =   "&Ftp"
         Begin VB.Menu mnuViewlvFTP 
            Caption         =   "Grandes icônes"
            Index           =   0
         End
         Begin VB.Menu mnuViewlvFTP 
            Caption         =   "Petites icônes"
            Index           =   1
         End
         Begin VB.Menu mnuViewlvFTP 
            Caption         =   "Liste"
            Index           =   2
         End
         Begin VB.Menu mnuViewlvFTP 
            Caption         =   "Liste Détaillé"
            Index           =   3
         End
      End
   End
   Begin VB.Menu zLocal 
      Caption         =   "&Local"
      Begin VB.Menu zLocalUpload 
         Caption         =   "Envoyer le fichier"
         Shortcut        =   ^T
      End
      Begin VB.Menu zLocalSep1 
         Caption         =   "-"
      End
      Begin VB.Menu zLocalOpen 
         Caption         =   "Ouvrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu zLocalEdit 
         Caption         =   "Editer dans le bloc-notes"
         Shortcut        =   ^E
      End
      Begin VB.Menu zLocalDelete 
         Caption         =   "Effacer"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu zLocalRename 
         Caption         =   "Renommer"
         Shortcut        =   {F2}
      End
      Begin VB.Menu zLocalMakeDir 
         Caption         =   "Créer un dossier"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu zLocalChangeDir 
         Caption         =   "Changer de dossier"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu zLocalSep2 
         Caption         =   "-"
      End
      Begin VB.Menu zLocalChangeDrive 
         Caption         =   "Changer de disque"
         Begin VB.Menu zLocalChangeDriveX 
            Caption         =   "X:\"
            Index           =   0
         End
      End
      Begin VB.Menu zLocalSep3 
         Caption         =   "-"
      End
      Begin VB.Menu zLocalInformation 
         Caption         =   "Informations"
      End
      Begin VB.Menu zLocalSep4 
         Caption         =   "-"
      End
      Begin VB.Menu zLocalProperties 
         Caption         =   "Propriétés"
      End
   End
   Begin VB.Menu zDistant 
      Caption         =   "&Distant"
      Begin VB.Menu zDistantDownload 
         Caption         =   "Télécharger le fichier"
      End
      Begin VB.Menu zDistantDelete 
         Caption         =   "Effacer les objets sélectionnés"
      End
      Begin VB.Menu zDistantRename 
         Caption         =   "Renommer le fichier/dossier"
      End
      Begin VB.Menu zDistantCreateFolder 
         Caption         =   "Créer un dossier"
      End
      Begin VB.Menu zDistantSep1 
         Caption         =   "-"
      End
      Begin VB.Menu zDistantChangeFolder 
         Caption         =   "Changer de dossier"
      End
      Begin VB.Menu zDistantSep2 
         Caption         =   "-"
      End
      Begin VB.Menu zDistantInformation 
         Caption         =   "Informations sur le dossier"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Aide"
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Page web de LeechSTRO"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&A propos de.."
      End
   End
   Begin VB.Menu zSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu zSystrayAbout 
         Caption         =   "A propos de.."
      End
      Begin VB.Menu zSystraySep1 
         Caption         =   "-"
      End
      Begin VB.Menu zSystrayOpen 
         Caption         =   "Ouvrir"
      End
      Begin VB.Menu zSystraySep2 
         Caption         =   "-"
      End
      Begin VB.Menu zSystrayClose 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ************************************************************
'    ..........................................................
'     Nom De L'Application : LeechSTRO
'     Développeur/Programmeur: Urgo (redacted@example.com)
'    ..........................................................
'     Merci à www.vbip.com pour ses modules et modules de classes
'     présents dans "FtpClient Library Example - Tutorial 6"
'    ..........................................................
'     Note: Soyez sûr que :               |   Et que :
'     - "scrrun.dll"                      |   - "COMDLG32.OCX"
'     - "SHELL32.dll"                     |   - "RICHTX32.OCX"
'     - "stdole2.tlb"                     |   - "MSCOMCTL.OCX"
'     - "VB6.OLB"                         |   - "TABCTL32.OCX"
'     - "msvbvm.dll"                      |   - "COMCT332.OCX"
'     - "msvbvm.dll\3"                    |   - "MSCOMCTL.OCX"
'     sont chargés dans Projet/Références |   sont chargés dans Projet/Composants
'    (Runtime References)                 |   (Design Time References)
'    ..........................................................
'   ************************************************************
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                    Variables pour la ListView du FTP (lvFTP)
'
Private WithEvents m_objFtpClient As CFtpClient
Attribute m_objFtpClient.VB_VarHelpID = -1
Private m_strRootDirectory As String
Private m_strFileName As String
Private m_lngFileSize As Long
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''       Pour les propriétés de la ListView locale (Private Sub zLocalProperties_Click)
'
Const SW_SHOWNORMAL = 1
Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long
Dim lsItem As ListItem
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''                    Pour changer la couleur de la ProgressBar
'
' Déclaration des constantes permettant d'agir sur la couleur de la ProgressBar
Const WM_USER = &H400
Const PBM_SETBARCOLOR = (WM_USER + 9)
''
' Déclaration des constantes permettant d'agir sur la couleur de fond de la ProgressBar
Const CCM_FIRST = &H2000
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
''
' Déclaration de l'API permettant d'appliquer les couleurs sur la ProgressBar
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
' Pour ouvrir un fichier et afficher une page web (deuxième solution)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' Pour savoir la vitesse de transfert en Kbps
Private BeginTransfer                   As Single
Private TransferRate                    As Single
' Pour ouvrir une page web (deuxième solution)
Private Declare Function GetActiveWindow Lib "user32" () As Long
' Pour changer de disque dur dans la ListView locale sans passer par un DriveListBox
Private MenuIndex As Long
' Variable pour stocker le Path de la ListView locale (lvLocal)
Dim CurPath As String
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Affiche la fenêtre pour fermer windows
Private Sub mnuCloseWin_Click()

    AppActivate ("Program Manager")
    ' Appui sur les touches ALT-F4
    SendKeys "%{F4}"
    
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

' Sert à placer LeechSTRO dans le systray
Public Sub mnuSystray_Click()

    frmMain.Hide
    ToolbarMain.Visible = False
    
    Systray.cbSize = Len(Systray)
    Systray.hwnd = Me.hwnd
    Systray.uId = vbNull
    Systray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Systray.uCallBackMessage = WM_MOUSEMOVE
    Systray.hIcon = Me.Icon
    Systray.szTip = "LeechSTRO" & vbNullChar

    Call Shell_NotifyIcon(NIM_ADD, Systray)
    Call Shell_NotifyIcon(NIM_MODIFY, Systray)
    
End Sub

Private Sub mnuViewlvLocal_Click(Index As Integer)
    
    Select Case Index
        Case 0
            lvLocal.View = lvwIcon
        Case 1
            lvLocal.View = lvwSmallIcon
        Case 2
            lvLocal.View = lvwList
        Case 3
            lvLocal.View = lvwReport
    End Select
    
End Sub

Private Sub mnuViewlvFTP_Click(Index As Integer)
    
    Select Case Index
        Case 0
            lvFTP.View = lvwIcon
        Case 1
            lvFTP.View = lvwSmallIcon
        Case 2
            lvFTP.View = lvwList
        Case 3
            lvFTP.View = lvwReport
    End Select
    
End Sub

' Permet de changer de disque sans passer par un DriveListBox
Private Sub zLocalChangeDriveX_Click(Index As Integer)

    CurPath = zLocalChangeDriveX(Index).Caption
    LoadLocalList CurPath
    
End Sub

' Permet d'obtenir les informations concernant le répertoire...
Private Sub zLocalInformation_Click()

    frmInformation.lblPath = Me.sbStatusBar.Panels(2).Text
    frmInformation.Show
    
End Sub

' Permet d'ouvrir un fichier ou un dossier dans la ListView locale
Private Sub zLocalOpen_Click()
    
    ShellExecute 0, vbNullString, CurPath & lvLocal.SelectedItem.Text, vbNullString, CurPath, SW_SHOWNORMAL

End Sub

' Permet d'ouvrir le fichier sélectionné de la ListView locale dans Notepad
Private Sub zLocalEdit_Click()

    Shell "Notepad " & Me.lvLocal.SelectedItem, vbNormalFocus
    
End Sub

' Permet d'obtenir les propriétés d'un fichier ou dossier de la ListView locale
Private Sub zLocalProperties_Click()

    Dim shInfo As SHELLEXECUTEINFO
    
    If lvLocal.SelectedItem Is Nothing Then
        MsgBox "Propriétés de quoi?"
        Exit Sub
    End If

    Set lsItem = lvLocal.SelectedItem
        With shInfo
            .cbSize = LenB(shInfo)
            .lpFile = CurPath & lsItem.Text
            .nShow = SW_SHOW
            .fMask = SEE_MASK_INVOKEIDLIST
            .lpVerb = "properties"
        End With
        ShellExecuteEx shInfo
'    frmInfo.Show vbModal, frmMain
'Dim Total, Libre As Long
'lpRootPathName = drvChange.Drive
'lpRootPathName = Left$(lpRootPathName, 2) & "\"
'Resultat = GetDiskFreeSpaceEx(lpRootPathName, lpFreeBytesAvailableToCaller, lpTotalNumberOfBytes, lpTotalNumberOfFreeBytes)
''Convertion des valeurs LARGE_INTEGER
'Total = CLargeInt(lpTotalNumberOfBytes.lowpart, lpTotalNumberOfBytes.highpart)
'Libre = CLargeInt(lpTotalNumberOfFreeBytes.lowpart, lpTotalNumberOfFreeBytes.highpart)
'MsgBox "Espace total sur " & lpRootPathName & " : " & Total & " octets (" & Format(Total / 1024 ^ 3, "0.00") & " Go) " & vbCr & "Espace libre sur " & lpRootPathName & " : " & Libre & " octets (" & Format(Libre / 1024 ^ 3, "0.00") & " Go) ", vbOKOnly, "Infos..."
End Sub

Private Sub lvLocal_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Si on fait un click droit sur la ListView locale -> on affiche le menu
    If Button = 2 Then
        ' Le Texte [ Ouvrir ] est affiché en gras
        Me.PopupMenu zLocal, 0, , , zLocalOpen
    End If
    
End Sub

Private Sub lvFTP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Si on fait un click droit sur la ListView FTP
    If Button = 2 Then
        ' Si on est pas connecté au ftp -> on quitte
        If lvFTP.BackColor = &H8000000F Then Exit Sub
        ' On affiche le menu
        Me.PopupMenu zDistant
    End If
    
End Sub

' Supprimer un dossier ou fichier dans la ListView locale
Private Sub zLocalDelete_Click()
  
On Error GoTo ErrHnd
  
    Dim uAns As VbMsgBoxResult
    Dim oLV As ListView
    Dim svFileFolderType As String
    Dim lErrCnt As Long
  
    Set oLV = Me.lvLocal

    ' Si la taille de l'item sélectionné est égale à zéro -> alors c'est un dossier
    If oLV.SelectedItem.SubItems(1) = "0" Then
        svFileFolderType = "dossier"
    Else
        svFileFolderType = "fichier"
    End If

    uAns = MsgBox("Etes-vous sûr de vouloir supprimer le " & svFileFolderType & " """ & oLV.SelectedItem & """ ?", vbQuestion + vbYesNo, "Suppression de " & svFileFolderType)
    If uAns = vbYes Then
        Me.MousePointer = vbHourglass
        If svFileFolderType = "dossier" Then
            RmDir CurPath & oLV.SelectedItem.Text
        Else
            Kill CurPath & oLV.SelectedItem.Text
        End If
        LoadLocalList CurPath
        Me.MousePointer = vbDefault
    End If

    If lErrCnt > 0 Then
        MsgBox "Le fichier ou dossier sélectionné ne peut pas être supprimé." & vbCrLf & vbCrLf & _
        "Les fichiers en cours d'utilisation, et les dossiers qui ne sont pas vides, ne peuvent pas être supprimés.", vbExclamation, "...::: LeechSTRO :::..."
    End If
    Exit Sub

ErrHnd:
  lErrCnt = lErrCnt + 1
  Resume Next
  
End Sub

' Renommer un fichier ou un dossier dans la ListView locale
Private Sub zLocalRename_Click()

    Dim svAns As String
  
    If Len(Me.lvLocal.SelectedItem) > 0 Then
    svAns = InputBox("Nouveau nom :", "Renommer", Me.lvLocal.SelectedItem)
    If Len(svAns) > 0 Then
        Name Me.lvLocal.SelectedItem As svAns
        ' Rafraîchi corectement la ListView locale
        LoadLocalList CurPath
    End If
    End If

End Sub

' Pour créer un dossier dans la ListView locale
Private Sub zLocalMakeDir_Click()
  
    Dim svAns As String
    
    svAns = InputBox("Nom du dossier :", "Créer un nouveau dossier")
    If Len(svAns) > 0 Then
        ' Créé le dossier
        MkDir CurPath & svAns
        ' Rafraîchi corectement la ListView locale
        LoadLocalList CurPath
    End If

End Sub

Private Sub HandleErrors(plErr As Long, psvSource As String, psvDescription As String)
  
    Dim lErr As Long
    Dim uAns As VbMsgBoxResult
  
    If plErr < 0 Or plErr > 65535 Then
        lErr = plErr - vbObjectError
    Else
        lErr = plErr
    End If
    
    Select Case lErr
        Case 12031        ' connection reset
        Case Else
            uAns = MsgBox("L'erreur suivante s'est produite. Choisissez [Annuler] pour quitter LeechSTRO." & vbCrLf & vbCrLf & _
            "Numéro de l'erreur: " & lErr & vbCrLf & _
            "Source de l'erreur: " & psvSource & vbCrLf & _
            "Description de l'erreur: " & psvDescription, _
            vbCritical + vbOKCancel, "Erreur LeechSTRO")
    End Select
  
    If uAns = vbCancel Then
        Unload frmMain
    End If
    
End Sub

' Renommer un fichier/dossier avec [F2]
Private Sub lvFTP_BeforeLabelEdit(Cancel As Integer)
  
On Error GoTo ErrHnd

  '  msvOldName = Me.lvFTP.SelectedItem.Text
  '  If msvOldName = ".." Then
  '      msvOldName = ""
  '      Cancel = True
  '  End If
  '  Exit Sub

ErrHnd:
    HandleErrors Err.Number, Err.Source, Err.Description
  
End Sub

' Add delete, rename and nav keys to listview
Private Sub lvLocal_BeforeLabelEdit(Cancel As Integer)
      
   ' msvOldName = Me.lvLocal.SelectedItem.Text
   ' If msvOldName = ".." Then
   '     msvOldName = ""
   '     Cancel = True
   ' End If
    
End Sub

Private Sub lvLocal_AfterLabelEdit(Cancel As Integer, NewString As String)

On Error GoTo ErrHnd

  '  Name msvOldName As NewString
  '  Me.lvLocal.SelectedItem.Key = NewString
  '  Exit Sub

ErrHnd:
    Cancel = True
    HandleErrors Err.Number, Err.Source, Err.Description
    
End Sub

' Si on apppui sur [entrer] -> on rentre dans le dossier
Private Sub lvLocal_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then lvLocal_DblClick
End Sub

' Ajouter supprimer, renommer and nav keys to listview
Private Sub lvLocal_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 And Shift = 0 Then
        KeyCode = 0
        zLocalDelete_Click              ' Supprimer
    ElseIf KeyCode = 113 And Shift = 0 Then
        KeyCode = 0
        lvLocal.StartLabelEdit      ' F2
    ElseIf KeyCode = 8 And Shift = 0 Then ' Backspace
        KeyCode = 0
        On Error Resume Next
        Me.MousePointer = vbHourglass
        ChDir ".."
 '   If Err.Number = 0 Then RefreshLocalLists
    '    Me.MousePointer = vbDefault
   '     On Error GoTo 0
    End If
  
End Sub

' Interdit d'éditer, d'envoyer, de supprimer, de renommer, etc...dans la ListView locale
Private Sub lvLocal_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Not (Item Is Nothing) Then Me.zLocalUpload.Enabled = (Item.SubItems(1) <> "0" And Item.Text <> "Répertoire Précédent")
    If Not (Item Is Nothing) Then Me.zLocalOpen.Enabled = (Item.Text <> "Répertoire Précédent")
    If Not (Item Is Nothing) Then Me.zLocalEdit.Enabled = (Item.SubItems(1) <> "0" And Item.Text <> "Répertoire Précédent")
    If Not (Item Is Nothing) Then Me.zLocalDelete.Enabled = (Item.Text <> "Répertoire Précédent")
    If Not (Item Is Nothing) Then Me.zLocalRename.Enabled = (Item.Text <> "Répertoire Précédent")
    If Not (Item Is Nothing) Then Me.zLocalInformation.Enabled = (Item.Text <> "Répertoire Précédent")
    If Not (Item Is Nothing) Then Me.zLocalProperties.Enabled = (Item.Text <> "Répertoire Précédent")
    
End Sub

' Changer de répertoire en faisant un double click
Private Sub lvLocal_DblClick()

    ' Si on double-clique pour revenir en arrière
    If lvLocal.SelectedItem = "Répertoire Précédent" Then
        ' Si y'a un "\" à la fin du Path alors le path courrant est le meme sans "\"
        If Right(CurPath, 1) = "\" Then CurPath = Left$(CurPath, Len(CurPath) - 1)
        Call GetRightWord(CurPath, "\", True)
        LoadLocalList CurPath
        
    ' Si on double-clique sur un fichier
    ElseIf lvLocal.SelectedItem.Tag = "file" Then
        ' METTRE CODE ICI
        
    ' Si on double-clique sur un dossier
    Else
        CurPath = CurPath & lvLocal.SelectedItem.Text & "\"
        LoadLocalList CurPath
    End If

End Sub

Private Sub Form_Load()

    '
    Me.sbStatusBar.Panels(3).Text = "Non Connecté"
    
    ' Affiche le tab de la file d'attente
    SSTabTransferts.Tab = 0

    ' Affiche le fond de la ProgressBar en blanc
    SendMessage ProgressBarTransfert.hwnd, PBM_SETBKCOLOR, 0, ByVal RGB(255, 255, 255)

    ' Les threads ne sont pas visibles
    FrameConnection.Visible = False
    FrameThread1.Visible = False

    Dim ctl As Control

    ' Utile pour frmConnection
    Set m_objFtpClient = New CFtpClient

    'Clear design time values


    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is Label Then
            If Left$(ctl.name, 3) = "lbl" Then
                ctl.Caption = ""
            End If
        End If
    Next

    ' Si l'exe est situé à la racine du disque dur -> on fait rien
    If Right(App.path, 2) = ":\" Then
        CurPath = App.path
    ' S'il ne l'est pas -> on rajoute un antislash
    Else
        CurPath = App.path & "\"
    End If
    
    LoadLocalList CurPath
    Exit Sub

Form_LoadErr:
  HandleErrors Err.Number, Err.Source, Err.Description
End Sub

Private Sub mnuQuickConnect_Click()
    frmQuickConnection.Show vbModal, frmMain
End Sub

Private Sub mnuHelpAbout_Click()

On Error Resume Next

    frmAbout.Show vbModal, frmMain
    
End Sub

Private Sub mnuConnect_Click()
    Call EstablishConnection
End Sub

Private Sub mnuQuitter_Click()
    End
End Sub

Private Sub mnuHelpWeb_Click()

    ' Permet d'ouvrir IE dans une nouvelle fenêtre et non pas dans celle en cours
    If Dir("C:\program files\internet explorer\iexplore.exe") <> vbNullString Then
        Shell ("C:\program files\internet explorer\iexplore http://www.urgo.fr.tc"), 1
    Else
        ShellExecute GetActiveWindow(), "Open", "http://www.urgo.fr.tc", "", 0&, 1
    End If
    
End Sub

Private Sub mnuFileClose_Click()
    Unload Me
End Sub

' EN COURS DE CREATION !!!
Private Sub picSeparateurV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbLeftButton Then
        picSeparateurV.BackColor = vbBlack
        picSeparateurV.Top = picSeparateurV.Top + y
    End If
    
End Sub

' EN COURS DE CREATION !!!
Private Sub picSeparateurV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        If picSeparateurV.Top < 650 Then picSeparateurV.Top = 650
            If picSeparateurV.Top > frmMain.Height - 4000 Then picSeparateurV.Top = frmMain.Height - 4000
                picSeparateurV.BackColor = &H8000000F
                
                RtbFTP.Height = picSeparateurV.Top - 610
                
                SSTabTransferts.Top = picSeparateurV.Top + 50
                lvLocal.Top = SSTabTransferts.Top
                picSeparateurH.Top = SSTabTransferts.Top
                lvFTP.Top = SSTabTransferts.Top
                
                SSTabTransferts.Height = Me.ScaleHeight - picSeparateurV.Top
                lvLocal.Height = Me.ScaleHeight - picSeparateurV.Top
                picSeparateurH.Height = Me.ScaleHeight - picSeparateurV.Top
                lvFTP.Height = Me.ScaleHeight - picSeparateurV.Top
        End If
    
End Sub

Private Sub picSeparateurH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        picSeparateurH.BackColor = vbBlack
        picSeparateurH.Left = picSeparateurH.Left + x
    End If

End Sub

Private Sub picSeparateurH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        If picSeparateurH.Left < 5800 Then picSeparateurH.Left = 5800
            If picSeparateurH.Left > frmMain.width - 2500 Then picSeparateurH.Left = frmMain.width - 2500
                picSeparateurH.BackColor = &H8000000F
                lvLocal.width = picSeparateurH.Left - SSTabTransferts.width - 60
             '   lblLocalPath.width = picSeparateurH.Left - SSTabTransferts.width - 60
                lvFTP.Left = picSeparateurH.Left + 60
                lvFTP.width = frmMain.ScaleWidth - lvFTP.Left
              '  lblFTPPath.Left = picSeparateurH.Left + 60
              '  lblFTPPath.width = frmMain.ScaleWidth - lvFTP.Left
                With sbStatusBar
                    .Panels(1).width = SSTabTransferts.width
                    .Panels(2).width = lvLocal.width + 50
                    .Panels(3).width = lvFTP.width + 50
                End With
        End If
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    ' Pour la largeur
    If Me.width < 8000 Then Me.width = 8000
    If Me.Height < 8000 Then Me.Height = 8000
    RtbFTP.width = Me.ScaleWidth
    picSeparateurV.width = RtbFTP.width
    lvLocal.width = ((Me.ScaleWidth - SSTabTransferts.width) / 2) - 60
    picSeparateurH.Left = lvLocal.Left + lvLocal.width + 5
    lvFTP.Left = lvLocal.Left + lvLocal.width + 60
    lvFTP.width = lvLocal.width
    
    ' Pour la hauteur
    SSTabTransferts.Height = Me.ScaleHeight - SSTabTransferts.Top - 330
    lvQueue.Height = SSTabTransferts.Height - lvQueue.Top - 105
    lvLocal.Height = Me.ScaleHeight - lvLocal.Top - 330
    lvFTP.Height = lvLocal.Height
    picSeparateurH.Height = lvLocal.Height
    
    ' Pour la StatusBar
    With sbStatusBar
        .Panels(1).width = SSTabTransferts.width
        .Panels(2).width = lvLocal.width + 50
        .Panels(3).width = lvFTP.width + 50
    End With
    
    ' On rafraîchi la ListView locale et FTP
    lvLocal.Refresh
    lvFTP.Refresh
    
End Sub

Private Sub ToolbarMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    Case "Connection"
        Call EstablishConnection
    Case "FastConnection"
        frmQuickConnection.Show vbModal, frmMain
    Case "AbortLoginProcess"
        Call AbortLoginProcess_Click
    Case "Disconnect"
        Call cmdCloseConnection_Click
    Case "Quit"
        End
    End Select
    
End Sub

Private Sub AbortLoginProcess_Click()

    If m_objFtpClient.Busy Then
        m_objFtpClient.CancelAsyncMethod
    End If
        
End Sub

Private Sub cmdCloseConnection_Click()

On Error GoTo ERORR_HANDLER

    m_objFtpClient.CloseControlConnection

    ' Efface le ListView FTP
    lvFTP.ListItems.Clear
    ' Efface le path du FTP dans le panneau 3 de la barre de statut
    Me.sbStatusBar.Panels(3).Text = "Déconnecté"
    ' Pour que la couleur de fond de la lvFTP soit en gris (sommet de bouton)
  '  lvFTP.BackColor = -2147483633
    lvFTP.BackColor = &H8000000F
    ' Remplace le titre de frmMain
    frmMain.Caption = "...::: LeechSTRO :::... Non connecté"
    ' Les threads deviennent invisibles
    FrameConnection.Visible = False
    FrameThread1.Visible = False

    Exit Sub

ERORR_HANDLER:
    With Err
        MsgBox "Error = " & .Number & vbCrLf & .Description, vbExclamation
    End With

End Sub

Private Sub EstablishConnection()

    ' F = frmConnection
    Dim F As New frmConnection

    ' Montrez la form
    F.Show vbModal

    ' Si le bouton OK a été cliqué
    If F.Action = comdOK Then

        ' Pour que la couleur de fond de la lvFTP soit en blanc
        lvFTP.BackColor = vbWhite
        ' Efface le ListView FTP
        lvFTP.ListItems.Clear
        ' Efface le path du FTP dans le panneau 3 de la barre de statut
        Me.sbStatusBar.Panels(3).Text = "Connection en cours..."

        With m_objFtpClient

            ' Inicialiser les propriétés objects
            .FtpServer = F.FtpServer
            .RemotePort = F.RemotePort
            .UserName = F.UserName
            .Password = F.Password
            .TransferMode = F.TransferMode
            .PassiveMode = F.PassiveMode
            .TimeOut = F.TimeOut
      '      .SetCurrentDirectory = F.RootDirectory

            ' Se connecter
            .Connect
             
            ' Ajoute dans le caption de frmMain le login, l'ip et le port auquelle on est connecté
            frmMain.Caption = "...::: LeechSTRO :::... " & .UserName & " @ " & .FtpServer & " : " & .RemotePort

            FrameConnection.Visible = True

        End With
    End If
End Sub

Private Sub cmdQuitSession_Click()

On Error GoTo ERORR_HANDLER

    If Not m_objFtpClient.Busy Then
        m_objFtpClient.QuitSession
    End If

    Exit Sub

ERORR_HANDLER:
    With Err
        MsgBox "Error = " & .Number & vbCrLf & .Description, vbExclamation
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_objFtpClient = Nothing
End Sub

Private Sub lvFTP_DblClick()

    Dim strDir      As String
    Dim intSlashPos As Integer

    ' S'il n'y a aucun item sélectionné dans la ListView -> on quitte
    If lvFTP.SelectedItem Is Nothing Then
        Exit Sub
    End If

    ' Si l'item sélectionnée n'est pas un dossier -> on quitte
    If lvFTP.SelectedItem.SmallIcon = 2 Or lvFTP.SelectedItem.SmallIcon = 4 Then
        Exit Sub
    End If

    ' Si on est occupé avec quelque chose d'autre -> on quitte
    If m_objFtpClient.Busy Then
        Exit Sub
    End If

    Select Case lvFTP.SelectedItem.SmallIcon
        ' Pour revenir en arrière
        Case 1
            strDir = m_objFtpClient.CurrentDirectory
            intSlashPos = InStrRev(strDir, "/")
            If intSlashPos > 1 Then
                strDir = Left$(strDir, intSlashPos - 1)
            Else
                strDir = "/"
            End If
        ' Pour rentrer dans le dossier
        Case 2
            strDir = lvFTP.SelectedItem.Text
        Case 3
            strDir = lvFTP.SelectedItem.Text
    End Select

    m_objFtpClient.SetCurrentDirectory strDir
    lvFTP.ListItems.Clear

    ' Sert à mettre le path du FTP dans le panneau 3 de la barre de statut
    Me.sbStatusBar.Panels(3).Text = m_objFtpClient.CurrentDirectory

End Sub

Private Sub m_objFtpClient_OnConnect(ByVal AsyncResultStatus As AsyncResultStatusConstants)

    m_strRootDirectory = m_objFtpClient.CurrentDirectory
    Call DisplayEventInfo("OnConnect", AsyncResultStatus)
    m_objFtpClient.EnumFiles

    ' Sert à mettre le path du FTP dans le panneau 3 de la barre de statut
    Me.sbStatusBar.Panels(3).Text = m_objFtpClient.CurrentDirectory

End Sub

Private Sub m_objFtpClient_OnDataTransferProgress(ByVal lngBytesTransferred As Long)

On Error Resume Next

    Dim lTotalBytes As Long ' Pour le pourcentage
    Dim j As Long ' Pour le pourcentage

    ' Pour savoir la vitesse de tranfert
    TransferRate = Format(Int(lngBytesTransferred / (Timer - BeginTransfer)) / 1024, "####.00")

    ' Pour le pourcentage
    ProgressBarTransfert.Max = lTotalBytes
    ProgressBarTransfert.Min = 0
    j = ProgressBarTransfert.value

    Dim strCaption As String

    Select Case m_objFtpClient.FtpSessionState
        Case ftpDownloadInProgress
            ' Couleur de la ProgressBar lors du download (bleue sur fond blanc)
            SendMessage ProgressBarTransfert.hwnd, PBM_SETBKCOLOR, 0, ByVal RGB(255, 255, 255) ' Fond blanc
            SendMessage ProgressBarTransfert.hwnd, PBM_SETBARCOLOR, 0, ByVal RGB(0, 0, 255) ' Barre bleue
            strCaption = "Téléchargement de " & m_strFileName
        Case ftpUploadInProgress
            ' Couleur de la ProgressBar lors de l'upload (verte sur fond blanc)
            SendMessage ProgressBarTransfert.hwnd, PBM_SETBKCOLOR, 0, ByVal RGB(255, 255, 255) ' Fond blanc
            SendMessage ProgressBarTransfert.hwnd, PBM_SETBARCOLOR, 0, ByVal RGB(0, 200, 0) ' Barre verte
            strCaption = "Envoi de " & m_strFileName
        Case ftpRetrievingDirectoryInfo
            strCaption = "Recherche de la liste du répertoire"
    End Select

    lInfosThread1.Caption = strCaption & ", " & Format$(lngBytesTransferred, "### ### ### ###") & " de " _
                            & Format$(m_lngFileSize, "### ### ### ###") & " bytes,"

    If Not m_objFtpClient.FtpSessionState = ftpRetrievingDirectoryInfo Then
        ProgressBarTransfert.value = lngBytesTransferred / (m_lngFileSize / 100)
    End If

    ' Pour connaître le pourcentage du transfert
    lPourcentage.Caption = "Pourcentage: " & Format$(CLng((j / ProgressBarTransfert.Max) * 100)) + "%"

    ' Pour connaître la vitesse du transfert en Kbps
    lVitesse.Caption = "Speed: " & Format(TransferRate, "##.#0#") & " Kbps"

    ' Pour connaître le temps restant avant la fin du transfert
    lTempsRestant.Caption = "Temps Restant: " & ConvertTime(Int(((ProgressBarTransfert.Max - ProgressBarTransfert.value) / 1024) / TransferRate))

    DoEvents

End Sub

Private Sub m_objFtpClient_OnDownloadFile(ByVal AsyncResultStatus As AsyncResultStatusConstants)

    ' Lorsqu'on a terminé le download -> on efface la frame du thread
    FrameThread1.Visible = False
    ProgressBarTransfert.value = 0.01
    Call DisplayEventInfo("OnDownloadFile", AsyncResultStatus)

    ' On efface aussi le premier item dans lvQueue que l'on vient de downloader
    lvQueue.ListItems.Remove 1

    LoadLocalList CurPath

    ' On appelle le transfert au cas où il y aurait plus d'un fichier à transférer
    Call CmdTransferer_Click

 '   If AsyncResultStatus = arStatusOk Then
 '       lblTransferInfo.Caption = "The file " & m_strFileName & " was downloaded successfully."
 '   Else
 '       lblTransferInfo.Caption = ""
 '   End If
    '
End Sub

Private Sub m_objFtpClient_OnEnumFiles(ByVal AsyncResultStatus As AsyncResultStatusConstants)

    Dim objFtpFile As CFtpFile
    Dim objListItem As ListItem
    Dim intIconIndex As Integer

 '   lblTransferInfo.Caption = ""

    'If the AsyncResultStatus argument of the OnEnumFiles event is arStatusOk,
    'the listing has been received and parsed successfully. This means that if
    'there are any files or subdirectories in the current FTP directory, the
    'CurrentDirectoryFiles property of the CFtpClient class contains an instance
    'of the CFtpFiles collection which we can read with the For...Next loop.
    If AsyncResultStatus = arStatusOk Then

        ' Si le répertoire racine est différent du répertoire courrant
        If m_strRootDirectory <> m_objFtpClient.CurrentDirectory Then
            ' On ajoute le dossier pour revenir en arrière
            lvFTP.ListItems.Add , "GoToParent", "Répertoire Précédent", , 1
            ' Permet d'enlever un bug de sélection si on met ces valeurs pour les SubItems
            With lvFTP.ListItems("GoToParent")
                .SubItems(1) = " "
                .SubItems(2) = " "
                .SubItems(3) = " "
            End With
        End If

        ' S'il y a quelques chose dans le répertoire courrant
        If m_objFtpClient.CurrentDirectoryFiles.Count > 0 Then

            'Walk through the files' collection
            For Each objFtpFile In m_objFtpClient.CurrentDirectoryFiles

                ' Obtenir l'index des icones de la liste d'images :
                If objFtpFile.IsDirectory Then
                    ' On lui donne l'icone n°3 de imlListViews (folder)
                    intIconIndex = 3
                ElseIf objFtpFile.Permissions = "lrwxrwxrwx" Then
                    ' On lui donne l'icone n°2 de imlListViews (folder_back)
                    intIconIndex = 2
                Else
                    ' On lui donne l'icone n°4 de imlListViews (file)
                    intIconIndex = 4
                    'intIconIndex = GetImageNumber(objFtpFile.FileName)
                End If

                ' Ajoute un item dans la ListView
                Set objListItem = lvFTP.ListItems.Add(, "F" & objFtpFile.FileName, objFtpFile.FileName, , intIconIndex)

                ' Ecrit la taille du fichier, sa date et ses autorisations
                objListItem.SubItems(1) = Format$(objFtpFile.FileSize, "### ### ### ###")
                objListItem.SubItems(2) = objFtpFile.LastWriteTime
                objListItem.SubItems(3) = objFtpFile.Permissions

            Next
        End If

    End If

End Sub

Private Sub m_objFtpClient_OnGetCurrentDirectory(ByVal AsyncResultStatus As AsyncResultStatusConstants)

    Call DisplayEventInfo("OnGetCurrentDirectory", AsyncResultStatus)
    Call m_objFtpClient.EnumFiles

    ' Sert à mettre le path du FTP dans le panneau 3 de la barre de statut
    Me.sbStatusBar.Panels(3).Text = m_objFtpClient.CurrentDirectory
    
End Sub

Private Sub m_objFtpClient_OnQuitSession(ByVal AsyncResultStatus As AsyncResultStatusConstants)

    Call DisplayEventInfo("OnQuitSession", AsyncResultStatus)

End Sub

Private Sub DisplayEventInfo(ByVal strEvent As String, ByVal AsyncResultStatus As AsyncResultStatusConstants)

    Dim strCaption As String

    strCaption = strEvent & " - "

    Select Case AsyncResultStatus
        Case arStatusOk: strCaption = strCaption & "arStatusOk"
        Case arStatusError: strCaption = strCaption & "arStatusError"
        Case arStatusCancel: strCaption = strCaption & "arStatusCancel"
        Case arStatusTimeOut: strCaption = strCaption & "arStatusTimeOut"
    End Select

   ' StatusBar1.Panels(2).Text = strCaption

End Sub

Private Sub m_objFtpClient_OnSetCurrentDirectory(ByVal AsyncResultStatus As AsyncResultStatusConstants)

    Call DisplayEventInfo("OnSetCurrentDirectory", AsyncResultStatus)

    If AsyncResultStatus = arStatusOk Then
        Call m_objFtpClient.GetCurrentDirectory
    End If

End Sub

' Savoir ce que LeechSTRO fait
Private Sub m_objFtpClient_OnStateChange(ByVal SessionState As FtpSessionStates)

    Dim strStatusString As String

    Select Case SessionState
        Case ftpFreeState
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Ready"
        Case ftpClosed
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "The control connection is closed"
        Case ftpConnecting
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Connecting to the " & m_objFtpClient.FtpServer & "..."
        Case ftpConnected
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Connected"
        Case ftpAuthentication
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Authentication in progress..."
        Case ftpUserLoggedIn
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "User has been logged in successfully"
        Case ftpChangingCurrentDirectory
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Changing current directory..."
        Case ftpDeletingFile
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Deleting file..."
        Case ftpRemovingDirectory
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Removing directory..."
        Case ftpCreatingDirectory
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Creating directory..."
        Case ftpRenamingFile
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Renaming file..."
        Case ftpEstablishingDataConnection
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Establishing data connection..."
        Case ftpDataConnectionEstablished
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Data connection established"
        Case ftpRetrievingDirectoryInfo
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Retrieving directory info..."
        Case ftpDirectoryInfoRetrieved
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Directory info retrieved"
        Case ftpDownloadInProgress
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Download in progress..."
        Case ftpDownloadCompleted
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Download complete"
        Case ftpUploadInProgress
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Upload in progress..."
        Case ftpUploadCompleted
            strStatusString = "[ThreadPrincipal] Hôte : " & m_objFtpClient.FtpServer & vbCrLf & _
                              "Upload complete"
    End Select

    lIp.Caption = strStatusString

End Sub

Private Sub m_objFtpClient_OnUploadFile(ByVal AsyncResultStatus As AsyncResultStatusConstants)

    ' Lorsqu'on a terminé l'upload -> on efface la frame du thread
    FrameThread1.Visible = False
    ProgressBarTransfert.value = 0.01
    Call DisplayEventInfo("OnUploadFile", AsyncResultStatus)

    ' On efface aussi le premier item dans lvQueue que l'on vient d'uploader
    lvQueue.ListItems.Remove 1

    LoadLocalList CurPath

    ' On appelle le transfert au cas où il y aurait plus d'un fichier à transférer
    Call CmdTransferer_Click

'    If AsyncResultStatus = arStatusOk Then
'        lblTransferInfo.Caption = "The file " & m_strFileName & " was uploaded successfully."
'    Else
'        lblTransferInfo.Caption = ""
'    End If
    '
    lvFTP.ListItems.Clear
    Call m_objFtpClient.EnumFiles

End Sub

' Permet de configurer la couleur les messages apparaîssant dans la RichTextBox
Private Sub m_objFtpClient_SessionProtocolMessage(ByVal strMessage As String, ByVal MessageType As SessionProtocolMessageTypes)

    Select Case MessageType
        Case FTP_USER_COMMAND
            RtbFTP.SelColor = &HC00000 ' Bleu foncé
        Case FTP_SERVER_RESPONSE
            RtbFTP.SelColor = &H8000& ' Vert
        Case FTP_SERVER_BAD_RESPONSE
            RtbFTP.SelColor = &HFF& ' Rouge
        Case FTP_APPLICATION_MESSAGE
            RtbFTP.SelColor = &H0& ' Noir
    End Select

    RtbFTP.SelText = strMessage & vbCrLf
    RtbFTP.SelStart = Len(RtbFTP.Text)

End Sub

' Si l'on souhaite effacer la queue ou la mettre en pause
Private Sub ToolbarQueue_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "Pause"
        Call CmdTransferer_Click
        Case "EffacerList"
        lvQueue.ListItems.Clear
    End Select
    
End Sub

' Lorsque l'on souhaite envoyer un fichier avec le click droit sur l'item
Private Sub zLocalUpload_Click()

On Error Resume Next

    Dim Item1 As ListItem

    Set Item1 = lvQueue.ListItems.Add(, , lvLocal.SelectedItem.Text)
    Item1.SmallIcon = 2
    Item1.SubItems(1) = m_objFtpClient.FtpServer
    Item1.SubItems(2) = Format$(FileLen(lvLocal.SelectedItem), "### ### ### ###")

    ' Démarrer l'upload
    Call CmdTransferer_Click

End Sub

' Lorsque l'on souhaite télécharger un fichier avec le click droit sur l'item
Private Sub zDistantDownload_Click()

On Error Resume Next

    Dim Item2 As ListItem

    Set Item2 = lvQueue.ListItems.Add(, , lvFTP.SelectedItem.Text)
    Item2.SmallIcon = 1
    Item2.SubItems(1) = m_objFtpClient.FtpServer
    Item2.SubItems(2) = Format$((lvFTP.SelectedItem.SubItems(1)), "### ### ### ###")

    ' Démarrer le download
    Call CmdTransferer_Click
    
End Sub

Private Sub zDistantDelete_Click()

    ' Si la connection au ftp n'est pas établise -> quitter
    If m_objFtpClient.FtpSessionState = ftpClosed Then
        Exit Sub
    End If

    ' Si on est occupé avec quelque chose d'autre -> quitter
    If m_objFtpClient.Busy Then
        Exit Sub
    End If

    Dim strFilePath As String
    If MsgBox("Etes-vous sûr de vouloir effacer """ & lvFTP.SelectedItem.Text & """ ?", vbQuestion + vbYesNo, "Effacement de fichier/dossier") = vbYes Then
    End If
           
End Sub

Private Sub zDistantRename_Click()
' Renommer
End Sub

Private Sub zDistantCreateFolder_Click()

    ' Si la connection au ftp n'est pas établise -> quitter
    If m_objFtpClient.FtpSessionState = ftpClosed Then
        Exit Sub
    End If

    ' Si on est occupé avec quelque chose d'autre -> quitter
    If m_objFtpClient.Busy Then
        Exit Sub
    End If

    ' On entre le nom du dossier
    Dim strDirectoryPath As String
    strDirectoryPath = Trim(InputBox("Veillez entrer le nom du dossier à créer:"))

    ' Appelez la méthode CreateDirectory afin de créer un répertoire
    Call m_objFtpClient.CreateDirectory(strDirectoryPath)

End Sub

Private Sub zDistantChangeFolder_Click()
' Changer de dossier
End Sub

Private Sub zDistantInformation_Click()
' Informations
End Sub

' Interdit de télécharger le dossier [Répertoire Précédent]
Private Sub lvFTP_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If Not (Item Is Nothing) Then Me.zDistantDownload.Enabled = (Item.Text <> "Répertoire Précédent")
    If Not (Item Is Nothing) Then Me.zDistantDelete.Enabled = (Item.Text <> "Répertoire Précédent")
    If Not (Item Is Nothing) Then Me.zDistantRename.Enabled = (Item.Text <> "Répertoire Précédent")
    
End Sub

Private Sub CmdTransferer_Click()

    Dim blnOverWrite    As Boolean
    Dim varResult       As VbMsgBoxResult
    Dim lngStartPos     As Long
    Dim objListItem     As ListItem
    Dim blnFileExists   As Boolean
    Dim strPrompt       As String
    Dim lngRemoteFileSize As Long
    Dim varMsgBoxResult As VbMsgBoxResult

' Si le transfer n'est pas en pause -> continuer
If ToolbarQueue.Buttons(1).value = 0 Then

    ' Si la connection au ftp n'est pas établise -> quitter
    If m_objFtpClient.FtpSessionState = ftpClosed Then
        Exit Sub
    End If

    ' Si on est occupé avec quelque chose d'autre -> quitter
    If m_objFtpClient.Busy Then
        Exit Sub
    End If

    ' S'il n'y a aucun fichier dans le file -> quitter
    If lvQueue.ListItems.Count = 0 Then
        Exit Sub
    End If

    ' Pour savoir la vitesse du transfert en Kbps
    BeginTransfer = Timer

    ' Affiche le tab de l'état du transfert
    SSTabTransferts.Tab = 1

    ' Affiche le thread
    FrameThread1.Visible = True

    lIp2.Caption = "Hôte : " & lvQueue.ListItems(1).SubItems(1)

      ' Si c'est un fichier à uploader
      If lvQueue.ListItems(1).SmallIcon = 2 Then
 
                ' Affiche l'icone d'upload
                PicThread1DL.Visible = False
                PicThread1UL.Visible = True

                ' Stockez le nom du fichier au niveau du module variable m_strFileName
                m_strFileName = Mid$(CurPath & lvQueue.ListItems(1).Text, InStrRev(CurPath & lvQueue.ListItems(1).Text, "\") + 1)
                ' Stockez la taille de fichier au niveau du module variable m_lngFileSize
                m_lngFileSize = FileLen(CurPath & lvQueue.ListItems(1).Text)

' A VERIFIER !!!
                ' Vérifiez l'existence du fichier
'                Set objListItem = lvFTP.ListItems("F" & m_strFileName)
'                If Not objListItem Is Nothing Then blnFileExists = True

                ' Si le fichier existe déjà
                If blnFileExists Then

                    lngRemoteFileSize = CLng(objListItem.SubItems(1))

                    If lngRemoteFileSize < m_lngFileSize Then

                        ' Si la longueur du fichier local est plus grande que la longueur du fichier distant,
                        ' probablement l'opération précédente d'upload était échouée.

                        ' Nous devons demander à l'utilisateur quoi faire.

                        strPrompt = "Le fichier " & m_strFileName & " existe déjà!" & _
                                    vbCrLf & vbCrLf & "Taille du fichier distant:" & vbTab & _
                                    Format$(lngRemoteFileSize, "### ### ### ###") & " bytes" & _
                                    vbCrLf & "Taille du fichier local:" & vbTab & _
                                    Format$(m_lngFileSize, "### ### ### ###") & " bytes" & _
                                    vbCrLf & vbCrLf & _
                                    "Aimeriez-vous résumer le transfer afin de terminer l'envoi?" & vbCrLf & vbCrLf & _
                                    "Note: Si vous choisissez NON, un nouveau fichier va être créé."

                        varMsgBoxResult = MsgBox(strPrompt, vbYesNoCancel + vbQuestion, "File already exists")

                        If varMsgBoxResult = vbYes Then

                            ' L'utilisateur aimerait apposer des données (ou remettre en marche le transfert de données).
                            ' Stockez la valeur de position de relancement dans la variable lngStartPos,
                            ' ce qui sera passé comme un argument à la méthode d'UploadFile.
                            lngStartPos = lngRemoteFileSize

                        ElseIf varMsgBoxResult = vbCancel Then

                            Exit Sub

                        End If

                    Else

                        strPrompt = "Le fichier " & m_strFileName & " existe déjà!" & _
                                    vbCrLf & vbCrLf & "Taille du fichier distant:" & vbTab & _
                                    Format$(lngRemoteFileSize, "### ### ### ###") & " bytes" & _
                                    vbCrLf & "Taille du fichier local:" & vbTab & _
                                    Format$(m_lngFileSize, "### ### ### ###") & " bytes" & _
                                    vbCrLf & vbCrLf & _
                                    "Voulez-vous de toute façon envoyer le nouveau fichier?" & vbCrLf & vbCrLf & _
                                    "Note: Si vous choisissez OUI, le vieux fichier va être remplacé avec le nouveau."

                        varMsgBoxResult = MsgBox(strPrompt, vbYesNo + vbQuestion, "File already exists")

                        If varMsgBoxResult = vbNo Then

                            Exit Sub

                        End If

                    End If

                End If

                ' Appelez la méthode UploadFile afin de commencer l'envoi
                Call m_objFtpClient.UploadFile(CurPath & lvQueue.ListItems(1).Text, m_strFileName, lngStartPos)

      End If

      ' Si c'est un fichier à downloader
      If lvQueue.ListItems(1).SmallIcon = 1 Then

                ' Affiche l'icone de download
                PicThread1UL.Visible = False
                PicThread1DL.Visible = True

                ' Stockez le nom du fichier au niveau du module variable m_strFileName
                m_strFileName = lvQueue.ListItems(1).Text
                ' Stockez la taille de fichier au niveau du module variable m_lngFileSize
                m_lngFileSize = CLng(lvQueue.ListItems(1).SubItems(2))

                ' Si le fichier existe déjà
                If FileExists(CurPath & lvQueue.ListItems(1)) Then

                    If m_lngFileSize > FileLen(CurPath & lvQueue.ListItems(1)) Then

                        strPrompt = "Le fichier " & lvQueue.ListItems(1) & " existe déjà!" & _
                                    vbCrLf & vbCrLf & "Taille du fichier distant:" & _
                                    vbTab & Format$(m_lngFileSize, "### ### ### ###") & _
                                    vbTab & "bytes" & vbCrLf & "Taille du fichier local:" & _
                                    vbTab & Format$(FileLen(lvQueue.ListItems(1)), "### ### ### ###") & _
                                    vbTab & "bytes" & vbCrLf & vbCrLf & _
                                    "Aimeriez-vous résumer le transfert afin de terminer le téléchargement?" & _
                                    vbCrLf & vbCrLf & "Note: Si vous choisissez NON, un nouveau fichier va être créé."

                        varResult = MsgBox(strPrompt, vbYesNoCancel + vbQuestion, "Le fichier existe déjà")

                        If varResult = vbNo Then
                            blnOverWrite = True
                        ElseIf varResult = vbCancel Then
                            Exit Sub
                        End If

                    Else

                        strPrompt = "Le fichier " & lvQueue.ListItems(1) & " existe déjà!" & _
                                    vbCrLf & vbCrLf & "Taille du fichier distant:" & _
                                    vbTab & Format$(m_lngFileSize, "#### ### ### ###") & _
                                    vbTab & "bytes" & vbCrLf & "Taille du fichier local:" & _
                                    vbTab & Format$(FileLen(lvQueue.ListItems(1)), "### ### ### ###") & _
                                    vbTab & "bytes" & vbCrLf & vbCrLf & _
                                    "Aimeriez-vous quitter le téléchargement?" & vbCrLf & vbCrLf & _
                                    "Note: Si vous choisissez NON, un nouveau fichier va être créé."

                        varResult = MsgBox(strPrompt, vbYesNo + vbQuestion, "File already exists")

                        If varResult = vbYes Then
                            Exit Sub
                        Else
                            blnOverWrite = True
                        End If

                    End If

                End If

                ' Appelez la méthode DownloadFile afin de commencer le téléchargement
                Call m_objFtpClient.DownloadFile(m_strFileName, CurPath & lvQueue.ListItems(1), blnOverWrite)

      End If

Else
    ' Si le transfert est en pause -> quitter
    Exit Sub
End If

End Sub

' Si l'on clique sur l'icone dans le systray
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Select Case x
        Case 7725: ' Double click gauche
            frmMain.Show
            ToolbarMain.Visible = True
            Call Shell_NotifyIcon(NIM_DELETE, Systray)
        Case 7755: ' Click droit
            ' Sert à afficher le PopupMenu et mettre le texte [ Ouvrir ] en gras
            PopupMenu zSystray, 0, , , zSystrayOpen
    End Select
    
End Sub

Private Sub zSystrayAbout_Click()
    frmAbout.Show
End Sub

Private Sub zSystrayOpen_Click()

    frmMain.Show
    ToolbarMain.Visible = True
    Call Shell_NotifyIcon(NIM_DELETE, Systray)
    
End Sub

' Si l'on souhaite quitter à partir du systray
Private Sub zSystrayClose_CLick()

    Call Shell_NotifyIcon(NIM_DELETE, Systray)
    End
    
End Sub

' Provisoire -> pour connaître le nombre de fichiers dans la liste
Private Sub TimerQueue_Timer()
    lNbFilesQueue.Caption = "Fichiers Dans La Queue: " & lvQueue.ListItems.Count
End Sub

' Ajoute chaque dossiers et fichiers à la ListView
Private Sub LoadLocalList(ByVal path As String)

On Error Resume Next
    
    Dim oLIs As ListItems
    Dim fsys As New FileSystemObject
    Dim drv As Drive
    Dim fld As Folder
    Dim sfld As Folder
    Dim fil As File
    Dim fils As Files
    
    ' Change le curseur de la souris en sablier
    Me.MousePointer = vbHourglass

    Set oLIs = Me.lvLocal.ListItems
    ' Efface la ListView Locale
    oLIs.Clear
    
    ' Si ce n'est pas le répertoire racine
    If Not Len(path) = 3 Then
        ' Rajoute le dossier pour revenir en arrière
        oLIs.Add , "GoToParent", "Répertoire Précédent", , "folder_back"
        ' Permet d'enlever un bug de sélection si on met ces valeurs pour les SubItems
        With oLIs("GoToParent")
            .SubItems(1) = " "
            .SubItems(2) = " "
            .SubItems(3) = " "
        End With
    End If
   
    ' Affiche le path courant
    Me.sbStatusBar.Panels(2).Text = path
    ' Si le path est trop long, on ajoute "..." devant le chiffre du disque dur
'    If Len(Me.lblLocalPath) > 70 Then
'        Me.lblLocalPath = Trim(Left(Me.lblLocalPath, 3)) & "..." & Trim(Right(Me.lblLocalPath, 70))
'    End If
  
    ' Permet d'ajouter tous les disques dans un popup menu
    If zLocalChangeDriveX.Item(0).Caption = "X:\" Then
        For Each drv In fsys.Drives
            Load zLocalChangeDriveX(MenuIndex)
            zLocalChangeDriveX(MenuIndex).Caption = drv.DriveLetter & ":\"
            zLocalChangeDriveX(MenuIndex).Visible = True
            MenuIndex = MenuIndex + 1
        Next
    End If
    
    ' Permet d'ajouter tous les dossiers
    Set fld = fsys.GetFolder(path)
    For Each sfld In fld.SubFolders
        oLIs.Add , , sfld.name, , "folder"
        oLIs.Item(lvLocal.ListItems.Count).Tag = "folder"
        oLIs.Item(lvLocal.ListItems.Count).SubItems(1) = 0 ' Taille
        oLIs.Item(lvLocal.ListItems.Count).SubItems(2) = sfld.DateLastModified ' Date
    Next
    
    ' Permet d'ajouter tous les fichiers
    Set fils = fsys.GetFolder(path).Files
    For Each fil In fils
        oLIs.Add , , fil.name, , "file"
        oLIs.Item(lvLocal.ListItems.Count).Tag = "file"
        oLIs.Item(lvLocal.ListItems.Count).SubItems(1) = Format$(fil.Size, "### ### ### ###") ' Taille
        oLIs.Item(lvLocal.ListItems.Count).SubItems(2) = fil.DateLastModified ' Date
    Next
    
    Me.lvLocal.Refresh
    ' Remet le curseur de défaut de la souris
    Me.MousePointer = vbDefault
    
End Sub

Public Function FormatFileSize(lFileSize As Long) As String

On Error GoTo ERROR_HANDLER
    
    If lFileSize >= 1024 Then
        FormatFileSize = Format$(CStr(lFileSize / 1024), "###,###,###,###KB")
    Else
        FormatFileSize = CStr(lFileSize) & "bytes"
    End If

    Exit Function
    
ERROR_HANDLER:
    Debug.Print Err.Number & " " & Err.Description
    
End Function

'Private Function GetImageNumber(strFileName As String) As Integer

'    Dim strExt As String

'    strExt = Mid$(strFileName, InStrRev(strFileName, ".") + 1)

'    On Error Resume Next

'    Select Case LCase(strExt)
'        Case "asp", "asa", "inc", "css", "shtml", "txt", "htm", "html", "lst", "log", "ini", "inf", ""
'            GetImageNumber = 3
'        Case Else
'            GetImageNumber = 4
'    End Select

'End Function

Private Function FileExists(strFileName As String) As Boolean
    
On Error GoTo ERROR_HANDLER
    
    FileExists = (GetAttr(strFileName) And vbDirectory) = 0

ERROR_HANDLER:
    
End Function

' Pour savoir le temps qu'il reste avant la fin du transfert
Public Function ConvertTime(ByVal TheTime As Single) As String

    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    
End Function
