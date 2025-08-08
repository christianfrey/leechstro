VERSION 5.00
Begin VB.Form frmInformation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "...::: LeechSTRO :::... Informations sur le répertoire local"
   ClientHeight    =   3180
   ClientLeft      =   5295
   ClientTop       =   4350
   ClientWidth     =   6075
   LinkTopic       =   "frmInfo"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.Label lblPath 
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Label Label12 
      Caption         =   "xXx"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label11 
      Caption         =   "xXx"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2640
      TabIndex        =   12
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label10 
      Caption         =   "xXx"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2640
      TabIndex        =   11
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "xXx"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2640
      TabIndex        =   10
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label8 
      Caption         =   "xXx"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2640
      TabIndex        =   9
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "xXx"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Fichiers Au Total :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   7
      Top             =   1200
      Width           =   2300
   End
   Begin VB.Label Label6 
      Caption         =   "Dossiers Au Total :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   6
      Top             =   1560
      Width           =   2300
   End
   Begin VB.Label Label5 
      Caption         =   "Taille Sélectionée :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   5
      Top             =   2040
      Width           =   2300
   End
   Begin VB.Label Label4 
      Caption         =   "Fichiers Sélectionnés :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   4
      Top             =   2400
      Width           =   2300
   End
   Begin VB.Label Label3 
      Caption         =   "Dossiers Sélectionnées :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   3
      Top             =   2760
      Width           =   2300
   End
   Begin VB.Label Label1 
      Caption         =   "Taille Totale :"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   2
      Top             =   840
      Width           =   2300
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
