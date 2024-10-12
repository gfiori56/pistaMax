VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "PistaMaxV2"
   ClientHeight    =   9870
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   17925
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   17925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_note 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   44
      Top             =   6720
      Width           =   12195
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MyGrid 
      Height          =   5895
      Index           =   0
      Left            =   7680
      TabIndex        =   41
      Top             =   720
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   10398
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmb_quante_macchine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   7440
      Width           =   4215
   End
   Begin VB.TextBox txt_seconds 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   37
      Top             =   9240
      Width           =   1215
   End
   Begin VB.ComboBox cmb_tipo_gara 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   8040
      Width           =   4215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11760
      TabIndex        =   35
      Top             =   8640
      Width           =   3015
   End
   Begin VB.TextBox txt_name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   33
      Text            =   "Macchina B"
      Top             =   6000
      Width           =   2000
   End
   Begin VB.TextBox txt_name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   32
      Text            =   "Macchina A"
      Top             =   6000
      Width           =   2000
   End
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   5040
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txt_giri_da_eseguire 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSimul 
      Caption         =   "AB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdSimul 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   12
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdSimul 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Inizio nuova corsa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   9
      Top             =   8640
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid MyGrid 
      Height          =   5895
      Index           =   1
      Left            =   13920
      TabIndex        =   42
      Top             =   720
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   10398
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1560
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Note personali"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   360
      TabIndex        =   43
      Top             =   6720
      Width           =   4935
   End
   Begin VB.Label lbl_advise 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5400
      TabIndex        =   39
      Top             =   8040
      Width           =   12255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Durata corsa (secondi)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   38
      Top             =   9240
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Indentificativo macchina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   360
      TabIndex        =   34
      Top             =   6000
      Width           =   4935
   End
   Begin VB.Label lbl_info 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "attesa partenza ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   31
      Top             =   7440
      Width           =   6015
   End
   Begin VB.Label lbl_elap 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   30
      Top             =   1320
      Width           =   1995
   End
   Begin VB.Label lbl_elap 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   29
      Top             =   1320
      Width           =   2000
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo totale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   360
      TabIndex        =   28
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label lbl_info 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "attesa partenza ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   27
      Top             =   7440
      Width           =   6015
   End
   Begin VB.Label lbl_tempo_max 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   26
      Top             =   4320
      Width           =   1995
   End
   Begin VB.Label lbl_tempo_min 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   25
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Label lbl_tempo_medio 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   24
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label lbl_tempo_ultimo_giro 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   23
      Top             =   2520
      Width           =   1995
   End
   Begin VB.Label lbl_tempo_giri_eseguiti 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   22
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Label lbl_giri_eseguiti 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   11640
      TabIndex        =   21
      Top             =   720
      Width           =   1995
   End
   Begin VB.Label lbl_simul 
      Alignment       =   2  'Center
      Caption         =   "Simulazione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lbl_tempo_max 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   18
      Top             =   4320
      Width           =   2000
   End
   Begin VB.Label lbl_tempo_min 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   17
      Top             =   3720
      Width           =   2000
   End
   Begin VB.Label lbl_tempo_medio 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   16
      Top             =   3120
      Width           =   2000
   End
   Begin VB.Label lbl_tempo_ultimo_giro 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   15
      Top             =   2520
      Width           =   2000
   End
   Begin VB.Label lbl_tempo_giri_eseguiti 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   14
      Top             =   1920
      Width           =   2000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Numero giri da eseguire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   8760
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo giro peggiore (#giro)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo giro migliore (#giro)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo medio giro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   3120
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo ultimo giro eseguito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo giri eseguiti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label lbl_giri_eseguiti 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   2000
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      Caption         =   "Macchina B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   11640
      TabIndex        =   2
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label lbl_name 
      Alignment       =   2  'Center
      Caption         =   "Macchina A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   1
      Top             =   240
      Width           =   2000
   End
   Begin VB.Label Label1 
      Caption         =   "Numero giri eseguiti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "Apri ..."
      End
      Begin VB.Menu save 
         Caption         =   "Salva ..."
      End
      Begin VB.Menu export 
         Caption         =   "Esporta ..."
      End
      Begin VB.Menu dummy 
         Caption         =   "---------"
      End
      Begin VB.Menu exit 
         Caption         =   "Esci"
      End
   End
   Begin VB.Menu strumenti 
      Caption         =   "&Strumenti"
      Begin VB.Menu classifica 
         Caption         =   "Classifica"
      End
   End
   Begin VB.Menu info 
      Caption         =   "&Informazioni"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_GIRI = 1000
Private Const DEFAULT_GIRI = 20
Private Const INFO_ANNI = "2023"
Private Const GRID_SCROLL_DEFAULT = 14
Private Const FILE_EXT = "pm2"

Private Const NG_SEP1      As String = " ("
Private Const NG_SEP2      As String = ")"

'questo ordine corrisponde a quello della combo
Private Enum E_QM
  CMB_QM_A
  CMB_QM_B
  CMB_QM_AB
  CMB_QMMAX 'always in the last position
End Enum
'questo ordine corrisponde a quello della combo
Private Enum E_TG
  CMB_TG_UNLIM
  CMB_TG_GIRI
  CMB_TG_TEMPO
  CMB_TGMAX 'always in the last position
End Enum
Private Enum E_RUN
  eWait
  eRun
  eStop
End Enum
Private Type T_MACCHINA
  lGiri       As Long
  lGiriTarget As Long
  dTime       As Double
  dTimeLatch  As Double
  dLastTime   As Double
  dAveTime    As Double
  dMaxTime    As Double
  dMinTime    As Double
  l_NGMax     As Long
  l_NGMin     As Long
  eRun        As E_RUN
  dStartTime As Double
  dTrace(1 To MAX_GIRI) As Double
End Type
Private Type T_PGAME
  bEnable      As Boolean
  eCmbQmSelect As E_QM
  eCmbTgSelect As E_TG
  iMaxCycles   As Integer
  dMaxSeconds  As Double
End Type

Private muMac(1 To N_MACCHINE) As T_MACCHINA 'dati di lavoro
Private muGame                 As T_PGAME    'parametri del gioco
Private miGRID_SCROLL As Integer

'configurazione
Private mi_com    As Integer
Private ms_baud   As String
Private mbSimul   As Boolean
Private mbPrvInst As Boolean

'---------- resize data >> ------------
Private Type T_ITEM
  l_Top      As Long
  l_Left     As Long
  l_Height   As Long
  l_Width    As Long
  i_FontSize As Integer
  o_ctl      As Control
End Type
Private ml_Start_Height As Long
Private ml_Start_Width  As Long
Private Const MAX_ITEMS = 1000
Private ml_ItemsNr      As Long
Private mu_Start_Item(1 To MAX_ITEMS) As T_ITEM

'---------- << resize data ------------
'------------------------------------------------
'           LOAD/UNLOAD
'------------------------------------------------
Private Sub Form_Load()
  Dim lbTest  As Boolean
  Dim liX     As Integer
  Dim llP     As Long
  
  If App.PrevInstance Then
    mbPrvInst = True
    Unload Me
    Exit Sub
  End If
  
  miGRID_SCROLL = GRID_SCROLL_DEFAULT
  Call sbAdvise
  
  '-- initialize combo's --
  Dim QM(0 To CMB_QMMAX - 1) As String
  QM(CMB_QM_A) = "Uso solo macchina A"
  QM(CMB_QM_B) = "Uso solo macchina B"
  QM(CMB_QM_AB) = "Uso macchine A e B"
  Call cmb_quante_macchine.Clear
  For liX = 0 To CMB_QMMAX - 1
    Call cmb_quante_macchine.AddItem(QM(liX))
  Next liX
  Dim TG(0 To CMB_TGMAX - 1) As String
  TG(CMB_TG_UNLIM) = "Prove"
  TG(CMB_TG_GIRI) = "Corsa a numero giri"
  TG(CMB_TG_TEMPO) = "Corsa a tempo"
  Call cmb_tipo_gara.Clear
  For liX = 0 To CMB_TGMAX - 1
    Call cmb_tipo_gara.AddItem(TG(liX))
  Next liX
  
  For liX = 0 To N_MACCHINE - 1
    Call InitGrid(liX)
  Next liX
  
  Call sbInitializeTimer
  Call sbRestart(False)
  Call sbFormInPrimoPiano(Me)
  Call sbItemsInit
  Call sbInitializeHomeDir
  If InStr(1, Command$, "/SETUP", vbTextCompare) > 0 Then
    'configurazione
    Dim lsValue1 As String
    Dim lsValue2 As String
    lsValue1 = ReadConfigFile("SETTINGS", "COM")
    lsValue2 = InputBox("Inserire COM o stringa di riconoscimento", "Setup", lsValue1)
    If Len(lsValue2) > 0 And lsValue2 <> lsValue1 Then
      Call WriteConfigFile("SETTINGS", "COM", lsValue2)
      Call MsgBox("Valore " & Chr(&H22) & lsValue2 & Chr(&H22) & " memorizzato", vbInformation)
    End If
  End If
  Call sbLoad
  
  If InStr(1, Command$, "/SIMUL", vbTextCompare) > 0 Then mbSimul = True
  If mbSimul = False Then
    mbSimul = Not Comm_Init()
  End If
  lbl_simul.Visible = mbSimul
  cmdSimul(1).Visible = mbSimul
  cmdSimul(2).Visible = mbSimul
  cmdSimul(3).Visible = mbSimul
  Call sbPresetParameters
  miGRID_SCROLL = LinesGrid(0)
  Timer1.Interval = 100
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If mbPrvInst = True Then Exit Sub
  Call Comm_Exit
  Call sbSave
End Sub
'------------------------------------------------
'           altri eventi
'------------------------------------------------
'modifica combo TG
Private Sub cmb_tipo_gara_Click()
   Call sbPresetParameters
End Sub
'simulazione seriale
Private Sub cmdSimul_Click(Index As Integer)
  Dim ldNow  As Double
  ldNow = fnReadTimer()
  Call cmdPassaggio(Index, ldNow)
End Sub

'stop
Private Sub sbStop(Optional ByVal abNoMsg As Boolean = False)
  cmb_quante_macchine.Enabled = True
  cmb_tipo_gara.Enabled = True
  cmdRestart.Enabled = True
  
  cmb_quante_macchine.Locked = False
  cmb_tipo_gara.Locked = False
  txt_seconds.Locked = False
  txt_giri_da_eseguire.Locked = False
  muGame.bEnable = False
  If abNoMsg = True Then Exit Sub
  
  Dim liX As Integer
  For liX = 1 To N_MACCHINE
    lbl_info(liX - 1).Caption = vbNullString
    lbl_info(liX - 1).ForeColor = vbBlack
    lbl_info(liX - 1).BackColor = GRAY_COLOR
  Next liX
  Call sbAdvise("STOP")
End Sub
Private Sub cmdStop_Click()
  Call sbStop
End Sub

'--------------------------------------------
'           MENU
'--------------------------------------------
Private Sub info_Click()
  If muGame.bEnable = True Then Exit Sub
  Call MsgBox("Versione del programma:" & vbCrLf & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "Sviluppo:" & vbCrLf & INFO_ANNI, vbInformation)
End Sub
Private Sub exit_Click()
  Unload Me
End Sub
 
Private Sub open_Click()
  If muGame.bEnable = True Then Exit Sub
  
  Dim lsFile As String
  On Error Resume Next
  CommonDialog1.CancelError = True
  CommonDialog1.FileName = "*." & FILE_EXT
  CommonDialog1.Filter = "PistaMax file (*." & FILE_EXT & ")|*." & FILE_EXT & "|Tutti i file (*.*)|*.*"
  CommonDialog1.ShowOpen
  
  lsFile = CommonDialog1.FileName
  If Len(lsFile) = 0 Then Exit Sub
  If InStr(1, lsFile, "*") > 0 Then Exit Sub
  
  'lettura file
  Dim liX As Integer
  Dim liY As Integer
  Dim lsLine As String
  Dim lsSection As String
  
  On Error GoTo local_err
  lsSection = "General"
  lsLine = ReadConfigFile(lsSection, "EXT", lsFile)
  If lsLine <> UCase(FILE_EXT) Then
    Call sbAdvise("FILE NON VALIDO", vbBlack, vbRed)
    Exit Sub
  End If
  txt_note.Text = ReadConfigFile(lsSection, "NOTE", lsFile)
  
  For liX = 1 To N_MACCHINE
    With muMac(liX)
      lsSection = "REPORT_" & Chr(64 + liX)
      lbl_giri_eseguiti(liX - 1).Caption = ReadConfigFile(lsSection, "GIRI_ESEGUITI", lsFile)
      lbl_tempo_giri_eseguiti(liX - 1).Caption = ReadConfigFile(lsSection, "TEMPO_GIRI_ESEGUITI", lsFile)
      lbl_tempo_ultimo_giro(liX - 1).Caption = ReadConfigFile(lsSection, "TEMPO_ULTIMO_GIRO", lsFile)
      lbl_tempo_medio(liX - 1).Caption = ReadConfigFile(lsSection, "TEMPO_MEDIO", lsFile)
      lbl_tempo_max(liX - 1).Caption = ReadConfigFile(lsSection, "TEMPO_MAX", lsFile)
      lbl_tempo_min(liX - 1).Caption = ReadConfigFile(lsSection, "TEMPO_MIN", lsFile)
      lbl_elap(liX - 1).Caption = ReadConfigFile(lsSection, "TEMPO_TOTALE", lsFile)
      txt_name(liX - 1).Text = ReadConfigFile(lsSection, "NOME", lsFile)
      lsSection = "DATI_" & Chr(64 + liX)
      lsLine = ReadConfigFile(lsSection, "N_GIRI", lsFile)
      .lGiri = Val(lsLine)
      If .lGiri > MAX_GIRI Then .lGiri = MAX_GIRI
      For liY = 1 To .lGiri
        lsLine = ReadConfigFile(lsSection, "GIRO_" & liY, lsFile)
        .dTrace(liY) = Val(lsLine)
      Next liY
    End With
  Next liX
  Call RefillGrid
  Call sbAdvise("FILE CARICATO")
  Exit Sub
  
local_err:
  Call sbAdvise("ERRORE LETTURA FILE codice=" & Err.Number, vbBlack, vbRed, Err.Description)
End Sub
Private Sub save_Click()
  If muGame.bEnable = True Then Exit Sub
  
  Dim lsFile As String
  Dim lsLongDate As String
  On Error Resume Next
  CommonDialog1.CancelError = True
  CommonDialog1.FileName = fnNow2Str(lsLongDate) & "." & FILE_EXT
  CommonDialog1.Filter = "PistaMax file (*." & FILE_EXT & ")|*." & FILE_EXT & "|Tutti i file (*.*)|*.*"
  CommonDialog1.ShowSave
  
  lsFile = CommonDialog1.FileName
  If Len(lsFile) = 0 Then Exit Sub
  If InStr(1, lsFile, "*") > 0 Then Exit Sub
  
  If fnFileExist(lsFile) Then
     If MsgBox("Il file esiste, vuoi sovrascriverlo?", vbExclamation Or vbYesNo) = vbNo Then Exit Sub
  End If
  
  'scrittura file
  Dim liX As Integer
  Dim liY As Integer
  Dim lsSection As String
  
  On Error GoTo local_err
  lsSection = "General"
  Call WriteConfigFile(lsSection, "EXT", UCase(FILE_EXT), lsFile)
  Call WriteConfigFile(lsSection, "NOTE", txt_note.Text, lsFile)
  
  For liX = 1 To N_MACCHINE
    With muMac(liX)
      lsSection = "REPORT_" & Chr(64 + liX)
      Call WriteConfigFile(lsSection, "GIRI_ESEGUITI", lbl_giri_eseguiti(liX - 1).Caption, lsFile)
      Call WriteConfigFile(lsSection, "TEMPO_GIRI_ESEGUITI", lbl_tempo_giri_eseguiti(liX - 1).Caption, lsFile)
      Call WriteConfigFile(lsSection, "TEMPO_ULTIMO_GIRO", lbl_tempo_ultimo_giro(liX - 1).Caption, lsFile)
      Call WriteConfigFile(lsSection, "TEMPO_MEDIO", lbl_tempo_medio(liX - 1).Caption, lsFile)
      Call WriteConfigFile(lsSection, "TEMPO_MAX", lbl_tempo_max(liX - 1).Caption, lsFile)
      Call WriteConfigFile(lsSection, "TEMPO_MIN", lbl_tempo_min(liX - 1).Caption, lsFile)
      Call WriteConfigFile(lsSection, "TEMPO_TOTALE", lbl_elap(liX - 1).Caption, lsFile)
      Call WriteConfigFile(lsSection, "NOME", txt_name(liX - 1).Text, lsFile)
      lsSection = "DATI_" & Chr(64 + liX)
      If .lGiri > MAX_GIRI Then .lGiri = MAX_GIRI
      Call WriteConfigFile(lsSection, "N_GIRI", .lGiri, lsFile)
      For liY = 1 To .lGiri
        Call WriteConfigFile(lsSection, "GIRO_" & liY, Str(.dTrace(liY)), lsFile)
      Next liY
    End With
  Next liX
  Call sbAdvise("FILE SALVATO")
  Exit Sub
local_err:
  Call sbAdvise("ERRORE SCRITTURA FILE codice=" & Err.Number, vbBlack, vbRed, Err.Description)
End Sub

Private Sub export_Click()
  If muGame.bEnable = True Then Exit Sub
  
  Dim SEP        As String
  Dim DEC        As String
  Dim loFrm      As frmExport
  
  Set loFrm = New frmExport
  Call loFrm.Show(vbModal)
  SEP = loFrm.msListSep
  DEC = loFrm.msDecSep
  Set loFrm = Nothing
  
  If SEP = vbNullString Or DEC = vbNullString Then Exit Sub
  
  Dim lsFile     As String
  Dim lsLongDate As String
  On Error Resume Next
  
  CommonDialog1.CancelError = True
  CommonDialog1.FileName = fnNow2Str(lsLongDate) & ".csv"
  CommonDialog1.Filter = "Excel csv files (*.csv)|*.csv"
  CommonDialog1.ShowSave
  
  lsFile = CommonDialog1.FileName
  If Len(lsFile) = 0 Then Exit Sub
  If InStr(1, lsFile, "*") > 0 Then Exit Sub
  
  If fnFileExist(lsFile) Then
    If MsgBox("Il file esiste, vuoi sovrascriverlo?", vbExclamation Or vbYesNo) = vbNo Then Exit Sub
  End If
  
  
  Dim liX        As Integer
  Dim liMaxLines As Integer
  Dim lsA        As String
  Dim lsB        As String
  Dim lsNomeA    As String
  Dim lsNomeB    As String
  
  lsNomeA = Trim(txt_name(0).Text)
  If Len(lsNomeA) = 0 Then lsNomeA = "A"
  lsNomeB = Trim(txt_name(1).Text)
  If Len(lsNomeB) = 0 Then lsNomeB = "B"
  
  If muMac(1).lGiri > MAX_GIRI Then muMac(1).lGiri = MAX_GIRI
  If muMac(2).lGiri > MAX_GIRI Then muMac(2).lGiri = MAX_GIRI
  
  liMaxLines = muMac(1).lGiri
  If liMaxLines < muMac(2).lGiri Then liMaxLines = muMac(2).lGiri
  
  On Error GoTo local_err
  Open lsFile For Output As #1
  Print #1, txt_note.Text
  Print #1, "#giro" & SEP & lsNomeA & SEP & lsNomeB
  For liX = 1 To liMaxLines
    lsA = vbNullString
    lsB = vbNullString
    If liX <= muMac(1).lGiri Then lsA = MyVal(muMac(1).dTrace(liX), DEC)
    If liX <= muMac(2).lGiri Then lsB = MyVal(muMac(2).dTrace(liX), DEC)
    Print #1, liX & SEP & lsA & SEP & lsB
  Next liX

local_exit:
  On Error Resume Next
  Close #1
  Exit Sub
local_err:
  Call sbAdvise("ERRORE EXPORT FILE codice=" & Err.Number, vbBlack, vbRed, Err.Description)
  Resume local_exit
End Sub

'--------------------------------------------
'    funzioni generiche
'--------------------------------------------
Private Sub sbAdvise(Optional ByVal asMsg As String = vbNullString, _
  Optional ByVal alFG As Long = vbWhite, _
  Optional ByVal alBG As Long = vbBlue, _
  Optional ByVal asTooltip As String = vbNullString)
  If asMsg = vbNullString Then
    alFG = vbBlack
    alBG = GRAY_COLOR
  End If
  lbl_advise.Caption = asMsg
  lbl_advise.ForeColor = alFG
  lbl_advise.BackColor = alBG
  lbl_advise.ToolTipText = asTooltip

End Sub


'--------------------------------------------
'    preset editable parameters
'--------------------------------------------
Private Sub sbPresetParameters()
  
  On Error Resume Next
  muGame.bEnable = False
  muGame.eCmbQmSelect = cmb_quante_macchine.ListIndex
  muGame.eCmbTgSelect = cmb_tipo_gara.ListIndex
  
  muGame.iMaxCycles = txt_giri_da_eseguire.Text
  If muGame.iMaxCycles < 1 Then muGame.iMaxCycles = 1
  If muGame.iMaxCycles > MAX_GIRI Then muGame.iMaxCycles = MAX_GIRI
  
  muGame.dMaxSeconds = txt_seconds.Text
  If muGame.dMaxSeconds < 0 Then muGame.dMaxSeconds = 0
  
  txt_giri_da_eseguire.Text = muGame.iMaxCycles
  txt_seconds.Text = Format(muGame.dMaxSeconds, TIME_FORMAT)
  
  If muGame.eCmbTgSelect = CMB_TG_GIRI Then
    Label1(6).Enabled = True
    txt_giri_da_eseguire.Enabled = True
    Label1(9).Enabled = False
    txt_seconds.Enabled = False
  ElseIf muGame.eCmbTgSelect = CMB_TG_TEMPO Then
    Label1(6).Enabled = False
    txt_giri_da_eseguire.Enabled = False
    Label1(9).Enabled = True
    txt_seconds.Enabled = True
  Else
    Label1(6).Enabled = False
    txt_giri_da_eseguire.Enabled = False
    Label1(9).Enabled = False
    txt_seconds.Enabled = False
  End If
   
End Sub
Private Sub Form_DblClick()
  Call OutputDebugString("W=" & Me.Width & vbCrLf & "H=" & Me.Height)
  If mbSimul = False Then Exit Sub
  'PC HP maximized
  Me.WindowState = vbNormal
  Me.Width = 20730
  Me.Height = 11160
End Sub

Private Sub cmdRestart_Click()
  Call sbPresetParameters
  cmdRestart.Enabled = False
  cmb_quante_macchine.Locked = True
  cmb_tipo_gara.Locked = True
  txt_seconds.Locked = True
  txt_giri_da_eseguire.Locked = True
  Call sbRestart
  If mbSimul = False Then Call sbPlaySound
End Sub

'------------------------------------------------------------------------
'              CHECK MACHINE ENABLE
'              aiMachine 1=A, 2=B
'------------------------------------------------------------------------
Private Function fnIsEnabled(ByVal aiMachine As Integer) As Boolean
  If muGame.bEnable = False Then Exit Function
  fnIsEnabled = True
  If muGame.eCmbQmSelect = CMB_QM_AB Then Exit Function
  If muGame.eCmbQmSelect = CMB_QM_A And aiMachine = 1 Then Exit Function
  If muGame.eCmbQmSelect = CMB_QM_B And aiMachine = 2 Then Exit Function
  fnIsEnabled = False
End Function
'------------------------------------------------------------------------
'              RESTART
'              Enable = false => configurazione
'              Enable = true => inizio nuova gara
'------------------------------------------------------------------------
Private Sub sbRestart(Optional ByVal abEnable As Boolean = True)
  Dim liX As Integer
  muGame.bEnable = abEnable
  For liX = 1 To N_MACCHINE
    Call InitGrid(liX - 1)
    With muMac(liX)
      .lGiri = 0
      .lGiriTarget = MAX_GIRI
      .dTime = 0
      .dLastTime = 0
      .dTimeLatch = 0
      .dAveTime = 0
      .dMaxTime = 0
      .dMinTime = 0
      .l_NGMax = 0
      .l_NGMin = 0
      .dStartTime = 0
      lbl_elap(liX - 1).Caption = vbNullString
      If fnIsEnabled(liX) Then
        .eRun = eWait
        lbl_info(liX - 1).Caption = "Attesa partenza ..."
        lbl_info(liX - 1).ForeColor = vbBlack
        lbl_info(liX - 1).BackColor = vbYellow
        If muGame.eCmbTgSelect = CMB_TG_GIRI Then
          .lGiriTarget = muGame.iMaxCycles
        End If
      Else
        .eRun = eStop
        lbl_info(liX - 1).Caption = vbNullString
        lbl_info(liX - 1).ForeColor = vbBlack
        lbl_info(liX - 1).BackColor = GRAY_COLOR
      End If
    End With
  Next liX
   
  If abEnable = False Then
    Call sbAdvise("Controllare i parametri prima di iniziare", vbBlack, vbYellow)
  Else
    Call sbAdvise
  End If
  Call sbRefresh
End Sub

'------------------------------------------------------------------------
'              INIT/EXIT SERIALE
'------------------------------------------------------------------------
Private Function Comm_Init() As Boolean
  'init
  On Error GoTo local_err
  With MSComm1
    If .PortOpen Then .PortOpen = False
    If mi_com > 0 Then
      .Settings = ms_baud
      .CommPort = mi_com
      .InputLen = 0
      .InputMode = comInputModeBinary
      .RTSEnable = True
      .DTREnable = True
      .RThreshold = 1
      .NullDiscard = False
      .EOFEnable = False
      .Handshaking = comNone
      .PortOpen = True
      Comm_Init = True
    Else
      Call sbAdvise("COM non valida, usare /SETUP", vbBlack, vbRed)
    End If
  End With
  Exit Function
local_err:
  Call sbAdvise("COM" & mi_com & ":" & ms_baud & " ERRORE codice=" & Err.Number, vbBlack, vbRed, Err.Description)
End Function
Private Sub Comm_Exit()
  On Error Resume Next
  With MSComm1
    If .PortOpen Then .PortOpen = False
  End With
End Sub
'----------------------------------------
'  LOAD/SAVE parameters
'----------------------------------------
Private Sub sbLoad()
  Dim ll_height    As Long
  Dim ll_width     As Long
  Dim li_MaxCycles As Integer
  Dim li_WState    As Integer
  Dim ld_seconds   As Double
  Dim lsLine       As String
  
  On Error Resume Next
  'com port settings
  lsLine = ReadConfigFile("SETTINGS", "COM")
  mi_com = Val(lsLine)
  If mi_com = 0 Then
    mi_com = findCOM(lsLine)
    If mi_com = 0 And Len(lsLine) = 0 Then
      mi_com = 7
      Call WriteConfigFile("SETTINGS", "COM", mi_com)
    End If
  End If
  ms_baud = ReadConfigFile("SETTINGS", "BAUD")
  If Len(ms_baud) = 0 Then
    ms_baud = "115200N81"
    Call WriteConfigFile("SETTINGS", "BAUD", ms_baud)
  End If
  
  'game integer parameters
  li_MaxCycles = Val(ReadConfigFile("RDS", "CYCLES"))
  If li_MaxCycles <= 0 Then li_MaxCycles = DEFAULT_GIRI
  If li_MaxCycles > MAX_GIRI Then li_MaxCycles = MAX_GIRI
  txt_giri_da_eseguire.Text = li_MaxCycles
  ld_seconds = ReadConfigFile("RDS", "MAXTIME")
  If ld_seconds <= 0 Then ld_seconds = 600#
  txt_seconds.Text = Format(ld_seconds, TIME_FORMAT)
  
  'combo
  Dim leCmbQmSelect As Integer
  leCmbQmSelect = Val(ReadConfigFile("RDS", "CMB_QM"))
  If leCmbQmSelect < 0 Or leCmbQmSelect >= CMB_QMMAX Then leCmbQmSelect = CMB_QM_A
  cmb_quante_macchine.ListIndex = leCmbQmSelect
  Dim leCmbTgSelect As Integer
  leCmbTgSelect = Val(ReadConfigFile("RDS", "CMB_TG"))
  If leCmbTgSelect < 0 Or leCmbTgSelect >= CMB_TGMAX Then leCmbTgSelect = CMB_TG_UNLIM
  cmb_tipo_gara.ListIndex = leCmbTgSelect
  
  'names
  Dim liX As Integer
  For liX = 1 To N_MACCHINE
    txt_name(liX - 1).Text = ReadConfigFile("RDS", "NAME_" & liX)
  Next liX
  
  'resize
  ll_height = Val(ReadConfigFile("RDS", "H"))
  ll_width = Val(ReadConfigFile("RDS", "W"))
  li_WState = Val(ReadConfigFile("RDS", "WSTATE"))
  If ll_height > 0 Then Me.Height = ll_height
  If ll_width > 0 Then Me.Width = ll_width
  If li_WState >= 0 And li_WState <= 2 Then Me.WindowState = li_WState
End Sub
Private Sub sbSave()
  Call WriteConfigFile("RDS", "H", Me.Height)
  Call WriteConfigFile("RDS", "W", Me.Width)
  Call WriteConfigFile("RDS", "CYCLES", txt_giri_da_eseguire.Text)
  Call WriteConfigFile("RDS", "MAXTIME", txt_seconds.Text)
  Call WriteConfigFile("RDS", "WSTATE", Me.WindowState)
  
  'combo
  Call WriteConfigFile("RDS", "CMB_QM", cmb_quante_macchine.ListIndex)
  Call WriteConfigFile("RDS", "CMB_TG", cmb_tipo_gara.ListIndex)
  'names
  Dim liX As Integer
  For liX = 1 To N_MACCHINE
    Call WriteConfigFile("RDS", "NAME_" & liX, txt_name(liX - 1).Text)
  Next liX
End Sub
'refresh visualizzazione statistiche
Private Sub sbRefresh()
   Dim liX As Integer
   For liX = 1 To N_MACCHINE
     With muMac(liX)
       If .lGiri > 0 Then
         lbl_giri_eseguiti(liX - 1).Caption = .lGiri
         lbl_tempo_giri_eseguiti(liX - 1).Caption = Format(.dTime, TIME_FORMAT) & TIME_SYMBOL
         lbl_tempo_ultimo_giro(liX - 1).Caption = Format(.dLastTime, TIME_FORMAT) & TIME_SYMBOL
         lbl_tempo_medio(liX - 1).Caption = Format(.dAveTime, TIME_FORMAT) & TIME_SYMBOL
         lbl_tempo_max(liX - 1).Caption = Format(.dMaxTime, TIME_FORMAT) & TIME_SYMBOL & NG_SEP1 & .l_NGMax & NG_SEP2
         lbl_tempo_min(liX - 1).Caption = Format(.dMinTime, TIME_FORMAT) & TIME_SYMBOL & NG_SEP1 & .l_NGMin & NG_SEP2
       Else
         lbl_giri_eseguiti(liX - 1).Caption = vbNullString
         lbl_tempo_giri_eseguiti(liX - 1).Caption = vbNullString
         lbl_tempo_ultimo_giro(liX - 1).Caption = vbNullString
         lbl_tempo_medio(liX - 1).Caption = vbNullString
         lbl_tempo_max(liX - 1).Caption = vbNullString
         lbl_tempo_min(liX - 1).Caption = vbNullString
       End If
     End With
   Next liX
End Sub
'---- resize items functions -----
'initialize item's size
Private Sub sbItemsInit()
  Dim ctl         As Control
  Dim sName       As String
  ml_Start_Height = Me.Height
  ml_Start_Width = Me.Width
  ml_ItemsNr = 0
  On Error Resume Next
  For Each ctl In Me.Controls
    sName = TypeName(ctl)
    If fnItemAccepted(sName) Then
      If ml_ItemsNr >= MAX_ITEMS Then
        Call MsgBox("SW ERROR: TOO MANY ITEMS")
        Exit For
      End If
      Call Err.Clear
      ml_ItemsNr = ml_ItemsNr + 1
      With mu_Start_Item(ml_ItemsNr)
         Set .o_ctl = ctl
        .l_Height = ctl.Height
        .l_Left = ctl.Left
        .l_Top = ctl.Top
        .l_Width = ctl.Width
        If sName <> "Shape" Then
          .i_FontSize = ctl.Font.Size
        End If
       If Err.Number <> 0 Then
         Call OutputDebugString(ml_ItemsNr & " INIT ERROR! " & sName)
         Call MsgBox("ITEM ERROR: " & sName)
       End If
      End With
    End If
  Next
  Call OutputDebugString("ml_ItemsNr=" & ml_ItemsNr)
End Sub
'item accepted
Private Function fnItemAccepted(ByVal asName As String) As Boolean
   'Call OutputDebugString("--------[" & asName & "]")
   fnItemAccepted = True
   If asName = "CommandButton" Then Exit Function
   If asName = "Label" Then Exit Function
   If asName = "TextBox" Then Exit Function
   If asName = "ComboBox" Then Exit Function
   If asName = "Shape" Then Exit Function
   If asName = "MSFlexGrid" Then Exit Function
   fnItemAccepted = False
End Function
'resize items
Private Sub Form_Resize()
  Dim l_index As Long
  Dim d_Height As Double
  Dim d_Width  As Double
  Dim d_PHeight As Double
  Dim d_PWidth  As Double
  Dim sName       As String
  
  d_Height = Me.Height
  d_Width = Me.Width
  
  d_PHeight = d_Height / ml_Start_Height
  d_PWidth = d_Width / ml_Start_Width
  
  On Error Resume Next
  For l_index = 1 To ml_ItemsNr
    With mu_Start_Item(l_index)
       Call Err.Clear
       sName = TypeName(.o_ctl)
       .o_ctl.Top = .l_Top * d_PHeight
       .o_ctl.Left = .l_Left * d_PWidth
       .o_ctl.Width = .l_Width * d_PWidth
       If sName <> "ComboBox" Then .o_ctl.Height = .l_Height * d_PHeight
       If sName <> "Shape" Then .o_ctl.Font.Size = fnUpdateFonSize(.i_FontSize, d_PHeight)
       If Err.Number <> 0 Then
         Call OutputDebugString(l_index & " RESIZE ERROR! " & sName)
       End If
    End With
  Next l_index
  miGRID_SCROLL = LinesGrid(0)
  Call OutputDebugString("Grid rows=" & miGRID_SCROLL)
End Sub
'compute font size
Private Function fnUpdateFonSize(ByVal ai_OriginalFontSize, ByVal ad_Scale As Double) As Double
  Dim liX             As Integer
  Dim liFontSize      As Integer
  Dim d_NewFontSize   As Double
  Dim d_Distance      As Double
  Dim d_Distance_Min  As Double
  Dim liIndex         As Integer
  
  Const FONT_SIZES = 6
  Dim d_Font_Size(1 To FONT_SIZES) As Double
  d_Font_Size(1) = 8
  d_Font_Size(2) = 10
  d_Font_Size(3) = 12
  d_Font_Size(4) = 14
  d_Font_Size(5) = 18
  d_Font_Size(6) = 24
  fnUpdateFonSize = ai_OriginalFontSize
  d_NewFontSize = ai_OriginalFontSize * ad_Scale
  
  fnUpdateFonSize = d_NewFontSize
  
'  For liX = 1 To FONT_SIZES
'    d_Distance = Abs(d_NewFontSize - d_Font_Size(liX))
'    If liX = 1 Then
'      liINDEX = liX
'      d_Distance_Min = d_Distance
'    ElseIf d_Distance_Min > d_Distance Then
'      liINDEX = liX
'      d_Distance_Min = d_Distance
'    End If
'  Next liX
'  fnUpdateFonSize = d_Font_Size(liINDEX)
End Function


'---------- ricezione caratteri --------------
Private Sub MSComm1_OnComm()
  Dim llCnt     As Long
  Dim lyFrame() As Byte

  On Error GoTo MSComm1_err
  If mbSimul = True Then Exit Sub  'for safety
  If MSComm1.CommEvent <> comEvReceive Then Exit Sub
  
  Dim ldNow  As Double
  ldNow = fnReadTimer()

  lyFrame = MSComm1.Input
  For llCnt = LBound(lyFrame) To UBound(lyFrame)
    If lyFrame(llCnt) >= &H31 And lyFrame(llCnt) <= &H33 Then
      Call cmdPassaggio(lyFrame(llCnt) - &H30, ldNow)
      Call OutputDebugString(llCnt & " GoodRx: 0x" & Hex(lyFrame(llCnt)))
    Else
      Call OutputDebugString(llCnt & " WrongRx: 0x" & Hex(lyFrame(llCnt)))
    End If
  Next llCnt
  Exit Sub

MSComm1_err:
  Call OutputDebugString("ERRORE codice=" & Err.Number & " - " & Err.Description)
End Sub


'---- visualizzazione tempo gara ----
Private Sub Timer1_Timer()
  Dim lbRefresh As Boolean
  Dim liX       As Integer
  Dim lgElap    As Double
  Dim ldNow     As Double
  
  If muGame.bEnable = False Then Exit Sub
  ldNow = fnReadTimer()
  For liX = 1 To N_MACCHINE
    If muMac(liX).eRun = eRun Then
      lgElap = ldNow - muMac(liX).dStartTime
      lbl_elap(liX - 1).Caption = Format(lgElap, TIME_FORMAT)
      If muGame.eCmbTgSelect = CMB_TG_TEMPO And lgElap >= muGame.dMaxSeconds Then
        'limite a tempo
        lbl_elap(liX - 1).Caption = Format(muGame.dMaxSeconds, TIME_FORMAT)
        muMac(liX).eRun = eStop
        lbl_info(liX - 1).Caption = "Fine corsa"
        lbl_info(liX - 1).ForeColor = vbBlack
        lbl_info(liX - 1).BackColor = vbGreen
        lbRefresh = True
        Call sbRefresh
        Call sbEsameConteggi
      End If
    End If
  Next liX
  If lbRefresh = False Then Exit Sub
  Call sbRefresh
  Call sbEsameConteggi
End Sub

'check all machines stopped
Private Function fnMacAllStopped() As Boolean
  Dim liX   As Integer
  Dim liCnt As Integer
  For liX = 1 To N_MACCHINE
    If muMac(liX).eRun = eStop Then
      liCnt = liCnt + 1
    End If
  Next liX
  If liCnt = N_MACCHINE Then fnMacAllStopped = True
End Function

'----------------------------------------
'  passaggio di almeno una macchina
'----------------------------------------
Private Sub cmdPassaggio(ByVal aiCode As Integer, ByVal adNow As Double)
  Dim liX    As Integer
  
  If fnMacAllStopped() Then Exit Sub
  For liX = 1 To N_MACCHINE
    With muMac(liX)
      If (aiCode And 2 ^ (liX - 1)) = 2 ^ (liX - 1) Then
        If .eRun = eWait Then
          .eRun = eRun
          .dStartTime = adNow
          .lGiri = 0
          Call OutputDebugString("PARTENZA MACCHINA " & Chr(64 + liX))
          lbl_info(liX - 1).Caption = "Macchina in corsa"
          lbl_info(liX - 1).ForeColor = vbWhite
          lbl_info(liX - 1).BackColor = vbBlue
        ElseIf .eRun = eRun Then
          'almeno un giro eseguito
          .dTime = adNow - .dStartTime
          .lGiri = .lGiri + 1
          If .lGiri = 1 Then
            'primo giro
            .dLastTime = .dTime
            .dTimeLatch = .dTime
            .dMaxTime = .dTime
            .dMinTime = .dTime
            .l_NGMax = 1
            .l_NGMin = 1
          Else
            'giri successivi al primo
            .dLastTime = .dTime - .dTimeLatch
            .dTimeLatch = .dTime
            If .dLastTime < .dMinTime Then .dMinTime = .dLastTime: .l_NGMin = .lGiri
            If .dLastTime > .dMaxTime Then .dMaxTime = .dLastTime: .l_NGMax = .lGiri
          End If
          If .lGiri <= MAX_GIRI Then .dTrace(.lGiri) = .dLastTime
          Call WriteGrid(liX - 1, .lGiri, Format(.dLastTime, TIME_FORMAT))
          
          If .lGiri >= .lGiriTarget Or .lGiri >= MAX_GIRI Then
            .eRun = eStop
            lbl_info(liX - 1).Caption = "Fine corsa"
            lbl_info(liX - 1).ForeColor = vbBlack
            lbl_info(liX - 1).BackColor = vbGreen
          End If
          Call OutputDebugString("MACCHINA " & Chr(64 + liX) & " GIRO #" & .lGiri & " TEMPO=" & .dLastTime)
          .dAveTime = .dTime / .lGiri
        End If
      End If
    End With
  Next liX
  Call sbRefresh
  Call sbEsameConteggi
End Sub

'-------------------------------------------------------------------------
'    check fine gara
'-------------------------------------------------------------------------
Private Sub sbEsameConteggi()
 
  If muGame.eCmbTgSelect = CMB_TG_UNLIM Then Exit Sub
  If fnMacAllStopped() = False Then Exit Sub
  
  Call sbStop(True)
  If muGame.eCmbQmSelect <> CMB_QM_AB Then Exit Sub
  
  If muMac(1).lGiri > muMac(2).lGiri Then
    Call sbAdvise("Corsa terminata: vince macchina A")
  ElseIf muMac(1).lGiri < muMac(2).lGiri Then
    Call sbAdvise("Corsa terminata: vince macchina B")
  ElseIf muGame.eCmbTgSelect = CMB_TG_GIRI Then
    'stesso numero di giri
    If muMac(1).dTime < muMac(2).dTime Then
      Call sbAdvise("Corsa terminata: vince macchina A con distacco=" & Format(muMac(2).dTime - muMac(1).dTime, TIME_FORMAT))
    ElseIf muMac(1).dTime > muMac(2).dTime Then
      Call sbAdvise("Corsa terminata: vince macchina B con distacco=" & Format(muMac(1).dTime - muMac(2).dTime, TIME_FORMAT))
    Else
      Call sbAdvise("Corsa terminata: pareggio")
    End If
  Else
    Call sbAdvise("Corsa terminata")
  End If
End Sub

'--------------------------------------------------------------------------------
'     GESTIONE GRID
'--------------------------------------------------------------------------------
Private Sub InitGrid(ByVal aiIndex As Integer)
  Dim i As Integer
  Dim j As Integer
  Const NR_COLS = 1
  Const NR_ROWS = DEFAULT_GIRI
  Const W_LABEL = 1500
  Const W_VALUE = 2000
  
  If aiIndex < 0 Or aiIndex >= N_MACCHINE Then Exit Sub
  On Error Resume Next
  With MyGrid(aiIndex)
    .AllowUserResizing = flexResizeColumns
    .TopRow = 1
    .Cols = NR_COLS + 1
    .Rows = NR_ROWS + 1
    .ColAlignment(0) = 3 '0=sx, 3=center
    .ColWidth(0) = W_LABEL
    .Enabled = True
    .TextMatrix(0, 0) = "#Giro"
    For j = 1 To NR_COLS
      .ColAlignment(j) = 3 '0=sx, 3=center
      .ColWidth(j) = W_VALUE
      .TextMatrix(0, j) = "Tempi"
    Next j
    For i = 1 To NR_ROWS
      .TextMatrix(i, 0) = i
      'NR_COLS
      .TextMatrix(i, 1) = vbNullString
    Next i
  End With
End Sub
Private Sub WriteGrid(ByVal aiIndex As Integer, ByVal aiRow As Integer, ByVal asValue As String)
  If aiIndex < 0 Or aiIndex >= N_MACCHINE Then Exit Sub
  With MyGrid(aiIndex)
    If .Rows <= aiRow Then
      .Rows = aiRow + 1
      .TextMatrix(aiRow, 0) = aiRow
    End If
    .TextMatrix(aiRow, 1) = asValue
    If (aiRow > miGRID_SCROLL) Then
      .TopRow = aiRow - miGRID_SCROLL + 1
    End If
  End With
End Sub
Private Sub RefillGrid()
  Dim liX As Integer
  Dim liY As Integer
  For liX = 1 To N_MACCHINE
    Call InitGrid(liX - 1)
    With muMac(liX)
      For liY = 1 To .lGiri
        Call WriteGrid(liX - 1, liY, Format(.dTrace(liY), TIME_FORMAT))
      Next liY
    End With
  Next liX
End Sub
Private Function LinesGrid(ByVal aiIndex As Integer) As Long
  Dim ll_rh As Long
  Dim ll_h  As Long
  
  On Error GoTo local_err
  With MyGrid(aiIndex)
     ll_h = .Height
     ll_rh = .RowHeight(1)
     LinesGrid = Int(ll_h / ll_rh)
  End With
  LinesGrid = LinesGrid - 1
  If LinesGrid < 1 Then LinesGrid = GRID_SCROLL_DEFAULT
  Exit Function
local_err:
  LinesGrid = GRID_SCROLL_DEFAULT
  Call OutputDebugString("LinesGrid() ERRORE codice=" & Err.Number & " - " & Err.Description)
End Function

'----------------------------------------------------------------------
'  versione 1.1 - classifica tramite lettura di piu' file .pm2
'----------------------------------------------------------------------
Private Sub classifica_Click()
  Dim liX      As Integer
  Dim liY      As Integer
  Dim liMaxGio As Integer
  
  If muGame.bEnable = True Then Exit Sub
  
  Dim lsFile As String
  On Error Resume Next
  CommonDialog2.CancelError = True
  CommonDialog2.FileName = "*." & FILE_EXT
  CommonDialog2.Filter = "PistaMax file (*." & FILE_EXT & ")|*." & FILE_EXT & "|Tutti i file (*.*)|*.*"
  CommonDialog2.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
  CommonDialog2.ShowOpen
  lsFile = CommonDialog2.FileName
  If Len(lsFile) = 0 Then Exit Sub
  If InStr(1, lsFile, "*") > 0 Then Exit Sub
  
  On Error GoTo local_err1
  
  Dim vFiles As Variant
  Dim lFile  As Long
  Dim lFiles As Long
  Dim uFiles() As T_CLASSIFICA
  Dim lsPath   As String
  
  Call sbAdvise
  vFiles = Split(lsFile, Chr(0))
  If UBound(vFiles) < 0 Or LBound(vFiles) <> 0 Then Exit Sub
  
  
  If UBound(vFiles) = 0 Then
    lsPath = fnExtractPath(lsFile)
    If Right(lsPath, 1) <> "\" Then lsPath = lsPath & "\"
    lFiles = 1
    ReDim uFiles(1 To lFiles)
    uFiles(1).sFile = fnExtractFileName(lsFile)
  Else
    lsPath = vFiles(0)
    If Right(lsPath, 1) <> "\" Then lsPath = lsPath & "\"
    lFiles = UBound(vFiles)
    ReDim uFiles(1 To lFiles)
    For lFile = 1 To lFiles
      uFiles(lFile).sFile = vFiles(lFile)
    Next lFile
  End If
  On Error GoTo local_err2
  
  'lettura file
  Dim lbError    As Boolean
  Dim liN        As Integer
  Dim lsLine     As String
  Dim lsSection  As String
  Dim lsFileShow As String
  Dim lsDebug    As String
    
  For lFile = 1 To lFiles
    lsFile = lsPath & uFiles(lFile).sFile
    lsFileShow = "[" & uFiles(lFile).sFile & "]"
    lsDebug = lsDebug & lsFileShow & " "
    Call OutputDebugString(lFile & ") " & lsFile)
    lsSection = "General"
    lsLine = ReadConfigFile(lsSection, "EXT", lsFile)
    If lsLine <> UCase(FILE_EXT) Then
      lsDebug = lsDebug & " FILE NON VALIDO"
      lbError = True
    Else
      For liX = 1 To N_MACCHINE
        lsSection = "REPORT_" & Chr(64 + liX)
        uFiles(lFile).uGiocatore(liX).sNome = UCase(ReadConfigFile(lsSection, "NOME", lsFile))
        uFiles(lFile).uGiocatore(liX).iGiri = Val(ReadConfigFile(lsSection, "GIRI_ESEGUITI", lsFile))
        uFiles(lFile).uGiocatore(liX).dTempo = Val(ReadConfigFile(lsSection, "TEMPO_GIRI_ESEGUITI", lsFile))
        If Len(uFiles(lFile).uGiocatore(liX).sNome) > 0 And uFiles(lFile).uGiocatore(liX).iGiri > 0 Then
          lsDebug = lsDebug & uFiles(lFile).uGiocatore(liX).sNome & "(giri=" & uFiles(lFile).uGiocatore(liX).iGiri & ") "
          If liN = 0 Then liN = uFiles(lFile).uGiocatore(liX).iGiri
          If liN <> uFiles(lFile).uGiocatore(liX).iGiri Then lbError = True
          liMaxGio = liMaxGio + 1
        End If
      Next liX
    End If
    lsDebug = lsDebug & vbCrLf
  Next lFile
  If lbError Or liMaxGio < 2 Then
    'diagnostica errore
    Dim loFrm As frmAvviso
    Set loFrm = New frmAvviso
    Call loFrm.Form_Fill(vbBlack, vbRed, "ERRORE DI COERENZA", lsDebug)
    Call loFrm.Show(vbModal)
    Set loFrm = Nothing
    Exit Sub
  End If
  '--- classifica ---
  'creazione lista concorrenti con tempi migliori
  Dim lbExist As Boolean
  ReDim guList(1 To liMaxGio)
  
  giNList = 0
  For lFile = 1 To lFiles
    For liX = 1 To N_MACCHINE
      If Len(uFiles(lFile).uGiocatore(liX).sNome) > 0 And uFiles(lFile).uGiocatore(liX).iGiri > 0 Then
         'ricerca esistenza ed eventuale inserimento o sostituzione
         lbExist = False
         For liY = 1 To giNList
           If guList(liY).sNome = uFiles(lFile).uGiocatore(liX).sNome Then
             lbExist = True
             If guList(liY).dTempo > uFiles(lFile).uGiocatore(liX).dTempo Then
               guList(liY).dTempo = uFiles(lFile).uGiocatore(liX).dTempo
             End If
             Exit For
           End If
         Next liY
         If lbExist = False Then
           'nuovo giocatore
           giNList = giNList + 1
           guList(giNList).sNome = uFiles(lFile).uGiocatore(liX).sNome
           guList(giNList).dTempo = uFiles(lFile).uGiocatore(liX).dTempo
           guList(giNList).iPos = 0
         End If
      End If
    Next liX
  Next lFile
  
  If giNList = 0 Then
    Call sbAdvise("NESSUN NOME DI GIOCATORE TROVATO", vbBlack, vbRed, Err.Description)
    Exit Sub
  End If
  
  'sort
  Dim liIndex As Integer
  Dim lsSaveN  As String
  Dim ldSaveT  As Double
  
  For liX = 1 To giNList
    liIndex = liX
    For liY = liX + 1 To giNList
      If guList(liY).dTempo < guList(liIndex).dTempo Then
         liIndex = liY
      End If
    Next liY
    If liIndex <> liX Then
      'xchg
      lsSaveN = guList(liX).sNome
      ldSaveT = guList(liX).dTempo
      guList(liX).sNome = guList(liIndex).sNome
      guList(liX).dTempo = guList(liIndex).dTempo
      guList(liIndex).sNome = lsSaveN
      guList(liIndex).dTempo = ldSaveT
    End If
    guList(liX).iPos = 0
  Next liX
  'classifiche
  guList(1).iPos = 1
  For liX = 2 To giNList
    If guList(liX).dTempo = guList(liX - 1).dTempo Then
      guList(liX).iPos = guList(liX - 1).iPos
    Else
      guList(liX).iPos = liX
    End If
  Next liX
  
  For liX = 1 To giNList
    Call OutputDebugString(guList(liX).iPos & " [" & guList(liX).sNome & "] --> " & guList(liX).dTempo)
  Next liX
  
  Dim loFrmC As frmClassifica
  Set loFrmC = New frmClassifica
  Call loFrmC.Show(vbModal)
  Set loFrmC = Nothing
  Exit Sub
  
local_err1:
  Call sbAdvise("ERRORE CREAZIONE LISTA FILE codice=" & Err.Number, vbBlack, vbRed, Err.Description)
  Exit Sub
local_err2:
  Call sbAdvise("ERRORE LETTURA FILE codice=" & Err.Number, vbBlack, vbRed, Err.Description)
  Exit Sub
End Sub

