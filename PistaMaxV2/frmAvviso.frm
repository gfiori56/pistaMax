VERSION 5.00
Begin VB.Form frmAvviso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Avviso"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   Icon            =   "frmAvviso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Chiudi"
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
      Left            =   6240
      TabIndex        =   1
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txt_avviso 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmAvviso.frx":0442
      Top             =   1080
      Width           =   9135
   End
   Begin VB.Label lbl_summary 
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
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   9135
   End
End
Attribute VB_Name = "frmAvviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlFG As Long
Private mlBG As Long
Private msText As String
Private msLong_Text As String

Public Sub Form_Fill(ByVal alFG As Long, ByVal alBG As Long, ByVal asText As String, ByVal asLong_Text As String)
  mlFG = alFG
  mlBG = alBG
  msText = asText
  msLong_Text = asLong_Text
End Sub
Private Sub Form_Load()
  Call sbFormInPrimoPiano(Me)
  lbl_summary.ForeColor = mlFG
  lbl_summary.BackColor = mlBG
  lbl_summary.Caption = msText
  txt_avviso.Text = msLong_Text
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub

