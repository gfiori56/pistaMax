VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmClassifica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Classifica"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   Icon            =   "frmClassifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9375
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
      Left            =   5760
      TabIndex        =   1
      Top             =   4080
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid CGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   4
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
End
Attribute VB_Name = "frmClassifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
  Call sbFormInPrimoPiano(Me)
  
  Dim i As Integer
  Dim j As Integer
  Const NR_COLS = 3
  Const W_LABEL = 1000
  Const W_NAME = 4000
  Const W_VALUE = 2000
  
  On Error Resume Next
  With CGrid
    .Width = W_LABEL + W_NAME + W_VALUE + W_VALUE + 100
    .AllowUserResizing = flexResizeColumns
    .TopRow = 1
    .Cols = NR_COLS + 1
    .Rows = giNList + 1
    
    For j = 0 To NR_COLS
      .ColAlignment(j) = 3 '0=sx, 3=center
    Next j
    .Enabled = True
    
    .TextMatrix(0, 0) = "Pos."
    .TextMatrix(0, 1) = "Giocatore"
    .TextMatrix(0, 2) = "Tempo"
    .TextMatrix(0, 3) = "Distacco"
    
    .ColWidth(0) = W_LABEL
    .ColWidth(1) = W_NAME
    .ColWidth(2) = W_VALUE
    .ColWidth(3) = W_VALUE
    
    For i = 1 To giNList
      .TextMatrix(i, 0) = i
      'NR_COLS+1
      .TextMatrix(i, 0) = guList(i).iPos
      .TextMatrix(i, 1) = guList(i).sNome
      .TextMatrix(i, 2) = Format(guList(i).dTempo, TIME_FORMAT)
      .TextMatrix(i, 3) = Format(guList(i).dTempo - guList(1).dTempo, TIME_FORMAT)
    Next i
  End With
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub

