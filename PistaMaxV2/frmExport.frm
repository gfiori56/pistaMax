VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametri di esportazione"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   ControlBox      =   0   'False
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4785
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Conferma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ComboBox cmb_sep 
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
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1680
      Width           =   4215
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
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'valori di return
Public msDecSep As String
Public msListSep As String

'questo ordine corrisponde a quello della combo
Private Enum E_SEP
  CMB_SEP_SYSTEM
  CMB_SEP_TAB_POINT
  CMB_SEP_TAB_COMMA
  CMB_SEP_COMMA_POINT
  CMB_SEP_SEMICOLON_POINT
  CMB_SEP_SEMICOLON_COMMA
  CMB_SEPMAX 'always in the last position
End Enum

Private Type T_SEP
  sMsg    As String
  sLstSep As String
  sDecSep As String
End Type
Private muCnf(0 To CMB_SEPMAX - 1) As T_SEP
Private INFO_MSG As String

Private Sub Form_Load()
  Dim liX                      As Integer
  
  msDecSep = vbNullString
  msListSep = vbNullString
  INFO_MSG = "Scegliere i separatori di" & vbCrLf & "Lista / Cifre decimali"
  
  Call sbFormInPrimoPiano(Me)
  Call sbAdvise(INFO_MSG)
  
  With muCnf(CMB_SEP_SYSTEM)
    .sMsg = "Impostazioni di sistema"
    .sLstSep = fnGetListSeparator()
    .sDecSep = fnGetDecimalSeparator()
  End With
  With muCnf(CMB_SEP_TAB_POINT)
    .sMsg = "TAB / Punto"
    .sLstSep = Chr(9)
    .sDecSep = "."
  End With
  With muCnf(CMB_SEP_TAB_COMMA)
    .sMsg = "TAB / Virgola"
    .sLstSep = Chr(9)
    .sDecSep = ","
  End With
  With muCnf(CMB_SEP_COMMA_POINT)
    .sMsg = "Virgola / Punto"
    .sLstSep = ","
    .sDecSep = "."
  End With
  With muCnf(CMB_SEP_SEMICOLON_POINT)
    .sMsg = "Punto e Virgola / Punto"
    .sLstSep = ";"
    .sDecSep = "."
  End With
  With muCnf(CMB_SEP_SEMICOLON_COMMA)
    .sMsg = "Punto e Virgola / Virgola"
    .sLstSep = ";"
    .sDecSep = ","
  End With
  
  Call cmb_sep.Clear
  For liX = 0 To CMB_SEPMAX - 1
    With muCnf(liX)
      Call cmb_sep.AddItem(.sMsg)
      Call OutputDebugString(.sMsg & " LIST[" & .sLstSep & "] DECIMAL[" & .sDecSep & "]")
    End With
  Next liX
  
  Dim leCmbSelect As E_SEP
  leCmbSelect = Val(ReadConfigFile("RDS", "CMB_SEP"))
  If leCmbSelect < 0 Or leCmbSelect >= CMB_SEPMAX Then leCmbSelect = CMB_SEP_SYSTEM
  cmb_sep.ListIndex = leCmbSelect
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
'        eventi azioni
'--------------------------------------------
Private Sub cmb_sep_Click()
  Call sbAdvise(INFO_MSG)
End Sub
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim leCmbSelect As E_SEP
  leCmbSelect = cmb_sep.ListIndex
  If leCmbSelect < 0 Or leCmbSelect >= CMB_SEPMAX Then
    Call sbAdvise("Scelta non valida", vbBlack, vbRed)
    Exit Sub
  End If
  
  If muCnf(leCmbSelect).sLstSep = vbNullString Or _
   muCnf(leCmbSelect).sDecSep = vbNullString Or _
   muCnf(leCmbSelect).sLstSep = muCnf(leCmbSelect).sDecSep Then
    Call sbAdvise("Separatori non validi", vbBlack, vbRed)
    Exit Sub
  End If
  
  msListSep = muCnf(leCmbSelect).sLstSep
  msDecSep = muCnf(leCmbSelect).sDecSep
  Call WriteConfigFile("RDS", "CMB_SEP", leCmbSelect)
  Unload Me
End Sub


