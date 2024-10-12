Attribute VB_Name = "modGlobals"
Option Explicit

Public Const N_MACCHINE   As Integer = 2
Public Const TIME_FORMAT As String = "0.000"
Public Const TIME_SYMBOL As String = ""

'strutture per classifica
Public Type T_GIOCATORE
  sNome  As String
  dTempo As Double
  iGiri  As Integer
End Type
Public Type T_CLASSIFICA
  sFile As String
  uGiocatore(1 To N_MACCHINE) As T_GIOCATORE
End Type
Public Type T_LISTA_GIOCATORI
  iPos   As Integer
  sNome  As String
  dTempo As Double
End Type

Public giNList  As Integer
Public guList() As T_LISTA_GIOCATORI

