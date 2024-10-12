Attribute VB_Name = "modUtilities"
Option Explicit

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Declare Sub OutputDebugString Lib "kernel32" _
  Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
                                          lpvDest As Any, _
                                          lpvSource As Any, _
                                          ByVal cbCopy As Long)

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
 "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
 ByVal lpKeyName As Any, ByVal lpDefault As String, _
 ByVal lpReturnedString As String, ByVal nSize As Long, _
 ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
 "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
 ByVal lpKeyName As Any, ByVal lpString As Any, _
 ByVal lpFileName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMiliseconds As Long)

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
 ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, _
 ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const MAX_LINE_FILE = 2048
Public Const GRAY_COLOR = &H8000000F

'HIGH TIMER
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef Count As Currency) As Boolean
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef Count As Currency) As Boolean
Private mcFrequency As Currency
Private msHomeDir   As String

'HOME DIRECTORY
'<========== TIME FUNCTIONS ===================================================>

Public Sub sbInitializeTimer()
  On Error Resume Next
  Call QueryPerformanceFrequency(mcFrequency)
End Sub
Public Sub sbInitializeHomeDir()
  On Error Resume Next
  msHomeDir = Environ("USERPROFILE")
  If Len(msHomeDir) = 0 Then msHomeDir = CurDir
  If Right(msHomeDir, 1) <> "\" Then msHomeDir = msHomeDir & "\"
End Sub
Public Function fnReadTimer() As Double
  Dim lcCounter As Currency
  Call QueryPerformanceCounter(lcCounter)
  fnReadTimer = lcCounter / mcFrequency
End Function
'----------------------------------------------------------
'    Play BEEP
'----------------------------------------------------------
Public Sub sbPlaySound()
  Call Beep(800, 200)
End Sub

'----------------------------------------------------------
'    Form  in foreground
'----------------------------------------------------------
Public Sub sbFormInPrimoPiano(frm As Form)
    Dim llX    As Long
    Dim llY    As Long
    Dim llCx   As Long
    Dim llCy   As Long
    Dim llFlag As Long
    Dim llR    As Long
    Dim llPos  As Long

    llPos = -1
    llX = 0
    llY = 0
    llCx = 0
    llCy = 0
    llFlag = &H3
    llR = SetWindowPos(frm.hwnd, llPos, llX, llY, llCx, llCy, llFlag)
End Sub

'----------------------------------------------------------------------
'            read a configuration parameter
'----------------------------------------------------------------------
Public Function ReadConfigFile(ByVal asSection As String, _
 ByVal asKey As String, Optional asConfigFile As String = "") As String
 
  Dim lsConfigFile As String
  Dim llStringLength As Long
  Dim lsParaString As String * MAX_LINE_FILE

  ReadConfigFile = ""
  On Error GoTo ReadIniLine_err

  If asConfigFile = "" Then
    lsConfigFile = msHomeDir & App.EXEName & ".ini"
  Else
    lsConfigFile = asConfigFile
  End If
  llStringLength = GetPrivateProfileString(asSection, asKey, "", lsParaString, MAX_LINE_FILE, lsConfigFile)
  ReadConfigFile = Left(lsParaString, llStringLength)
  ReadConfigFile = Trim(ReadConfigFile)

ReadIniLine_exit:
  Exit Function

ReadIniLine_err:
  Resume ReadIniLine_exit
End Function

'----------------------------------------------------------------------
'            save a configuration parameter
'----------------------------------------------------------------------
Public Sub WriteConfigFile(ByVal asSection As String, ByVal asKey As String, _
  ByVal asValue As String, Optional asConfigFile As String = "")
  Dim lsConfigFile As String
  Dim llStringLength As Long
  Dim lsParaString As String

  On Error GoTo WriteConfigFile_err
  If asConfigFile = "" Then
    lsConfigFile = msHomeDir & App.EXEName & ".ini"
  Else
    lsConfigFile = asConfigFile
  End If
  lsParaString = asValue
  Call WritePrivateProfileString(asSection, asKey, lsParaString, lsConfigFile)

WriteConfigFile_exit:
  Exit Sub

WriteConfigFile_err:
  Resume WriteConfigFile_exit
End Sub

'--- input double seconds, output string: "gg hh:mm:ss" -----
Public Function fnTimeElapsed(ByVal adSeconds As Double) As String
   Dim llDD As Long
   Dim llSS As Long
   Dim liHH As Integer
   Dim liMM As Integer
   Dim liSS As Integer
   
   On Error GoTo local_err
   
   llSS = adSeconds
   llDD = llSS \ 86400
   
   llSS = llSS - llDD * 86400
   liHH = llSS \ 3600&
   
   llSS = llSS - liHH * 3600
   liMM = llSS \ 60&
   
   llSS = llSS - liMM * 60
   liSS = llSS
   
   fnTimeElapsed = llDD & " " & Format(liHH, "00") & ":" & Format(liMM, "00") & ":" & Format(liSS, "00")
   Exit Function
local_err:
  fnTimeElapsed = "???"
End Function

'****************************************************************************************************************
' RICONOSCIMENTO COM dal nome del dispositivo USB
'****************************************************************************************************************
Public Function findCOM(ByVal asPattern As String) As Integer
  On Error Resume Next
  
  Dim strComputer As String
  Dim objWMIService As Object
  Dim colItems As Object
  Dim objItem As Object
  Dim strID As String
  Dim comName As String
  Dim strI As String
  
  strComputer = "."
  strID = ""
  Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  
 'Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort")   'In Windows7 a 64 bit non trova le COM Virtuali
  Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPEntity")    'Questa query estrae un contenuto molto ampio, più di 100 stringhe
                                                                             'è necessario scorrerle fino a trovare la com relativa al DATAMAN
  For Each objItem In colItems
    comName = objItem.Name
    If Left(comName, Len(asPattern)) = asPattern Then
      strI = Right(comName, Len(comName) - InStrRev(comName, "(COM"))
      strID = Mid(strI, InStr(strI, "M") + 1, InStr(strI, ")") - (InStr(strI, "M") + 1))
      Call OutputDebugString("FOUND COM" & strID & " " & objItem.Description & vbNewLine)
      Exit For
    End If
  Next
  findCOM = Int(strID)
End Function


'----------------------------------------------------------------------
'  Current date & time to String conversion
'  Short: "YYYYMMDD-hhmmss"
'----------------------------------------------------------------------
Public Function fnNow2Str(ByRef asLongDate As String) As String
  Dim ltNow  As Date
  Dim lsYYYY As String
  Dim lsMM   As String
  Dim lsDD   As String
  Dim lsHH   As String
  Dim lsMN   As String
  Dim lsSS   As String

  ltNow = Now()

  lsYYYY = fnPInt2Str(Year(ltNow), 4)
  lsMM = fnPInt2Str(Month(ltNow), 2)
  lsDD = fnPInt2Str(Day(ltNow), 2)
  lsHH = fnPInt2Str(Hour(ltNow), 2)
  lsMN = fnPInt2Str(Minute(ltNow), 2)
  lsSS = fnPInt2Str(Second(ltNow), 2)

  asLongDate = lsYYYY & "/" & lsMM & "/" & lsDD & "-" _
   & lsHH & ":" & lsMN & ":" & lsSS

  fnNow2Str = lsYYYY & lsMM & lsDD & "-" & lsHH & lsMN & lsSS

End Function

'----------------------------------------------------------------------
'            positive integer conversion into a string
'----------------------------------------------------------------------
Private Function fnPInt2Str(ByVal aiValue As Integer, ByVal aiLen As Integer) As String
  Dim liLen As Integer
  If aiValue > 0 Then
    fnPInt2Str = aiValue
    liLen = Len(fnPInt2Str)
    If liLen < aiLen Then
      fnPInt2Str = String(aiLen - liLen, "0") & fnPInt2Str
    ElseIf liLen > aiLen Then
      fnPInt2Str = String(aiLen, "0")
    End If
  Else
    fnPInt2Str = String(aiLen, "0")
  End If
End Function

'--------------------------------------------------------------------------------
' Name       : fnFileExist()
' Description: check file existance
' Input      : asFileName=file pathname
' Output     : True=exist, False=not exist
'--------------------------------------------------------------------------------
Public Function fnFileExist(asFileName As String) As Boolean
  Dim liAttr As Integer
  On Error GoTo ErrorWay
  If Len(asFileName) > 0 Then
    liAttr = GetAttr(asFileName)
    If (liAttr And vbDirectory) = 0 Then
      fnFileExist = True
    End If
  End If
ErrorWay:
End Function

'--------------------------------------------------------------------------------
' Name       : MyVal
'--------------------------------------------------------------------------------
Public Function MyVal(ByVal adValue As Double, ByVal asSep As String) As String
  MyVal = Str(adValue)
  MyVal = Replace(MyVal, ".", asSep)
End Function

' Get system list separator
Public Function fnGetListSeparator() As String
    Dim RegObj As Object
    Dim RegKey As String
    On Error GoTo local_err
    Set RegObj = CreateObject("WScript.Shell")
    RegKey = RegObj.RegRead("HKEY_CURRENT_USER\Control Panel\International\Slist")
    fnGetListSeparator = RegKey
    Set RegObj = Nothing
    Exit Function
local_err:
End Function

' Get system decimal separator
Public Function fnGetDecimalSeparator() As String
  Dim lgNumber As Single
  Dim lsNumber As String
  fnGetDecimalSeparator = "."
  lgNumber = 1.23
  lsNumber = lgNumber
  If InStr(1, lsNumber, ",") > 0 Then fnGetDecimalSeparator = ","
End Function

'-----------------------------------------------
'extract file name
'-----------------------------------------------
Public Function fnExtractFileName(ByVal asFull As String) As String
  Dim lLen As Long
  Dim lX   As Long
  
  fnExtractFileName = asFull
  lLen = Len(asFull)
  For lX = lLen To 1 Step -1
     If Mid(asFull, lX, 1) = "\" Then
        fnExtractFileName = Right(asFull, lLen - lX)
        Exit For
     End If
  Next lX
End Function

'-----------------------------------------------
'extract path name
'-----------------------------------------------
Public Function fnExtractPath(ByVal asFull As String) As String
  Dim lLen As Long
  Dim lX   As Long
  
  fnExtractPath = vbNullString
  lLen = Len(asFull)
  For lX = lLen To 1 Step -1
     If Mid(asFull, lX, 1) = "\" Then
        fnExtractPath = Left(asFull, lX)
        Exit For
     End If
  Next lX
End Function

