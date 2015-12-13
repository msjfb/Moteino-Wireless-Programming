Attribute VB_Name = "Module1"
Option Explicit
Public giPort As Integer, giTargetNode As Integer
Public glBaud As Long
Public gsFilePath As String
Public gbAutoMode As Boolean
Public gTcpIpMode As Boolean
Public gTcpIpInputLen As Long
Public gTcpIpDatInTxtBuffer As String
Public gTcpIpDatInBinBuffer As Variant
Public gTcpIpError As Long
Public gTcpIpInputMode As Long
Public gsGatewayIpAddress As String
Public glIpPort As Long
Public gsStartupError As String
    
    
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)

Sub Main()
     Dim sCommandLine As String, sArguments() As String
    'CHeck if application is already running
    If App.PrevInstance = True Then
        MsgBox "This application is already running...", vbCritical, "Error"
        End
    End If
    
    sCommandLine = Trim(Command())
    If sCommandLine <> "" Then
        sArguments = Split(sCommandLine, ",")   '<port>,<BaudRate>,<TargetNode>,<FilePath>
        If UBound(sArguments) = 3 Then  'must have 4 arguments
            giTargetNode = Val("" & sArguments(2))
            gsFilePath = "" & sArguments(3)
            If InStr(sArguments(1), ".") <> 0 Then 'REMOTE TCP CONNECTION
                glIpPort = Val("" & sArguments(0))
                gsGatewayIpAddress = sArguments(1)
                If (glIpPort > 0) And (gsGatewayIpAddress <> "") And (giTargetNode > 0) And (gsFilePath <> "") Then
                    gbAutoMode = True
                    gTcpIpMode = True
                Else
                    gsStartupError = "Problem with command line: TcpPort=" & glIpPort & " GatewayIp=" & gsGatewayIpAddress & " Node=" & giTargetNode & " HexFile=" & gsFilePath
                End If
            Else  'LOCAL SERIAL PORT
                giPort = Val("" & sArguments(0))
                glBaud = Val("" & sArguments(1))
                If (giPort > 0) And (glBaud > 0) And (giTargetNode > 0) And (gsFilePath <> "") Then
                    gbAutoMode = True
                Else
                    gsStartupError = "Problem with command line: SerialPort=" & giPort & " BaudRate=" & glBaud & " Node=" & giTargetNode & " HexFile=" & gsFilePath
                End If
            End If
        Else
                gsStartupError = "Problem with command line: " & sCommandLine & ". 4 parameters are required"
        End If
    End If
    
    Load frmAdmin
  
End Sub

Public Function LinesInHexFile(FileName As String) As Long
    Dim buff() As Byte
    Dim hF As Integer
    Dim i As Long, n As Long
    
    hF = FreeFile(0)
    
    Open FileName For Binary Access Read As #hF
    ReDim buff(LOF(hF) - 1)
    Get #hF, , buff()
    Close #hF
    
    For i = 0 To UBound(buff)
        If buff(i) = 58 Then n = n + 1  'count the number of colons (ASCII 58)
    Next
    
    LinesInHexFile = n
End Function

Public Function FileExist(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileExist = Dir$(FileName) <> ""
    If Err.Number <> 0 Then FileExist = False
    
    If FileName = "" Or InStr(FileName, ".") = 0 Then
        FileExist = False
    End If
    On Error GoTo 0
End Function

Function GetIni$(ByVal section As String, ByVal Item As String, _
                 Optional ByVal bUpperCase As Boolean = True)
    Dim sFname As String
    Dim sRetrieved As String
    
    sFname = App.Path & "\MoteinoWP.ini"
    sRetrieved = VBGetPrivateProfileString(section$, Item$, sFname)
    If bUpperCase = True Then
        GetIni$ = UCase$(sRetrieved)
    Else
        GetIni$ = sRetrieved
    End If

End Function
Function VBGetPrivateProfileString(section$, key$, file$) As String
    Dim KeyValue$
    #If Win32 Then
        Dim characters As Long
    #Else
        Dim characters As Integer
    #End If
    
    KeyValue$ = String$(4096, 0)
    
    characters = GetPrivateProfileStringByKeyName(section$, key$, "", KeyValue$, 4095, file$)

    If characters >= 1 Then
        KeyValue$ = Left$(KeyValue$, characters)
    Else
        KeyValue$ = ""
    End If
    
    VBGetPrivateProfileString = KeyValue$
End Function

Function WriteIni$(section$, Item$, Value$)
    Dim sFname As String
    Dim sReturned As String
    
    sFname = App.Path & "\MoteinoWP.ini"
    sReturned = VBWritePrivateProfileString(section$, Item$, Value$, sFname)
    WriteIni$ = UCase$(sReturned)

End Function

Function VBWritePrivateProfileString(section$, key$, Valeur$, file$) As String
    Dim KeyValue$
    #If Win32 Then
        Dim characters As Long
    #Else
        Dim characters As Integer
    #End If
    KeyValue$ = String$(128, 0)
    characters = WritePrivateProfileStringByKeyName(section$, key$, Valeur$, file$)
    KeyValue$ = characters
    VBWritePrivateProfileString = KeyValue$
End Function

Public Sub MyMsgBox(Message As String)
    If gbAutoMode = False Then
        MsgBox "This application is already running...", vbCritical, "Error"
    End If
End Sub
