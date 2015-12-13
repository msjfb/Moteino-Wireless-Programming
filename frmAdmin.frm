VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmin 
   Caption         =   "Moteino Wireless Programming"
   ClientHeight    =   9375
   ClientLeft      =   930
   ClientTop       =   345
   ClientWidth     =   15000
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   15000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Moteino Target Node"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   5775
      Begin VB.Frame Frame3 
         Caption         =   "Node Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtTarget 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            MaxLength       =   3
            TabIndex        =   16
            Text            =   "1"
            ToolTipText     =   "Enter the node number of the target Mote"
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Gateway Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   5775
      Begin VB.Frame frLinkStatus 
         Caption         =   "Port"
         Height          =   732
         Index           =   1
         Left            =   4680
         TabIndex        =   30
         Top             =   240
         Width           =   852
         Begin VB.Image YellowLed 
            Height          =   180
            Index           =   1
            Left            =   360
            Picture         =   "frmAdmin.frx":628A
            Top             =   360
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Image RedLed 
            Height          =   180
            Index           =   1
            Left            =   120
            Picture         =   "frmAdmin.frx":636C
            Top             =   360
            Width           =   180
         End
         Begin VB.Image GreenLed 
            Height          =   180
            Index           =   1
            Left            =   600
            Picture         =   "frmAdmin.frx":644E
            Top             =   360
            Visible         =   0   'False
            Width           =   180
         End
      End
      Begin VB.Frame fraTcpIp 
         Caption         =   "Tcp"
         Height          =   1095
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   5415
         Begin VB.Frame IpAddr 
            Caption         =   "IpAddress"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1440
            TabIndex        =   27
            Top             =   240
            Width           =   3615
            Begin VB.TextBox txtIpAddress 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               MaxLength       =   15
               TabIndex        =   29
               ToolTipText     =   "Enter the Ip address of the remote gateway"
               Top             =   240
               Width           =   3015
            End
         End
         Begin VB.Frame fraIpPort 
            Caption         =   "Port No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1215
            Begin VB.TextBox txtIpPort 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               MaxLength       =   5
               TabIndex        =   28
               ToolTipText     =   "Enter the TCP port number of the remote Gateway"
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Connection Type"
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton optTcpIp 
            Caption         =   "Tcp"
            Height          =   255
            Left            =   1680
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optSerial 
            Caption         =   "Serial"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame fraSerial 
         Caption         =   "Serial"
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   5415
         Begin VB.Frame Frame1 
            Caption         =   "Com Port"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1215
            Begin VB.TextBox txtCommPort 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               MaxLength       =   3
               TabIndex        =   21
               Text            =   "1"
               ToolTipText     =   "Select COM Port connected to your Moteino Gateway"
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1440
            TabIndex        =   18
            Top             =   240
            Width           =   2295
            Begin VB.ComboBox cboBaudRate 
               Height          =   315
               ItemData        =   "frmAdmin.frx":6530
               Left            =   240
               List            =   "frmAdmin.frx":6555
               TabIndex        =   19
               Text            =   "Select a Baud rate"
               Top             =   240
               Width           =   1815
            End
         End
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Width           =   4335
      Begin VB.CommandButton cmdStop 
         BackColor       =   &H000000FF&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H0000FF00&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "View Debug"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      TabIndex        =   7
      Top             =   6600
      Width           =   1215
      Begin VB.OptionButton optDebugOff 
         Caption         =   "OFF"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optDebugOn 
         Caption         =   "ON"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ListBox List1 
      Height          =   8055
      Left            =   6120
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   8535
   End
   Begin MSComDlg.CommonDialog DialogTXT 
      Left            =   2640
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      Filter          =   "HEX file (*.hex)|*.hex"
   End
   Begin VB.Frame FrameDebug 
      Caption         =   "HEX Image File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   5805
      Begin VB.CommandButton cmdFindFIle 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   5
         Top             =   360
         Width           =   405
      End
      Begin VB.TextBox txtHEXFile 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "FILE.HEX"
         Top             =   360
         Width           =   5055
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8880
      Visible         =   0   'False
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Status 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   0
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   996
      TabIndex        =   0
      Top             =   9135
      Width           =   15000
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Transfer of HEX Image File to a Moteino Node through a Moteino Gateway"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   7965
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLastUpdate As Date
Private mbFlagStop As Boolean
Private mbDebug As Boolean

Public WithEvents mMoteino As clsMoteino
Attribute mMoteino.VB_VarHelpID = -1

Private Sub cmdFindFIle_Click()
    On Error GoTo err_Handler

    DialogTXT.CancelError = True
    DialogTXT.ShowOpen
    txtHEXFile.Text = DialogTXT.FileName
    Exit Sub
    
'================
err_Handler:
    Select Case Err.Number
        Case 32755 'Cancel button
        Case Else
            MyMsgBox ("The following Error was generated: " & Err.Number & " " & Err.Description)
    End Select
End Sub

Private Sub cmdStart_Click()
    Dim bres As Boolean
    Dim iErrLevel As Integer
    Dim MySettings As String
    
    mbFlagStop = False
    
    gTcpIpInputMode = comInputModeBinary
    
    List1.Clear
    Do
        If gbAutoMode = True Then
            If Not FileExist(gsFilePath) Then
                iErrLevel = 60  'file not found
                Exit Do
            End If
        Else
            glBaud = Val("" & cboBaudRate.Text)
            giPort = Val("" & txtCommPort.Text)
            giTargetNode = Val("" & txtTarget.Text)
            gsFilePath = txtHEXFile.Text
            gsGatewayIpAddress = txtIpAddress.Text
            glIpPort = Val("" & txtIpPort.Text)
        End If
    
        If gsFilePath <> "" Then
            If FileExist(gsFilePath) Then
                cmdStart.Visible = False
                cmdStop.Visible = True
                If gTcpIpMode Then
                    StatusPrint Now & " " & "Opening remote IP Port... "
                    MySettings = gsGatewayIpAddress & "," & Format(glIpPort)  'xxx.xxx.xxx.xxx,yyyyy  (yyyy is port number)
                Else
                    If glBaud = 0 Then
                        cmdStop.Visible = False
                        cmdStart.Visible = True
                        StatusPrint Now & " " & "Please select a baud rate. Application NOT started!"
                        iErrLevel = 10
                        Exit Do
                    End If
                    StatusPrint Now & " " & "Opening Serial Port... "
                    MySettings = Format(glBaud) & "," & Format(giPort) 'xxxxx,yyy  (xxxxx is baud rate, yyy is port number)
                End If
 
                'If mMoteino.Startup(giPort, glBaud) = True Then
                If mMoteino.Startup(MySettings) = True Then
                    If mMoteino.CommOpened = True Then
                        RedLed(1).Visible = False
                        GreenLed(1).Visible = True
                    Else
                        iErrLevel = 20
                    End If
                Else
                    iErrLevel = 20
                End If
                If iErrLevel = 20 Then
                    If gTcpIpMode = True Then
                        MyMsgBox "Cannot open TCPIP port " & MySettings & " Application NOT started!"
                    Else
                        StatusPrint Now & " " & "Unable to Open Com Port. Application NOT started!"
                    End If
                    Exit Do
                End If
            Else
                StatusPrint Now & " " & "HEX file not found. Application STOPPED!"
                iErrLevel = 60
                Exit Do
            End If
        
               
            ' Connecting to Moteino Gateway
            StatusPrint Now & " " & "Connecting to Target node... "
            mMoteino.ConnectToMote (giTargetNode)
            If mMoteino.TargetSet = False Then
                StatusPrint Now & " " & "Cannot set Target on Gateway. Application NOT started!"
                iErrLevel = 30
                Exit Do
            End If
        
            StatusPrint Now & " " & "Processing HEX file " & gsFilePath
            ProgressBar1.Min = 0
            ProgressBar1.Max = LinesInHexFile(gsFilePath)
            ProgressBar1.Visible = True
            
            If ProcessHexFile(gsFilePath) = True Then
                StatusPrint Now & " " & "SUCCESS!!! Upload of HEX file to Gateway complete..."
                iErrLevel = 0
            Else
                If mbFlagStop = True Then
                    StatusPrint Now & " " & "Processing stopped by user. Application STOPPED!"
                ElseIf mMoteino.ImageAccepted = False Then
                    StatusPrint Now & " " & "HANDSHAKE NAK [IMG REFUSED BY TARGET]. Application STOPPED!"
                    iErrLevel = 40
                Else
                    StatusPrint Now & " " & "Problem in processing HEx file. Application STOPPED!"
                    iErrLevel = 50
                End If
            End If
        Else
            StatusPrint Now & " " & "HEX file not found. Application STOPPED!"
            iErrLevel = 60
        End If
        Exit Do  'dummy loop
    Loop

    Screen.MousePointer = vbDefault

    mMoteino.Shutdown
    ProgressBar1.Visible = False
    cmdStop.Visible = False
    cmdStart.Visible = True
    If mMoteino.CommOpened = False Then
        RedLed(1).Visible = True
        GreenLed(1).Visible = False
    End If

    If gbAutoMode = True Then
        Unload Me
        ExitProcess iErrLevel
    End If
End Sub

Private Sub StatusPrint(msg$)
    '-- This routine prints a message in the status box
    '   at the bottom of the form.
    Static Last$
    With Me
      !Status.Cls
      If Len(msg$) = 0 Then
        !Status.Print " " & Last$
      Else
        !Status.Print " " & msg$
      End If
    End With
    Last$ = msg$
    DoEvents
End Sub

Private Sub cmdStop_Click()
    mbFlagStop = True
End Sub

Private Sub Form_Load()
    Set mMoteino = New clsMoteino
    frmAdmin.Width = 8550

    If gbAutoMode = True Then
        txtHEXFile.Text = gsFilePath
        cboBaudRate.Text = Format(glBaud)
        txtCommPort.Text = Format(giPort)
        txtIpAddress.Text = gsGatewayIpAddress
        txtIpPort.Text = Format(glIpPort)
        txtTarget.Text = Format(giTargetNode)
        Me.Show
        Me.WindowState = vbMinimized
        Call cmdStart_Click
    Else
        txtHEXFile.Text = GetIni("FilePath", "HexFile")
        cboBaudRate.Text = Val("" & GetIni("Serial", "BaudRate"))
        txtCommPort.Text = Val("" & GetIni("Serial", "Port"))
        txtIpAddress = GetIni("TCP", "IpAddress")
        txtIpPort = Val("" & GetIni("TCP", "Port"))
        Me.Show
    End If
    If gsStartupError <> "" Then
        Call optDebugOn_Click 'to get full width for error message
        optDebugOn.Value = True
        StatusPrint gsStartupError
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next

    If gbAutoMode = False Then
        WriteIni "FilePath", "HexFile", txtHEXFile.Text
        WriteIni "Serial", "BaudRate", cboBaudRate.Text
        WriteIni "Serial", "Port", txtCommPort.Text
        WriteIni "TCP", "IpAddress", txtIpAddress.Text
        WriteIni "TCP", "Port", txtIpPort.Text
    End If
    
    Select Case UnloadMode
    Case vbFormControlMenu
        If MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            Cancel = False
            Call cmdStop_Click
            End
        Else
            Cancel = True
        End If
    Case vbFormCode
        Cancel = False
    Case vbAppWindows
        Cancel = False
        Call cmdStop_Click
        End
    End Select
    On Error GoTo 0
End Sub

Private Sub mMoteino_Received(Chars As String)
    'Event lifted by the class for incoming characteers
    'No processing done here, only display in Debug list
    List1.AddItem Chars
End Sub

Private Sub mMoteino_Sent(Chars As String)
    'event lifted by class for outgoing characters
    'No processing done here, only display in Debug list
    List1.AddItem Chars
End Sub

Private Function ProcessHexFile(FileName As String) As Boolean
    Dim sLine As String
    Dim lSeq As Long
    Dim sCommand As String
    Dim iNumFile As Integer, iRetry As Integer

    On Error GoTo err_Handler

    'Filename = txtHEXFile.Text
    iNumFile = FreeFile
    Open FileName For Input As #iNumFile
     
    'Initial Handshake
    mMoteino.WaitForHandshake (False)
    If mMoteino.GotHandshake = False Then
        ProcessHexFile = False
        Close #iNumFile
        Exit Function
    End If
    lSeq = 0
    
    'Send file line by line
    Do While Not EOF(iNumFile)
        DoEvents
        If mbFlagStop Then
            Exit Do
        End If
        ProgressBar1.Value = IIf(lSeq > ProgressBar1.Max, ProgressBar1.Max, lSeq)
        Line Input #iNumFile, sLine
        For iRetry = 1 To 3
            If InStr(sLine, "00000001FF") <> 0 Then 'last record in file. Do not send it but send Handshake
                StatusPrint Now & " " & "End of file. Waiting for final Handshake..."
                mMoteino.WaitForHandshake (True)
                If mMoteino.GotHandshake = True Then
                    If mMoteino.ImageAccepted Then
                        ProcessHexFile = True  'SUCCESS!!!
                        Exit Do
                    Else
                        'Error, exit with ProcessHexFile=False
                    End If
                    Exit Do
                End If
            Else
                mMoteino.SendLine lSeq, sLine
                If mMoteino.GotReply = True Then
                    If mMoteino.ReceivedSeq = lSeq Then
                        lSeq = lSeq + 1
                        Exit For
                    Else
                        'out of sequence
                        'retry sending same line
                        StatusPrint Now & " " & "out of sequence"
                    End If
                Else   ' error while sending line
                       ' retry sending same line
                      StatusPrint Now & " " & "sent line = false"
                End If
            End If
       Next iRetry
       If iRetry > 3 Then 'maxed out retries, exit with ProcessHexFile=False
           Exit Do
       End If
       
    Loop

    Close #iNumFile
    Exit Function
'================
err_Handler:
    Close #iNumFile
    ProcessHexFile = False
    MyMsgBox "The following Error was generated: " & Err.Number & " " & Err.Description


End Function

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    End
End Sub

Private Sub optDebugOff_Click()
    frmAdmin.Width = 8550
    mbDebug = False
    List1.Visible = False
End Sub

Private Sub optDebugOn_Click()
    frmAdmin.Width = 15285
    mbDebug = True
    List1.Visible = True
End Sub

Private Sub optSerial_Click()
    gTcpIpMode = False
    fraTcpIp.Visible = False
    fraSerial.Visible = True
End Sub

Private Sub optSerial_DblClick()
    gTcpIpMode = False
    fraTcpIp.Visible = False
    fraSerial.Visible = True
End Sub

Private Sub optTcpIp_Click()
    gTcpIpMode = True
    fraTcpIp.Visible = True
    fraSerial.Visible = False
    
End Sub

Private Sub optTcpIp_DblClick()
    gTcpIpMode = True
    fraTcpIp.Visible = True
    fraSerial.Visible = False
End Sub

Private Sub txtTarget_Validate(Cancel As Boolean)
If Val(txtTarget.Text) > 255 Then
    StatusPrint "ERROR! Target Node must be between 1-255"
    Cancel = True
End If

End Sub
