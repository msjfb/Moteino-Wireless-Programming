VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   3420
   ClientLeft      =   3150
   ClientTop       =   2100
   ClientWidth     =   6675
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3420
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   5865
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2400
         TabIndex        =   0
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ProductName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   180
         Width           =   5565
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "App.Major && App.Minor"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   600
         Width           =   5685
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   480
         TabIndex        =   5
         Top             =   1320
         Width           =   4845
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CompanyName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   5445
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   1800
         Width           =   4005
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trademark"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   2160
         Width           =   4005
      End
      Begin VB.Image Image2 
         Height          =   3225
         Left            =   60
         Picture         =   "About.frx":0442
         Stretch         =   -1  'True
         Top             =   120
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err_Handler
    'this is the only line you need to change
    'imgAppIcon.Picture = Me.Icon
                                                                    '
    
    Label1.Caption = App.ProductName
    Label2.Caption = "Version " & App.Major & "." & App.Minor & " Rev " & App.Revision
    Label3.Caption = App.Comments
    Label4.Caption = App.CompanyName
    Label5.Caption = App.LegalCopyright
    Label6.Caption = App.LegalTrademarks
    'also available:  App.Comments, App.Revision, App.Title
    Exit Sub
    
'================
err_Handler:
    Select Case Err
    Case 326 'no data
        Resume Next
    Case Else
        MyMsgBox Error
        End
    End Select
End Sub


