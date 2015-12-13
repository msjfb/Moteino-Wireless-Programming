VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmComm 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSCommLib.MSComm MSComm1 
      Left            =   720
      Top             =   360
      _ExtentX        =   794
      _ExtentY        =   794
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      RThreshold      =   1
      InputMode       =   1
   End
End
Attribute VB_Name = "frmComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
