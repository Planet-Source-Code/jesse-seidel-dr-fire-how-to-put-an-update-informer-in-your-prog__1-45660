VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Form"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVer 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Rate this code at Planet Source Code"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   360
      Picture         =   "frmMain.frx":0000
      Top             =   960
      Width           =   1965
   End
   Begin VB.Label lblVer 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Program Version:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":1F11
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
lblVer.Caption = App.Major & "." & App.Minor
txtVer.Text = Inet1.OpenURL("http://members.lycos.co.uk/spitfiremsn/prog/progver.txt")
If txtVer.Text = lblVer.Caption Then
Exit Sub
Else
frmPop.Visible = True
Me.WindowState = 1
frmPop.WindowState = 0
End If
End Sub

