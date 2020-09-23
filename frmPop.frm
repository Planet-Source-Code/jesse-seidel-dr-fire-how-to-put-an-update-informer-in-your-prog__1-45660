VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPop 
   BorderStyle     =   0  'None
   Caption         =   "Program Update!!!"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form2"
   Picture         =   "frmPop.frx":0000
   ScaleHeight     =   2415
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet3 
      Left            =   1560
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   2640
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   840
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Not Right Now"
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   3480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line5 
      X1              =   3480
      X2              =   3480
      Y1              =   2400
      Y2              =   0
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   3480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2400
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Click Here to download it)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Download Update"
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Program Description)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Program Name)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program Update!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo err1:
Label2.Caption = Inet2.OpenURL("http://members.lycos.co.uk/spitfiremsn/prog/progname.txt")
On Error GoTo err2:
Label3.Caption = Inet3.OpenURL("http://members.lycos.co.uk/spitfiremsn/prog/des.txt")
Exit Sub
err1:
MsgBox ("There was an error retrieving the program name!")
err2:
MsgBox ("Error getting file description")
End Sub

Private Sub Label4_Click()
On Error Resume Next
        Dim ie
        Set ie = CreateObject("INTERNETEXPLORER.application") 'creates internal explorer
        ie.navigate Inet1.OpenURL("http://members.lycos.co.uk/spitfiremsn/prog/download.exe") 'Put the address of your program here
        ie.Visible = True
End Sub

Private Sub Label5_Click()
Me.Visible = False
End Sub
