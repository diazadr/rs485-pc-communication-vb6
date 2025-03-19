VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form x 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12120
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFindTxt 
      Caption         =   "Send"
      Height          =   615
      Left            =   7920
      TabIndex        =   6
      Top             =   7200
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Helvetica"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   9015
   End
   Begin VB.TextBox LabelDest 
      Height          =   615
      Left            =   3960
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   7200
      Width           =   2055
   End
   Begin VB.TextBox Pesan 
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6360
      Width           =   5655
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6600
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   9015
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   600
      Picture         =   "Aplikasi Antar PC dengan RS485.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aplikasi Chat Antar PC Dengan RS485"
      BeginProperty Font 
         Name            =   "Helvetica"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005E3717&
      Height          =   570
      Left            =   2400
      TabIndex        =   0
      Top             =   720
      Width           =   8670
   End
End
Attribute VB_Name = "x"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim message, msg_in, msg_out, first, source, destination, last As String
Dim length As Integer

Private Sub cmdFindTxt_Click()
    msg_out = "*" & "1" & LabelDest.Text & Pesan.Text & "#"
    MSComm1.Output = msg_out
    List1.AddItem "Me To PC" + LabelDest.Text + " : " + Pesan.Text
End Sub

Private Sub Form_Load()
    If MSComm1.PortOpen = False Then
       MSComm1.PortOpen = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
    End If
End Sub

Private Sub MSComm1_OnComm()
    msg_in = MSComm1.Input
    length = Len(msg_in)
    first = Left(msg_in, 1)
    source = Mid(msg_in, 2, 1)
    destination = Mid(msg_in, 3, 1)
    message = Mid(msg_in, 4, length - 4)
    last = Right(msg_in, 1)
    
    If first = "*" And last = "#" Then
        If destination = "1" Or destination = "A" Then
        List1.AddItem "From PC" + source + " : " + message
        End If
    End If
End Sub
