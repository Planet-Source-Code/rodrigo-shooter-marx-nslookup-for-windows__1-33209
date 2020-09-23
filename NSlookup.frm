VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "NSLookUp By RShooter - ICQ 49265623"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   Icon            =   "NSlookup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   365
      Left            =   1200
      TabIndex        =   5
      Top             =   1920
      Width           =   2500
   End
   Begin VB.TextBox Text3 
      Height          =   365
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get &Host"
      Height          =   365
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   1000
   End
   Begin VB.TextBox Text2 
      Height          =   365
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      Height          =   365
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Get IP"
      Height          =   365
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   1000
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "Host"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Host"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   Dim sHostName As String
   
   If SocketsInitialize() Then
   
     'pass the host address to the function
      sHostName = Text1.Text
      Text2.Text = GetIPFromHostName(sHostName)
      
      SocketsCleanup
      
   Else
   
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "is not successfully responding."
   
   End If
   

   
End Sub

Private Sub Command2_Click()
Text4.Text = GetHostNameFromIP(Text3.Text)
   
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub
