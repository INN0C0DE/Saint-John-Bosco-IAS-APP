VERSION 5.00
Begin VB.Form frmlogin 
   Caption         =   "Saint John Bosco I.A.S (Application)"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox username 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   3600
      Width           =   5055
   End
   Begin VB.TextBox password 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4800
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(AMPID CAMPUS)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   17.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   4200
      Width           =   4575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   17.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saint John Bosco Institute of Arts and Sciences"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   7695
   End
   Begin VB.Image Image4 
      Height          =   1575
      Left            =   360
      Picture         =   "frmlogin.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   1560
      Left            =   10080
      Picture         =   "frmlogin.frx":1C7A42
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   7815
      Left            =   0
      Picture         =   "frmlogin.frx":343584
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12240
   End
   Begin VB.Image Image1 
      Height          =   7815
      Left            =   0
      Top             =   0
      Width           =   12255
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim user As String
Dim pass As String
user = "Raphael@student.com"
pass = "admin"
If (user = username.Text And pass = password.Text) Then
MsgBox "Congratulations! Login Successful.", vbInformation, "Logging Status:"
frmwelcome.Show
Me.Hide
Else
MsgBox "Sorry, Login Failed.", vbInformation, "Logging Status:"
End If
End Sub

Private Sub Command2_Click()
End
End Sub
