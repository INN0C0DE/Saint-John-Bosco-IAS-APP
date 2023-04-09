VERSION 5.00
Begin VB.Form frmquiz 
   BackColor       =   &H00FFFFC0&
   Caption         =   "SJB Quiz Bee 2018"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdplay 
      BackColor       =   &H00C0E0FF&
      Caption         =   "PLAY"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H00C0E0FF&
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   1455
      Left            =   9960
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1455
      Left            =   600
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Prepared by: Raphael Arnaldo Cruz"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   7320
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Test your IQ)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3840
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saint John Bosco I.A.S"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   21.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUIZ BEE 2018"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   65.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1320
      TabIndex        =   3
      Top             =   2520
      Width           =   9375
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   600
      Picture         =   "frmquiz.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   1485
      Left            =   9960
      Picture         =   "frmquiz.frx":1C7A42
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1560
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   4815
   End
End
Attribute VB_Name = "frmquiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdplay_Click()
Dim frmquiz As New frmsub
frmsub.Show
Me.Hide
End Sub

Private Sub cmdquit_Click()
Dim frmquiz As New frmwelcome
frmwelcome.Show
Me.Hide
End Sub
