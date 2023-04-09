VERSION 5.00
Begin VB.Form frmscore 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Score"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "COMPUTE AVERAGE GRADE NOW"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   6015
   End
   Begin VB.CommandButton cmdmnu2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CONTINUE?"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   2295
   End
   Begin VB.CommandButton cmdmnu3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   7455
      Left            =   120
      Top             =   120
      Width           =   11655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CORRECT ANSWERS:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "WRONG ANSWERS:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label lblright 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblwrong 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SCORES:"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1455
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmscore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdmnu2_Click()
Dim frmscore As New frmsub
frmsub.Show
Me.Hide
End Sub

Private Sub cmdmnu3_Click()
Dim frmscore As New frmquiz
frmquiz.Show
Me.Hide
End Sub

Private Sub Command1_Click()
Dim frmscore As New frmcal
frmcal.Show
Me.Hide
End Sub
