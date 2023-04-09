VERSION 5.00
Begin VB.Form frmcal 
   BackColor       =   &H00FFFF80&
   Caption         =   "Average Grade Calculator"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtfil 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   11
      Text            =   "0"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox Txtenglish 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Text            =   "0"
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Txtmath 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Text            =   "0"
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox Txtscience 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Text            =   "0"
      Top             =   3960
      Width           =   3735
   End
   Begin VB.CommandButton Cmdcompute 
      BackColor       =   &H0080C0FF&
      Caption         =   "COMPUTE"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Filipino:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   12
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   7575
      Left            =   120
      Top             =   120
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Average Grade Calculator"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   720
      TabIndex        =   9
      Top             =   240
      Width           =   10695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "English Grade:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Math Grade:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Science Grade:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "AVERAGE:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblaverage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   4680
      TabIndex        =   4
      Top             =   4800
      Width           =   4455
   End
End
Attribute VB_Name = "frmcal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdcompute_Click()
Dim average As Integer
Dim units As Integer
Const math = 4
Const science = 3
Const english = 3
Const filipino = 3
units = math + science + english + filipino
average = ((Val(txtfil) * filipino + Val(Txtenglish) * english + Val(Txtscience) * science + Val(Txtmath) * math)) / units
lblaverage.Caption = average

End Sub

Private Sub Command1_Click()
Dim frmcal As New frmwelcome
frmwelcome.Show
Me.Hide
End Sub
