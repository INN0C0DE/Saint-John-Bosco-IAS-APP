VERSION 5.00
Begin VB.Form frmave 
   BackColor       =   &H00FFFF80&
   Caption         =   "Quiz Average Grade"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmdcompute1 
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   4335
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
      TabIndex        =   0
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label lblsci 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   5040
      TabIndex        =   12
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label lblmath 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label lblfil 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lbleng 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   5040
      TabIndex        =   9
      Top             =   1320
      Width           =   3615
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
      Caption         =   "QUIZ AVERAGE GRADE"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      TabIndex        =   8
      Top             =   360
      Width           =   7815
   End
   Begin VB.Label lblaverage1 
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
      Left            =   4440
      TabIndex        =   7
      Top             =   4800
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "AVERAGE:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   0
      Left            =   2400
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
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
      Left            =   1680
      TabIndex        =   5
      Top             =   4080
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
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   2535
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "frmave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdcompute1_Click()
Dim average As Integer
Dim units As Integer
Const math = 4
Const science = 3
Const english = 3
Const filipino = 3
units = math + science + english + filipino
average = ((Val(lblfil) * filipino + Val(lbleng) * english + Val(lblsci) * science + Val(lblmath) * math)) / units
lblaverage1.Caption = average
End Sub

