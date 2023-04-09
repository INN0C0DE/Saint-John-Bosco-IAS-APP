VERSION 5.00
Begin VB.Form frmwelcome 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Saint John Bosco I.A.S (Application)"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Average Grade Calculator"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton cmdquiz 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quiz Bee 2018"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "What would you like to do now?"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   11535
      Begin VB.CommandButton cmdlogout 
         BackColor       =   &H008080FF&
         Caption         =   "LOGOUT"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4440
         Width           =   3255
      End
      Begin VB.CommandButton cmdreg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Student Profile Registration"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Test your IQ)"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Compute your grades FASTER)"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   9
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Register now!)"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   1800
         Width           =   3855
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Register now!)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Thank You for Logging In...)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome, BOSCONIAN!"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   36
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   855
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmwelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcal_Click()
Dim frmwelcome As New frmcal
frmcal.Show
Me.Hide
End Sub


Private Sub cmdlogout_Click()
Dim frmwelcome As New frmlogin
frmlogin.Show
Me.Hide
End Sub

Private Sub cmdquiz_Click()
Dim frmwelcome As New frmquiz
frmquiz.Show
Me.Hide
End Sub

Private Sub cmdreg_Click()
Dim frmwelcome As New frmreg
frmreg.Show
Me.Hide
End Sub
