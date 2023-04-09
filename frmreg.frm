VERSION 5.00
Begin VB.Form frmreg 
   BackColor       =   &H00FFFF80&
   Caption         =   "Registration Form"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14160
   BeginProperty Font 
      Name            =   "Lucida Fax"
      Size            =   9
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txtcnog 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   22
      Top             =   6000
      Width           =   4575
   End
   Begin VB.TextBox Txtnog 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   19
      Top             =   5280
      Width           =   4575
   End
   Begin VB.TextBox Txtcp 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   16
      Top             =   4560
      Width           =   4575
   End
   Begin VB.TextBox Txtadd 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   3840
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "MENU"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8160
      Width           =   2655
   End
   Begin VB.TextBox Txtname 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox Txtage 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Width           =   4575
   End
   Begin VB.TextBox Txtbirthday 
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   3120
      Width           =   4575
   End
   Begin VB.CommandButton cmdrefresh 
      BackColor       =   &H0080C0FF&
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   3255
   End
   Begin VB.CommandButton cmdsubmit 
      BackColor       =   &H0080FF80&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   9015
      Left            =   120
      Top             =   120
      Width           =   13935
   End
   Begin VB.Label Lblcnog 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      TabIndex        =   24
      Top             =   6000
      Width           =   4815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No. of Guardian:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   23
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Lblnog 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   21
      Top             =   5280
      Width           =   4815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Guardian:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label Lblcp 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   18
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Cellphone #:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   17
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Lbladd 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   15
      Top             =   3840
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT PROFILE REGISTRATION"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   840
      TabIndex        =   11
      Top             =   480
      Width           =   12615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Lblname 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   7
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Lblage 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   6
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label Lblbirthday 
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   5
      Top             =   3120
      Width           =   4815
   End
End
Attribute VB_Name = "frmreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdrefresh_Click()
Txtname.Enabled = True
Txtage.Enabled = True
Txtbirthday.Enabled = True
Txtadd.Enabled = True
Txtcp.Enabled = True
Txtnog.Enabled = True
Txtcnog.Enabled = True
cmdsubmit.Enabled = True
End Sub

Private Sub cmdsubmit_Click()
If Txtname = "" And Txtage = "" And Txtbirthday = "" Then
MsgBox "YOU NEED TO FILL UP THE NEEDED FIELDS", vbOKOnly = vbInformation, "Registration"
Else
If MsgBox("Are you sure do you want to submit your answer?", vbYesNo + vbQuestion, "Submitting....") = vbYes Then
Lblname.Caption = Txtname.Text
Lblage.Caption = Txtage.Text
Lblbirthday.Caption = Txtbirthday.Text
Lbladd.Caption = Txtadd.Text
Lblcp.Caption = Txtcp.Text
Lblnog.Caption = Txtnog.Text
Lblcnog.Caption = Txtcnog.Text
Txtname.Enabled = False
Txtage.Enabled = False
Txtbirthday.Enabled = False
Txtadd.Enabled = False
Txtcp.Enabled = False
Txtnog.Enabled = False
Txtcnog.Enabled = False
cmdsubmit.Enabled = False
End If
End If
End Sub

Private Sub Command1_Click()
Dim frmreg As New frmwelcome
frmwelcome.Show
Me.Hide
End Sub

