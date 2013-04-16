VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Astable multivibrator -Winston SRIT"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12795
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   4320
      TabIndex        =   33
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   720
      TabIndex        =   29
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Height          =   405
      Left            =   4080
      TabIndex        =   28
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   405
      Left            =   4920
      TabIndex        =   27
      ToolTipText     =   "Enter here"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   405
      Left            =   4920
      TabIndex        =   23
      ToolTipText     =   "Enter here"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   4920
      TabIndex        =   22
      ToolTipText     =   "Enter here"
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   4920
      TabIndex        =   21
      ToolTipText     =   "Enter here"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   1800
      TabIndex        =   13
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   240
      TabIndex        =   10
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      ToolTipText     =   "Enter here"
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Enter here"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Enter here"
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "(kHz)"
      Height          =   255
      Left            =   4200
      TabIndex        =   32
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "(kHz)"
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   26
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Rf in k"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   25
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Frequency"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   24
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   20
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   19
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Enter C in"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   17
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "Enter R2 in k"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   15
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Enter R1 in k"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   3330
      Left            =   3480
      Picture         =   "Form2.frx":0000
      ToolTipText     =   "Astable multivibrator using OP-AMP with 50 % duty cycle"
      Top             =   120
      Width           =   3330
   End
   Begin VB.Label Label8 
      Caption         =   "Frequency"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   12
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Duty Cycle(%)"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Duty cycle is the ratio of on time to sum of on time and off time"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "Enter C in"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   3840
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "Enter R2 in k"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Enter R1 in k"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   120
      Picture         =   "Form2.frx":343F
      ToolTipText     =   "Astable multivibrator with 555"
      Top             =   240
      Width           =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
R1 = Val(Text1.Text)
R2 = Val(Text2.Text)
c = Val(Text3.Text)
f = 1.44 / ((R1 + 2 * R2) * c)
d = (R1 + R2) / (R1 + 2 * R2)
Text5.Text = f
Text4.Text = d * 100
End Sub

Private Sub Command2_Click()
R11 = Val(Text6.Text)
R22 = Val(Text7.Text)
rf = Val(Text9.Text)
c1 = Val(Text8.Text)
T = 2 * rf * c1 * Log((2 * R22 + R11) / R11)
Text11.Text = 1 / T
End Sub

Private Sub Command3_Click()
T = MsgBox("Written by Jabez Winston.C,2nd year EEE,SRIT", 0, "About")
End Sub
