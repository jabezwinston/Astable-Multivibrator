VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Astable multivibrator"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   6120
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3720
      TabIndex        =   6
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "Enter value of C in micro farad"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "Enter the value of R2 in kilo-ohm"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      ToolTipText     =   "Enter the value of R1 in kilo-ohm"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Winston,SRIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Duty Cycle(%)"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Frequency(kHz)"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "C(uF)"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "R2(k-ohm)"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "R1(k-ohm)"
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   120
      Picture         =   "Form1.frx":0000
      ToolTipText     =   "Astable Multivibrator using IC555"
      Top             =   120
      Width           =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub VScroll1_Change()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Command1_Click()
r1 = Val(Text1.Text)
r2 = Val(Text2.Text)
c = Val(Text3.Text)
f = 1.44 / ((r1 + 2 * r2) * c)
d = (r1 + r2) / (r1 + 2 * r2)
Text4.Text = f
Text5.Text = d * 100
End Sub

