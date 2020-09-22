VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin Project1.TxtCtrl TxtCtrl4 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ValidationList  =   "Y,N,T"
      TypeOfTextBox   =   2
   End
   Begin Project1.TxtCtrl TxtCtrl3 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      CharacterCase   =   2
   End
   Begin Project1.TxtCtrl TxtCtrl2 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      TypeOfTextBox   =   2
      CharacterCase   =   1
   End
   Begin Project1.TxtCtrl TxtCtrl1 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      NoOfDecimals    =   2
      TypeOfTextBox   =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Validation Y/N"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Numeric"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Char Lower Case"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Only Character Upper"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TxtCtrl1_GotFocus()
'Set TxtCtrl1.frm = Me
End Sub

