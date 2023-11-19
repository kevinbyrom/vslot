VERSION 5.00
Begin VB.Form frmScoring 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scoring"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Reach 75 dollars to win!"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   4455
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   1200
      Picture         =   "frmScoring.frx":0000
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "50 Dollars"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "35 Dollars"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "25 Dollars"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "15 Dollars"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "10 Dollars"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "4 Dollars"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   1200
      Picture         =   "frmScoring.frx":7572
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   1200
      Picture         =   "frmScoring.frx":EAE4
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   1200
      Picture         =   "frmScoring.frx":16056
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   1200
      Picture         =   "frmScoring.frx":1D5C8
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1200
      Picture         =   "frmScoring.frx":24B3A
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "3 Slots of..."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PRIZE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1 Dollar"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Line Line3 
      X1              =   2400
      X2              =   2400
      Y1              =   600
      Y2              =   5280
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   4560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      Caption         =   "CONDITION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1200
      Picture         =   "frmScoring.frx":2C0AC
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "If 1st Slot ="
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ways To Score..."
      BeginProperty Font 
         Name            =   "Carlisle"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmScoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

