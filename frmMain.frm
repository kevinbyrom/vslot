VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB Slots by Kevin Byrom"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   4905
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdScoring 
      Caption         =   "&Scoring"
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Carlisle"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmMain.frx":E1042
      Top             =   360
      Width           =   6615
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End Game"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdSpin 
      Caption         =   "&Spin"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.PictureBox Slot 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1560
      Index           =   2
      Left            =   5880
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   1800
      Width           =   1560
   End
   Begin VB.PictureBox Slot 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1560
      Index           =   1
      Left            =   4080
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   1
      Top             =   1800
      Width           =   1560
   End
   Begin VB.PictureBox Slot 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1560
      Index           =   0
      Left            =   2280
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   1800
      Width           =   1560
   End
   Begin VB.Label lblPrize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   6615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub UpdateMoney()

    frmMain.txtDisplay.Text = "Current Money : $ " & CurrentMoney
    
End Sub


Private Sub cmdEnd_Click()

    End
    
End Sub

Private Sub cmdScoring_Click()

    frmScoring.Show 1 'modal
    
End Sub

Private Sub cmdSpin_Click()

    Screen.MousePointer = 11 'HourGlass
    
    Dim i, ii As Integer
    Dim prize As Integer
    
    CurrentMoney = CurrentMoney - 1
    
    For i = 0 To 2
        vSlot(i).value(1) = vSlot(i).value(0)
        vSlot(i).value(0) = Int(6 * Rnd)
        vSlot(i).value(-1) = vSlot(i).value(0) + 1
        If vSlot(i).value(-1) > 5 Then vSlot(i).value(-1) = 0
    Next
    
    vSlot(0).speed = (Int(2 * Rnd) + 1) * 10
    vSlot(1).speed = (Int(2 * Rnd) + 1) * 20
    vSlot(2).speed = (Int(2 * Rnd) + 1) * 30
    
    lblPrize.Caption = ""
    
    For i = 1 To 60
        For ii = 0 To 2
            'stop slots in order
            If i = 21 Then vSlot(0).speed = 0
            If i = 41 Then vSlot(1).speed = 0
            
            'move slots
            vSlot(ii).y = vSlot(ii).y + vSlot(ii).speed
            If vSlot(ii).y >= 100 Then
                vSlot(ii).y = vSlot(ii).y - 100
                vSlot(ii).value(1) = vSlot(ii).value(0)
                vSlot(ii).value(0) = vSlot(ii).value(-1)
                vSlot(ii).value(-1) = vSlot(ii).value(-1) + 1
                If vSlot(ii).value(-1) = 6 Then vSlot(ii).value(-1) = 0
            End If
            DrawSlots
            DoEvents
        Next
    Next
    
    prize = CalcPrize
    If prize > 0 Then
        lblPrize.Caption = "You won " & prize & " dollar(s)!"
        CurrentMoney = CurrentMoney + prize
        UpdateMoney
    End If
     
    UpdateMoney
    
    Screen.MousePointer = 0 'Default
    
    If CurrentMoney = 0 Then
        frmMain.Hide
        frmLose.Show
    ElseIf CurrentMoney >= 75 Then
        frmMain.Hide
        frmWin.Show
    End If
End Sub

Private Sub Form_Load()

    Dim res As Integer
    
    Randomize Timer
    
    If Screen.Width <> (640 * Screen.TwipsPerPixelX) Then
        res = MsgBox("You must have the display set to 640x480 pixels to run this game!", vbOKOnly, "Cannot continue")
        End
    End If
    
    frmTitle.Show MODAL
    
    CurrentMoney = 20
    UpdateMoney
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
    
End Sub


Private Sub txtDisplay_KeyPress(KeyAscii As Integer)

    'Make text box read only
    KeyAscii = 0
    
End Sub
