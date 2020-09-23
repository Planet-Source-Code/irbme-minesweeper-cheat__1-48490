VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MineSweeper"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      Height          =   435
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   960
   End
   Begin VB.CommandButton cmdRead 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Read Mem"
      Height          =   435
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   960
   End
   Begin VB.PictureBox picMine 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6195
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   4830
      Width           =   240
   End
   Begin VB.CommandButton cmdMine 
      BackColor       =   &H00C0C0C0&
      Height          =   330
      Index           =   0
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   420
      Width           =   330
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRead_Click()

  Dim hProcess As Long
  Dim Buff() As Byte
  Dim x As Long, y As Long

  Dim Width As Long, Height As Long, Mines As Long

    hProcess = GetProcessByName("winmine")

    ReadMemory hProcess, &H1005334, Buff, 1
    Width = Buff(0)

    ReadMemory hProcess, &H1005338, Buff, 1
    Height = Buff(0)

    ReadMemory hProcess, &H1005330, Buff, 1
    Mines = Buff(0)
    
    For x = 1 To cmdMine.ubound
        Unload cmdMine(x)
    Next
    
    For x = 0 To Width - 1
        For y = 0 To Height - 1

            Load cmdMine(cmdMine.ubound + 1)
            cmdMine(cmdMine.ubound).Left = cmdMine(0).Left + x * cmdMine(0).Width
            cmdMine(cmdMine.ubound).Top = cmdMine(0).Top + y * cmdMine(0).Height
            
            ReadMemory hProcess, &H1005340 + (32 * (y + 1)) + (x + 1), Buff, 1
    
            If Buff(0) = &H8F Then
                Set cmdMine(cmdMine.ubound).Picture = picMine.Picture
            End If
            
            cmdMine(cmdMine.ubound).Visible = True
    Next y, x
    
    Me.Width = (Width + 1) * cmdMine(0).Width
    Me.Height = cmdMine(0).Top + ((Height + 1) * cmdMine(0).Height)

End Sub

Private Sub Form_Load()
    cmdRead_Click
End Sub
