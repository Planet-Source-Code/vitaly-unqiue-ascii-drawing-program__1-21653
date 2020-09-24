VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Droll drawer"
   ClientHeight    =   6975
   ClientLeft      =   -30
   ClientTop       =   930
   ClientWidth     =   10125
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   29.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   84.375
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Index           =   0
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00008080&
      Height          =   315
      Index           =   1
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00C000C0&
      Height          =   315
      Index           =   2
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   1200
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   3600
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00800000&
      Height          =   315
      Index           =   6
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   3120
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00008000&
      Height          =   315
      Index           =   5
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00808000&
      Height          =   315
      Index           =   4
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000000C0&
      Height          =   315
      Index           =   3
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   1680
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000FFFF&
      Height          =   315
      Index           =   8
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   4080
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF00FF&
      Height          =   315
      Index           =   9
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   4560
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF0000&
      Height          =   315
      Index           =   13
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   6480
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000FF00&
      Height          =   315
      Index           =   12
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   6000
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Index           =   11
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   5520
      Width           =   375
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000000FF&
      Height          =   315
      Index           =   10
      Left            =   9600
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   5040
      Width           =   375
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   5760
      Left            =   360
      MousePointer    =   1  'Arrow
      ScaleHeight     =   23.9
      ScaleMode       =   0  'User
      ScaleWidth      =   75
      TabIndex        =   0
      Tag             =   "*"
      Top             =   960
      Width           =   9000
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4200
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Shape shpPercent 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   7  'Diagonal Cross
         Height          =   615
         Left            =   1440
         Top             =   2160
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00E0E0E0&
         Height          =   615
         Left            =   1440
         Top             =   2160
         Visible         =   0   'False
         Width           =   6000
      End
   End
   Begin VB.Label lblButtomRuler 
      BackColor       =   &H00000000&
      Caption         =   "'---------------------------------------------------------------------------'"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   6720
      Width           =   9255
   End
   Begin VB.Label lblRightRuler 
      BackColor       =   &H00000000&
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   5775
      Left            =   9360
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblLeftRuler 
      BackColor       =   &H00000000&
      Caption         =   "A|"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblTopRuler 
      BackColor       =   &H00000000&
      Caption         =   "|123456789012345678901234567890123456789012345678901234567890123456789012345|"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   75
      Width           =   9255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCLS 
         Caption         =   "&Clear Screen"
      End
      Begin VB.Menu mnuGenerateDraw 
         Caption         =   "&Generate draw"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Clicking As Boolean
Dim Pos(23, 75) As String

Private Sub Form_Load()
    Dim I As Integer
    Dim Char As Integer
    Char = 64
    lblTopRuler.Caption = _
    ".---------------------------------------------------------------------------." & vbCr & _
    "|         1         2         3         4         5         6         7     |" & vbCr & _
    "|123456789012345678901234567890123456789012345678901234567890123456789012345|" & vbCr & _
    "|---------------------------------------------------------------------------|"
    lblRightRuler.Caption = ""
    For I = 1 To 24
        Char = Char + 1
        lblLeftRuler.Caption = lblLeftRuler.Caption + Chr(Char) + "|" + vbCr
        lblRightRuler.Caption = lblRightRuler.Caption + "|" + vbCr
    Next I
    For I = 0 To 6
        picColor(I).BackColor = QBColor(7 - I)
        picColor(I + 7).BackColor = QBColor(15 - I)
    Next I
End Sub

Private Sub mnuCLS_Click()
    Dim I As Integer
    Dim I2 As Integer
    picDraw.Cls
    For I = 0 To 23
        For I2 = 0 To 74
            Pos(I, I2) = 0
        Next I2
    Next I
End Sub

Private Sub mnuGenerateDraw_Click()
    Dim I As Integer
    Dim I2 As Integer
    Dim strColor As String
    frmOutput.txtOutput.Text = ""
    lblPercent.Visible = True
    shpBorder.Visible = True
    shpPercent.Visible = True
    For I = 0 To 23
        For I2 = 0 To 74
            If Pos(I, I2) <> "" And Pos(I, I2) <> "0" Then
                frmOutput.txtOutput.Text = frmOutput.txtOutput.Text + "dot " & Chr(65 + I) & I2 + 1
                Select Case Pos(I, I2)
                Case 12632256
                    strColor = " white"
                Case 32896
                    strColor = " orange"
                Case 8388736
                    strColor = " magenta"
                Case 128
                    strColor = " red"
                Case 8421376
                    strColor = " cyan"
                Case 32768
                    strColor = " green"
                Case 8388608
                    strColor = " blue"
                Case 16777215
                    strColor = " bold white"
                Case 65535
                    strColor = " bold yellow"
                Case 16711935
                    strColor = " bold magenta"
                Case 255
                    strColor = " bold red"
                Case 16776960
                    strColor = " bold cyan"
                Case 65280
                    strColor = " bold green"
                Case 16711680
                    strColor = " bold blue"
                End Select
                frmOutput.txtOutput.Text = frmOutput.txtOutput.Text + strColor + vbCrLf
                If lblPercent.Caption <> Int((I + 1) * 100 / 24) Then
                    frmMain.Refresh
                    lblPercent.Caption = Int((I + 1) * 100 / 24)
                    shpPercent.Width = ((I + 1) * 100 / 24) / 2
                End If
            End If
            If I > (shpPercent.Width - 1) / 100 * 24 * 2 Then
                frmMain.Refresh
                lblPercent.Caption = Int((I + 1) * 100 / 24)
                shpPercent.Width = ((I + 1) * 100 / 24) / 2
            End If
        Next I2
    Next I
    lblPercent.Visible = False
    shpBorder.Visible = False
    shpPercent.Visible = False
    frmOutput.Visible = True
    frmOutput.SetFocus
End Sub

Private Sub picColor_Click(Index As Integer)
    Dim I As Integer
    For I = 0 To 13
        picColor(I).Cls
    Next I
    picColor(Index).CurrentX = 130
    picColor(Index).Print "!"
    picDraw.ForeColor = picColor(Index).BackColor
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDraw.CurrentX = (X - 0.5) \ 1
    picDraw.CurrentY = (Y - 0.5) \ 1
    Clicking = True
    If Button = 2 Then picDraw.Tag = picDraw.ForeColor: picDraw.ForeColor = &H0:
    Call Draw(X, Y, Clicking)
End Sub
'And picDraw.CurrentX > 0 And picDraw.CurrentY > 0 _
'      And picDraw.CurrentX <= 75 And picDraw.CurrentY <= 24
Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDraw.CurrentX = (X - 0.5) \ 1
    picDraw.CurrentY = (Y - 0.5) \ 1
    picDraw.ToolTipText = Chr(65 + (Y - 0.5) \ 1) & (X - 0.5) \ 1 + 1
    Call Draw(X, Y, Clicking)
End Sub
Sub Draw(X, Y, Clicking)
    If Clicking And picDraw.CurrentX <= 74 And picDraw.CurrentY <= 23 _
      And picDraw.CurrentX >= 0 And picDraw.CurrentY >= 0 Then
            Pos((Y - 0.5) \ 1, (X - 0.5) \ 1) = picDraw.ForeColor
            Debug.Print (Y - 0.5) \ 1 & "x" & (X - 0.5) \ 1 & " " & "Color: " & Pos(picDraw.CurrentY, picDraw.CurrentX)
            picDraw.Print "@"
    End If
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clicking = False
    If Button = 2 Then picDraw.ForeColor = picDraw.Tag
End Sub


