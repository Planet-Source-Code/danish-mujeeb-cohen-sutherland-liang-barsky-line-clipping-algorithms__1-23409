VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "The Clipper"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Clipping Mode"
      Height          =   1095
      Left            =   3600
      TabIndex        =   14
      Top             =   6720
      Width           =   2895
      Begin VB.OptionButton op4 
         Caption         =   "Liang Barsky"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton op3 
         Caption         =   "Cohen Sutherland"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.PictureBox color 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7830
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   7440
      Width           =   255
   End
   Begin VB.PictureBox color 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   7830
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   6960
      Width           =   255
   End
   Begin VB.PictureBox color 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   7350
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   7440
      Width           =   255
   End
   Begin VB.PictureBox color 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7350
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   6960
      Width           =   255
   End
   Begin VB.PictureBox color 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6870
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   9
      Top             =   7440
      Width           =   255
   End
   Begin VB.PictureBox color 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6870
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
      Top             =   6960
      Width           =   255
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   8160
      TabIndex        =   5
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mode"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "Define clipping area"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Draw Lines"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   8160
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      Height          =   6495
      Left            =   120
      ScaleHeight     =   429
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label t2 
      Caption         =   "Label3"
      Height          =   375
      Left            =   8400
      TabIndex        =   18
      Top             =   7440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label t1 
      Caption         =   "Label3"
      Height          =   375
      Left            =   8400
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   345
      Left            =   6825
      Top             =   6915
      Width           =   345
   End
   Begin VB.Label Label2 
      Caption         =   "List of lines after clipping"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "List of lines befor clipping"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Point
Dim tmp As line
Dim mouseDown As Boolean
Dim clipping As Boolean


Private Sub color_Click(Index As Integer)
    Shape1.Top = color(Index).Top - 3
    Shape1.Left = color(Index).Left - 3
    
    tmp.c = Index + 9
End Sub

Private Sub Form_Load()
Me.ScaleMode = 3
Option1.Value = True
op3.Value = True

For i = 0 To 5
    color(i).BackColor = QBColor(i + 9)
Next i

tmp.c = 9
End Sub

Private Sub Option1_Click()
    clipping = False
    refreshDisplay Picture1
End Sub

Private Sub Option2_Click()
    clipping = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouseDown = True
Picture1.DrawMode = 6
tmp.p1.X = X
tmp.p1.Y = Y

tmp.p2.X = X
tmp.p2.Y = Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (mouseDown) Then
    If (clipping) Then
        drawBox tmp, Picture1
        tmp.p2.X = X
        tmp.p2.Y = Y
        drawBox tmp, Picture1
    Else
        drawLine tmp, 0, Picture1
        tmp.p2.X = X
        tmp.p2.Y = Y
        drawLine tmp, 0, Picture1
    End If
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouseDown = False
Picture1.DrawMode = 13
tmp.p2.X = X
tmp.p2.Y = Y

'putting the line into the list of lines
If (clipping) Then
    If op3.Value = True Then clipCohSuth tmp, Picture1
    If op4.Value = True Then clipLiangBarsky tmp, Picture1
    'fixRegion tmp
    'MsgBox toString(tmp)
Else
    addLine tmp
    drawLine tmp, 0, Picture1
End If
End Sub
