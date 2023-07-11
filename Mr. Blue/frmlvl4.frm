VERSION 5.00
Begin VB.Form frmlvl4 
   Caption         =   "Level 4"
   ClientHeight    =   2715
   ClientLeft      =   3420
   ClientTop       =   2760
   ClientWidth     =   3435
   Icon            =   "frmlvl4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   3435
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Black Chancery"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "`"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   14
      Left            =   2640
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   13
      Left            =   2640
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   12
      Left            =   2640
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   11
      Left            =   2040
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   10
      Left            =   2040
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   9
      Left            =   2040
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   8
      Left            =   1440
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   7
      Left            =   1440
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   6
      Left            =   1440
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   840
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   840
      Top             =   720
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   240
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   240
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmlvl4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BY: KYLE HESPE
Dim cindex As Integer
Dim valwin As Integer
Dim valmove As Integer
Dim valpowerup As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
For i = 0 To 14
    If Shape1(i).FillColor = vbBlue Then
        cindex = i
    End If
    'sets cindex
Next i
'gets keypress and moves color
If KeyCode = vbKeyUp Then
    If cindex = 0 Or cindex = 3 Or cindex = 6 Or cindex = 9 Or cindex = 12 Then Exit Sub
    If Shape1(cindex - 1).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex - 1).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex - 1
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex + 1).FillColor = vbWhite
ElseIf KeyCode = vbKeyDown Then
    If cindex = 2 Or cindex = 5 Or cindex = 8 Or cindex = 11 Or cindex = 14 Then Exit Sub
    If Shape1(cindex + 1).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex + 1).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex + 1
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex - 1).FillColor = vbWhite
ElseIf KeyCode = vbKeyRight Then
    If cindex = 12 Or cindex = 13 Or cindex = 14 Then Exit Sub
    If Shape1(cindex + 3).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex + 3).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex + 3
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex - 3).FillColor = vbWhite
ElseIf KeyCode = vbKeyLeft Then
    If cindex = 0 Or cindex = 1 Or cindex = 2 Then Exit Sub
    If Shape1(cindex - 3).FillColor = vbBlack Then Exit Sub
    If Shape1(cindex - 3).FillColor = vbGreen Then
        MsgBox "You won!", vbInformation, "Winner!"
        MsgBox "Next Level", vbInformation, "Winner!"
        valwin = 1
    End If
    cindex = cindex - 3
    Shape1(cindex).FillColor = vbBlue
    Shape1(cindex + 3).FillColor = vbWhite
End If
If valwin = 1 Then
    frmlvl5.Show 1
    Unload frmmaze
End If
End Sub


Private Sub Form_Load()
Unload frmlvl3

Shape1(4).FillColor = vbBlue
Shape1(10).FillColor = vbGreen
End Sub


Private Sub Label1_Click()
frmlvl5.Show 1
End Sub


Private Sub Label2_Click()
End
End Sub


