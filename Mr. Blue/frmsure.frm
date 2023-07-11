VERSION 5.00
Begin VB.Form frmsure 
   Caption         =   "Reset"
   ClientHeight    =   1320
   ClientLeft      =   4125
   ClientTop       =   3795
   ClientWidth     =   5040
   Icon            =   "frmsure.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   5040
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Are you sure?"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmsure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BY: KYLE HESPE

Private Sub Command1_Click()
Unload Me
Unload frmlvl7
frmlvl7.Show 1
End Sub


Private Sub Command2_Click()
Unload Me
End Sub


