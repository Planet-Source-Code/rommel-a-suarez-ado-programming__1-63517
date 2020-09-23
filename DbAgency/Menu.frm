VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5835
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Caption         =   "Hackmel@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   4920
      Picture         =   "Menu.frx":0000
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      Height          =   2775
      Left            =   240
      Top             =   960
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      FillColor       =   &H00008000&
      Height          =   4335
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Applicant Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Category Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo err

conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DbAgency.mdb;Persist Security Info=False"

Exit Sub
err:
MsgBox err.Description, vbInformation
End Sub



Private Sub Label1_Click(Index As Integer)
If Index = 0 Then
   FrmCategory.Show vbModal
Else
   FrmDataEntry.Show vbModal
End If
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Integer

For i = 0 To 1
   If i = Index Then
      Me.Label1(i).ForeColor = &H80&
      Me.Label1(i).FontBold = True
    Else
       Me.Label1(i).ForeColor = &HFF0000
       Me.Label1(i).FontBold = False
   End If
Next
End Sub


Private Sub Label2_Click()
FrmAbout.Show vbModal
End Sub
