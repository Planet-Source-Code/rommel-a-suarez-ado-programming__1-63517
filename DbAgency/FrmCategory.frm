VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Delete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   7680
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CommandButton Save 
      Caption         =   "&Save"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton New 
      Caption         =   "&New"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Code 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtCategory 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Category:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub clear()
Me.Code.Text = ""
Me.TxtCategory.Text = ""
End Sub




Sub fillList()
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim Item As ListItem
Me.ListView1.ListItems.clear

Set rs = conn.Execute("Select * from TblCategory  order by Code")
If Not rs.EOF Then
   
   Do Until rs.EOF
      Set Item = Me.ListView1.ListItems.Add(, , rs(0))
      Item.SubItems(1) = rs(1)
      rs.MoveNext
   Loop
End If
End Sub


Private Sub Code_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Delete_Click()
On Error GoTo err

If Me.Code.Text = "" Then
   MsgBox "Code is null", vbInformation
   Exit Sub
End If
If MsgBox("Warning: You are about to remove this record, do you wish to continue?", vbYesNo + vbQuestion) = vbYes Then
  conn.Execute "Delete from tblCategory where Code='" & Me.Code.Text & "'"
  clear
End If

fillList


Exit Sub
err:
MsgBox err.Description, vbInformation
End Sub

Private Sub Form_Load()
fillList
End Sub



Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.Code.Text = Item.Text
Me.TxtCategory.Text = Item.SubItems(1)
End Sub

Private Sub New_Click()
Me.clear
End Sub

Private Sub Save_Click()
On Error GoTo err
Dim rs As New ADODB.Recordset

If Me.Code.Text = "" Then
   MsgBox "Code is null", vbInformation
   Exit Sub
End If
 
If Me.TxtCategory.Text = "" Then
   MsgBox "Category is null", vbInformation
   Exit Sub
End If
 

rs.Open "Select * from TblCategory where Code='" & Me.Code.Text & "'", conn, 1, 3
If rs.EOF Then
   rs.AddNew
Else
   If MsgBox("Record already exist,do you want to replace?", vbYesNo + vbQuestion) = vbNo Then
      Exit Sub
   End If
End If

rs!Code = Me.Code.Text
rs!Category = Me.TxtCategory.Text
rs.Update
fillList


Exit Sub
err:
MsgBox err.Description, vbInformation
End Sub


Private Sub TxtCategory_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
