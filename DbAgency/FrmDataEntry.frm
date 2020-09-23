VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmDataEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Applicant's Information Management Form"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Delete 
      Caption         =   "&Delete"
      Height          =   615
      Left            =   12480
      TabIndex        =   46
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton Save 
      Caption         =   "&Save"
      Height          =   615
      Left            =   9000
      TabIndex        =   45
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton New 
      Caption         =   "&New"
      Height          =   615
      Left            =   6720
      TabIndex        =   44
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   6720
      TabIndex        =   42
      Top             =   7200
      Width           =   8055
      Begin MSForms.ComboBox Category 
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Tag             =   "x"
         Top             =   240
         Width           =   6735
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "11880;661"
         ColumnCount     =   2
         cColumnInfo     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0;35"
      End
      Begin VB.Label Label19 
         Caption         =   "Category:"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3255
      Left            =   6720
      TabIndex        =   36
      Top             =   3960
      Width           =   8055
      Begin VB.TextBox Years_Of_Exp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   26
         Tag             =   "x"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox isLicensed 
         Caption         =   "Licensed"
         Height          =   255
         Left            =   6960
         TabIndex        =   27
         Tag             =   "x"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox School 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2760
         TabIndex        =   25
         Tag             =   "x"
         Top             =   2280
         Width           =   5175
      End
      Begin VB.ComboBox Highest_Educ_attain 
         Height          =   315
         ItemData        =   "FrmDataEntry.frx":0000
         Left            =   2760
         List            =   "FrmDataEntry.frx":0016
         TabIndex        =   24
         Tag             =   "x"
         Top             =   1920
         Width           =   5175
      End
      Begin VB.TextBox Curr_Position 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Tag             =   "x"
         Top             =   840
         Width           =   6255
      End
      Begin VB.TextBox Curr_Employer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Tag             =   "x"
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label18 
         Caption         =   "Years of Experience:"
         Height          =   255
         Left            =   1080
         TabIndex        =   41
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "School:"
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Highest Educational Attainment:"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label15 
         Caption         =   "Current Position:"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Current Employer:"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   6720
      TabIndex        =   14
      Top             =   2400
      Width           =   8055
      Begin VB.TextBox Weight 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   21
         Tag             =   "x"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox XHeight 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6480
         TabIndex        =   20
         Tag             =   "x"
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox Gender 
         Height          =   315
         ItemData        =   "FrmDataEntry.frx":0080
         Left            =   6480
         List            =   "FrmDataEntry.frx":008A
         TabIndex        =   19
         Tag             =   "x"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Age 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Tag             =   "x"
         Top             =   720
         Width           =   735
      End
      Begin MSMask.MaskEdBox Bday 
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Tag             =   "x"
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Telephone 
         Height          =   255
         Left            =   3720
         TabIndex        =   17
         Tag             =   "x"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   12
         Mask            =   "###-###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cell 
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Tag             =   "x"
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   13
         Mask            =   "####-###-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Cell#:"
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Weight:"
         Height          =   255
         Left            =   5760
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Height:"
         Height          =   255
         Left            =   5760
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Gender:"
         Height          =   255
         Left            =   5760
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "B-Day:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Age:"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   6720
      TabIndex        =   3
      Top             =   360
      Width           =   8055
      Begin VB.CommandButton Find 
         Caption         =   "Find"
         Height          =   315
         Left            =   3000
         TabIndex        =   47
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Address 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Tag             =   "x"
         Top             =   1440
         Width           =   6855
      End
      Begin VB.TextBox MI 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         MaxLength       =   1
         TabIndex        =   8
         Tag             =   "x"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox First 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Tag             =   "x"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox Last 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Tag             =   "x"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Control 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Tag             =   "x"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Address:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "First Name"
         Height          =   255
         Left            =   4200
         TabIndex        =   11
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "MI"
         Height          =   255
         Left            =   7320
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Control #:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.TextBox KeySearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   13785
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
         Text            =   "Control#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Key Search"
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
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "FrmDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub clear()
Dim c As Control

For Each c In Me.Controls
  If c.Tag = "x" Then
   
    If c.Name = "Telephone" Then
       c.Mask = ""
       c.Text = ""
       c.Mask = "###-###-####"
    ElseIf c.Name = "Cell" Then
       c.Mask = ""
       c.Text = ""
       c.Mask = "####-###-####"
   ElseIf c.Name = "Bday" Then
       c.Mask = ""
       c.Text = ""
       c.Mask = "##/##/####"
   ElseIf c.Name = "isLicensed" Then
       c.Value = False
   Else
          c.Text = ""
   End If
 
 End If
 
Next
Me.Age.Text = ""
End Sub

Function getCateogory(Code As String) As String
On Error GoTo err

Dim rs As New ADODB.Recordset, x As Integer
Set rs = conn.Execute("Select * from TblCategory where Code='" & Code & "'")
If Not rs.EOF Then
    getCateogory = rs!Category
End If


Exit Function
err:
MsgBox err.Description, vbInformation
End Function

Sub fillCategory()
On Error GoTo err

Dim rs As New ADODB.Recordset, x As Integer
Set rs = conn.Execute("Select * from TblCategory")
Me.Category.clear
x = 0
If Not rs.EOF Then
   Do Until rs.EOF
      Category.AddItem rs(0)
      Category.Column(1, x) = rs(1)
      x = x + 1
      rs.MoveNext
   Loop
End If


Exit Sub
err:
MsgBox err.Description, vbInformation
End Sub
Sub fillList(key As String)
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim Item As ListItem
Me.ListView1.ListItems.clear

Set rs = conn.Execute("Select [ctrl#],[Last Name] +','+ [First Name] +' '+ Mi as Name from TblApplicants where [Last Name] +','+ [First Name] +' '+ Mi Like '" + "%" + key + "%" + "'")
If Not rs.EOF Then
   
   Do Until rs.EOF
      Set Item = Me.ListView1.ListItems.Add(, , rs(0))
      Item.SubItems(1) = rs(1)
      rs.MoveNext
   Loop
End If
End Sub

Sub Search(ByVal cont As String)
On Error GoTo err

Dim rs As New ADODB.Recordset

rs.Open "Select * from TblApplicants where [Ctrl#]='" & cont & "'", conn, 1, 3
If Not rs.EOF Then

 Me.Control.Text = rs![Ctrl#]
 Me.Category.Text = getCateogory(rs!Category)
 Me.Last.Text = rs![Last Name]
 Me.First.Text = rs![First Name]
 Me.MI.Text = rs!MI
 Me.Address.Text = rs!Address
 Me.Telephone.Text = rs!Telephone
 
 Me.Cell.Text = rs!Cell
 Me.Age.Text = rs!Age
 Me.Bday.Text = Format(rs!Bday, "mm/dd/yyyy")
 Me.Gender.Text = rs!Gender
 Me.XHeight.Text = rs!Height
 Me.Weight.Text = rs!Weight
 Me.Curr_Employer.Text = rs!Curr_Employer
 Me.Curr_Position.Text = rs!Curr_Position
 Me.Highest_Educ_attain.Text = rs!Highest_Educ_attain
 Me.School.Text = rs!School
 Me.isLicensed.Value = CInt(rs!isLicensed) * -1
 Me.Years_Of_Exp.Text = rs!Years_Of_Exp

Else
 MsgBox "Record not found", vbInformation
 clear

End If



Exit Sub
err:
MsgBox err.Description, vbInformation
End Sub


Private Sub Address_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Bday_Change()
If IsDate(Me.Bday.Text) Then
   Me.Age.Text = DateDiff("yyyy", Me.Bday.Text, Now())
End If
End Sub





Private Sub Category_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub Control_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Curr_Employer_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Curr_Position_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Delete_Click()
On Error GoTo err

If MsgBox("Warning: You are about to remove this record, do you wish to continue?", vbYesNo + vbQuestion) = vbYes Then
  conn.Execute ("Delete from TblApplicants where [Ctrl#]='" & Me.Control.Text & "'")
  clear
End If
fillList Me.KeySearch.Text


Exit Sub
err:
MsgBox err.Description, vbInformation
End Sub

Private Sub Find_Click()
Me.Search Me.Control.Text
End Sub


Private Sub First_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
fillList Me.KeySearch.Text
fillCategory
End Sub





Private Sub Gender_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Highest_Educ_attain_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub KeySearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  fillList Me.KeySearch.Text
End If
End Sub

Private Sub Last_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.Search Item.Text
End Sub


Private Sub MI_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub New_Click()
clear
End Sub

Private Sub Save_Click()
On Error GoTo err

Dim rs As New ADODB.Recordset

If Me.Control.Text = "" Then
   MsgBox "Control is null", vbInformation
   Exit Sub
End If

If Me.Last.Text = "" Then
   MsgBox "Last Name is null", vbInformation
   Me.Last.SetFocus
   Exit Sub
End If

If Me.First.Text = "" Then
   MsgBox "First Name is null", vbInformation
   Me.First.SetFocus
   Exit Sub
End If

If Me.MI.Text = "" Then
   MsgBox "Middle Name is null", vbInformation
   Me.MI.SetFocus
   Exit Sub
End If

If Me.Address.Text = "" Then
   MsgBox "Address is null", vbInformation
   Me.Address.SetFocus
   Exit Sub
End If

If Me.Gender.Text = "" Then
   MsgBox "Please select a gender", vbInformation
   Me.Gender.SetFocus
   Exit Sub
End If


If Me.Highest_Educ_attain.Text = "" Then
   MsgBox "Please select an item for educational attainment", vbInformation
   Me.Highest_Educ_attain.SetFocus
   Exit Sub
End If

If Me.School.Text = "" Then
   MsgBox "School is null", vbInformation
   Me.School.SetFocus
   Exit Sub
End If


If Me.Category.Text = "" Then
   MsgBox "Please select a category", vbInformation
   Me.Category.SetFocus
   Exit Sub
End If

If Not IsDate(Me.Bday) Then
   MsgBox "Invalid Birth Day", vbInformation
   Me.Bday.SetFocus
   Exit Sub
End If




rs.Open "Select * from TblApplicants where [Ctrl#]='" & Me.Control.Text & "'", conn, 1, 3
If rs.EOF Then
   rs.AddNew
Else
  If MsgBox("This control number already exist, do you want to replace?", vbYesNo + vbQuestion) = vbNo Then
     Exit Sub
  End If
End If


rs![Ctrl#] = Me.Control.Text
rs!Category = Me.Category.Value
rs![Last Name] = Me.Last.Text
rs![First Name] = Me.First.Text
rs!MI = Me.MI.Text
rs!Address = Me.Address.Text
rs!Telephone = Me.Telephone.Text
rs!Cell = Me.Cell.Text
rs!Age = Val(Me.Age.Text)
rs!Bday = Me.Bday.Text
rs!Gender = Me.Gender.Text
rs!Height = Me.XHeight.Text
rs!Weight = Me.Weight.Text
rs!Curr_Employer = Me.Curr_Employer.Text
rs!Curr_Position = Me.Curr_Position.Text
rs!Highest_Educ_attain = Me.Highest_Educ_attain.Text
rs!School = Me.School.Text
rs!isLicensed = Me.isLicensed.Value
rs!Years_Of_Exp = Val(Me.Years_Of_Exp.Text)

rs.Update

fillList Me.KeySearch.Text


Exit Sub
err:
MsgBox err.Description, vbInformation
End Sub

Private Sub School_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Weight_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub XHeight_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Years_Of_Exp_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
