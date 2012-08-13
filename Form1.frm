VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTot1 
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtTot2 
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtNo2 
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtNo1 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton btnFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   10920
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ListBox lstTables2 
      Height          =   4740
      Left            =   6120
      TabIndex        =   7
      Top             =   3360
      Width           =   4575
   End
   Begin VB.ListBox lstTables1 
      Height          =   4740
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   4575
   End
   Begin VB.CommandButton btnDB2 
      Caption         =   "Select"
      Height          =   375
      Left            =   10800
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox lstDB2 
      Height          =   2595
      Left            =   6120
      TabIndex        =   4
      Top             =   600
      Width           =   4575
   End
   Begin VB.TextBox txtDB2 
      Height          =   285
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   4575
   End
   Begin VB.CommandButton btnDB1 
      Caption         =   "Select"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox lstDB1 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.TextBox txtDB1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim cnnDB1 As New ADODB.Connection
    Dim cnnDB2 As New ADODB.Connection
    Dim constr As String
    Dim temSQL As String

Private Sub btnDB1_Click()
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "MDB|*.mdb"
    On Error GoTo eh
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtDB1.Text = CommonDialog1.FileName
    End If
    Exit Sub
eh:
    MsgBox "Error loading the image"
End Sub

Private Sub btnDB2_Click()
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "MDB|*.mdb"
    On Error GoTo eh
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtDB2.Text = CommonDialog1.FileName
    End If
    Exit Sub
eh:
    MsgBox "Error loading the image"
End Sub

Public Function NonSystemTables(dbPath As String) As Collection

'Input: Full Path to an Access Database

'Returns: Collection of the names
'of non-system tables in that database
'or Nothing if there is an error

'Requires: a reference to data access
'objects (DAO) in your project

'On Error GoTo ErrHandler

Dim td As DAO.TableDef
Dim db As DAO.Database
Dim colTables As Collection

Set db = Workspaces(0).OpenDatabase(dbPath, False, False, ";pwd =Bud7Nil")

Set colTables = New Collection

 For Each td In db.TableDefs

    If td.Attributes >= 0 And td.Attributes <> dbHiddenObject _
         And td.Attributes <> 2 Then
   
          colTables.Add td.Name
    End If
  Next
db.Close
Set NonSystemTables = colTables

Exit Function
ErrHandler:
On Error Resume Next
If Not db Is Nothing Then db.Close

Set NonSystemTables = Nothing

End Function

Private Sub Form_Load()
    Call GetSettings
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, txtDB1.Name, txtDB1.Text
    SaveSetting App.EXEName, Me.Name, txtDB2.Name, txtDB2.Text
End Sub

Private Sub GetSettings()
    txtDB1.Text = GetSetting(App.EXEName, Me.Name, txtDB1.Name, "")
    txtDB2.Text = GetSetting(App.EXEName, Me.Name, txtDB2.Name, "")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub lstDB1_Click()
    lstTables1.Clear
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    Dim Tot As Long
    With rsTem
        temSQL = "Select * from " & lstDB1.Text
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        For i = 0 To .Fields.Count - 1
            lstTables1.AddItem .Fields(i).Name & vbTab & .Fields(i).Type
            Tot = Tot + .Fields(i).Type
        Next i
    End With
    txtNo1.Text = i
    txtTot1.Text = Tot
    lstDB2.ListIndex = lstDB1.ListIndex
End Sub

Private Sub lstDB2_Click()
    lstTables2.Clear
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    Dim Tot As Long
    With rsTem
        temSQL = "Select * from " & lstDB2.Text
        .Open temSQL, cnnDB2, adOpenStatic, adLockReadOnly
        For i = 0 To .Fields.Count - 1
            lstTables2.AddItem .Fields(i).Name & vbTab & .Fields(i).Type
            Tot = Tot + .Fields(i).Type
        Next i
    End With
    txtNo2.Text = i
    txtTot2.Text = Tot
End Sub

Private Sub lstTables1_Click()
    lstTables2.ListIndex = lstTables1.ListIndex
End Sub

Private Sub txtDB1_Change()
    Dim MyTables As New Collection
    Set MyTables = NonSystemTables(txtDB1.Text)
    lstDB1.Clear
    Dim i As Integer
    For i = 1 To MyTables.Count
        lstDB1.AddItem MyTables(i)
    Next i
    If cnnDB1.State = 1 Then cnnDB1.Close
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & txtDB1.Text & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=Bud7Nil"
    cnnDB1.Open constr
End Sub

Private Sub txtDB2_Change()
    Dim MyTables As New Collection
    Set MyTables = NonSystemTables(txtDB2.Text)
    lstDB2.Clear
    Dim i As Integer
    For i = 1 To MyTables.Count
        lstDB2.AddItem MyTables(i)
    Next i
    If cnnDB2.State = 1 Then cnnDB1.Close
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & txtDB2.Text & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=Bud7Nil"
    cnnDB2.Open constr
End Sub
