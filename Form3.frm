VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form3 
   Caption         =   "Sql"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   13635
   Begin VB.TextBox txtVB1 
      Height          =   6855
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SQL"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   7560
      Width           =   1215
   End
   Begin VB.ListBox lstTables1 
      Height          =   6810
      Left            =   3120
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtDB1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8055
   End
   Begin VB.CommandButton btnDB1 
      Caption         =   "Select"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstDB1 
      Height          =   6885
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtSql 
      Height          =   6855
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim cnnDB1 As New ADODB.Connection
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

Public Function NonSystemTables(dbPath As String) As Collection
    Dim td As DAO.TableDef
    Dim db As DAO.Database
    Dim colTables As Collection
'    Set db = Workspaces(0).OpenDatabase(dbPath, False, False, ";pwd =Bud7Nil")
    Set db = Workspaces(0).OpenDatabase(dbPath, False, False)
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

Private Sub Command1_Click()
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    Dim Tot As Long
    txtVB1.Text = Empty
    txtSql.Text = Empty
    With rsTem
        temSQL = "Select * from " & lstDB1.Text
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        txtSql.Text = "Select "
        For i = 0 To .Fields.Count - 1
            If lstTables1.Selected(i) = True Then
                txtSql.Text = txtSql.Text & " Sum(" & .Fields(i).Name & ") as SumOf" & .Fields(i).Name & ", "
            End If
            If lstTables1.Selected(i) = True Then
                txtVB1.Text = txtVB1.Text & vbNewLine & " Dim " & .Fields(i).Name & " as " & .Fields(i).Type & vbTab
            End If
        Next i
        For i = 0 To .Fields.Count - 1
            If lstTables1.Selected(i) = True Then
                txtVB1.Text = txtVB1.Text & vbNewLine & "If IsNull(!SumOf" & .Fields(i).Name & ") = False then " & .Fields(i).Name & " = !SumOf" & .Fields(i).Name & vbNewLine
            End If
        Next i
        txtSql.Text = txtSql & vbTab & " from " & lstDB1.Text
    
    End With

End Sub

Private Sub Form_Load()
    Call GetSettings
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, txtDB1.Name, txtDB1.Text
End Sub

Private Sub GetSettings()
    txtDB1.Text = GetSetting(App.EXEName, Me.Name, txtDB1.Name, "")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
'
'Private Sub lstDB2_Click()
'    lstTables2.Clear
'    Dim rsTem As New ADODB.Recordset
'    Dim i As Integer
'    Dim Tot As Long
'    With rsTem
'        temsql = "Select * from " & lstDB2.Text
'        .Open temsql, cnnDB2, adOpenStatic, adLockReadOnly
'        For i = 0 To .Fields.Count - 1
'            lstTables2.AddItem .Fields(i).Name & vbTab & .Fields(i).Type
'            Tot = Tot + .Fields(i).Type
'        Next i
'    End With
'    txtNo2.Text = i
'    txtTot2.Text = Tot
'End Sub

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
'
'Private Sub txtDB2_Change()
'    Dim MyTables As New Collection
'    Set MyTables = NonSystemTables(txtDB2.Text)
'    lstDB2.Clear
'    Dim i As Integer
'    For i = 1 To MyTables.Count
'        lstDB2.AddItem MyTables(i)
'    Next i
'    If cnnDB2.State = 1 Then cnnDB1.Close
'    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & txtDB2.Text & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=Bud7Nil"
'    cnnDB2.Open constr
'End Sub
'
'
'
'
