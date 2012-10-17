VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCreateCodeSS 
   Caption         =   "Get Data From Tables"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   14205
   Begin VB.TextBox txtType 
      Height          =   4815
      Left            =   9360
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtConnName 
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox txtSQLServer 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtDatabase 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox txtSql 
      Height          =   4815
      Left            =   11760
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtVB1 
      Height          =   4815
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2760
      Width           =   3975
   End
   Begin VB.ListBox lstDB1 
      Height          =   4740
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnDB1 
      Caption         =   "Select"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SQL 2005 Instance"
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   825
   End
End
Attribute VB_Name = "frmCreateCodeSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim cnnDB1 As New ADODB.Connection
    Dim constr As String
    Dim temSQL As String
    Dim connectionName As String
    
    
    ' " & connectionName & "
    ' " & connectionName & "
    
Private Function connectToDatabase() As Boolean
    On Error GoTo eh:
    Dim constr As String
    connectToDatabase = False
    constr = "Provider=MSDataShape.1;Persist Security Info=True;Data Source=" & txtServer.Text & _
        "\" & txtSQLServer.Text & _
        ";User ID=" & txtUserName.Text & _
        ";Password=" & txtPassword.Text & _
        ";Initial Catalog=" & txtDatabase.Text & _
        ";Data Provider=SQLOLEDB.1"
    If cnnDB1.State = 1 Then cnnDB1.Close
    cnnDB1.Open constr
    connectToDatabase = True
    Exit Function
eh:
    connectToDatabase = False
End Function

    
    
Private Function getDataType(AccessDataType As Integer, IsAutoIncrement As Boolean) As MyDataType
    Select Case AccessDataType
        Case 20:
            If (IsAutoIncrement) Then
                getDataType = MyDataType.MyID
            Else
                getDataType = MyDataType.MyLong
            End If
        Case 3:
            If (IsAutoIncrement) Then
                getDataType = MyDataType.MyID
            Else
                getDataType = MyDataType.MyLong
            End If
            
        Case 5:
                getDataType = MyDataType.MyDouble
        Case 202:
                getDataType = MyDataType.MyText
        Case 203:
                getDataType = MyDataType.MyMemo
        Case 11:
                getDataType = MyDataType.MyBoolean
                
        Case 135:
                getDataType = MyDataType.MyDate
        Case Else:
                getDataType = MyDataType.MyOther
    End Select

End Function

Private Function createAsType(myDataField As DataField) As String
    Dim temStr As String
    Select Case myDataField.FieldType
        Case MyDataType.MyBoolean:
            temStr = " as Boolean"
        Case MyDataType.MyDate:
            temStr = " as Date"
        Case MyDataType.MyDouble:
            temStr = " as Double"
        Case MyDataType.MyID:
            temStr = " as Long"
        Case MyDataType.MyLong:
            temStr = " as Long"
        Case MyDataType.MyMemo:
            temStr = " as String"
        Case MyDataType.MyOther:
            temStr = " "
        Case MyDataType.MyText:
            temStr = " as String"
        Case Else:
            temStr = ""
    End Select
    createAsType = temStr
End Function

Private Function createDeclerations(myDataField As DataField) As String
    createDeclerations = "    Private var" & myDataField.FieldName & createAsType(myDataField) & vbNewLine
End Function

Private Function createGetAndLet(myDataField As DataField) As String
    Dim temStr As String
    temStr = temStr & "Public Property Let " & myDataField.FieldName & "(ByVal v" & myDataField.FieldName & " " & createAsType(myDataField) & ")" & vbNewLine
    temStr = temStr & vbTab & "var" & myDataField.FieldName & " = v" & myDataField.FieldName & vbNewLine
    temStr = temStr & "End Property" & vbNewLine
    temStr = temStr & vbNewLine
    
    temStr = temStr & "Public Property Get " & myDataField.FieldName & "() " & createAsType(myDataField) & vbNewLine
    temStr = temStr & vbTab & myDataField.FieldName & " = var" & myDataField.FieldName & vbNewLine
    temStr = temStr & "End Property" & vbNewLine
    temStr = temStr & vbNewLine
    
    createGetAndLet = temStr
End Function



Private Sub btnDB1_Click()
    connectionName = txtConnName.Text
    If connectToDatabase Then
        MsgBox "Successfully Connected"
    Else
        MsgBox "Connection Failure"
        Exit Sub
    End If
    lstDB1.Clear
    listTables
End Sub

Private Sub listTables()
    Dim rsTem As New ADODB.Recordset
    temSQL = "select * from INFORMATION_SCHEMA.TABLES order by table_name"
    With rsTem
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        While .EOF = False
            lstDB1.AddItem .Fields(2).Value
            .MoveNext
        Wend
        .Close
    End With
End Sub

Public Function NonSystemTables(dbPath As String) As Collection
    
    Dim td As DAO.TableDef
    Dim db As DAO.Database
    Dim colTables As Collection
    Dim passWord As String

    passWord = InputBox("Password?")
    
    If Trim(passWord) = "" Then
        Set db = Workspaces(0).OpenDatabase(dbPath, False, False)
    Else
        Set db = Workspaces(0).OpenDatabase(dbPath, False, False, ";pwd =" & passWord)
    End If
    
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
    SaveSetting App.EXEName, Me.Name, txtDatabase.Name, txtDatabase.Text
    SaveSetting App.EXEName, Me.Name, txtPassword.Name, txtPassword.Text
    SaveSetting App.EXEName, Me.Name, txtServer.Name, txtServer.Text
    SaveSetting App.EXEName, Me.Name, txtSQLServer.Name, txtSQLServer.Text
    SaveSetting App.EXEName, Me.Name, txtUserName.Name, txtUserName.Text
    SaveSetting App.EXEName, Me.Name, txtConnName.Name, txtConnName.Text

End Sub

Private Sub GetSettings()
    txtDatabase.Text = GetSetting(App.EXEName, Me.Name, txtDatabase.Name)
    txtPassword.Text = GetSetting(App.EXEName, Me.Name, txtPassword.Name)
    txtServer.Text = GetSetting(App.EXEName, Me.Name, txtServer.Name)
    txtSQLServer.Text = GetSetting(App.EXEName, Me.Name, txtSQLServer.Name)
    txtUserName.Text = GetSetting(App.EXEName, Me.Name, txtUserName.Name)
    txtConnName.Text = GetSetting(App.EXEName, Me.Name, txtConnName.Name, "cnnStores")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub lstDB1_Click()


'    txtVB1.Text = Empty
'    Dim rsTem As New ADODB.Recordset
'    Dim allFields As New Collection
'    Dim temField As DataField
'    Dim i As Integer
'    Dim Tot As Long
'    With rsTem
'        temSQL = "Select * from " & lstDB1.Text
'        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
'        For i = 0 To .Fields.Count - 1
'            Set temField = New DataField
'            temField.FieldName = .Fields(i).Name
'            temField.FieldType = getDataType(.Fields(i).Type, .Fields(i).Properties(2).Value)
'            allFields.Add temField
'        Next i
'
'    End With
'    txtVB1.Text = createVBClass(allFields)


    txtVB1.Text = Empty
    txtType.Text = Empty
    Dim rsTem As New ADODB.Recordset
    Dim allFields As New Collection
    Dim temField As DataField
    Dim i As Integer
    Dim Tot As Long
    With rsTem
        temSQL = "Select * from " & lstDB1.Text
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        For i = 0 To .Fields.Count - 1
            Set temField = New DataField
            temField.FieldName = .Fields(i).Name
            temField.FieldType = getDataType(.Fields(i).Type, .Fields(i).Properties(5).Value)
            txtType.Text = txtType.Text & vbNewLine & .Fields(i).Name & vbTab & .Fields(i).Type
            allFields.Add temField
        Next i
        
    End With
    txtVB1.Text = createVBClass(allFields)
End Sub

Private Function createVBClass(fieldCollection As Collection)
    Dim temStr As String
    Dim temField As DataField
    temStr = "Option Explicit" & vbNewLine
    temStr = temStr & "    Dim temSQL As String" & vbNewLine
    
    For Each temField In fieldCollection
        temStr = temStr & createDeclerations(temField)
    Next
    
    temStr = temStr + vbNewLine
    
    txtSql.Text = Empty
    
    Dim temStr1 As String
    
    temStr1 = InputBox("Id Field Name")
    
    For Each temField In fieldCollection
        txtSql.Text = txtSql.Text & vbNewLine & temField.FieldName & vbTab & temField.FieldType
        If temField.FieldName = temStr1 Then
            temStr = temStr & createGetDetailsByID(fieldCollection, temField)
        Else
            temStr = temStr & createGetAndLet(temField)
        End If
    Next
    
    createVBClass = temStr
    
End Function

Private Function createGetDetailsByID(fieldCollection As Collection, iDField As DataField) As String
    Dim temStr As String
    Dim temField As DataField
    Dim temStr1 As String
    
    
    temStr = temStr & "Public Sub saveData()" & vbNewLine
    temStr = temStr & " " & vbNewLine
    temStr = temStr & "    Dim rsTem As New ADODB.Recordset" & vbNewLine
    temStr = temStr & "    Dim newEntry As Boolean" & vbNewLine
    temStr = temStr & "    With rsTem" & vbNewLine
    temStr = temStr & "        temSQL = " & Chr(34) & "SELECT * FROM " & lstDB1.Text & " Where " & iDField.FieldName & " = " & Chr(34) & " & var" & iDField.FieldName & vbNewLine
    temStr = temStr & "        If .State = 1 Then .Close " & vbNewLine
    temStr = temStr & "        .Open temSQL, " & connectionName & ", adOpenStatic, adLockOptimistic" & vbNewLine
    temStr = temStr & "        If .RecordCount  <= 0 Then" & vbNewLine
    temStr = temStr & "            .addnew" & vbNewLine
    temStr = temStr & "            newEntry = true" & vbNewLine
    temStr = temStr & "        Else" & vbNewLine
    temStr = temStr & "            newEntry = false" & vbNewLine
    temStr = temStr & "        End If" & vbNewLine
    
    
    For Each temField In fieldCollection
        If temField.FieldName <> iDField.FieldName And temField.FieldName <> "upsize_ts" Then
            temStr = temStr & "        !" & temField.FieldName & " = var" & temField.FieldName & vbNewLine
        End If
    Next

    
    temStr = temStr & "        .update" & vbNewLine
    temStr = temStr & "        if newENtry = true then" & vbNewLine
    temStr = temStr & "            .close" & vbNewLine
    temStr = temStr & "            temSQL = " & Chr(34) & "SELECT @@IDENTITY AS NewID" & Chr(34) & vbNewLine
    temStr = temStr & "           .Open temSQL, " & connectionName & ", adOpenStatic, adLockReadOnly" & vbNewLine
    temStr = temStr & "            var" & iDField.FieldName & " = !NewID" & vbNewLine
    temStr = temStr & "        Else" & vbNewLine
    temStr = temStr & "            var" & iDField.FieldName & " = !" & iDField.FieldName & vbNewLine
    temStr = temStr & "        End if" & vbNewLine
    temStr = temStr & "        if .state =1 then .close" & vbNewLine
    temStr = temStr & "    end with        " & vbNewLine
    temStr = temStr & "    " & vbNewLine
    
    temStr = temStr & "End Sub" & vbNewLine
    
    
    temStr = temStr & "Public Sub loadData()" & vbNewLine
    temStr = temStr & " " & vbNewLine
    temStr = temStr & "    Dim rsTem As New ADODB.Recordset" & vbNewLine
    temStr = temStr & "    With rsTem" & vbNewLine
    temStr = temStr & "        temSQL = " & Chr(34) & "SELECT * FROM " & lstDB1.Text & " WHERE " & iDField.FieldName & " = " & Chr(34) & " & var" & iDField.FieldName & vbNewLine
    temStr = temStr & "        If .State = 1 Then .Close " & vbNewLine
    temStr = temStr & "        .Open temSQL, " & connectionName & ", adOpenStatic, adLockOptimistic" & vbNewLine
    temStr = temStr & "        If .RecordCount > 0 then " & vbNewLine
    
    For Each temField In fieldCollection
        temStr = temStr & "            If not isnull(!" & temField.FieldName & ") Then" & vbNewLine
        temStr = temStr & "               var" & temField.FieldName & " = !" & temField.FieldName & vbNewLine
        temStr = temStr & "            End If" & vbNewLine
    Next
    temStr = temStr & "        End If " & vbNewLine
    temStr = temStr & "    if .state =1 then .close" & vbNewLine
    temStr = temStr & "    end with        " & vbNewLine
    temStr = temStr & "    " & vbNewLine
    
    temStr = temStr & "End Sub" & vbNewLine
    
    temStr = temStr & "Public Sub clearData()" & vbNewLine
    For Each temField In fieldCollection
        Select Case temField.FieldType
            Case MyDataType.MyBoolean:
                temStr1 = " = False"
            Case MyDataType.MyDate:
                temStr1 = " = Empty"
            Case MyDataType.MyDouble:
                temStr1 = " = 0"
            Case MyDataType.MyID:
                temStr1 = " = 0"
            Case MyDataType.MyLong:
                temStr1 = " = 0"
            Case MyDataType.MyMemo:
                temStr1 = " = Empty"
            Case MyDataType.MyOther:
                temStr1 = " = Empty"
            Case MyDataType.MyText:
                temStr1 = " = Empty"
            Case Else:
                temStr1 = " = Empty"
        End Select
        temStr = temStr & "    var" & temField.FieldName & temStr1
        temStr = temStr & vbNewLine
    Next
    temStr = temStr & "End Sub" & vbNewLine

    temStr = temStr & vbNewLine
    
    temStr = temStr & "Public Property Let " & iDField.FieldName & "(ByVal v" & iDField.FieldName & " " & createAsType(iDField) & ")" & vbNewLine
    temStr = temStr & "    call clearData" & vbNewLine
    temStr = temStr & vbTab & "var" & iDField.FieldName & " = v" & iDField.FieldName & vbNewLine
    temStr = temStr & "    call loadData" & vbNewLine
    
    
    
    temStr = temStr & "End Property" & vbNewLine
    temStr = temStr & vbNewLine
    
    temStr = temStr & "Public Property Get " & iDField.FieldName & "() " & createAsType(iDField) & vbNewLine
    temStr = temStr & vbTab & iDField.FieldName & " = var" & iDField.FieldName & vbNewLine
    temStr = temStr & "End Property" & vbNewLine
    temStr = temStr & vbNewLine

    createGetDetailsByID = temStr
    
End Function


