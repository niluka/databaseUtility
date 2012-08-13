VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreateCode 
   Caption         =   "Get Data From Tables"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14205
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   14205
   Begin VB.TextBox txtSql 
      Height          =   6855
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox txtVB1 
      Height          =   6855
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
   Begin VB.ListBox lstDB1 
      Height          =   6885
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnDB1 
      Caption         =   "Select"
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtDB1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frmCreateCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim cnnDB1 As New ADODB.Connection
    Dim constr As String
    Dim temSQL As String
    
Private Function getDataType(AccessDataType As Integer, IsAutoIncrement As Boolean) As MyDataType
    Select Case AccessDataType
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
                
        Case 7:
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
    SaveSetting App.EXEName, Me.Name, txtDB1.Name, txtDB1.Text
End Sub

Private Sub GetSettings()
    txtDB1.Text = GetSetting(App.EXEName, Me.Name, txtDB1.Name, "")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub lstDB1_Click()
    txtVB1.Text = Empty
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
            temField.FieldType = getDataType(.Fields(i).Type, .Fields(i).Properties(2).Value)
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
    
    For Each temField In fieldCollection
        If temField.FieldType = MyID Then
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
    temStr = temStr & "    With rsTem" & vbNewLine
    temStr = temStr & "        temSQL = " & Chr(34) & "SELECT * FROM " & lstDB1.Text & " Where " & iDField.FieldName & " = " & Chr(34) & " & var" & iDField.FieldName & vbNewLine
    temStr = temStr & "        If .State = 1 Then .Close " & vbNewLine
    temStr = temStr & "        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic" & vbNewLine
    temStr = temStr & "        If .RecordCount  <= 0 Then .addnew" & vbNewLine
    
    For Each temField In fieldCollection
        If temField.FieldType <> MyID Then
            temStr = temStr & "        !" & temField.FieldName & " = var" & temField.FieldName & vbNewLine
        End If
    Next

    
    temStr = temStr & "        .update" & vbNewLine
    temStr = temStr & "        var" & iDField.FieldName & " = !" & iDField.FieldName & vbNewLine
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
    temStr = temStr & "        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic" & vbNewLine
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



