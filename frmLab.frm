VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLab 
   Caption         =   "Get Data From Tables"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   15240
   Begin VB.ListBox lstValues 
      Height          =   5325
      Left            =   9480
      TabIndex        =   8
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox txtCentre 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7800
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox lstIxItem 
      Height          =   5325
      Left            =   3960
      TabIndex        =   3
      Top             =   1320
      Width           =   3735
   End
   Begin VB.ListBox lstIx 
      Height          =   5325
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
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
   Begin VB.Label Label1 
      Caption         =   "txtCenter"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmLab"
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


Private Sub ListIx()
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    With rsTem
        temSQL = "Select * from tblIx order by Ix"
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        i = 0
        While .EOF = False
            lstIx.AddItem !Ix
            lstIx.ItemData(i) = !IxID
            i = i + 1
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub ListIxItem()
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    With rsTem
        temSQL = "Select * from tblIxItem where IsValue = true AND IxID = " & Val(lstIx.ItemData(lstIx.ListIndex))
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        i = lstIxItem.ListCount
        While .EOF = False
            lstIxItem.AddItem !IxItem
            lstIxItem.ItemData(i) = !IxItemID
            i = i + 1
            .MoveNext
        Wend
        .Close
    End With
End Sub


Private Sub btnRemove_Click()
    Dim i As Integer
    i = lstIxItem.ListIndex
    lstIxItem.RemoveItem (i)
    lstIxItem.ListIndex = i - 1
End Sub

Private Sub btnRun_Click()
    
    Dim excelApp As Excel.Application
    Dim excelWB As Excel.Workbook
    Dim excelWS As Excel.Worksheet
    
    Dim myRow As Long
    Dim myCol As Long
    
    Dim temStr As String
    
    Dim temFor As String
    Dim temVal As String
    
    Dim temDic As String
    Dim temFormulas As String
    Dim temGetValues As String
    Dim temSetValues As String
    Dim temNames As String
    Dim temBlank As String
    
                Dim i As Integer


    Set excelApp = New Excel.Application
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    excelApp.Visible = True
    
    myRow = 1

    myCol = 1
    excelWS.Cells(myRow, myCol).Value = "Centre"
    myCol = 2
    excelWS.Cells(myRow, myCol).Value = "ID"
    myCol = 3
    excelWS.Cells(myRow, myCol).Value = "Year"
    myCol = 4
    excelWS.Cells(myRow, myCol).Value = "Month"
    myCol = 5
    excelWS.Cells(myRow, myCol).Value = "Date"
    myCol = 6
    excelWS.Cells(myRow, myCol).Value = "Age in Years"
    myCol = 7
    excelWS.Cells(myRow, myCol).Value = "Investigation"
    myCol = 8
    excelWS.Cells(myRow, myCol).Value = "Sex"

    For i = 0 To lstIxItem.ListCount - 1
        myCol = myCol + 1

        excelWS.Cells(myRow, myCol).Value = lstIxItem.List(i)
        
    Next

    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblPatientIx.PatientIxID, tblPatientIxBill.Date, tblPatient.DateOfBirth, tblSex.Sex, tblIx.Ix " & _
            "FROM (((tblPatientIx LEFT JOIN tblPatient ON tblPatientIx.PatientID = tblPatient.PatientID) LEFT JOIN tblPatientIxBill ON tblPatientIx.PatientIxBillID = tblPatientIxBill.PatientIxBillID) LEFT JOIN tblSex ON tblPatient.SexID = tblSex.SexID) LEFT JOIN tblIx ON tblPatientIx.IxID = tblIx.IxID " & _
            "WHERE tblPatientIx.IxID = " & lstIx.ItemData(lstIx.ListIndex)
        myRow = 2
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        While .EOF = False
                        
            myCol = 1
            excelWS.Cells(myRow, myCol).Value = txtCentre.Text
            myCol = 2
            excelWS.Cells(myRow, myCol).Value = !PatientIxID
            myCol = 3
            excelWS.Cells(myRow, myCol).Value = Format(!Date, "yyyy")
            myCol = 4
            excelWS.Cells(myRow, myCol).Value = Format(!Date, "M")
            myCol = 5
            excelWS.Cells(myRow, myCol).Value = Format(!Date, "d")
            myCol = 6
            excelWS.Cells(myRow, myCol).Value = DateDiff("yyyy", !DateOfBirth, !Date)
            myCol = 7
            excelWS.Cells(myRow, myCol).Value = !Ix
            myCol = 8
            excelWS.Cells(myRow, myCol).Value = !Sex
                        
            listIxItemValues !PatientIxID, excelWS, myCol, myRow
                        
            myRow = myRow + 1
            .MoveNext
        Wend
        .Close
    End With
    
    excelApp.UserControl = True
    
    
    Set excelWS = Nothing
    Set excelWB = Nothing
    Set excelApp = Nothing
    
    
End Sub

Private Sub listIxItemValues(PatientID As Long, excelWS As Excel.Worksheet, myCol As Long, myRow As Long)
    Dim rsTem As New ADODB.Recordset
    
    Dim i As Integer
    Dim temCol As Integer
    
    With rsTem
        temSQL = "SELECT tblPatientIxItem.IxItemID, tblPatientIxItem.Value " & _
                    "From tblPatientIxItem " & _
                    "WHERE (((tblPatientIxItem.PatientIxID)=" & PatientID & "))"
        If .State = 1 Then .Close
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        While .EOF = False
        
            temCol = myCol
            For i = 0 To lstIxItem.ListCount - 1
                temCol = temCol + 1
                If !IxItemID = lstIxItem.ItemData(i) Then
                    excelWS.Cells(myRow, temCol).Value = !Value
                End If
            Next
        
            .MoveNext
        Wend
        .Close
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

Private Sub lstIx_DblClick()
    ListIxItem
End Sub

Private Sub lstIxItem_Click()
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    lstValues.Clear
    With rsTem
        temSQL = "SELECT tblPatientIxItem.Value & ' ' & Count(tblPatientIxItem.PatientIxItemID) AS Summery " & _
                    "From tblPatientIxItem Where (((tblPatientIxItem.IxItemID) = " & Val(lstIxItem.ItemData(lstIxItem.ListIndex)) & ")) " & _
                    "GROUP BY tblPatientIxItem.Value ORDER BY Count(tblPatientIxItem.PatientIxItemID) DESC"
        .Open temSQL, cnnDB1, adOpenStatic, adLockReadOnly
        While .EOF = False
            lstValues.AddItem !Summery
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub txtDB1_Change()
    If cnnDB1.State = 1 Then cnnDB1.Close
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & txtDB1.Text & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=Bud7Nil"
    cnnDB1.Open constr
    ListIx
End Sub




