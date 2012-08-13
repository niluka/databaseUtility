VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   9810
   Begin VB.CommandButton btnNAUCKD 
      Caption         =   "Non Albumin Uric CKD"
      Height          =   495
      Left            =   5640
      TabIndex        =   25
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton btnAgeAtDCat 
      Caption         =   "Age at D Cat"
      Height          =   495
      Left            =   2880
      TabIndex        =   24
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton btnAlbuminuria 
      Caption         =   "Albuminuria"
      Height          =   495
      Left            =   2880
      TabIndex        =   23
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton btnBPCombined 
      Caption         =   "BP Combined"
      Height          =   495
      Left            =   2880
      TabIndex        =   22
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton btnDysLipiCom 
      Caption         =   "Dyslipidemia (Combined)"
      Height          =   495
      Left            =   2880
      TabIndex        =   21
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton btnGlobalObesity 
      Caption         =   "GlobalObesity"
      Height          =   495
      Left            =   2880
      TabIndex        =   20
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton btnIPPH 
      Caption         =   "IPPH"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton btnDys 
      Caption         =   "Dyslipidemia"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CommandButton btnCentralObesity 
      Caption         =   "CentralObesity"
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton btnBP 
      Caption         =   "BP"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton btnEGFRStatus 
      Caption         =   "EGFR Status"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton btnAlbuminStatus 
      Caption         =   "AlbuminStatus"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Values"
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8175
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   4920
         TabIndex        =   11
         Top             =   840
         Width           =   3135
         Begin VB.OptionButton optFromNumeric 
            Caption         =   "Numeric"
            Height          =   195
            Left            =   1680
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optFromString 
            Caption         =   "String"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtTable 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton btnConvert 
         Caption         =   "Convert"
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtToField 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtFromField 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Table"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "To Field"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "From Field"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDB1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   8055
   End
   Begin VB.CommandButton btnDB1 
      Caption         =   "Select"
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim cnnDB1 As New ADODB.Connection
    Dim constr As String
    Dim temSQL As String
    
Private Sub btnAgeAtDCat_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If IsNull(!AgeAtD) = True Then
                !AgeAtOnsetCategory = "NA"
            ElseIf !AgeAtD = 0 Then
                !AgeAtOnsetCategory = "NA"
            ElseIf !AgeAtD < 30 Then
                !AgeAtOnsetCategory = "<30"
            Else
                !AgeAtOnsetCategory = "=>30"
            End If
            .Update
            .MoveNext
        Wend
        .Close
    
    End With


End Sub

Private Sub btnAlbuminStatus_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If !UTI = "Yes" Then
                !AlbuminStatus = "UTI Present"
                !AlbuminStatusDetail = "UTI Present"
            Else
                If !UFRAlbuminPresent = "Present" Then
                    !AlbuminStatus = "Macroalbuminuria"
                    !AlbuminStatusDetail = "Macroalbuminuria"
                ElseIf !UFRAlbuminPresent = "Nill" Then
                    If !UMA > 30 Then
                        !AlbuminStatus = "Microalbuminuria"
                        !AlbuminStatusDetail = "Microalbuminuria"
                    Else
                        !AlbuminStatus = "Normal Excretion"
                        !AlbuminStatusDetail = "Normal Excretion"
                    End If
                ElseIf !UFRAlbuminPresent = "Not Available" Then
                    If !UMA > 30 Then
                        !AlbuminStatus = "Microalbuminuria"
                        !AlbuminStatusDetail = "Microalbuminuria"
                    Else
                        !AlbuminStatus = "Not Available"
                        !AlbuminStatusDetail = "Not Available"
                    End If
                End If
            End If
            .Update
            .MoveNext
        Wend
        .Close
    
    End With
End Sub

Private Sub btnAlbuminuria_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If !UTI = "Yes" Then
                !Albuminuria = "NA"
                !Albuminuria = "NA"
            Else
                If !UFRAlbuminPresent = "Present" Then
                    !Albuminuria = "Present"
                ElseIf !UFRAlbuminPresent = "Nill" Then
                    If !UMA > 300 Then
                        !Albuminuria = "Present"
                    Else
                        !Albuminuria = "Absent"
                    End If
                ElseIf !UFRAlbuminPresent = "Not Available" Then
                    If !UMA > 300 Then
                        !Albuminuria = "Present"
                    ElseIf !UMA = 0 Then
                        !Albuminuria = "NA"
                    Else
                        !Albuminuria = "Absent"
                    End If
                End If
            End If
            .Update
            .MoveNext
        Wend
        .Close
    
    End With

End Sub

Private Sub btnBP_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If IsNull(!SBP) = True Or IsNull(!DBP) = True Then
                !BPStatus = "NA"
            ElseIf (!SBP) = 0 Or (!DBP) = 0 Then
                !BPStatus = "NA"
            ElseIf (!SBP) > 130 And (!DBP) > 85 Then
                !BPStatus = "Both SPB>130 & DBP>85"
            ElseIf (!SBP) > 130 Then
                !BPStatus = "SPB>130"
            ElseIf (!DBP) > 85 Then
                !BPStatus = "DBP>85"
            Else
                !BPStatus = "Normal"
            End If
            
            If IsNull(!SBP) = True Or IsNull(!DBP) = True Then
                !BPStatus = "NA"
            ElseIf (!SBP) = 0 Or (!DBP) = 0 Then
                !BPStatus = "NA"
            ElseIf (!SBP) > 160 And (!DBP) <= 90 Then
                !ISHStatus = "ISH Present"
            Else
                !ISHStatus = "ISH Absent"
            End If
            
            .Update
            .MoveNext
        Wend
        .Close
    
    End With


End Sub

Private Sub btnBPCombined_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If IsNull(!SBP) = True Or IsNull(!DBP) = True Then
                !Hypertension = "NA"
            ElseIf (!SBP) = 0 Or (!DBP) = 0 Then
                !Hypertension = "NA"
            ElseIf (!SBP) > 130 And (!DBP) > 85 Then
                !Hypertension = "Present"
            ElseIf (!SBP) > 130 Then
                !Hypertension = "Present"
            ElseIf (!DBP) > 85 Then
                !Hypertension = "Present"
            Else
                !Hypertension = "Absent"
            End If
            .Update
            .MoveNext
        Wend
        .Close
    
    End With

End Sub

Private Sub btnCentralObesity_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temInt As Integer
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            temInt = 0
            If !Sex = "Male" Then
                If IsNull(!Wst) = True Then
                    !CentralObesity = "NA"
                ElseIf (!Wst) > 90 Then
                    !CentralObesity = "Present"
                Else
                    !CentralObesity = "Absent"
                End If
            ElseIf !Sex = "Female" Then
                If IsNull(!Wst) = True Then
                    !CentralObesity = "NA"
                ElseIf (!Wst) > 80 Then
                    !CentralObesity = "Present"
                Else
                    !CentralObesity = "Absent"
                End If
            Else
                !CentralObesity = "NA"
            End If
            .Update
            .MoveNext
        Wend
        .Close
    End With


End Sub

Private Sub btnConvert_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temTxt As String
    With rsTem
        If chkClear.Value = 1 Then
            If .State = 1 Then .Close
            temSQL = "SELECT * from " & txtTable.Text
            .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
            While .EOF = False
                .Fields(txtToField.Text) = Null
                .Update
                .MoveNext
            Wend
            .Close
        End If
searchAgain:
        If .State = 1 Then .Close
        temSQL = "SELECT " & txtFromField.Text & ", " & txtToField.Text & "  from " & txtTable.Text & " GROUP BY  " & txtFromField.Text & ", " & txtToField.Text
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If IsNull(.Fields(txtToField.Text)) = True Then
                temTxt = InputBox("New Value for " & .Fields(txtFromField.Text))
                updateFields .Fields(txtFromField.Text), temTxt
                GoTo searchAgain
            End If
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub updateFields(FromValue As String, ToValue As String)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        If optFromString.Value = True Then
            temSQL = "SELECT * from " & txtTable.Text & " Where " & txtFromField.Text & " = '" & FromValue & "'"
        Else
            temSQL = "SELECT * from " & txtTable.Text & " Where " & txtFromField.Text & " = " & FromValue
        End If
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            .Fields(txtToField.Text) = ToValue
            .Update
            .MoveNext
        Wend
        .Close
    End With
End Sub

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



Private Sub btnDys_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temInt As Integer
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            temInt = 0
            If IsNull(!LDL) = True Or IsNull(!TG) = True Or IsNull(!HDL) = True Then
                !DyslipidemiaStatus = "NA"
            ElseIf (!LDL) = 0 Or (!TG) = 0 Or (!HDL) = 0 Then
                !DyslipidemiaStatus = "NA"
            ElseIf IsNull(!Sex) Then
                !DyslipidemiaStatus = "NA"
            ElseIf (!Sex <> "Male") And !Sex <> "Female" Then
                !DyslipidemiaStatus = "NA"
            ElseIf !Sex = "" Then
                !DyslipidemiaStatus = "NA"
            Else
                If (!LDL) >= 100 Then temInt = temInt + 1
                If (!TG) >= 150 Then temInt = temInt + 1
                If (!HDL) < 40 And (!Sex) = "Male" Then temInt = temInt + 1
                If (!HDL) < 50 And (!Sex) = "Female" Then temInt = temInt + 1
                Select Case temInt
                    Case 0: !DyslipidemiaStatus = "Normal"
                    Case 1: !DyslipidemiaStatus = "One Abnormal"
                    Case 2: !DyslipidemiaStatus = "Two Abnormal"
                    Case 3: !DyslipidemiaStatus = "Three Abnormal"
                    Case Else: !DyslipidemiaStatus = "No"
                End Select
                
            End If
            .Update
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub btnDysLipiCom_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temInt As Integer
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            temInt = 0
            If IsNull(!LDL) = True Or IsNull(!TG) = True Or IsNull(!HDL) = True Then
                !Dyslipidemia = "NA"
            ElseIf (!LDL) = 0 Or (!TG) = 0 Or (!HDL) = 0 Then
                !Dyslipidemia = "NA"
            ElseIf IsNull(!Sex) Then
                !Dyslipidemia = "NA"
            ElseIf (!Sex <> "Male") And !Sex <> "Female" Then
                !Dyslipidemia = "NA"
            ElseIf !Sex = "" Then
                !Dyslipidemia = "NA"
            Else
                If (!LDL) > 100 Then temInt = temInt + 1
                If (!TG) > 150 Then temInt = temInt + 1
                If (!HDL) < 40 And (!Sex) = "Male" Then temInt = temInt + 1
                If (!HDL) < 50 And (!Sex) = "Female" Then temInt = temInt + 1
                Select Case temInt
                    Case 0: !Dyslipidemia = "Absent"
                    Case Else: !Dyslipidemia = "Present"
                End Select
                
            End If
            .Update
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub btnEGFRStatus_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If IsNull(!EGFR) = True Then
                !EGFRStatus = "NA"
            ElseIf !EGFR = 0 Then
                !EGFRStatus = "NA"
            ElseIf !EGFR < 15 Then
                !EGFRStatus = "< 15"
            ElseIf !EGFR < 30 Then
                !EGFRStatus = "15 - 29"
            ElseIf !EGFR < 61 Then
                !EGFRStatus = "30 - 60"
            Else
                !EGFRStatus = "> 60"
            End If
            .Update
            .MoveNext
        Wend
        .Close
    
    End With

End Sub

Private Sub btnGlobalObesity_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temInt As Integer
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            temInt = 0
            If IsNull(!BMI) = True Then
                !GlobalObesity = "NA"
            ElseIf (!BMI) < 5 Then
                !GlobalObesity = "NA"
            ElseIf (!BMI) > 23 Then
                !GlobalObesity = "Present"
            Else
                !GlobalObesity = "Absent"
            End If
            .Update
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub btnIPPH_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If IsNull(!FBS) = True Or IsNull(!PPBS) = True Or IsNull(!HbA1c) = True Then
                !IsolatedPPHStuatus = "NA"
            ElseIf (!FBS) = 0 Or (!PPBS) = 0 Or (!HbA1c) = 0 Then
                !IsolatedPPHStuatus = "NA"
            ElseIf (!FBS) < 125 And (!PPBS) > 180 And (!HbA1c) < 7.5 Then
                !IsolatedPPHStuatus = "Isolated PP Hypeglycemia"
            Else
                !IsolatedPPHStuatus = "No"
            End If
            .Update
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub btnNAUCKD_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * from NewAllFirstVisits"
        .Open temSQL, cnnDB1, adOpenStatic, adLockOptimistic
        While .EOF = False
            If IsNull(!EGFR) = True Then
                !NACKD = ""
            ElseIf !EGFR = 0 Then
                !NACKD = ""
            ElseIf !EGFR >= 60 Then
                !NACKD = ""
            Else
                If !Albuminuria = "NA" Then
                    !NACKD = ""
                ElseIf !Albuminuria = "Present" Then
                    !NACKD = "ACKD"
                ElseIf !Albuminuria = "Absent" Then
                    !NACKD = "NACKD"
                Else
                    !NACKD = ""
                End If
            End If
            .Update
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


Private Sub txtDB1_Change()
    If cnnDB1.State = 1 Then cnnDB1.Close
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & txtDB1.Text & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=Bud7Nil"
    cnnDB1.Open constr
End Sub





