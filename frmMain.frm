VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form2"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   8790
   Begin VB.TextBox txtTem 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmMain.frx":0000
      Top             =   3240
      Width           =   4575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnChrishantha 
      Caption         =   "Chrishantha Database Import"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   4575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "!FiledName = "
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SQL"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Data From Tables"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find Database Changes"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnChrishantha_Click()
    Dim dbFile As String
    Dim constr  As String
    dbFile = getFile
    Dim cnnEMA As New ADODB.Connection
    Dim cnnTem As New ADODB.Connection
    
    Dim temSQL As String
    
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & dbFile & ";Mode=ReadWrite;Persist Security Info=True;Jet OLEDB:System database=False;Jet OLEDB:Database Password=Bud7Nil"
    cnnEMA.Open constr
    
    
    Dim rsTem As New ADODB.Recordset
    Dim rsPt As New ADODB.Recordset
    
    
    With rsTem
        temSQL = "SELECT * from temData"
        .Open temSQL, cnnEMA, adOpenStatic, adLockOptimistic
        While .EOF = False
            Dim i As Integer
            Dim temComments As String
            temComments = ""
'            For i = 3 To 10
'                If Trim(.Fields(i)) <> "" Then
'                    temComments = Trim(.Fields(i)) & vbNewLine
'                    txtTem.Text = temComments
'                    DoEvents
'                End If
'            .Fields(i) = ""
'            Next i
'            !f3 = temComments
            .Update
            .MoveNext
        Wend
        .Close
    
    End With
    
End Sub

Private Sub makeComments()

End Sub

Private Function getFile() As String
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "MDB|*.mdb"
    On Error GoTo eh
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        getFile = CommonDialog1.FileName
    End If
    Exit Function
eh:
    MsgBox "Error loading the image"
End Function


Private Sub openExcel()

End Sub



Private Sub Command3_Click()
    Form3.Show
End Sub

Private Sub Command4_Click()
    Form4.Show
End Sub
