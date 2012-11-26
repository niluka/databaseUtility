VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsCompareDatabases 
         Caption         =   "Compare Databases"
      End
      Begin VB.Menu mnuToolsCreateCode 
         Caption         =   "Create Code for SQL Server"
      End
      Begin VB.Menu mnuAccessCode 
         Caption         =   "Create code for MS-Access"
      End
      Begin VB.Menu mnuLab 
         Caption         =   "Lab"
      End
      Begin VB.Menu mnuDC 
         Caption         =   "DC"
      End
   End
   Begin VB.Menu mnuTem 
      Caption         =   "Tem"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuAccessCode_Click()
    frmCreateCode.Show
    frmCreateCode.ZOrder 0
End Sub

Private Sub mnuDC_Click()
    frmDC.Show
    frmDC.ZOrder 0
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuLab_Click()
    frmLab.Show
    frmLab.ZOrder 0
End Sub

Private Sub mnuTem_Click()
    Form3.Show
    Form4.Show
    MDIForm1.Show
End Sub

Private Sub mnuToolsCompareDatabases_Click()
    frmCompareDatabases.Show
    frmCompareDatabases.ZOrder 0
End Sub

Private Sub mnuToolsCreateCode_Click()
    frmCreateCodeSS.Show
    frmCreateCodeSS.ZOrder 0
End Sub
