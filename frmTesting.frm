VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmTesting 
   Caption         =   "Testing"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   8070
   Begin MSDataListLib.DataCombo cmbList 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox txtNewName 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtNewPrice 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim allItems As New Collection
    Dim selectedItem As New Item
    
