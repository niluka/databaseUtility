VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarItemName As String 'local copy
Private mvarItemPrice As Double 'local copy
Public Property Let ItemPrice(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemPrice = 5
    mvarItemPrice = vData
End Property


Public Property Get ItemPrice() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemPrice
    ItemPrice = mvarItemPrice
End Property



Public Property Let ItemName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ItemName = 5
    mvarItemName = vData
End Property


Public Property Get ItemName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ItemName
    ItemName = mvarItemName
End Property



