VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarFieldName As String 'local copy
Private mvarFieldType As MyDataType 'local copy

Public Property Let FieldType(ByVal vData As MyDataType)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.FieldType = Form1
    mvarFieldType = vData
End Property


Public Property Get FieldType() As MyDataType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldType
    FieldType = mvarFieldType
End Property



Public Property Let FieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldName = 5
    mvarFieldName = vData
End Property


Public Property Get FieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldName
    FieldName = mvarFieldName
End Property



