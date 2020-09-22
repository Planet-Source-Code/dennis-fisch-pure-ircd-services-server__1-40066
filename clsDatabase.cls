VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarFileName As String

Public Property Let FileName(ByVal FileName As String)
Attribute FileName.VB_Description = "Specify a FileName to be used when reading/writing"
    mvarFileName = FileName
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Sub WriteEntry(Section As String, Key As String, Items As String)
WritePrivateProfileString Section, Key, Items, mvarFileName
End Sub

Public Function ReadEntry(Section As String, Key As String, Default As String) As String
Dim buf As String
buf = Space$(255)
GetPrivateProfileString Section, Key, Default, buf, Len(buf), mvarFileName
 ReadEntry = StripTerminator(buf)
End Function
