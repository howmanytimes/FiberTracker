VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"CWord"
Attribute VB_Ext_KEY = "Member0" ,"CWord"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Sub CountWords(ByVal str As String)
    Dim sWords() As String
    Dim n As Long, nCount As Long
    Dim obj As CWord
    
    'remove all new line characters
    str = Replace$(LCase$(str), vbCrLf, " ")
    'remove all tabs
    str = Replace$(str, vbTab, " ")
    'remove double spaces
    Do While InStr(str, "  ")
        str = Replace$(str, "  ", " ")
    Loop
    'remove other delimiters
    str = Replace$(str, ".", "")
    str = Replace$(str, ",", "")
    str = Replace$(str, "!", "")
    str = Replace$(str, "?", "")
    str = Replace$(str, ":", "")
    str = Replace$(str, ";", "")
    str = Replace$(str, """", "")
    str = Replace$(str, "'", "")
    str = Replace$(str, "(", "")
    str = Replace$(str, ")", "")
    'if you come up with any other word delimiters make sure you replace them
    sWords = Split(Trim$(str), " ")
    nCount = UBound(sWords)
    'clear the collection
    Do While mCol.Count
        mCol.Remove 1
    Loop
    On Error Resume Next
    For n = 0 To nCount
        Err.Clear
        Set obj = New CWord
        obj.Word = sWords(n)
        obj.Count = 1
        mCol.Add obj, sWords(n)
        If Err Then
            mCol(sWords(n)).Count = mCol(sWords(n)).Count + 1
        End If
    Next
End Sub

Public Property Get Item(vntIndexKey As Variant) As CWord
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub
