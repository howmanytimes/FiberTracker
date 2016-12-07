VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_sWord As String
Private m_nCount As Long

Public Property Let Count(ByVal nValue As Long)
    m_nCount = nValue
End Property

Public Property Get Count() As Long
    Count = m_nCount
End Property

Public Property Let Word(ByVal sData As String)
    m_sWord = sData
End Property

Public Property Get Word() As String
    Word = m_sWord
End Property



