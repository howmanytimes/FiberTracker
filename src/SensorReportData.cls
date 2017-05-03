VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SensorReportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Average As Double
Private m_AverageTemperature As Double
Private m_Cv As Double

Public Property Get Average() As Double
    Average = m_Average
End Property

Public Property Let Average(value As Double)
    m_Average = value
End Property

Public Property Get cv() As Double
    cv = m_Cv
End Property

Public Property Let cv(value As Double)
    m_Cv = value
End Property

Public Property Get AverageTemperature() As Double
    AverageTemperature = m_AverageTemperature
End Property

Public Property Let AverageTemperature(value As Double)
    Dim newValue As Double
    newValue = (((value / CInt(GetIniSetting("Application", "TempForumlaDivideBy"))) - CDbl(GetIniSetting("Application", "TempForumlaSubtractBy"))) * CInt(GetIniSetting("Application", "TempForumlaMultipleBy")))
    m_AverageTemperature = newValue
End Property
