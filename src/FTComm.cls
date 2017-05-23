VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private o_MSComm As MSCommLib.MSComm

Public Property Set MSComm(ByRef RHS As MSComm)
    Set o_MSComm = RHS
End Property

Public Sub ClearInputBuffer()
    log "CLEARING INPUT BUFFER."
    o_MSComm.InBufferCount = 0
End Sub

Public Sub ClosePort()
    log "CLOSING COM PORT"
    o_MSComm.PortOpen = False
End Sub

Public Sub HandleEvent()
    Dim ERMsg$
    
    'Handle error conditions.
    ' Branch according to the CommEvent property.
    Select Case o_MSComm.CommEvent
        Case comBreak
            ERMsg$ = "Break Received"
        Case comCDTO
            ERMsg$ = "Carrier Detect Timeout"
        Case comCTSTO
            ERMsg$ = "CTS Timeout"
        Case comDCB
            ERMsg$ = "Error retrieving DCB"
        Case comDSRTO
            ERMsg$ = "DSR Timeout"
        Case comFrame
            ERMsg$ = "Framing Error"
        Case comOverrun
            ERMsg$ = "Overrun Error"
        Case comRxOver
            ERMsg$ = "Receive Buffer Overflow"
        Case comRxParity
            ERMsg$ = "Parity Error"
        Case comTxFull
            ERMsg$ = "Transmit Buffer Full"
        Case Else
            ERMsg$ = "Unknown error or event"
    End Select

    If Len(ERMsg$) Then
        RxComm.error = True
        
        log "COMM ERROR/EVENT " & "(" & o_MSComm.CommEvent & "): " & ERMsg$
        
        'Display any error messages to the user
        MsgBox "Communications Error " & CStr(o_MSComm.CommEvent) & ": " & ERMsg$, vbOKOnly Or vbCritical, "FiberTrack FTComm"
    End If
End Sub

Public Function OpenPort(portNumber As Integer) As Boolean
On Error GoTo ErrHan
    Debug.Assert portNumber >= 1 And portNumber <= 16
    
    log "OPENING COM PORT " & portNumber
    
    If Not o_MSComm.PortOpen Then
        o_MSComm.Settings = "19200,n,8,1"    'Open comm port
        o_MSComm.CommPort = portNumber
        o_MSComm.PortOpen = True
    End If
    
    OpenPort = True
    
    Exit Function
ErrHan:
    OpenPort = False
    'MSDN says error 62 for invalid port; trial and error indicates 8002 received
    If Err.Number = 62 Or Err.Number = 8002 Then
        log "Error " & Err.Number & ": " & Err.Description
        MsgBox "Error opening communications channel.  Please try another COM Port or check your device configuration.", vbCritical Or vbOKOnly, "FiberTrack FTComm"
    Else
        MsgBox "Unexpected error occurred while opening communications channel." & vbCrLf & "Please try another COM Port or check your device configuration." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical Or vbOKOnly, "FiberTrack FTComm"
    End If
        
End Function

Public Function ReadData() As Variant
    Dim v_Temp As Variant
    v_Temp = o_MSComm.Input
    ReadData = v_Temp
   
    If IsByteArray(v_Temp) Then
        Debug.Print "IN: " & ByteArrayToString(v_Temp)
    Else
        Debug.Print "IN: " & v_Temp
    End If
End Function

Public Sub WriteData(ByVal v As Variant)
    o_MSComm.Output = v
        
    If IsByteArray(v) Then
        log "OUT: " & ByteArrayToString(v)
    Else
        log "OUT: " & v
    End If
End Sub

'Expecting 26 byte packet from sensor.
Public Sub WaitForPacket()
    Dim Dummy
    
    log "WAITING FOR PACKET."
    
    Do Until SystemStatus.IsWait2Seconds = False Or RxComm.done
        
        'Suspend task - Give time to OS
        Dummy = DoEvents()
        
        'Check for end of xmission
        If o_MSComm.InBufferCount >= 26 Then
            RxComm.done = True
        End If
    Loop
    
    log "PACKET RECEIVED OR TIMEOUT."
End Sub

Private Function IsByteArray(ByVal v As Variant) As Boolean
    IsByteArray = (VarType(v) = vbArray Or vbByte)
End Function

Private Function ByteArrayToString(ByVal v As Variant)
    Debug.Assert IsByteArray(v)
    
    Dim l       As Long
    Dim s_Dbg   As String
    
    For l = LBound(v) To UBound(v)
        s_Dbg = s_Dbg & Hex$(v(l)) & ","
    Next l

    ByteArrayToString = s_Dbg
End Function
