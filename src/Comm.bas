Attribute VB_Name = "modComm"
Option Explicit

'***Comm parameters
Public Host_send(13)            As Byte         'Comm routine output buffer (14 bytes)
Public Command_Packet(4)        As Byte         'Comm routine host command packet (5 bytes)
'JW 2/27/00 Created alternate command packet as 17 bytes.
Public Command_Packet2(16)      As Byte         'Comm routine host command packet (17 bytes)

Private Const HR_PKTLEN         As Integer = (26 * LAST_SENSOR) - 1
Public Host_rcv(HR_PKTLEN)      As Byte         'RxComm input buffer (26 bytes each)

Public Const RcvBUF_LENGTH      As Byte = 26

Public Type RxComm_info                 'Contains Rxcomm parameters
    Sensor          As Byte             'Current sensor using RxComm
    Index           As Byte             'Index into Rx buffer
    done            As Boolean          'Indicates Xmission complete
    error           As Boolean          'Error in Rx packet
End Type

Public RxComm                   As RxComm_info

Public Const FF_BYTE            As Byte = &HED        'Comm routine ff_byte
Public Const ALL_CHANNELS       As Byte = &H41

'Define host to sensor commands
Public Const INIT_COMMAND       As Byte = &H41
Public Const ZERO_COMMAND       As Byte = &H42
Public Const CAL_COMMAND        As Byte = &H43
Public Const POLL_COMMAND       As Byte = &H46
Public Const STARTDATA_COMMAND  As Byte = &H44
Public Const STOPDATA_COMMAND   As Byte = &H47

'Define sensor to host commands
Public Const INIT_COMPLETE      As Byte = &H69
Public Const ZERO_PACKET        As Byte = &H7A
Public Const CAL_PACKET         As Byte = &H63
Public Const SEND_DATA          As Byte = &H64

'Define PIC sensor status word
Public Const PICSTATUS_B0       As Byte = &H1   'bit 0 = 1 = initialization complete
Public Const PICSTATUS_B1       As Byte = &H2   'bit 1 = 1 = zero complete
Public Const PICSTATUS_B2       As Byte = &H4   'bit 2 = 1 = calibration complete
Public Const PICSTATUS_B3       As Byte = &H8   'bit 3 = 1 = poll command received
Public Const PICSTATUS_B4       As Byte = &H10  'bit 4 = 1 = stop command received
Public Const PICSTATUS_B5       As Byte = &H20  'bit 5 = 1 = data collection in progress
Public Const PICSTATUS_B6       As Byte = &H40  'bit 6 = 1 = error detected
Public Const PICSTATUS_B7       As Byte = &H80  'bit 7 = 1 = not defined

'Definition of receive packet format
Public Const S_PIC              As Byte = 1     'sensor pic number
Public Const S_PKT              As Byte = 2     'sensor packet number
Public Const S_CVS3             As Byte = 3     'sensor coefficient of variance sum, cvSummary3
Public Const S_CVS2             As Byte = 4     'sensor coefficient of variance sum, cvSummary2
Public Const S_CVS1             As Byte = 5     'sensor coefficient of variance sum, cvSummary1
Public Const S_CVS0             As Byte = 6     'sensor coefficient of variance sum, cvSummary0
Public Const S_DAT_HI           As Byte = 7     'sensor ave diameter
Public Const S_DAT_LO           As Byte = 8     'sensor ave diameter
Public Const S_L1S              As Byte = 9     'sensor level 1 slub count
Public Const S_L1T              As Byte = 10    'sensor level 1 thin spot count
Public Const S_L2S              As Byte = 11    'sensor level 2 slub count
Public Const S_L2T              As Byte = 12    'sensor level 2 thin spot count
Public Const S_STATUS           As Byte = 13    'sensor system status byte
Public Const Z_DAT_HI           As Byte = 14    'sensor zero datran value MSB
Public Const Z_DAT_LO           As Byte = 15    'sensor zero datran value LSB
Public Const C_DAT_HI           As Byte = 16    'sensor calibration datran value MSB
Public Const C_DAT_LO           As Byte = 17    'sensor claibration datran value LSB
Public Const E_CODE             As Byte = 18    'sensor error code
Public Const E_ERR1             As Byte = 19    'sensor error byte
Public Const E_ERR2             As Byte = 20    'sensor error byte
Public Const E_SUM              As Byte = 21    'sensor error summary count
Public Const S_SEQ_NUM          As Byte = 22    'sensor sequence number
Public Const S_IDLE             As Byte = 23    'sensor idle time in main17 wait loop
Public Const F_RES3             As Byte = 24    'sensor reserved for future use 3
Public Const S_CKS              As Byte = 25    'sensor checksum
'***End of Comm parameters


'Compute checksum for given packet byte array.  First and last bytes are not computed
'in the checksum, which is stored in the last byte of the array.
Public Sub PacketChecksum(ByRef Packet() As Byte)
    Dim temp    As Byte
    Dim i       As Integer
    
    'The first and last bytes are not computed in the checksum.
    temp = 0
    
    For i = (LBound(Packet) + 1) To (UBound(Packet) - 1)
        temp = temp Xor Packet(i)
    Next i
    
    Packet(UBound(Packet)) = temp           'Last byte gets checksum
    
    log "Checksum for " & (UBound(Packet) + 1) & " byte packet: " & Hex$(temp)
End Sub

'Insert word into 2 successive bytes of comm. output buffer.
Public Sub InsertWord(ByRef Packet() As Byte, InputWord As Long, ByteIndex As Byte)
    Debug.Assert UBound(Packet) > ByteIndex And LBound(Packet) < ByteIndex
    
    If InputWord < 256 Then
        Packet(ByteIndex) = 0                           'High byte
        Packet(ByteIndex + 1) = InputWord               'Low byte
    Else
        Packet(ByteIndex) = InputWord \ 256            'High byte
        Packet(ByteIndex + 1) = InputWord Mod 256      'Low byte
    End If
End Sub
