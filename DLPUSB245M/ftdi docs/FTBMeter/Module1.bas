Attribute VB_Name = "Module1"
Option Explicit
Public RegKey As String                             ' name of registry key
Public ZerodBm As Single                            ' calibration level of 0dBm for current range
Public ZerodBmHF As Single                          ' calibration level of 0dBm for HF
Public ZerodBmVHF As Single                         ' calibration level of 0dBm for VHF
Public ZerodBmUHF As Single                         ' calibration level of 0dBm for UHF
Public Minus40dBm As Single                         ' calibration level of -40dBm for current range
Public Minus40dBmHF As Single                       ' calibration level of -40dBm for HF
Public Minus40dBmVHF As Single                      ' calibration level of -40dBm for VHF
Public Minus40dBmUHF As Single                      ' calibration level of -40dBm for UHF
Public Slope As Single                              ' number of reading units per dBm
Public Const Green = &HFF00&
Public Const Red = &HFF&
Public Const White = &HFFFFFF
Public Const Yellow = &HFFFF&
Public Const ButtonFace = &H8000000F
Public NumberOfReadings                             ' number of readings taken when continuous
Public StopReading As Boolean                       ' true = stop continuous readings
Public OurDevice As String                          ' the name of our DLP2232 device
Public Reading As Single                            ' actual value of reading from ADC
Public Saved_Port_Value As Byte                     ' the setting of the first 8 data lines
Public OutIndex As Long                             ' position within the output buffer
Public PortAIsOpen As Boolean                       ' true = the DLP2232M chan A is open

'==============================
'CLASSIC INTERFACE DECLARATIONS
'==============================
Public Declare Function FT_ListDevices Lib "FTD2XX.DLL" ( _
                                    ByVal arg1 As Long, _
                                    ByVal arg2 As String, _
                                    ByVal dwFlags As Long) As Long
                                    
Public Declare Function FT_GetNumDevices Lib "FTD2XX.DLL" Alias "FT_ListDevices" ( _
                                    ByRef arg1 As Long, _
                                    ByVal arg2 As String, _
                                    ByVal dwFlags As Long) As Long
                                    
Public Declare Function FT_Open Lib "FTD2XX.DLL" ( _
                                    ByVal intDeviceNumber As Integer, _
                                    ByRef lngHandle As Long) As Long
                                    
Public Declare Function FT_OpenEx Lib "FTD2XX.DLL" ( _
                                    ByVal arg1 As String, _
                                    ByVal arg2 As Long, _
                                    ByRef lngHandle As Long) As Long
                                    
Public Declare Function FT_Close Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long) As Long
                                    
Public Declare Function FT_Read Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal lpszBuffer As String, _
                                    ByVal lngBufferSize As Long, _
                                    ByRef lngBytesReturned As Long) As Long
                                    
Public Declare Function FT_Write Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal lpszBuffer As String, _
                                    ByVal lngBufferSize As Long, _
                                    ByRef lngBytesWritten As Long) As Long
                                    
Public Declare Function FT_WriteByte Lib "FTD2XX.DLL" Alias "FT_Write" ( _
                                    ByVal lngHandle As Long, _
                                    ByRef lpszBuffer As Any, _
                                    ByVal lngBufferSize As Long, _
                                    ByRef lngBytesWritten As Long) As Long
                                    
Public Declare Function FT_SetBaudRate Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal lngBaudRate As Long) As Long
                                    
Public Declare Function FT_SetDataCharacteristics Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal byWordLength As Byte, _
                                    ByVal byStopBits As Byte, _
                                    ByVal byParity As Byte) As Long
                                    
Public Declare Function FT_SetFlowControl Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal intFlowControl As Integer, _
                                    ByVal byXonChar As Byte, _
                                    ByVal byXoffChar As Byte) As Long
                                    
Public Declare Function FT_SetDtr Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long) As Long
                                    
Public Declare Function FT_ClrDtr Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long) As Long
                                    
Public Declare Function FT_SetRts Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long) As Long
                                    
Public Declare Function FT_ClrRts Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long) As Long
                                    
Public Declare Function FT_GetModemStatus Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByRef lngModemStatus As Long) As Long
                                    
Public Declare Function FT_SetChars Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal byEventChar As Byte, _
                                    ByVal byEventCharEnabled As Byte, _
                                    ByVal byErrorChar As Byte, _
                                    ByVal byErrorCharEnabled As Byte) As Long
                                    
Public Declare Function FT_Purge Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal lngMask As Long) As Long
                                    
Public Declare Function FT_SetTimeouts Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal lngReadTimeout As Long, _
                                    ByVal lngWriteTimeout As Long) As Long
                                    
Public Declare Function FT_GetQueueStatus Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByRef lngRxBytes As Long) As Long
                                    
Public Declare Function FT_SetBreakOn Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long) As Long
                                    
Public Declare Function FT_SetBreakOff Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long) As Long
                                    
Public Declare Function FT_GetStatus Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByRef lngRxBytes As Long, _
                                    ByRef lngTxBytes As Long, _
                                    ByRef lngEventsDWord As Long) As Long
                                    
Public Declare Function FT_SetEventNotification Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal dwEventMask As Long, _
                                    ByVal pVoid As Long) As Long
                                    
Public Declare Function FT_ResetDevice Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long) As Long
                                    
Public Declare Function FT_GetBitMode Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByRef intData As Any) As Long
                                    
Public Declare Function FT_SetBitMode Lib "FTD2XX.DLL" ( _
                                    ByVal lngHandle As Long, _
                                    ByVal intMask As Byte, _
                                    ByVal intMode As Byte) As Long
                                    
Public Declare Function FT_SetLatencyTimer Lib "FTD2XX.DLL" ( _
                                    ByVal Handle As Long, _
                                    ByVal pucTimer As Byte) As Long
                                    
Public Declare Function FT_GetLatencyTimer Lib "FTD2XX.DLL" ( _
                                    ByVal Handle As Long, _
                                    ByRef ucTimer As Long) As Long
                                    
' Return codes
Public Const FT_OK = 0
Public Const FT_INVALID_HANDLE = 1
Public Const FT_DEVICE_NOT_FOUND = 2
Public Const FT_DEVICE_NOT_OPENED = 3
Public Const FT_IO_ERROR = 4
Public Const FT_INSUFFICIENT_RESOURCES = 5
Public Const FT_INVALID_PARAMETER = 6
Public Const FT_INVALID_BAUD_RATE = 7
Public Const FT_DEVICE_NOT_OPENED_FOR_ERASE = 8
Public Const FT_DEVICE_NOT_OPENED_FOR_WRITE = 9
Public Const FT_FAILED_TO_WRITE_DEVICE = 10
Public Const FT_EEPROM_READ_FAILED = 11
Public Const FT_EEPROM_WRITE_FAILED = 12
Public Const FT_EEPROM_ERASE_FAILED = 13
Public Const FT_EEPROM_NOT_PRESENT = 14
Public Const FT_EEPROM_NOT_PROGRAMMED = 15
Public Const FT_INVALID_ARGS = 16
Public Const FT_NOT_SUPPORTED = 17
Public Const FT_OTHER_ERROR = 18

' Flags for FT_OpenEx
Public Const FT_OPEN_BY_SERIAL_NUMBER = 1
Public Const FT_OPEN_BY_DESCRIPTION = 2

' Flags for FT_ListDevices
Public Const FT_LIST_NUMBER_ONLY = &H80000000
Public Const FT_LIST_BY_INDEX = &H40000000
Public Const FT_LIST_ALL = &H20000000

' IO buffer sizes
Public Const FT_In_Buffer_Size = 1024
Public Const FT_Out_Buffer_Size = 1024

Public FT_In_Buffer As String * FT_In_Buffer_Size
Public FT_Out_Buffer As String * FT_Out_Buffer_Size
Public FT_IO_Status As Long
Public FT_Result As Long
Public FT_Device_Count As Long
Public FT_Device_String_Buffer As String * 50
Public FT_Device_String As String

Global lngHandle As Long

Public FT_HANDLE As Long
Public PV_Device As Integer
Public FT_Q_Bytes As Long

Public Sub AddToBuffer(I As Long)
' add a character to the output buffer
    
    Mid(FT_Out_Buffer, OutIndex + 1, 1) = Chr(I)
    OutIndex = OutIndex + 1
    
End Sub

Public Function Close_USB_Device() As Long
' close the module

    FT_Result = FT_Close(FT_HANDLE)
    If FT_Result <> FT_OK Then
        FT_Error_Report "FT_Close", FT_Result
    End If
    Close_USB_Device = FT_Result
    
End Function

Public Sub FormatReading(Reading As Single)
' format a Reading in dBm
Dim Difference As Single
Dim dB As Single

    If Reading = 0 Then Exit Sub                    ' if 0 then exit
    
    Difference = Abs(ZerodBm - Reading)             ' calc reading difference to 0dBm
    dB = Difference / Slope                         ' calc how many dBm
    If Reading < ZerodBm Then
        Form1.lblDBM.Caption = Format(dB, "-#0.0dBm") ' format with a -
    Else
        Form1.lblDBM.Caption = Format(dB, "#0.0dBm") ' format without a -
    End If

End Sub

Public Sub FT_Error_Report(ErrStr As String, PortStatus As Long)
' show an error message
Dim Str As String

    Select Case PortStatus
        Case FT_INVALID_HANDLE
            Str = ErrStr & " - Invalid Handle"
        Case FT_DEVICE_NOT_FOUND
            Str = ErrStr & " - Device Not Found"
        Case FT_DEVICE_NOT_OPENED
            Str = ErrStr & " - Device Not Opened"
        Case FT_IO_ERROR
            Str = ErrStr & " - General IO Error"
        Case FT_INSUFFICIENT_RESOURCES
            Str = ErrStr & " - Insufficient Resources"
        Case FT_INVALID_PARAMETER
            Str = ErrStr & " - Invalid Parameter"
        Case FT_INVALID_BAUD_RATE
            Str = ErrStr & " - Invalid Baud Rate"
        Case FT_DEVICE_NOT_OPENED_FOR_ERASE
            Str = ErrStr & " - Device not opened for Erase"
        Case FT_DEVICE_NOT_OPENED_FOR_WRITE
            Str = ErrStr & " - Device not opened for Write"
        Case FT_FAILED_TO_WRITE_DEVICE
            Str = ErrStr & " - Failed to write Device"
        Case FT_EEPROM_READ_FAILED
            Str = ErrStr & " - EEPROM read failed"
        Case FT_EEPROM_WRITE_FAILED
            Str = ErrStr & " - EEPROM write failed"
        Case FT_EEPROM_ERASE_FAILED
            Str = ErrStr & " - EEPROM erase failed"
        Case FT_EEPROM_NOT_PRESENT
            Str = ErrStr & " - EEPROM not present"
        Case FT_EEPROM_NOT_PROGRAMMED
            Str = ErrStr & " - EEPROM not programmed"
        Case FT_INVALID_ARGS
            Str = ErrStr & " - Invalid Arguments"
        Case FT_NOT_SUPPORTED
            Str = ErrStr & " - not supported"
        Case FT_OTHER_ERROR
            Str = ErrStr & " - other error"
    End Select
    
    Form1.shpOK.BackColor = Red                     ' turn status indicator red
    StopReading = True                              ' turn off continuous readings
    Form1.lblStatus.Caption = Str                   ' show the message in the status area
    MsgBox Str                                      ' display the message
    
End Sub

Public Function Get_USB_Device_QueueStatus() As Long
' return the number of bytes waiting to be read

    FT_Result = FT_GetQueueStatus(FT_HANDLE, FT_Q_Bytes)
    If FT_Result <> FT_OK Then
        FT_Error_Report "FT_GetQueueStatus", FT_Result
    End If
    Get_USB_Device_QueueStatus = FT_Result

End Function

Public Function GetDeviceString() As String
' get the device name

    GetDeviceString = Left(FT_Device_String_Buffer, InStr(FT_Device_String_Buffer, Chr(0)) - 1)
    
End Function

Public Function GetFTDeviceCount() As Long
' get the number of connected devices
    
    FT_Result = FT_GetNumDevices(FT_Device_Count, 0, FT_LIST_NUMBER_ONLY)
    If FT_Result = FT_OK Then
        GetFTDeviceCount = FT_Device_Count          ' return the number of devices
    Else
        FT_Error_Report "GetFTDeviceCount", FT_Result ' show error message
        GetFTDeviceCount = 0                        ' return 0 devices
    End If
    
End Function

Public Function GetFTDeviceDescription(DeviceIndex As Long) As String
' get the device description of a specific device
    
    FT_Result = FT_ListDevices(DeviceIndex, FT_Device_String_Buffer, (FT_OPEN_BY_DESCRIPTION Or FT_LIST_BY_INDEX))
    If FT_Result = FT_OK Then
        FT_Device_String = GetDeviceString          ' strip off trailing nulls
        GetFTDeviceDescription = FT_Device_String   ' return the character part
    Else
        FT_Error_Report "GetFTDeviceDescription", FT_Result
        GetFTDeviceDescription = ""                 ' init to null
    End If
    
End Function

Public Function GetFTDeviceSerialNo(DeviceIndex As Long) As String
' get the serial number of a specific device
    
    FT_Result = FT_ListDevices(DeviceIndex, FT_Device_String_Buffer, (FT_OPEN_BY_SERIAL_NUMBER Or FT_LIST_BY_INDEX))
    If FT_Result = FT_OK Then
        FT_Device_String = GetDeviceString          ' strip off trailing nulls
        GetFTDeviceSerialNo = FT_Device_String      ' return the character part
    Else
        FT_Error_Report "GetFTDeviceSerialNo", FT_Result
        GetFTDeviceSerialNo = ""                    ' init to null
    End If
    
End Function

Public Function Init_Controller(DName As String) As Boolean
' initialise the controller on port DName

    Init_Controller = OpenPort(DName)               ' open the port

End Function

Public Sub InitialiseVariables()
' initialise variables

    RegKey = "FTBMeter"
    OurDevice = "DLP2232M A"                        ' set the name of our DLP2232M
    ZerodBmHF = GetSetting(RegKey, "Settings", "ZerodBmHF", 2556)
    ZerodBmVHF = GetSetting(RegKey, "Settings", "ZerodBmVHF", 2519)
    ZerodBmUHF = GetSetting(RegKey, "Settings", "ZerodBmUHF", 2501)
    Minus40dBmHF = GetSetting(RegKey, "Settings", "Minus40dBmHF", 915)
    Minus40dBmVHF = GetSetting(RegKey, "Settings", "Minus40dBmVHF", 913)
    Minus40dBmUHF = GetSetting(RegKey, "Settings", "Minus40dBmUHF", 872)
    ZerodBm = ZerodBmVHF
    Minus40dBm = Minus40dBmVHF
    Slope = (ZerodBm - Minus40dBm) / 40
    Form1.cmdHF.BackColor = ButtonFace
    Form1.cmdVHF.BackColor = Green
    Form1.cmdUHF.BackColor = ButtonFace

End Sub

Public Function Open_USB_Device_By_Description(Device_Description As String) As Long

    SetDeviceString Device_Description
    FT_Result = FT_OpenEx(FT_Device_String_Buffer, FT_OPEN_BY_DESCRIPTION, FT_HANDLE)
    If FT_Result <> FT_OK Then
        FT_Error_Report "Open_USB_Device_By_Description", FT_Result
    End If
    
End Function

Public Sub OpenDevice()
' open the DLP2232M module by name. The A port is the only one that can be used for MPSSE SPI
' communications.
Dim I As Long
Dim X As Long
Dim DeviceDescription As String
Dim FoundDevice As Boolean
Dim Res As Long

    ' if the port is already open then close it
    If PortAIsOpen Then
        Res = Close_USB_Device
        If FT_Result <> FT_OK Then
            PortAIsOpen = False
            Form1.shpOK.BackColor = Red
            Form1.lblStatus.Caption = "Attempt to close DLP2232M failed."
            StopReading = True
            Exit Sub
        End If
    End If
    
    ' set port A not open
    PortAIsOpen = False
    
    ' see if anything connected
    X = GetFTDeviceCount
    Debug.Print X
    If X = 0 Then
        Form1.shpOK.BackColor = Yellow
        Form1.lblStatus.Caption = "No FTDI devices found. Please connect the meter and re-try"
        Exit Sub
    End If
    
    ' get the descriptions and look for DLP module channel A
    For I = 0 To FT_Device_Count - 1
        DeviceDescription = GetFTDeviceDescription(I)
        Debug.Print DeviceDescription
        If FT_Result = FT_OK Then
            If DeviceDescription = "DLP-USB245M" Then
                FoundDevice = True
                Exit For
            End If
        End If
    Next

    ' check we have a DLP A module found
    If Not (FoundDevice) Then
        Form1.shpOK.BackColor = Yellow
        Form1.lblStatus.Caption = "No DLP2232M A device found. Please re-connect the meter and re-try"
        Exit Sub
    End If

    ' open by the device description
    Open_USB_Device_By_Description DeviceDescription
    If FT_Result <> FT_OK Then
        Form1.shpOK.BackColor = Red
        StopReading = True
        Form1.lblStatus.Caption = "The open for the meter did not complete successfully."
        Exit Sub
    End If

'    ' try a command
'    Res = Get_USB_Device_QueueStatus
'    If FT_Result <> FT_OK Then
'        Form1.shpOK.BackColor = Red
'        StopReading = True
'        Form1.lblStatus.Caption = "Get USB Device QueuStatus command failed in procedure OpenDevice"
'        Exit Sub
'    End If
'    PortAIsOpen = True
'
'    ' set the latency
'    FT_Result = Set_USB_Device_LatencyTimer(16)
'    If FT_Result <> FT_OK Then
'        Form1.shpOK.BackColor = Red
'        StopReading = True
'        Form1.lblStatus.Caption = "Set USB Device Latency Timer failed"
'        Exit Sub
'    End If
'
'    ' reset the controller
'    FT_Result = Set_USB_Device_BitMode(&H0, &H0) ' reset the controller
'    If FT_Result <> FT_OK Then
'        Form1.shpOK.BackColor = Red
'        StopReading = True
'        Form1.lblStatus.Caption = "Device reset failed in procedure OpenDevice."
'        Exit Sub
'    End If
'
'    ' set the module to MPSSE mode
'    FT_Result = Set_USB_Device_BitMode(&H0, &H2) ' set to MPSSE mode
'    If FT_Result <> FT_OK Then
'        Form1.shpOK.BackColor = Red
'        StopReading = True
'        Form1.lblStatus.Caption = "Set to MPSSE mode failed in procedure OpenDevice."
'        Exit Sub
'    End If
'
'    ' sync MPSSE mode
'    If Not (Sync_To_MPSSE) Then
'        Form1.shpOK.BackColor = Red
'        StopReading = True
'        Form1.lblStatus.Caption = "Unable to synchronise the MPSSE write/read cycle in procedure OpenDevice."
'        Exit Sub
'    End If
'
'    ' initialise the port
'    OutIndex = 0                                ' point to the start of output buffer
'    Saved_Port_Value = &H8                      ' set the initial state of the first 8 lines
'    ' set the low byte
'    AddToBuffer &H80                            ' Set data bits low byte command
'    AddToBuffer &H8                             ' set CS=high, DI=low, DO=low, SK=low
'    AddToBuffer &HB                             ' CS=output, DI=input, DO=output, SK=output
'    ' set the clock divisor
'    AddToBuffer &H86                            ' set clock divisor command to 1MHz
'    AddToBuffer &H5                             ' low byte
'    AddToBuffer &H0                             ' high byte
'    AddToBuffer &H85                            ' turn off loopback
'    SendBytes OutIndex                          ' send to command processor
'
'    ' check for a bad command being echoed back
'    Res = Get_USB_Device_QueueStatus
'    If FT_Q_Bytes > 0 Or Res <> 0 Then
'        Form1.shpOK.BackColor = Yellow
'        Form1.lblStatus.Caption = "Possible bad command detected in procedure OpenDevice."
'        Exit Sub
'    End If
'
    Form1.shpOK.BackColor = Green               ' set status to green
    Form1.lblStatus.Caption = "Opened"              ' set OK
    
End Sub

Public Function OpenPort(PortName As String) As Boolean
' to open the port named PortName
Dim Res As Long
Dim NoOfDevs As Long
Dim I As Long
Dim Name As String
Dim DualName As String

    PortAIsOpen = False                         ' init to port not open
    OpenPort = False                            ' init to failure to open port
    Name = ""                                   ' set name to null
    DualName = PortName                         ' set which port we want to open
    NoOfDevs = GetFTDeviceCount                 ' get the number of devices
    If FT_Result <> FT_OK Then Exit Function    ' exit if failure
    
    ' try to find the requested port
    For I = 0 To NoOfDevs - 1
       Name = GetFTDeviceDescription(I)         ' get the device desctiption
       If Name = DualName Then Exit For         ' exit if this is the one
    Next
    
    If Name <> DualName Then Exit Function      ' exit if not found
    
    Res = Open_USB_Device_By_Description(DualName) ' open the device by its description
    If FT_Result <> FT_OK Then Exit Function    ' exit if failure
    
    Res = Get_USB_Device_QueueStatus            ' perform a test function on the port
    If FT_Result <> FT_OK Then Exit Function    ' exit if failure
    PortAIsOpen = True                          ' flag port as open
    OpenPort = True                             ' return open OK

End Function

Public Function Read_USB_Device_Buffer(Read_Count As Long) As Long
' Reads Read_Count bytes or less from the USB device to the FT_In_Buffer
' The function returns the number of bytes actually received which may range from zero
' to the actual number of bytes requested, depending on how many have been received
' at the time of the request + the read timeout value.
Dim Read_Result As Long

    If Read_Count = 1 Then Read_Result = Read_Count
    
    FT_IO_Status = FT_Read(FT_HANDLE, FT_In_Buffer, Read_Count, Read_Result)
    If FT_IO_Status <> FT_OK Then
        FT_Error_Report "FT_Read", FT_IO_Status
    End If
    Read_USB_Device_Buffer = Read_Result
    
End Function

Public Sub SendBytes(NumberOfBytes As Long)
Dim I As Long

    I = Write_USB_Device_Buffer(NumberOfBytes)
    OutIndex = OutIndex - I
    
End Sub

Public Function Set_USB_Device_BitMode(ucMask As Byte, ucEnable As Byte) As Long

    Set_USB_Device_BitMode = FT_SetBitMode(FT_HANDLE, ucMask, ucEnable)
    
End Function

Public Function Set_USB_Device_LatencyTimer(ucLatency As Byte) As Long

    Set_USB_Device_LatencyTimer = FT_SetLatencyTimer(FT_HANDLE, ucLatency)
    
End Function

Public Sub SetDeviceString(S As String)
' set the device name

    FT_Device_String_Buffer = S & Chr(0)
    
End Sub

Public Function Sync_To_MPSSE() As Boolean
' uses &HAA and &HAB commands which are invalid so that the MPSSE processor should
' echo these back to use preceded with &HFA
Dim Res As Long
Dim I As Long
Dim J As Long

    Sync_To_MPSSE = False
    
    ' clear anything in the input buffer
    Res = Get_USB_Device_QueueStatus
    If Res <> FT_OK Then Exit Function
    If FT_Q_Bytes > 0 Then
        ' read chunks of 'input buffer size'
        Do While FT_Q_Bytes > FT_In_Buffer_Size
            I = Read_USB_Device_Buffer(FT_In_Buffer_Size) ' read a chunk
            FT_Q_Bytes = FT_Q_Bytes - FT_In_Buffer_Size ' calculate bytes left
        Loop
        I = Read_USB_Device_Buffer(FT_Q_Bytes) ' read the final bytes
    End If
    
    ' put a bad command to the command processor
    OutIndex = 0 ' point to start of buffer
    AddToBuffer &HAA ' add a bad command
    SendBytes OutIndex  ' send to command processor
    ' wait for a response
    Do
        Res = Get_USB_Device_QueueStatus
    Loop Until (FT_Q_Bytes > 0) Or (Res <> FT_OK)
    If Res <> FT_OK Then Exit Function
    
    ' read the input queue
    I = Read_USB_Device_Buffer(FT_Q_Bytes) ' read the bytes
    For J = 1 To I
        If Mid(FT_In_Buffer, J, 1) = Chr(&HAA) Then
            Sync_To_MPSSE = True
            Exit Function
        End If
    Next
        
End Function

Public Sub TakeReading()
' take a single read of the ADC
Dim BitTest As Byte
Dim Res As Long
Dim Byte0 As Byte
Dim Byte1 As Byte
Dim I As Long
Dim Reading0 As Integer
Dim Reading1 As Integer
Dim LoopLimit As Integer

    ' set CS low to initiate a conversion in the MAX187 ADC
    Saved_Port_Value = Saved_Port_Value And &HF7    ' set CS=low
    AddToBuffer &H80                                ' Set data bits low byte command
    AddToBuffer CLng(Saved_Port_Value)
    AddToBuffer &HB                                 ' CS=output, DI=input, DO=output, SK=output
    SendBytes OutIndex                              ' send to command processor
    
    ' check for bad command
    Res = Get_USB_Device_QueueStatus
    If FT_Q_Bytes > 0 Or Res <> 0 Then
        Form1.shpOK.BackColor = Yellow
        Form1.lblStatus.Caption = "Possible bad command detected in procedure TakeReading when initiating an ADC conversion."
    End If
    
    ' wait for DI to go high - raised by DO on the MAX187 to signal conversion complete
    LoopLimit = 0                                   ' clear the limit counter
    Do
        AddToBuffer &H81                            ' read data bits low byte
        AddToBuffer &H87                            ' send back results immediately
        SendBytes OutIndex                          ' send to command processor
        LoopLimit = LoopLimit + 1
        Do
            Res = Get_USB_Device_QueueStatus '
        Loop Until (FT_Q_Bytes > 0) Or (Res <> FT_OK) ' wait for answer to be available
        If Res <> FT_OK Then
            Form1.shpOK.BackColor = Red
            StopReading = True
            Form1.lblStatus.Caption = "Get USB device queue status failed in procedureTakeReading."
            Exit Sub
        End If
        ' read the input queue
        I = Read_USB_Device_Buffer(FT_Q_Bytes)      ' read the byte
        BitTest = CByte(Asc(Mid(FT_In_Buffer, 1, 1))) And &H4 ' check if conversion complete
    Loop Until BitTest <> &H0 Or LoopLimit > 100
    
    If LoopLimit > 100 Then
        Form1.shpOK.BackColor = Yellow
        StopReading = True
        Form1.lblStatus.Caption = "No reading received - please check the ADC power is turned on."
    Else
        Form1.shpOK.BackColor = Green
        Form1.lblStatus.Caption = "OK"
    End If
    
    ' Clock data in. 2 bytes on -ve clock MSB first, no write
    AddToBuffer &H24                                ' read bytes on -ve clock MSB
    AddToBuffer &H1                                 ' LSB value 2
    AddToBuffer &H0                                 ' MSB value 0
    AddToBuffer &H87                                ' do it now
    SendBytes OutIndex
    ' wait for data to become available
    Do
        Res = Get_USB_Device_QueueStatus '
    Loop Until (FT_Q_Bytes > 0) Or (Res <> FT_OK)   ' wait for answer to be available
    If Res <> FT_OK Then
        Form1.shpOK.BackColor = Red
        StopReading = True
        Form1.lblStatus.Caption = "Get USB device queue status failed while waiting to read an ADC conversion."
        Exit Sub
    End If
    ' read the input queue
    I = Read_USB_Device_Buffer(FT_Q_Bytes)          ' read the bytes
    ' the MAX187 sends 1 start bit followed by 7 data bits in the first byte, then the
    ' remaining 5 data bits in the second byte. We must join the 2 together...
    Byte0 = CByte(Asc(Mid(FT_In_Buffer, 1, 1)))     ' convert to byte format
    Byte1 = CByte(Asc(Mid(FT_In_Buffer, 2, 1)))     ' convert to byte format
    Byte0 = Byte0 And &H7F                          ' drop the start bit put there by the MAX187
    Reading0 = Reading0 Or Byte0                    ' convert the MSB to integer
    Reading0 = Reading0 * 32                        ' shift left 5 bits
    Reading1 = Reading1 Or Byte1                    ' convert the LSB to integer
    Reading1 = Reading1 \ 8                         ' shift right 3 bits
    Reading = Reading0 + Reading1                   ' add both together
    
    ' turn CS high
    Saved_Port_Value = Saved_Port_Value Or &H8      ' set CS=high
    AddToBuffer &H80                                ' Set data bits low byte command
    AddToBuffer CLng(Saved_Port_Value)
    AddToBuffer &HB                                 ' CS=output, DI=input, DO=output, SK=output
    SendBytes OutIndex                              ' send to command processor
    
    ' check got a reading
    If Reading = 0 Then
        Form1.shpOK.BackColor = Yellow
        StopReading = True
        Form1.lblStatus.Caption = "No reading received - please check the ADC power is turned on."
        Exit Sub
    Else
        Form1.shpOK.BackColor = Green
        Form1.lblStatus.Caption = "OK"
    End If

End Sub

Public Function Write_USB_Device_Buffer(Write_Count As Long) As Long
Dim Write_Result As Long

    FT_IO_Status = FT_Write(FT_HANDLE, FT_Out_Buffer, Write_Count, Write_Result)
    If FT_IO_Status <> FT_OK Then FT_Error_Report "FT-Write", FT_IO_Status
    Write_USB_Device_Buffer = Write_Result
    
End Function
