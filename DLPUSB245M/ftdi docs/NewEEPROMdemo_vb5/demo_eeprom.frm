VERSION 5.00
Begin VB.Form DEMO_EEPROM 
   Caption         =   "EEPROM FUNCTION DEMO (FTD2XX Ver. 1.03.20 or greater)"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton btnReadEEUA 
         Caption         =   "Read EEPROM-UA"
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CommandButton btnProgUA 
         Caption         =   "Program EEPROM-UA"
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CommandButton btnGetUASize 
         Caption         =   "Get EEPROM-UA size"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton btnWrite 
         Caption         =   "EEPROM &WRITE"
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   5400
         Width           =   1935
      End
      Begin VB.CommandButton btnREAD 
         Caption         =   "EEPROM &READ"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   4800
         Width           =   1935
      End
      Begin VB.ListBox LoggerList 
         Height          =   4350
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   " (EEPROM WRITE changes the DESCRIPTION - Field to ""EEPROM WRITTEN!"" and  then back to the original value)"
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   6120
         Width           =   3135
      End
   End
End
Attribute VB_Name = "DEMO_EEPROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FT_Open Lib "FTD2XX.DLL" (ByVal intDeviceNumber As Integer, ByRef lngHandle As Long) As Long
Private Declare Function FT_Close Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Private Declare Function FT_Read Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lpszBuffer As String, ByVal lngBufferSize As Long, ByRef lngBytesReturned As Long) As Long
Private Declare Function FT_Write Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lpszBuffer As String, ByVal lngBufferSize As Long, ByRef lngBytesWritten As Long) As Long
Private Declare Function FT_SetBaudRate Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lngBaudRate As Long) As Long
Private Declare Function FT_SetDataCharacteristics Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal byWordLength As Byte, ByVal byStopBits As Byte, ByVal byParity As Byte) As Long
Private Declare Function FT_SetFlowControl Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal intFlowControl As Integer, ByVal byXonChar As Byte, ByVal byXoffChar As Byte) As Long
Private Declare Function FT_ResetDevice Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Private Declare Function FT_SetDtr Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Private Declare Function FT_ClrDtr Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Private Declare Function FT_SetRts Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Private Declare Function FT_ClrRts Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Private Declare Function FT_GetModemStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngModemStatus As Long) As Long
Private Declare Function FT_Purge Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lngMask As Long) As Long
Private Declare Function FT_GetStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngRxBytes As Long, ByRef lngTxBytes As Long, ByRef lngEventsDWord As Long) As Long
Private Declare Function FT_GetQueueStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngRxBytes As Long) As Long
Private Declare Function FT_GetEventStatus Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lngEventsDWord As Long) As Long
Private Declare Function FT_SetChars Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal byEventChar As Byte, ByVal byEventCharEnabled As Byte, ByVal byErrorChar As Byte, ByVal byErrorCharEnabled As Byte) As Long
Private Declare Function FT_SetTimeouts Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal lngReadTimeout As Long, ByVal lngWriteTimeout As Long) As Long
Private Declare Function FT_SetBreakOn Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long
Private Declare Function FT_SetBreakOff Lib "FTD2XX.DLL" (ByVal lngHandle As Long) As Long

'==============================================================
'Declarations for the EEPROM-accessing functions in FTD2XX.dll:
'==============================================================
Private Declare Function FT_EE_Program Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lpData As FT_PROGRAM_DATA) As Long
Private Declare Function FT_EE_Read Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lpData As FT_PROGRAM_DATA) As Long
Private Declare Function FT_EE_UASize Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByRef lpdwSize As Long) As Long
Private Declare Function FT_EE_UAWrite Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal pucData As String, ByVal dwDataLen As Long) As Long
Private Declare Function FT_EE_UARead Lib "FTD2XX.DLL" (ByVal lngHandle As Long, ByVal pucData As String, ByVal dwDataLen As Long, ByRef lpdwBytesRead As Long) As Long

'*********************************************************************************
'*********************************************************************************
        'Visual Basic Supports getting the addresses of pointers,
        'However the functions to do so are undocumented. for more information
        'on how to get pointers to variables in Visual Basic, see
        'Microsoft Knowledge Base Article - Q199824
        '
        'VarPtr                 Returns the address of a variable.
        'VarPtrArray            Returns the address of an array.
        'StrPtr                 Returns the address of the UNICODE string buffer.
        'VarPtrStringArray      Returns the address of an array of strings.
        'ObjPtr                 Returne the address of an object.
        

'Not Needed because of build in function in Visual Basic
'Private Declare Function agGetAddressForObject& Lib "apigid32.dll" (object As Any)
'*********************************************************************************
'*********************************************************************************


'====================================================================
'Type definition as equivalent for C-structure "ft_program_data" used
'in FT_EE_READ and FT_EE_WRITE;
'ATTENTION! The variables "Manufacturer", "ManufacturerID",
'"Description" and "SerialNumber" are passed as POINTERS to
'locations of Bytearrays. Each Byte in these arrays will be
'filled with one character of the whole string.
'(See below, calls to "agGetAddressForObject")
'=====================================================================
Private Type FT_PROGRAM_DATA
    VendorId As Integer                 '0x0403
    ProductId As Integer                '0x6001
    Manufacturer As Long                '32 "FTDI"
    ManufacturerId As Long              '16 "FT"
    Description As Long                 '64 "USB HS Serial Converter"
    SerialNumber As Long                '16 "FT000001" if fixed, or NULL
    MaxPower As Integer                 ' // 0 < MaxPower <= 500
    PnP As Integer                      ' // 0 = disabled, 1 = enabled
    SelfPowered As Integer              ' // 0 = bus powered, 1 = self powered
    RemoteWakeup As Integer             ' // 0 = not capable, 1 = capable
    ' Rev4 extensions:
    Rev4 As Boolean                     ' // true if Rev4 chip, false otherwise
    IsoIn As Boolean                    ' // true if in endpoint is isochronous
    IsoOut As Boolean                   ' // true if out endpoint is isochronous
    PullDownEnable As Boolean           ' // true if pull down enabled
    SerNumEnable As Boolean             ' // true if serial number to be used
    USBVersionEnable As Boolean         ' // true if chip uses USBVersion
    USBVersion As Integer               ' // BCD (0x0200 => USB2)
End Type

' Return codes
Const FT_OK = 0
Const FT_INVALID_HANDLE = 1
Const FT_DEVICE_NOT_FOUND = 2
Const FT_DEVICE_NOT_OPENED = 3
Const FT_IO_ERROR = 4
Const FT_INSUFFICIENT_RESOURCES = 5
Const FT_INVALID_PARAMETER = 6
Const FT_INVALID_BAUD_RATE = 7

Const FT_DEVICE_NOT_OPENED_FOR_ERASE = 8
Const FT_DEVICE_NOT_OPENED_FOR_WRITE = 9
Const FT_FAILED_TO_WRITE_DEVICE = 10
Const FT_EEPROM_READ_FAILED = 11
Const FT_EEPROM_WRITE_FAILED = 12
Const FT_EEPROM_ERASE_FAILED = 13
Const FT_EEPROM_NOT_PRESENT = 14
Const FT_EEPROM_NOT_PROGRAMMED = 15
Const FT_INVALID_ARGS = 16
Const FT_OTHER_ERROR = 17


'Bytearrays as "string-containers":
Dim bManufacturer(32) As Byte
Dim bManufacturerID(16) As Byte
Dim bDescription(64) As Byte
Dim bSerialNumber(16) As Byte


Private Sub btnGetUASize_Click()
'****************************************************
'Get the available size of the user accessible EEPROM
'in bytes:
'****************************************************
Dim lngSize
Dim plngSize
Dim lngRetVal As Long
Dim lngHandle As Long

LoggerList.AddItem "------------------------------------"

' Open the device
If FT_Open(0, lngHandle) <> FT_OK Then
    LoggerList.AddItem "Open Failed"
    Exit Sub
End If

lngRetVal = FT_EE_UASize(lngHandle, lngSize)
If lngRetVal <> FT_OK Then
    LoggerList.AddItem "Read UASIZE Failed: code " & Str(lngRetVal)
Else
    LoggerList.AddItem "UASIZE = " & Str(lngSize) & " bytes"
End If

If FT_Close(lngHandle) <> FT_OK Then
    LoggerList.AddItem "Close Failed"
End If


End Sub

Private Sub btnProgUA_Click()
'*******************************************************
'DEMO of writing data into the UA-EEPROM area;
'ATTENTION! If you attempt to write more bytes than
'the number of free (usable) bytes, an error will occur!
'*******************************************************
Dim lngSize
Dim lngRetVal As Long
Dim lngHandle As Long
Dim strBytesToWrite As String * 16          'String containing the bytes to be written
Dim lngBytesToWrite As Long                 'number of bytes to be written

LoggerList.AddItem "------------------------------------"

' Open the device
If FT_Open(0, lngHandle) <> FT_OK Then
    LoggerList.AddItem "Open Failed"
    Exit Sub
End If

'Determine the bytes to be written and how many:
strBytesToWrite = "Hello, world"
lngBytesToWrite = 2

'Write bytes:
lngRetVal = FT_EE_UAWrite(lngHandle, strBytesToWrite, lngBytesToWrite)
If lngRetVal <> FT_OK Then
    LoggerList.AddItem "Write EEPROM-UA Failed: code " & Str(lngRetVal)
Else
    LoggerList.AddItem "BytesWritten = " & Str(lngBytesToWrite)
End If

If FT_Close(lngHandle) <> FT_OK Then
    LoggerList.AddItem "Close Failed"
End If

End Sub

Private Sub btnREAD_Click()
'***********************************************************
'Reads and displays the whole structure "EEData" from EEPROM
'Pay Attention to the way of handling the strings as bytearrays
'in this routine!
'APIGID32.DLL (by DESAWARE, www.desaware.com) must be located
'in your system-directory!
'***********************************************************
Dim lngHandle As Long
Dim lngRetVal As Long
Dim lngCount As Long
Dim EEData As FT_PROGRAM_DATA
'result strings:
Dim strManufacturer As String
Dim strManufacturerID As String
Dim strDescription As String
Dim strSerialNumber As String

LoggerList.AddItem "------------------------------------"

' Open the device
If FT_Open(0, lngHandle) <> FT_OK Then
    LoggerList.AddItem "Open Failed"
    Exit Sub
End If

'Prepare EEData structure: assign the addresses of the
'beginning of the bytearrays:
'(The FT_PROGRAM_DATA structure contains only POINTERS to
'the bytearrays!)

        'These lines were replaced with functions built into Visual Basic See Below.
        'EEData.Manufacturer = agGetAddressForObject(bManufacturer(0))
        'EEData.ManufacturerId = agGetAddressForObject(bManufacturerID(0))
        'EEData.Description = agGetAddressForObject(bDescription(0))
        'EEData.SerialNumber = agGetAddressForObject(bSerialNumber(0))

'These lines of code function the as the above lines do. However, they do
'use the internal function built into visual basic for returning a pointer
'to a variable.

EEData.Manufacturer = VarPtr(bManufacturer(0))      'Use undocumented function to return pointer
EEData.ManufacturerId = VarPtr(bManufacturerID(0))  'Use undocumented function to return pointer
EEData.Description = VarPtr(bDescription(0))        'Use undocumented function to return pointer
EEData.SerialNumber = VarPtr(bSerialNumber(0))      'Use undocumented function to return pointer



'Read EEPROM data:
lngRetVal = FT_EE_Read(lngHandle, EEData)
If RetVal <> FT_OK Then
    LoggerList.AddItem "EE_Read Failed: code " & Str(lngRetVal)
    Exit Sub
End If

'Convert resulting bytearrays to strings
'(NULL-characters at the end are cut off):
strManufacturer = StrConv(bManufacturer, vbUnicode)
strManufacturer = Left(strManufacturer, InStr(strManufacturer, Chr(0)) - 1)

strManufacturerID = StrConv(bManufacturerID, vbUnicode)
strManufacturerID = Left(strManufacturerID, InStr(strManufacturerID, Chr(0)) - 1)

strDescription = StrConv(bDescription, vbUnicode)
strDescription = Left(strDescription, InStr(strDescription, Chr(0)) - 1)

strSerialNumber = StrConv(bSerialNumber, vbUnicode)
strSerialNumber = Left(strSerialNumber, InStr(strSerialNumber, Chr(0)) - 1)

'Display results:
LoggerList.AddItem "Manufacturer  : '" & strManufacturer & "'"
LoggerList.AddItem "ManufacturerID: '" & strManufacturerID & "'"
LoggerList.AddItem "Description   : '" & strDescription & "'"
LoggerList.AddItem "Serialnumber  : '" & strSerialNumber & "'"
LoggerList.AddItem "VendorID      : '" & Format(Hex(EEData.VendorId), "0000") & "'"
LoggerList.AddItem "ProductID     : '" & Format(Hex(EEData.ProductId), "0000") & "'"
LoggerList.AddItem "Max Power     : '" & EEData.MaxPower & "'mA"
LoggerList.AddItem "Plug-and-Play : '" & EEData.PnP & "'"
LoggerList.AddItem "Self-Powered  : '" & EEData.SelfPowered & "'"
LoggerList.AddItem "IsoIn         : '" & EEData.IsoIn & "'"
LoggerList.AddItem "IsoOut        : '" & EEData.IsoOut & "'"
LoggerList.AddItem "PullDownEnable: '" & EEData.PullDownEnable & "'"
LoggerList.AddItem "SerNumEnable  : '" & EEData.SerNumEnable & "'"
LoggerList.AddItem "USBVersion    : '" & EEData.USBVersion & "'"
LoggerList.AddItem "USBVersionEnable:'" & EEData.USBVersionEnable & "'"
LoggerList.AddItem "Rev4          : '" & EEData.Rev4 & "'"

'Close the device:
If FT_Close(lngHandle) <> FT_OK Then
    LoggerList.AddItem "Close Failed"
End If


End Sub

Private Sub btnReadEEUA_Click()
'*******************************************************
'Reads assigned number of bytes from the UA-EEPROM-area.
'*******************************************************
Dim lngSize
Dim lngRetVal As Long
Dim lngHandle As Long
Dim bBytesRead As String * 16
Dim lngBytesToRead As Long
Dim lngBytesRead As Long
Dim pBytesRead As Long
Dim lngN As Long

LoggerList.AddItem "------------------------------------"

' Open the device
If FT_Open(0, lngHandle) <> FT_OK Then
    LoggerList.AddItem "Open Failed"
    Exit Sub
End If

lngBytesToRead = 16

lngRetVal = FT_EE_UARead(lngHandle, bBytesRead, lngBytesToRead, lngBytesRead)
If lngRetVal <> FT_OK Then
    LoggerList.AddItem "Read EEPROM-UA Failed: code " & Str(lngRetVal)
Else
    LoggerList.AddItem "BytesRead = " & Str(lngBytesRead) & ":"
End If

'Display the result:
LoggerList.AddItem "'" & bBytesRead & "'"

'Close device:
If FT_Close(lngHandle) <> FT_OK Then
    LoggerList.AddItem "Close Failed"
End If

End Sub

Private Sub btnWrite_Click()
'**********************************************************
'DEMO: changes the device DESCRIPTION to "EEPROM WRITTEN!",
'reads it back and then restores the original settings.
'(Pay attention to the handling of strings as bytearrays and
'the use of pointers in this routine!)
'**********************************************************
Dim lngHandle As Long
Dim lngRetVal As Long
Dim lngCount As Long
Dim EEData As FT_PROGRAM_DATA
'result strings:
Dim strManufacturer As String
Dim strManufacturerID As String
Dim strDescription As String
Dim strSerialNumber As String

Dim bOLDDescription(64) As Byte

LoggerList.AddItem "------------------------------------"

'First Part: READ actual EEPROM-Settings:
'========================================
' Open the device
If FT_Open(0, lngHandle) <> FT_OK Then
    LoggerList.AddItem "Open Failed"
    Exit Sub
End If

'Prepare EEData structure: assign the addresses of the
'beginning of the bytearrays:
       
        'These lines were replaced with functions built into Visual Basic See Below.
        'EEData.Manufacturer = agGetAddressForObject(bManufacturer(0))
        'EEData.ManufacturerId = agGetAddressForObject(bManufacturerID(0))
        'EEData.Description = agGetAddressForObject(bDescription(0))
        'EEData.SerialNumber = agGetAddressForObject(bSerialNumber(0))


'These lines of code function the as the above lines do. However, they do
'use the internal function built into visual basic for returning a pointer
'to a variable.
EEData.Manufacturer = VarPtr(bManufacturer(0))      'Use undocumented function to return pointer
EEData.ManufacturerId = VarPtr(bManufacturerID(0))  'Use undocumented function to return pointer
EEData.Description = VarPtr(bDescription(0))        'Use undocumented function to return pointer
EEData.SerialNumber = VarPtr(bSerialNumber(0))      'Use undocumented function to return pointer







'Read EEPROM data:
lngRetVal = FT_EE_Read(lngHandle, EEData)
If RetVal <> FT_OK Then
    LoggerList.AddItem "EE_Read Failed: " & Str(lngRetVal)
    Exit Sub
End If

'store OLD description:
CopyByteArray bDescription, bOLDDescription

'Convert new description to ByteArray:
StringToByteArray "EEPROM WRITTEN!", bDescription

'Now write the complete set of EEPROM data
'(pointers are already set above before the read instruction...):
lngRetVal = FT_EE_Program(lngHandle, EEData)
If lngRetVal <> FT_OK Then
    LoggerList.AddItem "EE_Program FAILED: code " & Str(lngRetVal)
    Exit Sub
End If

LoggerList.AddItem "EEPROM successfully written!"

'Intermediately close the device:
If FT_Close(lngHandle) <> FT_OK Then
    LoggerList.AddItem "Close Failed"
End If

LoggerList.AddItem "------------------------------------"

LoggerList.AddItem "Read back starting...:"

'Read back actual values in EEPROM:
btnREAD_Click

'Finally, write back the original value of the description:
CopyByteArray bOLDDescription, bDescription

LoggerList.AddItem "------------------------------------"

'Re-Open the device
If FT_Open(0, lngHandle) <> FT_OK Then
    LoggerList.AddItem "Open Failed"
    Exit Sub
End If

'Restore original setting:
lngRetVal = FT_EE_Program(lngHandle, EEData)
If lngRetVal <> FT_OK Then
    LoggerList.AddItem "EE_Program (writing back original values) FAILED: code " & Str(lngRetVal)
    Exit Sub
End If

LoggerList.AddItem "Original Values written back!"

'Close device:
If FT_Close(lngHandle) <> FT_OK Then
    LoggerList.AddItem "Close Failed"
End If

End Sub

'=================================================
'TWO FUNCTIONS FOR THE HANDLING OF THE BYTEARRAYS:
'=================================================

Private Sub CopyByteArray(bArray, bCopy)
'***********************************************
'bArray and bCopy must be exactly the same SIZE!
'***********************************************
Dim lngN As Long
For lngN = 0 To UBound(bArray)
    bCopy(lngN) = bArray(lngN)
Next
End Sub

Private Sub StringToByteArray(strString, bByteArray)
Dim lngN As Long
'Fill bByteArray with "0":
For lngN = 0 To UBound(bByteArray)
    bByteArray(lngN) = 0
Next
For lngN = 1 To Len(strString)
    bByteArray(lngN - 1) = Asc(Mid(strString, lngN, 1))
Next
End Sub
