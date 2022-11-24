VERSION 5.00
Begin VB.Form DEMO_EEPROM 
   Caption         =   "EEPROM FUNCTION DEMO (FTD2XX Ver. 1.03.20 or greater)"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   Icon            =   "DEMO_EEPROM.frx":0000
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
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   495
         Left            =   4920
         TabIndex        =   10
         Top             =   6480
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   3960
         TabIndex        =   9
         Top             =   6480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   2880
         TabIndex        =   8
         Top             =   6480
         Width           =   855
      End
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
         ItemData        =   "DEMO_EEPROM.frx":08CA
         Left            =   240
         List            =   "DEMO_EEPROM.frx":08CC
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



'Bytearrays as "string-containers":
Dim bManufacturer(32) As Byte
Dim bManufacturerID(16) As Byte
Dim bDescription(64) As Byte
Dim bSerialNumber(16) As Byte
Dim Hndl As Long






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
lngBytesToWrite = 12

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
Private Sub Command1_Click()

' Open the device
If FT_Open(0, Hndl) <> FT_OK Then
    Debug.Print "Not Opened!"
    Exit Sub
Else
    Debug.Print "Opened!"
End If

End Sub

Private Sub Command2_Click()
' Open the device
If FT_Close(Hndl) <> FT_OK Then
    Debug.Print "Not Closed!"
    Exit Sub
Else
    Debug.Print "Closed!"
End If
End Sub


Private Sub Command3_Click()
Dim wrt As Long
If FT_Write(Hndl, "SAJJAD HUSSAIN", 14, wrt) <> FT_OK Then
    Debug.Print "Not Written!"
    Exit Sub
Else
    Debug.Print "Written!"
    Debug.Print wrt
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

EEData.signature1 = &H0
EEData.signature2 = &HFFFFFFFF
EEData.version = 0

EEData.Manufacturer = agGetAddressForObject(bManufacturer(0))
EEData.ManufacturerId = agGetAddressForObject(bManufacturerID(0))
EEData.Description = agGetAddressForObject(bDescription(0))
EEData.SerialNumber = agGetAddressForObject(bSerialNumber(0))

'Read EEPROM data:
lngRetVal = FT_EE_Read(lngHandle, EEData)
If lngRetVal <> FT_OK Then
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

EEData.signature1 = &H0
EEData.signature2 = &HFFFFFFFF
EEData.version = 0

EEData.Manufacturer = agGetAddressForObject(bManufacturer(0))
EEData.ManufacturerId = agGetAddressForObject(bManufacturerID(0))
EEData.Description = agGetAddressForObject(bDescription(0))
EEData.SerialNumber = agGetAddressForObject(bSerialNumber(0))

'Read EEPROM data:
lngRetVal = FT_EE_Read(lngHandle, EEData)
If lngRetVal <> FT_OK Then
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

