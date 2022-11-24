VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FTD2XX Loopback Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.CommandButton TestBtn 
         Caption         =   "&Test"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ListBox LoggerList 
         Height          =   1425
         ItemData        =   "Form1.frx":08CA
         Left            =   120
         List            =   "Form1.frx":08CC
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FT_Open Lib "FTD2XX.DLL" (ByVal intDeviceNumber As Integer, ByRef lngHandle As Long) As Long
Private Declare Function FT_OpenEx Lib "FTD2XX.DLL" (ByVal arg1 As String, ByVal arg2 As Long, ByRef lngHandle As Long) As Long
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
Private Declare Function FT_ListDevices Lib "FTD2XX.DLL" (ByVal arg1 As Long, ByVal arg2 As String, ByVal dwFlags As Long) As Long
Private Declare Function FT_GetNumDevices Lib "FTD2XX.DLL" Alias "FT_ListDevices" (ByRef arg1 As Long, ByVal arg2 As String, ByVal dwFlags As Long) As Long

' Return codes
Const FT_OK = 0
Const FT_INVALID_HANDLE = 1
Const FT_DEVICE_NOT_FOUND = 2
Const FT_DEVICE_NOT_OPENED = 3
Const FT_IO_ERROR = 4
Const FT_INSUFFICIENT_RESOURCES = 5
Const FT_INVALID_PARAMETER = 6
Const FT_INVALID_BAUD_RATE = 7


' Word Lengths
Const FT_BITS_8 = 8
Const FT_BITS_7 = 7

' Stop Bits
Const FT_STOP_BITS_1 = 0
Const FT_STOP_BITS_1_5 = 1
Const FT_STOP_BITS_2 = 2

' Parity
Const FT_PARITY_NONE = 0
Const FT_PARITY_ODD = 1
Const FT_PARITY_EVEN = 2
Const FT_PARITY_MARK = 3
Const FT_PARITY_SPACE = 4

' Flow Control
Const FT_FLOW_NONE = &H0
Const FT_FLOW_RTS_CTS = &H100
Const FT_FLOW_DTR_DSR = &H200
Const FT_FLOW_XON_XOFF = &H400

' Purge rx and tx buffers
Const FT_PURGE_RX = 1
Const FT_PURGE_TX = 2

' Flags for FT_OpenEx
Const FT_OPEN_BY_SERIAL_NUMBER = 1
Const FT_OPEN_BY_DESCRIPTION = 2

' Flags for FT_ListDevices
Const FT_LIST_BY_NUMBER_ONLY = &H80000000
Const FT_LIST_BY_INDEX = &H40000000
Const FT_LIST_ALL = &H20000000


Private Sub TestBtn_Click()

Dim lngHandle As Long
Dim strWriteBuffer As String * 256
Dim lngBytesWritten As Long
Dim strReadBuffer As String * 256
Dim lngBytesRead As Long
Dim lngTotalBytesRead As Long
Dim strLoggerBuffer As String
Dim flFailed As Boolean
Dim flTimedout As Boolean
Dim flFatalError As Boolean
Dim ftStatus As Long
Dim lngNumDevices As Long
Dim strSerialNumber As String * 256
Dim strDescription As String * 256
Dim strSerialNumber2 As String * 256
Dim strDescription2 As String * 256


Const HELLO_WORLD_LENGTH = 12   ' length of "hello, world"


flFailed = True

' Example assumes that 2 devices are connected

' Get device serial number

If FT_GetNumDevices(lngNumDevices, vbNullString, FT_LIST_BY_NUMBER_ONLY) <> FT_OK Then
    LoggerList.AddItem ("FT_GetNumDevices failed")
    Exit Sub
Else
    LoggerList.AddItem ("NumDevices " & lngNumDevices)
End If

If FT_ListDevices(0, strDescription, FT_LIST_BY_INDEX Or FT_OPEN_BY_DESCRIPTION) <> FT_OK Then
    LoggerList.AddItem ("ListDevices failed")
    Exit Sub
Else
    LoggerList.AddItem ("Device Description " & strDescription)
End If

If FT_ListDevices(0, strSerialNumber, FT_LIST_BY_INDEX Or FT_OPEN_BY_SERIAL_NUMBER) <> FT_OK Then
    LoggerList.AddItem ("ListDevices failed")
    Exit Sub
Else
    LoggerList.AddItem ("Serial Number " & strSerialNumber)
End If

' Get serial number and description of device 2

If FT_ListDevices(1, strDescription2, FT_LIST_BY_INDEX Or FT_OPEN_BY_DESCRIPTION) <> FT_OK Then
    LoggerList.AddItem ("ListDevices failed")
'    Exit Sub
Else
    LoggerList.AddItem ("Device Description " & strDescription2)
End If

If FT_ListDevices(1, strSerialNumber2, FT_LIST_BY_INDEX Or FT_OPEN_BY_SERIAL_NUMBER) <> FT_OK Then
    LoggerList.AddItem ("ListDevices failed")
'    Exit Sub
Else
    LoggerList.AddItem ("Serial Number " & strSerialNumber2)
End If

' Open the device

If FT_OpenEx(strDescription, FT_OPEN_BY_DESCRIPTION, lngHandle) <> FT_OK Then
    LoggerList.AddItem "Open Failed"
    Exit Sub
End If

' Set baud rate to 115200

ftStatus = FT_SetBaudRate(lngHandle, 115200)
If ftStatus <> FT_OK Then
    LoggerList.AddItem "SetBaudRate Failed"
    GoTo CloseHandle
End If

' 8 data bits, 1 stop bit, no parity

If FT_SetDataCharacteristics(lngHandle, FT_BITS_8, FT_STOP_BITS_1, FT_PARITY_NONE) <> FT_OK Then
    LoggerList.AddItem "SetDataCharacteristics Failed"
    GoTo CloseHandle
End If

' no flow control

If FT_SetFlowControl(lngHandle, FT_FLOW_NONE, 0, 0) <> FT_OK Then
    LoggerList.AddItem "SetFlowControl Failed"
    GoTo CloseHandle
End If

' 5 second read timeout

If FT_SetTimeouts(lngHandle, 5000, 0) <> FT_OK Then
    LoggerList.AddItem "SetFlowControl Failed"
    GoTo CloseHandle
End If

If FT_SetChars(lngHandle, 126, 1, 0, 0) <> FT_OK Then
    LoggerList.AddItem "SetFlowControl Failed"
    GoTo CloseHandle
End If

' write the string "hello, world"

strWriteBuffer = "Hello, world"
lngBytesWritten = 0
If FT_Write(lngHandle, strWriteBuffer, HELLO_WORLD_LENGTH, lngBytesWritten) <> FT_OK Then
    LoggerList.AddItem "Write Failed"
    GoTo CloseHandle
End If

' Loop until until the string has been read back.
' Note that FT_Read returns FT_IO_ERROR on timeout.
' We break out of the loop when we get a timeout and
' nothing has been returned from FT_Read.

flTimedout = False
flFatalError = False

lngTotalBytesRead = 0
Do
    lngBytesRead = 0
    ftStatus = FT_Read(lngHandle, strReadBuffer, HELLO_WORLD_LENGTH, lngBytesRead)
    If (ftStatus = FT_OK) Or (ftStatus = FT_IO_ERROR) Then
        If lngBytesRead > 0 Then
            strLoggerBuffer = strLoggerBuffer + Left(strReadBuffer, lngBytesRead)
            lngTotalBytesRead = lngTotalBytesRead + lngBytesRead
        Else
            flTimedout = True
        End If
    Else
        flFatalError = True
    End If
Loop Until (lngTotalBytesRead = HELLO_WORLD_LENGTH) Or (flTimedout = True) Or (flFatalError = True)

If (flTimedout = False) And (flFatalError = False) Then
    LoggerList.AddItem strLoggerBuffer
    flFailed = False
ElseIf flTimedout = True Then
    LoggerList.AddItem "FT_Read timeout ftStatus=" & ftStatus
Else
    LoggerList.AddItem "FT_Read error ftStatus=" & ftStatus
End If

CloseHandle:

' close the device

If FT_Close(lngHandle) <> FT_OK Then
    LoggerList.AddItem "Close Failed"
End If

If flFailed = True Then
    LoggerList.AddItem "Test Failed"
End If

End Sub
