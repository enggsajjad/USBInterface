VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTBMeter"
   ClientHeight    =   6165
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   4200
      TabIndex        =   20
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   1800
      TabIndex        =   19
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Height          =   2655
      Left            =   60
      TabIndex        =   5
      Top             =   780
      Width           =   5985
      Begin VB.CommandButton cmdReadSingle 
         Caption         =   "Single Reading"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   2475
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Re-open the module"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   2130
         Width           =   2475
      End
      Begin VB.CommandButton cmdAverage4 
         Caption         =   "Average 4 Readings"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2475
      End
      Begin VB.CommandButton cmdContinuous 
         Caption         =   "Start Continuous Reading"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1350
         Width           =   2475
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop Continuous Reading"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1740
         Width           =   2475
      End
      Begin VB.Label lblDBM 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3060
         TabIndex        =   17
         Top             =   1830
         Width           =   2025
      End
      Begin VB.Label lblReading 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   225
         Left            =   3210
         TabIndex        =   16
         Top             =   630
         Width           =   675
      End
      Begin VB.Label lblAverage4 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   255
         Left            =   3210
         TabIndex        =   15
         Top             =   990
         Width           =   675
      End
      Begin VB.Label lblContinuous 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   255
         Left            =   3210
         TabIndex        =   14
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label lblPerSecond 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   285
         Left            =   4170
         TabIndex        =   13
         Top             =   1380
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Reading"
         Height          =   285
         Left            =   3330
         TabIndex        =   12
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Per sec"
         Height          =   345
         Left            =   4830
         TabIndex        =   11
         Top             =   1380
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frequency range"
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Top             =   3510
      Width           =   5985
      Begin VB.CommandButton cmdHF 
         Caption         =   "Measure HF"
         Height          =   315
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   300
         Width           =   1725
      End
      Begin VB.CommandButton cmdVHF 
         Caption         =   "Measure VHF"
         Height          =   315
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   300
         Width           =   1725
      End
      Begin VB.CommandButton cmdUHF 
         Caption         =   "Measure UHF"
         Height          =   315
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   1725
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3780
      Top             =   3840
   End
   Begin VB.Label lblStatus 
      Caption         =   "DLP2232M module status"
      Height          =   495
      Left            =   660
      TabIndex        =   0
      Top             =   270
      Width           =   5355
      WordWrap        =   -1  'True
   End
   Begin VB.Shape shpOK 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   285
      Left            =   180
      Shape           =   3  'Circle
      Top             =   210
      Width           =   285
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalibrate 
         Caption         =   "&Calibrate"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FTBMeter - a program to provide a power meter display. It uses an FTDI FT2232 chip on a DLP
' Design USB module DLP2232M to talk to a MAX187 A/D converter which reads the DC o/p voltage
' of an Analog Devices AD8307 logarithmic amplifier. The MAX187 has an external voltage reference
' of 2.5V applied so that the full range of the chip is used when reading the AD8307 (which has
' an output voltage range from 0.4V to 2.5V.

' This code draws on some sample code available on the FTDI web site that is written
' in Delphi. The Delphi project was originally downloaded from the FTDI web site here:
'   http://www.ftdichip.com/Projects/MPSSE.htm#SPI
' A number of the procedures and functions in this program are translations of similarly named
' procedures in the Delphi code.

' This program was developed in VB6 and uses the FTDI DLLs which became available in June 2006,
' the combined driver model versions (Microsoft WHQL certified).

' The program uses a DLP2232M USB module to provide USB communications and onwards to a MAX187
' analog to digital converter using the SPI protocol. The DLP module DLP-2232M-G was acquired
' through the FTDI Web-Shop but is also available from DLP here:
'  http://www.dlpdesign.com/usb/2232m.shtml

' If you customise the name of the A channel then change the variable OurDevice in
' procedure InitialiseVariables.

' G R Freeth (G4HFQ) July 2006
' http://www.g4hfq.co.uk

Option Explicit

Private Sub cmdAverage4_Click()
' average 4 readings
Dim I As Long
Dim AverageReading As Single

    If Not (PortAIsOpen) Then Exit Sub
    
    StopReading = True
    DoEvents
    
    For I = 1 To 4
        TakeReading
        AverageReading = AverageReading + Reading
    Next
    
    lblAverage4.Caption = Format(AverageReading / 4, "###0.0")
    FormatReading AverageReading / 4
    
End Sub

Private Sub cmdContinuous_Click()
' continuous readings
Static Reading1 As Single
Static Reading2 As Single
Static Reading3 As Single
Static Reading4 As Single
Dim AverageReading As Single

    If Not (PortAIsOpen) Then Exit Sub
    
    StopReading = False
    Timer1.Enabled = True                           ' start the one second timer
    
    Do
        TakeReading
        Reading4 = Reading3
        Reading3 = Reading2
        Reading2 = Reading1
        Reading1 = Reading
        AverageReading = (Reading4 + Reading3 + Reading2 + Reading1) / 4
        lblContinuous.Caption = Format(AverageReading, "###0.0")
        FormatReading AverageReading
        NumberOfReadings = NumberOfReadings + 1
        DoEvents
    Loop Until StopReading

End Sub

Private Sub cmdHF_Click()
' set up for HF values

    ZerodBm = ZerodBmHF
    Minus40dBm = Minus40dBmHF
    Slope = (ZerodBm - Minus40dBm) / 40
    Form1.cmdHF.BackColor = Green
    Form1.cmdVHF.BackColor = ButtonFace
    Form1.cmdUHF.BackColor = ButtonFace

End Sub

Private Sub cmdOpen_Click()
' open the DLP2232M module

    OpenDevice

End Sub

Private Sub cmdReadSingle_Click()
' take a single readig

    If Not (PortAIsOpen) Then Exit Sub
    StopReading = True
    DoEvents
    TakeReading
    lblReading = Format(Reading, "###0.0")
    FormatReading Reading
    
End Sub

Private Sub cmdStop_Click()
' stop continuous readings

    StopReading = True                              ' say we want to stop
    Timer1.Enabled = False                          ' stop the timer
    
End Sub

Private Sub cmdUHF_Click()
' set up for UHF

    ZerodBm = ZerodBmUHF
    Minus40dBm = Minus40dBmUHF
    Slope = (ZerodBm - Minus40dBm) / 40
    Form1.cmdHF.BackColor = ButtonFace
    Form1.cmdVHF.BackColor = ButtonFace
    Form1.cmdUHF.BackColor = Green

End Sub

Private Sub cmdVHF_Click()
' set up for VHF

    ZerodBm = ZerodBmVHF
    Minus40dBm = Minus40dBmVHF
    Slope = (ZerodBm - Minus40dBm) / 40
    Form1.cmdHF.BackColor = ButtonFace
    Form1.cmdVHF.BackColor = Green
    Form1.cmdUHF.BackColor = ButtonFace

End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command1_Click()
OpenDevice
End Sub

Private Sub Command2_Click()
Dim Res As Long
'If PortAIsOpen Then
        Res = Close_USB_Device
        If FT_Result <> FT_OK Then
            PortAIsOpen = False
            Form1.shpOK.BackColor = Red
            Form1.lblStatus.Caption = "Attempt to close DLP2232M failed."
            StopReading = True
            Exit Sub
        End If
    'End If
Form1.lblStatus.Caption = "Closed!"
End Sub


Private Sub Command3_Click()
FT_Out_Buffer = "SAJJAD"
Write_USB_Device_Buffer (6)
End Sub

Private Sub Form_Load()
' initialise the program

'    InitialiseVariables                             ' init variables
'
'    OpenDevice                                      ' open the DLP2232M
'
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'' unload form tidy up
'Dim Res As Long
'
'    If PortAIsOpen Then
'        Res = Close_USB_Device
'        If FT_Result <> FT_OK Then
'            PortAIsOpen = False
'            Form1.shpOK.BackColor = Red
'            StopReading = True
'            Form1.lblStatus.Caption = "Close device failed in procedure Form Unload."
'            Exit Sub
'        End If
'    End If
'
'End Sub

Private Sub mnuCalibrate_Click()
' show the calibrate form

    Form2.Show vbModal
    
End Sub

Private Sub mnuExit_Click()
' end program

    Unload Me
    
End Sub

Private Sub Timer1_Timer()
' show readings per second

    lblPerSecond.Caption = CStr(NumberOfReadings)   ' show how many
    NumberOfReadings = 0                            ' reset count

End Sub
