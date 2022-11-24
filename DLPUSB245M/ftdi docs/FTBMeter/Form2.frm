VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calibration"
   ClientHeight    =   2025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStarterValues 
      Caption         =   "Reset to approximate starter values"
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   1650
      Width           =   5685
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calibration"
      Height          =   1485
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   5985
      Begin VB.CommandButton cmdCal0dB 
         Caption         =   "Calibrate 0dB"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   570
         Width           =   1725
      End
      Begin VB.CommandButton cmdCalMinus40dB 
         Caption         =   "Calibrate -40dB"
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Top             =   1020
         Width           =   1725
      End
      Begin VB.CommandButton cmdCal0dBVHF 
         Caption         =   "Calibrate 0dB"
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   570
         Width           =   1725
      End
      Begin VB.CommandButton cmdCalMinus40dBVHF 
         Caption         =   "Calibrate -40dB"
         Height          =   315
         Left            =   2190
         TabIndex        =   5
         Top             =   1020
         Width           =   1725
      End
      Begin VB.CommandButton cmdCal0dBUHF 
         Caption         =   "Calibrate 0dB"
         Height          =   315
         Left            =   4140
         TabIndex        =   4
         Top             =   570
         Width           =   1725
      End
      Begin VB.CommandButton cmdCalMinus40dBUHF 
         Caption         =   "Calibrate -40dB"
         Height          =   315
         Left            =   4140
         TabIndex        =   3
         Top             =   1020
         Width           =   1725
      End
      Begin VB.Label Label4 
         Caption         =   "14MHz"
         Height          =   255
         Left            =   780
         TabIndex        =   11
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "144MHz"
         Height          =   255
         Left            =   2670
         TabIndex        =   10
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "440MHz"
         Height          =   255
         Left            =   4710
         TabIndex        =   9
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6210
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6210
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblReading 
      Height          =   315
      Left            =   6540
      TabIndex        =   13
      Top             =   1680
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Reading"
      Height          =   285
      Left            =   6450
      TabIndex        =   12
      Top             =   1320
      Width           =   915
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
' close form

    ' re-load original values
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
    
    ' close form
    Unload Me
    
End Sub

Private Sub cmdCal0dB_Click()
' calibrate this setting
Dim I As Long
Dim AverageReading As Single

    If Not (PortAIsOpen) Then Exit Sub
    
    StopReading = True
    DoEvents
    
    For I = 1 To 8
        TakeReading
        AverageReading = AverageReading + Reading
    Next
    
    ZerodBmHF = AverageReading / 8
    lblReading.Caption = Format(AverageReading / 8, "###0.0")

End Sub

Private Sub cmdCal0dBUHF_Click()
' calibrate this setting
Dim I As Long
Dim AverageReading As Single

    If Not (PortAIsOpen) Then Exit Sub
    
    StopReading = True
    DoEvents
    
    For I = 1 To 8
        TakeReading
        AverageReading = AverageReading + Reading
    Next
    
    ZerodBmUHF = AverageReading / 8
    lblReading.Caption = Format(AverageReading / 8, "###0.0")

End Sub

Private Sub cmdCal0dBVHF_Click()
' calibrate this setting
Dim I As Long
Dim AverageReading As Single

    If Not (PortAIsOpen) Then Exit Sub
    
    StopReading = True
    DoEvents
    
    For I = 1 To 8
        TakeReading
        AverageReading = AverageReading + Reading
    Next
    
    ZerodBmVHF = AverageReading / 8
    lblReading.Caption = Format(AverageReading / 8, "###0.0")

End Sub

Private Sub cmdCalMinus40dB_Click()
' calibrate this setting
Dim I As Long
Dim AverageReading As Single

    If Not (PortAIsOpen) Then Exit Sub
    
    StopReading = True
    DoEvents
    
    For I = 1 To 8
        TakeReading
        AverageReading = AverageReading + Reading
    Next
    
    Minus40dBmHF = AverageReading / 8
    lblReading.Caption = Format(AverageReading / 8, "###0.0")

End Sub

Private Sub cmdCalMinus40dBUHF_Click()
' calibrate this setting
Dim I As Long
Dim AverageReading As Single

    If Not (PortAIsOpen) Then Exit Sub
    
    StopReading = True
    DoEvents
    
    For I = 1 To 8
        TakeReading
        AverageReading = AverageReading + Reading
    Next
    
    Minus40dBmUHF = AverageReading / 8
    lblReading.Caption = Format(AverageReading / 8, "###0.0")

End Sub

Private Sub cmdCalMinus40dBVHF_Click()
' calibrate this setting
Dim I As Long
Dim AverageReading As Single

    If Not (PortAIsOpen) Then Exit Sub
    
    StopReading = True
    DoEvents
    
    For I = 1 To 8
        TakeReading
        AverageReading = AverageReading + Reading
    Next
    
    Minus40dBmVHF = AverageReading / 8
    lblReading.Caption = Format(AverageReading / 8, "###0.0")

End Sub

Private Sub cmdStarterValues_Click()
' reset to approximate starter values measured when the meter was originally built

    ZerodBmHF = 2556
    ZerodBmVHF = 2519
    ZerodBmUHF = 2501
    Minus40dBmHF = 915
    Minus40dBmVHF = 913
    Minus40dBmUHF = 872

End Sub

Private Sub OKButton_Click()
' save new values

    SaveSetting RegKey, "Settings", "ZerodBmHF", ZerodBmHF
    SaveSetting RegKey, "Settings", "ZerodBmVHF", ZerodBmVHF
    SaveSetting RegKey, "Settings", "ZerodBmUHF", ZerodBmUHF
    SaveSetting RegKey, "Settings", "Minus40dBmHF", Minus40dBmHF
    SaveSetting RegKey, "Settings", "Minus40dBmVHF", Minus40dBmVHF
    SaveSetting RegKey, "Settings", "Minus40dBmUHF", Minus40dBmUHF
    ZerodBm = ZerodBmVHF
    Minus40dBm = Minus40dBmVHF
    Slope = (ZerodBm - Minus40dBm) / 40
    Form1.cmdHF.BackColor = ButtonFace
    Form1.cmdVHF.BackColor = Green
    Form1.cmdUHF.BackColor = ButtonFace

    Unload Me
    
End Sub
