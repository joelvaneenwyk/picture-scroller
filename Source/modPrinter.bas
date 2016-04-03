Attribute VB_Name = "modPrinter"
Option Explicit

' Global constants for Win32 API
Private Const GMEM_ZEROINIT = &H40

Private Const DM_COPIES = &H100&
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
Private Const DM_COLOR = &H800&
Private Const DM_PAPERWIDTH = &H8&
Private Const DM_PAPERSIZE = &H2&
Private Const DM_PAPERLENGTH = &H4&
Private Const DM_PRINTQUALITY = &H400&

Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NOPAGENUMS = &H8
Public Const PD_NOSELECTION = &H4
Public Const PD_NOWARNING = &H80

'type definitions:
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long

    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'API declarations:
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Public Sub ShowPrinter(frmOwner As Form, Optional lPrintFlags As Long)

Dim tPrintDlg As PRINTDLG_TYPE
Dim tDevMode As DEVMODE_TYPE
Dim tDevName As DEVNAMES_TYPE

Dim lDevMode As Long, lDevName As Long
Dim iReturn As Integer
Dim objPrinter As Printer, sNewPrinterName As String
Dim sSetting As String

' Use PrintDialog to get the handle to a memory
' block with a tDevMode and tDevName structures
With tPrintDlg
    .lStructSize = Len(tPrintDlg)
    .hwndOwner = frmOwner.hwnd
    .flags = lPrintFlags
End With

'Set the current orientation and duplex setting
With tDevMode
    .dmSize = Len(tDevMode)
    .dmDeviceName = Printer.DeviceName
    .dmFields = DM_ORIENTATION Or DM_DUPLEX Or DM_PAPERSIZE Or DM_PAPERWIDTH Or DM_PAPERLENGTH Or DM_PRINTQUALITY Or DM_COLOR Or DM_COPIES
    .dmOrientation = Printer.Orientation
    .dmPaperSize = Printer.PaperSize
    .dmCopies = Printer.Copies
    If Printer.PaperSize <> 256 Then
        .dmPaperLength = Printer.Height
        .dmPaperWidth = Printer.Width
    End If
    .dmPrintQuality = Printer.PrintQuality
    .dmColor = Printer.ColorMode

    On Error Resume Next
    tDevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0
End With

'Allocate memory for the initialization hDevMode structure
'and copy the settings gathered above into this memory
tPrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(tDevMode))
lDevMode = GlobalLock(tPrintDlg.hDevMode)
If lDevMode > 0 Then
    CopyMemory ByVal lDevMode, tDevMode, Len(tDevMode)
    iReturn = GlobalUnlock(tPrintDlg.hDevMode)
End If

'Set the current driver, device, and port name strings
With tDevName
    .wDriverOffset = 8
    .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
    .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
    .wDefault = 0
End With

With Printer
    tDevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
End With

'Allocate memory for the initial hDevName structure
'and copy the settings gathered above into this memory
tPrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(tDevName))
lDevName = GlobalLock(tPrintDlg.hDevNames)
If lDevName > 0 Then
    CopyMemory ByVal lDevName, tDevName, Len(tDevName)
    iReturn = GlobalUnlock(lDevName)
End If

'Call the print dialog up and let the user make changes
If PrintDialog(tPrintDlg) Then
    'First get the tDevName structure.
    lDevName = GlobalLock(tPrintDlg.hDevNames)
    CopyMemory tDevName, ByVal lDevName, 45
    iReturn = GlobalUnlock(lDevName)
    GlobalFree tPrintDlg.hDevNames

    'Next get the tDevMode structure and set the printer
    'properties appropriately
    lDevMode = GlobalLock(tPrintDlg.hDevMode)
    CopyMemory tDevMode, ByVal lDevMode, Len(tDevMode)
    iReturn = GlobalUnlock(tPrintDlg.hDevMode)
    GlobalFree tPrintDlg.hDevMode
    sNewPrinterName = UCase$(Left(tDevMode.dmDeviceName, InStr(tDevMode.dmDeviceName, Chr$(0)) - 1))

    If Printer.DeviceName <> sNewPrinterName Then
        For Each objPrinter In Printers
           If UCase$(objPrinter.DeviceName) = sNewPrinterName Then
                Set Printer = objPrinter
           End If
        Next
    End If

    On Error Resume Next

    'Set printer object properties according to selections made
    'by user
    DoEvents
    With Printer
        .Copies = tDevMode.dmCopies
        .Duplex = tDevMode.dmDuplex
        .Orientation = tDevMode.dmOrientation
        .PaperSize = tDevMode.dmPaperSize
        If tDevMode.dmPaperSize = 0 Then
            .Width = tDevMode.dmPaperWidth
            .Height = tDevMode.dmPaperLength
        End If
        .PrintQuality = tDevMode.dmPrintQuality
        .ColorMode = tDevMode.dmColor
    End With
    On Error GoTo 0
End If

End Sub
