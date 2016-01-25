Attribute VB_Name = "MWinAPIGDI32"
' ==========================================================================
' Module      : MWinAPIGDI32
' Type        : Module
' Description : Support for Graphics Device Interface (GDI) functions
' --------------------------------------------------------------------------
' Procedures  : GetDPI                      Long
'               PixelsToPoints              Single
'               PointsPerPixel              Single
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

' ----------------
' Module Level
' ----------------

Private Const msMODULE          As String = "MWinAPIGDI32"

Private Const mlPOINTS_PER_INCH As Long = 72

' Device Parameters for GetDeviceCaps()
' -------------------------------------
Private Const DRIVERVERSION     As Long = 0     ' Device driver version
Private Const TECHNOLOGY        As Long = 2     ' Device classification
Private Const HORZSIZE          As Long = 4     ' Horizontal size in millimeters
Private Const VERTSIZE          As Long = 6     ' Vertical size in millimeters
Private Const HORZRES           As Long = 8     ' Horizontal width in pixels
Private Const VERTRES           As Long = 10    ' Vertical height in pixels
Private Const BITSPIXEL         As Long = 12    ' Number of bits per pixel
Private Const PLANES            As Long = 14    ' Number of planes
Private Const NUMBRUSHES        As Long = 16    ' Number of brushes the device has
Private Const NUMPENS           As Long = 18    ' Number of pens the device has
Private Const NUMMARKERS        As Long = 20    ' Number of markers the device has
Private Const NUMFONTS          As Long = 22    ' Number of fonts the device has
Private Const NUMCOLORS         As Long = 24    ' Number of colors the device supports
Private Const PDEVICESIZE       As Long = 26    ' Size required for device descriptor
Private Const CURVECAPS         As Long = 28    ' Curve capabilities
Private Const LINECAPS          As Long = 30    ' Line capabilities
Private Const POLYGONALCAPS     As Long = 32    ' Polygonal capabilities
Private Const TEXTCAPS          As Long = 34    ' Text capabilities
Private Const CLIPCAPS          As Long = 36    ' Clipping capabilities
Private Const RASTERCAPS        As Long = 38    ' Bitblt capabilities
Private Const ASPECTX           As Long = 40    ' Length of the X leg
Private Const ASPECTY           As Long = 42    ' Length of the Y leg
Private Const ASPECTXY          As Long = 44    ' Length of the hypotenuse

Private Const LOGPIXELSX        As Long = 88    ' Logical pixels/inch in X
Private Const LOGPIXELSY        As Long = 90    ' Logical pixels/inch in Y

Private Const SIZEPALETTE       As Long = 104    ' Number of entries in physical palette
Private Const NUMRESERVED       As Long = 106    ' Number of reserved entries in palette
Private Const COLORRES          As Long = 108    ' Actual color resolution

' Printing related DeviceCaps.
' These replace the appropriate Escapes
' -------------------------------------
Private Const PHYSICALWIDTH     As Long = 110    ' Physical Width in device units
Private Const PHYSICALHEIGHT    As Long = 111    ' Physical Height in device units
Private Const PHYSICALOFFSETX   As Long = 112    ' Physical Printable Area x margin
Private Const PHYSICALOFFSETY   As Long = 113    ' Physical Printable Area y margin
Private Const SCALINGFACTORX    As Long = 114    ' Scaling factor x
Private Const SCALINGFACTORY    As Long = 115    ' Scaling factor y

' Display driver specific
' -----------------------
Private Const VREFRESH          As Long = 116   ' Current vertical refresh rate of the
                                                ' display device (for displays only) in Hz
Private Const DESKTOPVERTRES    As Long = 117   ' Vertical height of entire desktop in pixels
Private Const DESKTOPHORZRES    As Long = 118   ' Horizontal width of entire desktop in pixels
Private Const BLTALIGNMENT      As Long = 119   ' Preferred blt alignment

Private Const SHADEBLENDCAPS    As Long = 120   ' Shading and blending caps
Private Const COLORMGMTCAPS     As Long = 121   ' Color Management caps

' Device Capability Masks:

' Device Technologies
' -------------------
Private Const DT_PLOTTER        As Long = 0   ' Vector plotter
Private Const DT_RASDISPLAY     As Long = 1   ' Raster display
Private Const DT_RASPRINTER     As Long = 2   ' Raster printer
Private Const DT_RASCAMERA      As Long = 3   ' Raster camera
Private Const DT_CHARSTREAM     As Long = 4   ' Character-stream, PLP
Private Const DT_METAFILE       As Long = 5   ' Metafile, VDM
Private Const DT_DISPFILE       As Long = 6   ' Display-file

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuDeviceCapsIndex
    dciLogicalPixelsX = LOGPIXELSX
    dciLogicalPixelsY = LOGPIXELSY
End Enum

Public Enum enuScreenAxis
    saX
    saY
End Enum

' -----------------------------------
' External Function declarations
' -----------------------------------

' The GetDC function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd144871(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetDC _
            Lib "user32.dll" (ByVal hWnd As LongPtr) _
            As LongPtr
#Else
    Private Declare _
            Function GetDC _
            Lib "user32.dll" (ByVal hWnd As Long) _
            As Long
#End If

' The ReleaseDC function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd162920(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function ReleaseDC _
            Lib "user32.dll" (ByVal hWnd As LongPtr, _
                          ByVal hDC As LongPtr) _
            As LongPtr
#Else
    Private Declare _
            Function ReleaseDC _
            Lib "user32.dll" (ByVal hWnd As Long, _
                          ByVal hDC As Long) _
            As Long
#End If

' The GetDeviceCaps function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/dd144877(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetDeviceCaps _
            Lib "Gdi32" (ByVal hWnd As LongPtr, _
                         ByVal nIndex As enuDeviceCapsIndex) _
            As LongPtr
#Else
    Private Declare _
            Function GetDeviceCaps _
            Lib "Gdi32" (ByVal hWnd As Long, _
                         ByVal nIndex As enuDeviceCapsIndex) _
            As Long
#End If

Public Function GetDPI(Optional Axis As enuScreenAxis = saX) As Long
' ==========================================================================
' Description : Return the current DPI setting for Windows
'
' Parameters  : Axis        Optionally specify which axis to use
'
' Returns     : Long
' ==========================================================================

    Const sPROC As String = "GetDPI"

    #If VBA7 Then
        Dim hDC As LongPtr
    #Else
        Dim hDC As Long
    #End If

    Dim lDPI    As Long
    Dim lRtn    As Long


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the device context
    ' ----------------------
    hDC = GetDC(0)

    ' Get the DPI
    ' -----------
    Select Case Axis
    Case saX
        lDPI = GetDeviceCaps(hDC, dciLogicalPixelsX)
    Case saY
        lDPI = GetDeviceCaps(hDC, dciLogicalPixelsY)
    Case Else
        lDPI = GetDeviceCaps(hDC, dciLogicalPixelsX)
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetDPI = lDPI

    ' Release the context
    ' -------------------
    If (hDC <> 0) Then
        lRtn = ReleaseDC(0, hDC)
    End If

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function PixelsToPoints(ByVal Axis As enuScreenAxis, _
                               ByVal Pixels As Long) As Single
' ==========================================================================
' Description : Convert pixels to points
'
' Parameters  : Axis        Identifies which axis to use (X or Y)
'               Pixels      The number of pixels to convert
'
' Returns     : Single
' ==========================================================================

    Const sPROC As String = "PixelsToPoints"

    Const lTPI  As Long = 1440  ' Twips Per Inch


    Dim lRtn    As Long
    Dim lDPI    As Long
    Dim sngRtn  As Single


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Determine the number of pixels per inch
    ' ---------------------------------------
    Select Case Axis
    Case enuScreenAxis.saX
        lDPI = GetDPI(saX)
    Case enuScreenAxis.saY
        lDPI = GetDPI(saY)
    End Select

    sngRtn = Pixels * lTPI / 20 / lDPI

    ' ----------------------------------------------------------------------

PROC_EXIT:

    PixelsToPoints = sngRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function

Public Function PointsPerPixel(ByVal Axis As enuScreenAxis) As Single
' ==========================================================================
' Description : Return the height or width of a pixel
'
' Parameters  : Axis    Determines if measuring height (Y) or width (X)
'
' Returns     : Single
' ==========================================================================

    Const sPROC As String = "PointsPerPixelX"

    Dim lDPI    As Long

    Dim sngRtn  As Single


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the current DPI setting
    ' ---------------------------
    lDPI = GetDPI(Axis)

    sngRtn = mlPOINTS_PER_INCH / lDPI

    ' ----------------------------------------------------------------------

PROC_EXIT:

    PointsPerPixel = sngRtn

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Function

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Function
