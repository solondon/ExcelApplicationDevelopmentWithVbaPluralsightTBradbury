Attribute VB_Name = "MWinAPIClipboardPicture"
' ==========================================================================
' Module      : MWinAPIClipboardPicture
' Type        : Module
' Description : Support for graphics on the Windows clipboard
' --------------------------------------------------------------------------
' Procedures  : CreatePicture       IPicture
'               PastePicture        IPicture
' --------------------------------------------------------------------------
' Dependencies: MWinAPIClipboard
'               MWinAPIError
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Constant declarations
' -----------------------------------
' Module Level
' ----------------

Private Const msMODULE As String = "MWinAPIClipboardPicture"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuImageType
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum

' OLE Picture types
' -----------------
Public Enum enuPicType
    PICTYPE_UNINITIALIZED = (-1)
    PICTYPE_NONE = 0
    PICTYPE_BITMAP = 1
    PICTYPE_METAFILE = 2
    PICTYPE_ICON = 3
    PICTYPE_ENHMETAFILE = 4
End Enum

' -----------------------------------
' Type declarations
' -----------------------------------
' Module Level
' ----------------

' Declare a UDT to store
' the bitmap information
' ----------------------
#If VBA7 Then
    Private Type TPICTDESC
        cbSizeofstruct  As Long
        PicType         As enuPicType
        hPict           As LongPtr
        hPal            As LongPtr
    End Type
#Else
    Private Type TPICTDESC
        cbSizeofstruct  As Long
        PicType         As enuPicType
        hPict           As Long
        hPal            As Long
    End Type
#End If

' -----------------------------------
' External Function declarations
' -----------------------------------
' Module Level
' ----------------

' The OpenClipboard function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms649048(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function OpenClipboard _
            Lib "User32" (ByVal hWnd As LongPtr) _
            As Boolean
#Else
    Private Declare _
            Function OpenClipboard _
            Lib "User32" (ByVal hWnd As Long) _
            As Boolean
#End If

' The EmptyClipboard function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms649037(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function EmptyClipboard _
            Lib "User32" () _
            As Boolean
#Else
    Private Declare _
            Function EmptyClipboard _
            Lib "User32" () _
            As Boolean
#End If

' The CloseClipboard function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms649035(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function CloseClipboard _
            Lib "User32" () _
            As Boolean
#Else
    Private Declare _
            Function CloseClipboard _
            Lib "User32" () _
            As Boolean
#End If

' The GetClipboardData function is described on MSDN at
' http://msdn.microsoft.com/en-us/office/ms649039(v=vs.90).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function GetClipboardData _
            Lib "User32" (ByVal uFormat As enuClipboardFormat) _
            As Long
#Else
    Private Declare _
            Function GetClipboardData _
            Lib "User32" (ByVal uFormat As enuClipboardFormat) _
            As Long
#End If

' The IsClipboardFormatAvailable function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms649047(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function IsClipboardFormatAvailable _
            Lib "User32" (ByVal Format As enuClipboardFormat) _
            As Boolean
#Else
    Private Declare _
            Function IsClipboardFormatAvailable _
            Lib "User32" (ByVal Format As enuClipboardFormat) _
            As Boolean
#End If

'Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As enuClipboardFormat) As Long

' The CopyEnhMetaFile function is described on MSDN at
' http://msdn.microsoft.com/en-us/office/dd183479(v=vs.80)
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function CopyEnhMetaFile _
            Lib "Gdi32" _
            Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, _
                                      ByVal lpszFile As String) _
            As Long
#Else
    Private Declare _
            Function CopyEnhMetaFile _
            Lib "Gdi32" _
            Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, _
                                      ByVal lpszFile As String) _
            As Long
#End If

' The CopyImage function is described on MSDN at
' http://msdn.microsoft.com/en-us/library/ms648031(v=vs.85).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function CopyImage _
            Lib "User32" (ByVal hImage As LongPtr, _
                          ByVal uType As enuImageType, _
                          ByVal cxDesired As Long, _
                          ByVal cyDesired As Long, _
                          ByVal fuFlags As enuLoadResourceFlag) _
            As LongPtr
#Else
    Private Declare _
            Function CopyImage _
            Lib "User32" (ByVal hImage As Long, _
                          ByVal uType As enuImageType, _
                          ByVal cxDesired As Long, _
                          ByVal cyDesired As Long, _
                          ByVal fuFlags As enuLoadResourceFlag) _
            As Long
#End If

' The OleCreatePictureIndirect function is described on MSDN at
' http://msdn.microsoft.com/en-us/office/ms694511(v=vs.90).aspx
' -----------------------------------
#If VBA7 Then
    Private Declare PtrSafe _
            Function OleCreatePictureIndirect _
            Lib "OlePro32" (ByRef lpPictDesc As TPICTDESC, _
                            ByRef riid As TGUID, _
                            ByVal fOwn As Boolean, _
                            ByRef lplpvObj As IPicture) _
            As Long
#Else
    Private Declare _
            Function OleCreatePictureIndirect _
            Lib "OlePro32" (ByRef lpPictDesc As TPICTDESC, _
                            ByRef riid As TGUID, _
                            ByVal fOwn As Boolean, _
                            ByRef lplpvObj As IPicture) _
            As Long
#End If

#If VBA7 Then
    Public Function CreatePicture(ByVal hPict As LongPtr, _
                                  ByVal hPal As LongPtr, _
                                  ByVal Format As enuClipboardFormat) _
           As IPicture
#Else
    Public Function CreatePicture(ByVal hPict As Long, _
                                  ByVal hPal As Long, _
                                  ByVal Format As enuClipboardFormat) _
           As IPicture
#End If
' ==========================================================================
' Description : Convert an image (and palette) handle into a Picture object
'
' Parameters  : hPict
'               hPal
'               Format
'
' Returns     : IPicture
' ==========================================================================

    Const sPROC         As String = "CreatePicture"


    Dim lRtn            As Long
    Dim sDescription    As String

    Dim udtRIID         As TGUID
    Dim udtPictDesc     As TPICTDESC

    Dim IPic            As IPicture


    On Error GoTo PROC_ERR
    Call Trace(tlNormal, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Create the Interface GUID (for the IPicture interface)
    ' ------------------------------------------------------
    With udtRIID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With

    ' Populate the structure
    ' ----------------------
    With udtPictDesc
        .cbSizeofstruct = Len(udtPictDesc)          ' Length of structure.
        .PicType = IIf(Format = CF_BITMAP, _
                       PICTYPE_BITMAP, _
                       PICTYPE_ENHMETAFILE)         ' Type of Picture
        .hPict = hPict                              ' Handle to image.
        .hPal = IIf(Format = CF_BITMAP, hPal, 0)    ' Handle to palette
    End With                                        ' (if bitmap)

    ' Convert the handle into
    ' an OLE IPicture interface
    ' -------------------------
    lRtn = OleCreatePictureIndirect(udtPictDesc, udtRIID, True, IPic)

    ' If an error occured, show the description
    ' -----------------------------------------
    If (lRtn <> 0) Then
        sDescription = HResultErrorToString(lRtn)
        Call Err.Raise(ERR_OLE_AUTOMATION, _
                       Concat(".", msMODULE, sPROC), _
                       sDescription)
        Debug.Print "Create Picture: " & sDescription
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    ' Return the new Picture object
    ' -----------------------------
    Set CreatePicture = IPic

    On Error GoTo 0
    Call Trace(tlNormal, msMODULE, sPROC, gsPROC_EXIT)

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

Public Function PastePicture(Optional ByVal Format _
                                         As enuImageType = IMAGE_BITMAP) _
       As IPicture
' ==========================================================================
' Description : Get a Picture object showing whatever's on the clipboard.
'
' Parameters  : Format      The type of picture to create.
'                           Possible values:
'                           - IMAGE_BITMAP to create a bitmap (default)
'                           - IMAGE_ENHMETAFILE to create a metafile
'
' Returns   : IPicture
' ==========================================================================

    Const sPROC             As String = "PastePicture"

    Const ORIGINAL_HEIGHT   As Long = 0
    Const ORIGINAL_WIDTH    As Long = 0

    Dim bAvailable          As Boolean
    Dim bRtn                As Boolean
    
    Dim eFormat             As enuClipboardFormat
    
    Dim IPic                As IPicture

    #If VBA7 Then
        Dim hPtr            As LongPtr
        Dim hCopy           As LongPtr
    #Else
        Dim hPtr            As Long
        Dim hCopy           As Long
    #End If


    On Error GoTo PROC_ERR
    Call Trace(tlNormal, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Convert the type of picture requested
    ' -------------------------------------------
    Select Case Format
    Case IMAGE_ENHMETAFILE
        eFormat = CF_ENHMETAFILE
    Case Else
        eFormat = CF_BITMAP
    End Select

    ' Check if the clipboard contains the required format
    ' ---------------------------------------------------
    bAvailable = IsClipboardFormatAvailable(eFormat)

    If bAvailable Then

        bRtn = OpenClipboard(0&)

        If bRtn Then

            ' Get a handle to the image data
            ' ------------------------------
            hPtr = GetClipboardData(eFormat)

            ' Create a copy of the image on the
            ' clipboard in the appropriate format
            ' -----------------------------------
            If (eFormat = CF_ENHMETAFILE) Then
                hCopy = CopyEnhMetaFile(hPtr, vbNullString)
            Else
                hCopy = CopyImage(hPtr, _
                                  IMAGE_BITMAP, _
                                  ORIGINAL_WIDTH, _
                                  ORIGINAL_HEIGHT, _
                                  LR_COPYRETURNORG)
            End If

            ' Release the clipboard to other programs
            ' ---------------------------------------
            bRtn = CloseClipboard()

            ' If a valid handle to the image is returned,
            ' convert it to a Picture object and return
            ' -------------------------------------------
            If (hPtr <> 0) Then
                Set PastePicture = CreatePicture(hCopy, 0, eFormat)
            End If

        End If
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    On Error GoTo 0
    Call Trace(tlNormal, msMODULE, sPROC, gsPROC_EXIT)

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
