Attribute VB_Name = "MMSOfficeProperties"
' ==========================================================================
' Module      : MMSOfficeProperties
' Type        : Module
' Description : Support for working with Office document properties
' --------------------------------------------------------------------------
' Procedures  : DeleteDocumentProperty
'               DocumentPropertyExists              Boolean
'               DocumentPropertyTypeToString        String
'               GetDocumentProperty                 Variant
'               GetDocumentPropertyType             Variant
'               ListProperties
'               SetDocumentProperty                 Boolean
' --------------------------------------------------------------------------
' References  : Microsoft Office Object Library
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Constant declarations
' -----------------------------------
' Global Level
' ----------------

Public Const gsDOCPROP_AUTHOR               As String = "Author"
Public Const gsDOCPROP_TITLE                As String = "Title"
Public Const gsDOCPROP_SUBJECT              As String = "Subject"
Public Const gsDOCPROP_KEYWORDS             As String = "Keywords"
Public Const gsDOCPROP_CATEGORY             As String = "Category"
Public Const gsDOCPROP_STATUS               As String = "Content status"
Public Const gsDOCPROP_COMMENTS             As String = "Comments"

Public Const gsDOCPROP_HYPERLINK_BASE       As String = "Hyperlink base"

Public Const gsDOCPROP_MANAGER              As String = "Manager"
Public Const gsDOCPROP_COMPANY              As String = "Company"

Public Const gsDOCPROP_CREATION             As String = "Creation date"
Public Const gsDOCPROP_LAST_AUTHOR          As String = "Last author"
Public Const gsDOCPROP_LAST_SAVE            As String = "Last save time"
Public Const gsDOCPROP_VERSION              As String = "Document version"

Public Const gsDOCPROP_REVISION             As String = "Revision number"
Public Const gsDOCPROP_TEMPLATE             As String = "Template"

Public Const gsDOCPROP_NUM_BYTES            As String = "Number of bytes"
Public Const gsDOCPROP_NUM_CHARS            As String = "Number of characters"
Public Const gsDOCPROP_NUM_CHAR_SPACES      As String = "Number of characters (with spaces)"
Public Const gsDOCPROP_NUM_WORDS            As String = "Number of words"
Public Const gsDOCPROP_NUM_LINES            As String = "Number of lines"
Public Const gsDOCPROP_NUM_PARAGRAPHS       As String = "Number of paragraphs"
Public Const gsDOCPROP_NUM_PAGES            As String = "Number of pages"

Public Const gsDOCPROP_NUM_SLIDES           As String = "Number of slides"
Public Const gsDOCPROP_NUM_NOTES            As String = "Number of notes"
Public Const gsDOCPROP_NUM_HIDDEN_SLIDES    As String = "Number of hidden Slides"
Public Const gsDOCPROP_NUM_MULTIMEDIA_CLIPS As String = "Number of multimedia clips"

Public Const gsDOCPROP_CST_VERSION          As String = "Version"
Public Const gsDOCPROP_CST_BUILDDATE        As String = "Date completed"

' ----------------
' Module Level
' ----------------

Private Const msMODULE                      As String = "MMSOfficeProperties"

' -----------------------------------
' Enumeration declarations
' -----------------------------------
' Global Level
' ----------------

Public Enum enuDocumentPropertyGroup
    dpgBuiltIn = 1
    dpgCustom = 2
    dpgBoth = dpgBuiltIn Or dpgCustom
End Enum

Public Sub DeleteDocumentProperty(ByVal PropertyName As String, _
                         Optional ByVal PropertyGroup _
                                     As enuDocumentPropertyGroup _
                                      = dpgCustom, _
                         Optional ByRef Document As Object)
' ==========================================================================
' Description : Delete an entry from a property group
'
' Parameters  : PropertyName  The name of the property
'               PropertyGroup The group the property belongs to
'               Book          The workbook the property is located in
' ==========================================================================

    Const sPROC As String = "DeleteDocumentProperty"

    Dim dDate   As Date

    Dim objDoc  As Object
    Dim eGroup  As enuDocumentPropertyGroup
    
    Dim oProp   As Office.DocumentProperty
    Dim oProps  As Office.DocumentProperties


    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get a reference to the Workbook
    ' -------------------------------
    If (Document Is Nothing) Then
        Set objDoc = GetDocument()
    Else
        Set objDoc = Document
    End If

    ' Determine which set of properties to look at
    ' --------------------------------------------
    eGroup = PropertyGroup
    Select Case PropertyGroup
    Case dpgBuiltIn
        Set oProps = objDoc.BuiltinDocumentProperties

    Case dpgCustom
        Set oProps = objDoc.CustomDocumentProperties

    Case dpgBoth
        eGroup = dpgCustom
        Set oProps = objDoc.CustomDocumentProperties

    End Select

    ' Test if the property exists
    ' ---------------------------
    Set oProp = oProps(PropertyName)

    If ((oProp Is Nothing) And (PropertyGroup = dpgBoth)) Then
        Set oProps = objDoc.BuiltinDocumentProperties
        Set oProp = oProps(PropertyName)
    End If
    
    ' Delete it if found
    ' ------------------
    If (Not oProp Is Nothing) Then

        ' Built-in properties cannot be deleted so
        ' their values are set to uninitialized state
        ' -------------------------------------------
        If (eGroup = dpgBuiltIn) Then
            Select Case oProp.type
            Case msoPropertyTypeString
                oProp = vbNullString
            Case msoPropertyTypeBoolean
                oProp = False
            Case Else
                oProp = 0
            End Select
        Else
            oProp.Delete
        End If
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set oProp = Nothing
    Set oProps = Nothing
    Set objDoc = Nothing

    Call Trace(tlVerbose, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

End Sub

Public Function DocumentPropertyExists(ByVal PropertyName As String, _
                              Optional ByVal PropertyGroup _
                                          As enuDocumentPropertyGroup _
                                           = dpgBoth, _
                              Optional ByRef Document As Object) _
       As Boolean
' ==========================================================================
' Description : Determines if a given property exists in the the document
'
' Parameters  : PropertyName    The name of the property to look for
'               PropertyGroup   The group to look in
'               Book            The workbook to look in
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC As String = "DocumentPropertyExists"

    Dim bRtn    As Boolean

    Dim objDoc  As Object

    Dim oProp   As Office.DocumentProperty
    Dim oProps  As Office.DocumentProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get a reference to the Workbook
    ' -------------------------------
    If (Document Is Nothing) Then
        Set objDoc = GetDocument()
    Else
        Set objDoc = Document
    End If

    ' Determine which set of properties to look at
    ' --------------------------------------------
    Select Case PropertyGroup
    Case dpgBuiltIn
        Set oProps = objDoc.BuiltinDocumentProperties

    Case dpgCustom
        Set oProps = objDoc.CustomDocumentProperties

    Case dpgBoth
        Set oProps = objDoc.BuiltinDocumentProperties

    End Select

    On Error Resume Next

    ' Test if the property exists
    ' ---------------------------
    Set oProp = oProps(PropertyName)

    ' Return the value if found
    ' -------------------------
    If (Err.Number = ERR_SUCCESS) Then
        bRtn = True
        GoTo PROC_EXIT
    End If

    ' If not found and both are selected
    ' search the other set of properties
    ' ----------------------------------
    If (PropertyGroup = dpgBoth) Then
        Set oProps = objDoc.CustomDocumentProperties
        Set oProp = oProps(PropertyName)
        bRtn = (Err.Number = ERR_SUCCESS)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    DocumentPropertyExists = bRtn

    Set oProp = Nothing
    Set oProps = Nothing
    Set objDoc = Nothing

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

Public Function DocumentPropertyTypeToString(ByVal PropertyType _
                                                As MsoDocProperties) _
       As String
' ==========================================================================
' Description : Convert an enumeration to a string
'
' Parameters  : PropertyType    The value to convert
'
' Returns     : String
' ==========================================================================

    Const sPROC As String = "DocumentPropertyTypeToString"

    Dim sRtn    As String


    On Error GoTo PROC_ERR

    ' ----------------------------------------------------------------------
    
    Select Case PropertyType
    Case msoPropertyTypeBoolean
        sRtn = "msoPropertyTypeBoolean"

    Case msoPropertyTypeDate
        sRtn = "msoPropertyTypeDate"

    Case msoPropertyTypeString
        sRtn = "msoPropertyTypeString"

    Case msoPropertyTypeNumber
        sRtn = "msoPropertyTypeNumber"

    Case msoPropertyTypeFloat
        sRtn = "msoPropertyTypeFloat"

    Case Else
        sRtn = vbNullString
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    DocumentPropertyTypeToString = sRtn

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

Public Function GetDocumentProperty(ByVal PropertyName As String, _
                           Optional ByVal PropertyGroup _
                                       As enuDocumentPropertyGroup _
                                        = dpgBoth, _
                           Optional ByRef Document As Object) _
       As Variant
' ==========================================================================
' Description : Get a property value from the document
'
' Parameters  : PropertyName  The name of the property
'               PropertyGroup The group the property belongs to
'               Document      The document the property is located in
'
' Returns     : Variant
' ==========================================================================

    Const sPROC As String = "GetDocumentProperty"

    Dim vRtn    As Variant
    Dim objDoc  As Object
    
    Dim oProp   As Office.DocumentProperty
    Dim oProps  As Office.DocumentProperties


    On Error GoTo PROC_ERR
    Call Trace(tlVerbose, msMODULE, sPROC, PropertyName)

    ' ----------------------------------------------------------------------
    ' Get a reference to the Workbook
    ' -------------------------------
    If (Document Is Nothing) Then
        Set objDoc = GetDocument
    Else
        Set objDoc = Document
    End If

    ' Determine which set of properties to look at
    ' --------------------------------------------
    Select Case PropertyGroup
    Case dpgBuiltIn
        Set oProps = objDoc.BuiltinDocumentProperties

    Case dpgCustom
        Set oProps = objDoc.CustomDocumentProperties

    Case dpgBoth
        Set oProps = objDoc.BuiltinDocumentProperties

    End Select

    On Error Resume Next

    ' Test if the property exists
    ' ---------------------------
    Set oProp = oProps(PropertyName)

    ' Return the value if found
    ' -------------------------
    If (Err.Number = ERR_SUCCESS) Then
        vRtn = oProp.Value
        GoTo PROC_EXIT
    End If

    ' If not found and both are selected
    ' search the other set of properties
    ' ----------------------------------
    If (PropertyGroup = dpgBoth) Then
        Set oProps = objDoc.CustomDocumentProperties
        Set oProp = oProps(PropertyName)

        If (Err.Number = ERR_SUCCESS) Then
            vRtn = oProp.Value
            GoTo PROC_EXIT
        End If
    End If

    vRtn = Null

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetDocumentProperty = vRtn

    Set oProp = Nothing
    Set oProps = Nothing
    Set objDoc = Nothing

    Call Trace(tlVerbose, msMODULE, sPROC, vRtn)
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

Public Function GetDocumentPropertyType(ByRef PropertyValue As Variant) _
       As Variant
' ==========================================================================
' Description : Determine the type of a property value
'
' Parameters  : PropertyValue The value to test for type
'
' Returns     : Variant       Contains a member of Office.MsoDocProperties
'                             or empty if it is not valid
' ==========================================================================

    Const sPROC As String = "GetOfficeDocumentPropertyType"

    Dim vRtn    As Variant


    On Error GoTo PROC_ERR
    '  Call Trace(tlVerbose, msMODULE, sPROC, PropertyValue)

    ' ----------------------------------------------------------------------

    Select Case VarType(PropertyValue)
    Case vbBoolean
        vRtn = msoPropertyTypeBoolean

    Case vbDate
        vRtn = msoPropertyTypeDate

    Case vbString
        vRtn = msoPropertyTypeString

    Case vbInteger, vbLong
        vRtn = msoPropertyTypeNumber

    Case vbSingle, vbDouble
        vRtn = msoPropertyTypeFloat

    Case Else
        vRtn = Null
    End Select

    ' ----------------------------------------------------------------------

PROC_EXIT:

    GetDocumentPropertyType = vRtn

    '  Call Trace(tlVerbose, msMODULE, sPROC, DocumentPropertyTypeToString(vRtn))
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

Public Sub ListProperties(Optional ByVal PropertyGroup _
                                      As enuDocumentPropertyGroup _
                                       = dpgBoth, _
                          Optional ByRef Document As Object, _
                          Optional ByVal CopyToClipboard As Boolean)
' ==========================================================================
' Description : List the document properties in the immediate window.
'
' Parameters  : PropertyGroup   Identifies whether to use the
'                               Built-In or Custom properties group
'               Document        The document to report on
' ==========================================================================

    Const sPROC     As String = "ListProperties"

    Const lLINE_LEN As Long = 70
    Const lTAB_SIZE As Long = 36

    Dim lIdx        As Long
    Dim lTab        As Long

    Dim sLine       As String
    Dim sLines      As String
    Dim lMaxLen     As Long

    Dim objDoc      As Object

    Dim eGroup      As enuDocumentPropertyGroup

    Dim oProp       As Office.DocumentProperty
    Dim oProps      As Office.DocumentProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_ENTER)

    ' ----------------------------------------------------------------------
    ' Get the document
    ' ----------------
    If Document Is Nothing Then
        Set objDoc = GetDocument()
    Else
        Set objDoc = Document
    End If

    Debug.Print String$(lLINE_LEN, "=")


    ' ----------------------------------------------------------------------
    ' Get the length of the
    ' longest property name
    ' ---------------------
    Set oProps = objDoc.BuiltinDocumentProperties
    For Each oProp In oProps
        lMaxLen = MaxVal(lMaxLen, Len(oProp.Name))
    Next oProp
    Set oProps = objDoc.CustomDocumentProperties
    For Each oProp In oProps
        lMaxLen = MaxVal(lMaxLen, Len(oProp.Name))
    Next oProp
    lMaxLen = lMaxLen + 7
    Do Until ((lMaxLen Mod 4) = 0)
        lMaxLen = lMaxLen - 1
    Loop

    ' ----------------------------------------------------------------------

    If (PropertyGroup = dpgCustom) Then
        GoTo CUSTOM
    End If

    ' ----------------------------------------------------------------------

BUILTIN:

    Set oProps = objDoc.BuiltinDocumentProperties
    eGroup = dpgBuiltIn

    sLine = "Built-In Properties for " & objDoc.Name
    sLines = Concat(vbNewLine, sLines, sLine)
    Debug.Print sLine

    sLine = String$(lLINE_LEN, "=")
    sLines = Concat(vbNewLine, sLines, sLine)
    Debug.Print sLine

    GoSub PRINT_PROPS

    ' ----------------------------------------------------------------------

    If (PropertyGroup = dpgBuiltIn) Then
        GoTo PROC_EXIT
    End If

    ' ----------------------------------------------------------------------

CUSTOM:

    Set oProps = objDoc.CustomDocumentProperties
    eGroup = dpgCustom

    sLine = "Custom Properties for " & objDoc.Name
    sLines = Concat(vbNewLine, sLines, sLine)
    Debug.Print sLine

    sLine = String$(lLINE_LEN, "=")
    sLines = Concat(vbNewLine, sLines, sLine)
    Debug.Print sLine

    GoSub PRINT_PROPS

    ' Make sure a vbNewLine is at the end
    ' -----------------------------------
    sLines = sLines & vbNewLine

    If CopyToClipboard Then
        Call SetClipboardText(sLines)
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    Set objDoc = Nothing
    Set oProp = Nothing
    Set oProps = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, gsPROC_EXIT)
    On Error GoTo 0

    Exit Sub

    ' ----------------------------------------------------------------------

PROC_ERR:

    If ErrorHandler(msMODULE, sPROC) Then
        Stop
        Resume
    Else
        Resume PROC_EXIT
    End If

    ' ----------------------------------------------------------------------

PRINT_PROPS:

    ' Some properties may throw an error if not set
    ' ---------------------------------------------
    On Error Resume Next

    For Each oProp In oProps
        sLine = PadR(oProp.Name, lMaxLen) _
              & Replace(oProp.Value, vbLf, gsMULTILINE_SEP)
        Debug.Print sLine

        ' If an error occurs
        ' only print the name
        ' -------------------
        If (Err.Number <> ERR_SUCCESS) Then
            sLine = oProp.Name
            Debug.Print sLine
            Err.Clear
        End If
        sLines = Concat(vbNewLine, sLines, sLine)
    Next oProp

    On Error GoTo PROC_ERR

    sLine = String$(lLINE_LEN, "-")
    sLines = Concat(vbNewLine, sLines, sLine)
    Debug.Print sLine

    sLine = CStr(oProps.Count) & " " _
                & IIf(eGroup = dpgBuiltIn, "Built-In", "Custom") _
                & " Properties."
    sLines = Concat(vbNewLine, sLines, sLine)
    Debug.Print sLine

    sLine = String$(lLINE_LEN, "=")
    sLines = Concat(vbNewLine, sLines, sLine)
    Debug.Print String$(lLINE_LEN, "=")

    Return

End Sub

Public Function SetDocumentProperty(ByVal PropertyName As String, _
                                    ByVal PropertyGroup _
                                       As enuDocumentPropertyGroup, _
                                    ByVal PropertyValue As Variant, _
                           Optional ByRef Document As Object, _
                           Optional ByVal ContentLink As Boolean = False) _
       As Boolean
' ==========================================================================
' Description : Set the value of PropertyName to PropertyValue.
'               If PropertyName does not exist, it will be created if
'               PropertyGroup is either Custom or Both.
'
' Parameters  : PropertyName  The name of the property
'               PropertyGroup The group the property belongs to
'               PropertyValue The value to assign to the property
'               Document      The document the property is located in
'               ContentLink   Creates a live link to a range address (Excel)
'
' Returns     : Boolean
' ==========================================================================

    Const sPROC     As String = "SetDocumentProperty"

    Dim bRtn        As Boolean
    Dim ePropType   As MsoDocProperties
    Dim objDoc      As Object

    Dim oProp       As Office.DocumentProperty
    Dim oProps      As Office.DocumentProperties


    On Error GoTo PROC_ERR
    Call Trace(tlMaximum, msMODULE, sPROC, PropertyName)

    ' ----------------------------------------------------------------------
    ' Test for valid inputs
    ' ---------------------
    If (PropertyName = vbNullString) Then
        GoTo PROC_EXIT
    ElseIf IsArray(PropertyValue) Then
        GoTo PROC_EXIT
    End If

    ' Get the document
    ' ----------------
    If (Document Is Nothing) Then
        Set objDoc = GetDocument
    Else
        Set objDoc = Document
    End If

    ' If it is a BuiltIn property just assign the value
    ' Linked content cannot be added to a BuiltIn property
    ' ----------------------------------------------------
    If (PropertyGroup = dpgBuiltIn) Or (PropertyGroup = dpgBoth) Then
        Set oProps = objDoc.BuiltinDocumentProperties

        Err.Clear

        Set oProp = oProps(PropertyName)

        If (Err.Number = ERR_SUCCESS) Then
            ' The Property exists.
            ' Set the value and exit.
            '------------------------
            oProp.Value = PropertyValue
            bRtn = True
            GoTo PROC_EXIT
        End If

        ' The property doesn't exist
        ' --------------------------
        GoTo PROC_EXIT

    End If

    If (PropertyGroup = dpgCustom) Or (PropertyGroup = dpgBoth) Then
        ' Delete the existing CustomProperty and
        ' replace it with a new CustomProperty with
        ' the same name. This allows us to
        ' change a LinkedContent property to
        ' an unlinked content property and vice-versa.
        '---------------------------------------------
        On Error Resume Next
        Err.Clear
        Set oProps = objDoc.CustomDocumentProperties
        Set oProp = oProps(PropertyName)
        On Error GoTo PROC_ERR

        ' If the property exists, delete it
        '----------------------------------
        If (Not oProp Is Nothing) Then
            oProp.Delete
        End If

        Err.Clear

        If ((Application.Name = gsOFFICE_APPNAME_EXCEL) _
        And ContentLink) Then
            ' If ContentLink is True, then PropertyValue
            ' is the defined name to which the property
            ' will be linked. In this case, PropertyValue
            ' must be a String and the Name must exist.
            ' -------------------------------------------

            If IsObject(PropertyValue) Then
                ' See if it is an Excel.Name
                ' --------------------------
                If TypeOf PropertyValue Is Excel.Name Then
                    ' Set the link and exit
                    ' ---------------------
                    Err.Clear

                    oProps.Add Name:=PropertyName, _
                               LinkToContent:=True, _
                               type:=msoPropertyTypeString, _
                               LinkSource:=PropertyValue.Name
                    bRtn = (Err.Number = 0)
                    GoTo PROC_EXIT
                Else
                    ' PropertyValue is an object
                    ' but is not a Name. Exit.
                    ' --------------------------
                    bRtn = False
                    GoTo PROC_EXIT
                End If

            ElseIf (VarType(PropertyValue) = vbString) Then
                If (Not NameExists(CStr(PropertyValue), objDoc)) Then
                    ' Name doesn't exist. Exit.
                    ' -------------------------
                    bRtn = False
                    GoTo PROC_EXIT
                End If
                ' Name exists. Create the link.
                ' -----------------------------
                oProps.Add Name:=PropertyName, _
                           type:=msoPropertyTypeString, _
                           LinkSource:=PropertyValue, _
                           LinkToContent:=True
            Else
                ' PropertyValue is neither
                ' a Name nor a String. Exit.
                ' --------------------------
                bRtn = False
                GoTo PROC_EXIT
            End If
        Else
            ' Not linking content. Just create
            ' the property, set the value, and exit.
            ' --------------------------------------
            Err.Clear
            ePropType = GetDocumentPropertyType(PropertyValue)

            If IsNull(ePropType) Then
                ' Illegal data type.
                ' ------------------
                bRtn = False
                GoTo PROC_EXIT
            End If

            Err.Clear
            oProps.Add Name:=PropertyName, _
                       LinkToContent:=False, _
                       type:=ePropType, _
                       Value:=PropertyValue
            bRtn = (Err.Number = 0)
        End If
    End If

    ' ----------------------------------------------------------------------

PROC_EXIT:

    SetDocumentProperty = bRtn

    Set oProp = Nothing
    Set oProps = Nothing
    Set objDoc = Nothing

    Call Trace(tlMaximum, msMODULE, sPROC, PropertyValue)
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
