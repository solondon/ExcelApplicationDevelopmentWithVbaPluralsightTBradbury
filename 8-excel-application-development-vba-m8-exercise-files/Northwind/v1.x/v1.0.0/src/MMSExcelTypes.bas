Attribute VB_Name = "MMSExcelTypes"
' ==========================================================================
' Module      : MMSExcelTypes
' Type        : Module
' Description : Support for identifying Excel object and variable types
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit

' -----------------------------------
' Public Constant declarations
' -----------------------------------
' Global Level
' ----------------

' Standard types
' --------------
Public Const gsTYPENAME_BYTE            As String = "Byte"
Public Const gsTYPENAME_INTEGER         As String = "Integer"
Public Const gsTYPENAME_LONG            As String = "Long"
Public Const gsTYPENAME_SINGLE          As String = "Single"
Public Const gsTYPENAME_DOUBLE          As String = "Double"
Public Const gsTYPENAME_CURRENCY        As String = "Currency"
Public Const gsTYPENAME_DECIMAL         As String = "Decimal"
Public Const gsTYPENAME_DATE            As String = "Date"
Public Const gsTYPENAME_STRING          As String = "String"
Public Const gsTYPENAME_BOOLEAN         As String = "Boolean"
Public Const gsTYPENAME_ERROR           As String = "Error"
Public Const gsTYPENAME_EMPTY           As String = "Empty"
Public Const gsTYPENAME_NULL            As String = "Null"
Public Const gsTYPENAME_OBJECT          As String = "Object"
Public Const gsTYPENAME_UNKNOWN         As String = "Unknown"
Public Const gsTYPENAME_NOTHING         As String = "Nothing"

' Array types
' -----------
Public Const gsTYPENAME_BYTE_ARRAY      As String = "Byte()"
Public Const gsTYPENAME_INTEGER_ARRAY   As String = "Integer()"
Public Const gsTYPENAME_LONG_ARRAY      As String = "Long()"
Public Const gsTYPENAME_SINGLE_ARRAY    As String = "Single()"
Public Const gsTYPENAME_DOUBLE_ARRAY    As String = "Double()"
Public Const gsTYPENAME_CURRENCY_ARRAY  As String = "Currency()"
Public Const gsTYPENAME_DECIMAL_ARRAY   As String = "Decimal()"
Public Const gsTYPENAME_DATE_ARRAY      As String = "Date()"
Public Const gsTYPENAME_STRING_ARRAY    As String = "String()"
Public Const gsTYPENAME_BOOLEAN_ARRAY   As String = "Boolean()"
Public Const gsTYPENAME_EMPTY_ARRAY     As String = "Empty()"
Public Const gsTYPENAME_NULL_ARRAY      As String = "Null()"
Public Const gsTYPENAME_OBJECT_ARRAY    As String = "Object()"

' Excel types
' -----------
Public Const gsTYPENAME_EXCEL_CHART     As String = "Chart"
Public Const gsTYPENAME_EXCEL_DIALOG    As String = "DialogSheet"
Public Const gsTYPENAME_EXCEL_RANGE     As String = "Range"
Public Const gsTYPENAME_EXCEL_WKB       As String = "Workbook"
Public Const gsTYPENAME_EXCEL_WKS       As String = "Worksheet"
