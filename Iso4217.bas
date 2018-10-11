Attribute VB_Name = "Iso4217"
Option Compare Database
Option Explicit

' CurrencyCode V1.0.2
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.CurrencyCode

' Source:
'   https://www.iso.org/iso-4217-currency-codes.html

    
' Compiler constants.
'
' Select Early Binding (True) or Late Binding (False).
#Const EarlyBinding = True

' Operation constants.
'
' Base URL for currency lists at ISO.
Private Const ServiceUrl    As String = "https://www.currency-iso.org/dam/downloads/lists/"
' File to import.
Private Const Filename      As String = "list_one.xml"
' Imported table name of Filename.
Private Const TableName     As String = "CcyNtry"
    
' Enums.
'
' HTTP status codes, reduced.
Private Enum HttpStatus
    OK = 200
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
End Enum
    
' Create or update a table holding the current and complete list of
' currency codes and numbers according to ISO 4217.
' Data are retrieved directly from the source.
'
' A list of unique codes and numbers can be retrieved with this query:
'
'   SELECT DISTINCT
'       Ccy AS Code, CcyNbr AS [Number], CcyNm AS Name
'   FROM
'       CcyNtry
'   WHERE
'       Ccy Is Not Null;
'
'
' 2018-08-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UpdateIso4217() As Boolean

    Dim TableDef        As DAO.TableDef
    
    Dim ImportOptions   As AcImportXMLOption
    Dim Sql             As String
    Dim Url             As String
    Dim LastPublished   As Date
    Dim PublishingDate  As Date
    Dim Result          As Boolean
    
    ' Retrieve current publishing date.
    PublishingDate = Iso4217PublishingDate
    ' Retrive publishing date of table.
    LastPublished = LastPublishingDate()
    
    ' Check if new data have been published.
    If DateDiff("d", LastPublished, PublishingDate) = 0 Then
        ' Currency code table is current.
        Result = True
    Else
        ' Update currency table.
        For Each TableDef In CurrentDb.TableDefs
            If TableDef.Name = TableName Then
                ImportOptions = acAppendData
                Exit For
            End If
        Next
        If ImportOptions = acAppendData Then
            ' Clear current list.
            Sql = "Delete From " & TableName
            CurrentDb.Execute Sql
        Else
            ' First time import.
            ImportOptions = acStructureAndData
        End If
    
        ' Fetch the current list and append it to the (empty) table.
        Url = ServiceUrl & Filename
        On Error Resume Next
        Application.ImportXML Url, ImportOptions
        
        ' Return success if no error.
        If Not CBool(Err.Number) Then
            Result = True
            ' Store the current publishing date to avoid repeated calls.
            LastPublishingDate PublishingDate
        End If
    End If
    
    UpdateIso4217 = Result

End Function

' Retrieve the current publishing date for the ISO 4217 currency codes.
'
' 2018-08-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Iso4217PublishingDate() As Date

    ' Function constants.
    '
    ' Async setting.
    Const Async         As Variant = False
    ' XML node and attribute names.
    Const RootNodeName  As String = "ISO_4217"
    Const DateItemName  As String = "Pblshd"
  
#If EarlyBinding Then
    ' Microsoft XML, v6.0.
    Dim Document        As MSXML2.DOMDocument60
    Dim XmlHttp         As MSXML2.XMLHTTP60
    Dim RootNodeList    As MSXML2.IXMLDOMNodeList
    Dim RootNode        As MSXML2.IXMLDOMNode

    Set Document = New MSXML2.DOMDocument60
    Set XmlHttp = New MSXML2.XMLHTTP60
#Else
    Dim Document        As Object
    Dim XmlHttp         As Object
    Dim RootNodeList    As Object
    Dim RootNode        As Object

    Set Document = CreateObject("MSXML2.DOMDocument")
    Set XmlHttp = CreateObject("MSXML2.XMLHTTP")
#End If

    Static LastChecked  As Date
    Static ValueDate    As Date
    
    Dim Url             As String
    
    If DateDiff("d", LastChecked, Date) <= 0 Then
        ' ValueDate has been retrieved recently.
        ' Don't check again until tomorrow.
    Else
        ' Retrieve current status.
        
        Url = ServiceUrl & Filename
        
        ' Retrieve data.
        XmlHttp.Open "GET", Url, Async
        XmlHttp.send
        
        If XmlHttp.status = HttpStatus.OK Then
            ' File retrieved successfully.
            Document.loadXML XmlHttp.ResponseText
        
            Set RootNodeList = Document.getElementsByTagName(RootNodeName)
            ' Find root node.
            For Each RootNode In RootNodeList
                If RootNode.nodeName = RootNodeName Then
                    Exit For
                Else
                    Set RootNode = Nothing
                End If
            Next
            
            If Not RootNode Is Nothing Then
                ' Set update date.
                ValueDate = CDate(RootNode.Attributes.getNamedItem(DateItemName).nodeValue)
                ' Set check date.
                LastChecked = Date
            End If
        End If
    End If
    
    Set XmlHttp = Nothing
    Set Document = Nothing
    
    Iso4217PublishingDate = ValueDate

End Function

' Retrieve the ISO 4217 currency code matching an ISO 4217 currency number.
'
' An empty string will be returned is the currency number is not found, or
' a default currency code can be specified for not found currency numbers.
'
' Examples:
'   ? CurrencyCode("978")           -> "EUR"
'   ? CurrencyCode("000")           -> ""
'   ? CurrencyCode("000", "XXX")    -> "XXX"
'
' 2018-08-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyCode( _
    ByVal CurrencyNumber As String, _
    Optional ByVal DefaultCode As String) _
    As String
    
    ' Field names.
    Const CodeFieldName     As String = "Ccy"
    Const NumberFieldName   As String = "CcyNbr"

    Static Number           As String
    Static Code             As String
    
    If Number <> CurrencyNumber & DefaultCode Then
        Code = Nz(DLookup(CodeFieldName, TableName, NumberFieldName & " = '" & CurrencyNumber & "'"), DefaultCode)
        Number = CurrencyNumber & DefaultCode
    End If
    
    CurrencyCode = Code
    
End Function

' Retrieve the ISO 4217 currency number matching an ISO 4217 currency code.
'
' An empty string will be returned is the currency code is not found, or
' a default currency number can be specified for not found currency codes.
'
' Examples:
'   ? CurrencyNumber("EUR")         -> "978"
'   ? CurrencyNumber("ZZZ")         -> ""
'   ? CurrencyNumber("ZZZ", "999")  -> "999"
'
' 2018-08-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyNumber( _
    ByVal CurrencyCode As String, _
    Optional ByVal DefaultNumber As String) _
    As String
    
    ' Field names.
    Const CodeFieldName     As String = "Ccy"
    Const NumberFieldName   As String = "CcyNbr"

    Static Number           As String
    Static Code             As String
    
    If Code <> CurrencyCode & DefaultNumber Then
        Number = Nz(DLookup(NumberFieldName, TableName, CodeFieldName & " = '" & CurrencyCode & "'"), DefaultNumber)
        Code = CurrencyCode & DefaultNumber
    End If
    
    CurrencyNumber = Number
    
End Function

' Set or get the date of the last published list of ISO 4217 currency codes
' using a property of CurrentProject.
'
' Example:
'   PublishingDate = #2020/01/10#
'   ' Set
'   ? LastPublishingDate(PublishingDate)    -> 2020-01-10 00:00:00
'   ' Get
'   ? LastPublishingDate()                  -> 2020-01-10 00:00:00
'
' 2018-08-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function LastPublishingDate( _
    Optional ByVal NewPublishingDate As Date) _
    As Date

    Const PropertyName  As String = "Iso4217PublishingDate"
    
    Dim StoredUpdate    As AccessObjectProperty
    
    Dim Index           As Integer
    Dim PublishingDate  As Date
    Dim PublishingValue As String
    
    ' The property cannot hold a Date value.
    ' Convert NewPublishingDate to a string expression.
    PublishingValue = Format(NewPublishingDate, "yyyy\-mm\-dd hh\:nn\:ss")
    
    For Index = 0 To CurrentProject.Properties.Count - 1
        If CurrentProject.Properties(Index).Name = PropertyName Then
            ' The property exists.
            Set StoredUpdate = CurrentProject.Properties(Index)
        End If
    Next
    If StoredUpdate Is Nothing Then
        ' This property has not be created.
        ' Create it with the value of PublishingValue.
        CurrentProject.Properties.Add PropertyName, PublishingValue
        Set StoredUpdate = CurrentProject.Properties(PropertyName)
    ElseIf CDate(PublishingValue) > #12:00:00 AM# Then
        ' Set value of property.
        StoredUpdate.Value = PublishingValue
    ElseIf Not IsDate(StoredUpdate.Value) Then
        ' For some reason, the property is not holding a date expression.
        ' Reset the value.
        StoredUpdate.Value = PublishingValue
    End If
    
    ' Read the stored string expression and convert to a date value.
    PublishingDate = CDate(StoredUpdate.Value)
    
    LastPublishingDate = PublishingDate
    
End Function
