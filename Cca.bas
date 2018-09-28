Attribute VB_Name = "Cca"
Option Compare Database
Option Explicit

' CurrencyCode V1.1.1
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.CurrencyCode


' API id or key. Guid string, 0, 24, or 32 characters.
'
' Currency Converter API:           "00000000-0000-0000-0000-000000000000"
' Leave empty for the free plan:    ""
Public Const CcaApiId   As String = ""

' Enums.
'
' Dimensions of array holding parameters.
Private Enum ParameterDetail
    Name = 0
    Value = 1
End Enum
'
' Dimensions of array holding codes.
Private Enum CodeDetail
    Code = 0
    Sign = 1
    Name = 2
End Enum
'
' HTTP status codes, reduced.
Private Enum HttpStatus
    OK = 200
    BadRequest = 400
    Unauthorized = 401
    Forbidden = 403
End Enum

' Currency code for neutral currency.
Public Const NeutralCode        As String = "XXX"
' Currency name for neutral currency.
Public Const NeutralName        As String = "No currency"
' Currency sign for neutral currency.
Public Const NeutralSign        As String = "¤"

' Retrieve the current currency code list from Currency Converter API.
' The list is returned as an array and cached until the next update.
'
' Source:
'   https://currencyconverterapi.com/
'   https://currencyconverterapi.com/docs
'
' Note:
'   The services are provided as is and without warranty.
'
' Example:
'   Dim Codes As Variant
'   Codes = ExchangeRatesCca()
'   Codes(101, 0)   -> CHF              ' Currency code.
'   Codes(101, 1)   -> "Fr."            ' Currency name.
'   Codes(101, 2)   -> "Swiss Franc"    ' Currency name.
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CurrencyCodesCca() As Variant
    
    ' Operational constants.
    '
    ' API endpoint.
    Const FreeSubdomain As String = "free"
    Const PaidSubdomain As String = "api"
    Const TempSubdomain As String = "xxx"
    ' API version must be 3 or higher.
    Const ApiVersion    As String = "6"
    Const ServiceUrl    As String = "https://" & TempSubdomain & ".currencyconverterapi.com/api/v" & ApiVersion & "/currencies"
    ' Update interval in minutes.
    Const UpdatePause   As Integer = 24 * 60
    
    ' Function constants.
    '
    ' Node names in retrieved collection.
    Const RootNodeName  As String = "root"
    Const ListNodeName  As String = "results"
    ' ResponseText when invalid currency code is passed.
    Const EmptyResponse As String = "{}"
    ' Field names.
    Const CodeId        As String = "id"
    Const CodeName      As String = "currencyName"
    Const CodeSymbol    As String = "currencySymbol"
    
    Static CodePairs    As Collection
    
    Static Codes()      As Variant
    Static LastCall     As Date
    
    Dim DataCollection  As Collection
    Dim CodeCollection  As Collection
    
    Dim Parameter()     As String
    Dim Parameters()    As String
    Dim UrlParts(1)     As String
    
    Dim Subdomain       As String
    Dim CodeCount       As Integer
    Dim Index           As Integer
    Dim Item            As Integer
    Dim Value           As String
    Dim FieldCount      As Integer
    Dim Url             As String
    Dim ResponseText    As String
    Dim ValueDate       As Date
    Dim ThisCall        As Date
    Dim IsCurrent       As Boolean
        
    ' Is the current collection of Codes up-to-date?
    IsCurrent = DateDiff("n", LastCall, Now) < UpdatePause
    
    If IsCurrent Then
        ' Return cached codes.
    Else
        ' Retrieve the code pair and add it to the collection of code pairs.
        
        ' Set subdomain to call.
        If CcaApiId = "" Then
            ' Free plan is used.
            Subdomain = FreeSubdomain
        Else
            ' Paid plan is used.
            Subdomain = PaidSubdomain
        End If
        
        ' Define parameter array.
        ' Redim for two dimensions: name, value.
        ReDim Parameter(0 To 0, 0 To 1)
        ' Parameter names.
        Parameter(0, ParameterDetail.Name) = "apiKey"
        ' Parameter values.
        Parameter(0, ParameterDetail.Value) = CcaApiId
        
        ' Assemble parameters.
        ReDim Parameters(LBound(Parameter, 1) To UBound(Parameter, 1))
        For Index = LBound(Parameters) To UBound(Parameters)
            Parameters(Index) = Parameter(Index, 0) & "=" & Parameter(Index, 1)
        Next
        
        ' Assemble URL.
        UrlParts(0) = Replace(ServiceUrl, TempSubdomain, Subdomain)
        UrlParts(1) = Join(Parameters, "&")
        Url = Join(UrlParts, "?")
        ' Uncomment for debugging.
        Debug.Print Url
        
        ' Define a no-result array.
        ' Redim for three dimensions: code, symbol, name.
        ReDim Codes(0, 0 To 2)
        ' Set "not found" return values.
        Codes(0, CodeDetail.Code) = NeutralCode
        Codes(0, CodeDetail.Name) = NeutralName
        Codes(0, CodeDetail.Sign) = NeutralSign
        
        If RetrieveDataResponse(Url, ResponseText) = True Then
            Set DataCollection = CollectJson(ResponseText)
        End If
    
        If DataCollection Is Nothing Then
            ' Error. ResponseText holds the error code.
            ' Optional error handling.
            Select Case ResponseText
                Case HttpStatus.BadRequest
                    ' Typical for invalid api key, or API limit reached.
                Case EmptyResponse
                    ' Invalid currency code.
                Case Else
                    ' Other error.
            End Select
        End If
        
        If Not DataCollection Is Nothing Then
            If DataCollection(RootNodeName)(CollectionItem.Data)(1)(CollectionItem.Name) = ListNodeName Then
                ' The code list was retrieved.
                ' Get count of codes.
                CodeCount = DataCollection(RootNodeName)(CollectionItem.Data)(ListNodeName)(CollectionItem.Data).Count
                ReDim Codes(0 To CodeCount - 1, 0 To 2)
                For Index = 1 To CodeCount
                    ' The code information is a collection.
                    Set CodeCollection = DataCollection(RootNodeName)(CollectionItem.Data)(1)(CollectionItem.Data)(Index)(CollectionItem.Data)
                    FieldCount = CodeCollection.Count
                    ' Fill one array item.
                    For Item = 1 To FieldCount
                        Value = CodeCollection(Item)(CollectionItem.Data)
                        Select Case CodeCollection(Item)(CollectionItem.Name)
                            Case CodeId
                                Codes(Index - 1, CodeDetail.Code) = Value
                            Case CodeName
                                Codes(Index - 1, CodeDetail.Name) = Value
                            Case CodeSymbol
                                Codes(Index - 1, CodeDetail.Sign) = Value
                        End Select
                    Next
                Next
                ' Round the call time down to the start of the update interval.
                ThisCall = CDate(Fix(Now * 24 * 60 / UpdatePause) / (24 * 60 / UpdatePause))
                ' Record hour of retrieval.
                LastCall = ThisCall
            End If
        End If
    End If
    
    CurrencyCodesCca = Codes

End Function

' Retrieve and update the table holding the list of currency codes
' published by Currency Code API.
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function UpdateCurrencyCodes() As Boolean

    ' Table and field names of table holding currency codes.
    Const TableName As String = "CurrencyCode"
    Const Field1    As String = "Code"
    Const Field2    As String = "Name"
    Const Field3    As String = "Symbol"
    Const Field4    As String = "Assigned"
    Const Field5    As String = "Unassigned"
    
    Dim Records     As DAO.Recordset
    
    Dim Codes       As Variant
    Dim Item        As Integer
    Dim Sql         As String
    Dim Criteria    As String
    Dim Unassigned  As Boolean
    
On Error GoTo Err_UpdateCurrencyCodes

    ' Retrieve array of current currency codes.
    Codes = CurrencyCodesCca
    
    Sql = "Select * From " & TableName & ""
    Set Records = CurrentDb.OpenRecordset(Sql)
    
    ' Add new currency codes.
    For Item = LBound(Codes, 1) To UBound(Codes, 1)
        Criteria = "Code = '" & Codes(Item, CodeDetail.Code) & "'"
        Records.FindFirst Criteria
        If Records.NoMatch Then
            ' New currency code.
            Records.AddNew
                Records.Fields(Field1).Value = Codes(Item, CodeDetail.Code)
                Records.Fields(Field2).Value = Codes(Item, CodeDetail.Name)
                Records.Fields(Field3).Value = Codes(Item, CodeDetail.Sign)
                Records.Fields(Field4).Value = Date
            Records.Update
        ElseIf Not IsNull(Records.Fields(Field5).Value) Then
            ' Existing currency code, marked as unassigned.
            ' Reassign.
            Records.Edit
                Records.Fields(Field4).Value = Date
                Records.Fields(Field5).Value = Null
            Records.Update
        End If
    Next
    
    ' Mark retracted currency codes as unassigned.
    Records.MoveFirst
    While Not Records.EOF
        Unassigned = True
        For Item = LBound(Codes, 1) To UBound(Codes, 1)
            If Records.Fields("Code").Value = Codes(Item, CodeDetail.Code) Then
                Unassigned = False
                Exit For
            End If
        Next
        If Unassigned Then
            Records.Edit
                Records.Fields("Unassigned").Value = Date
            Records.Update
        End If
        Records.MoveNext
    Wend
    Records.Close
    
    UpdateCurrencyCodes = True

Exit_UpdateCurrencyCodes:
    Exit Function
    
Err_UpdateCurrencyCodes:
    MsgBox "Error" & Str(Err.Number) & ": " & Err.Description, vbCritical + vbOKOnly, "Update Currency Codes"
    Resume Exit_UpdateCurrencyCodes
    
End Function

' Check if a currency code is one of the listed currency codes
' published by Currency Code API.
'
' 2018-09-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function IsCurrencyCode( _
    ByVal Code As String) _
    As Boolean
    
    ' Table (or query) and field names of table holding currency codes.
    Const TableName As String = "CcaCurrencyCode"
    Const Field1    As String = "Code"
    
    Dim Criteria    As String
    Dim Result      As Boolean
    
    Criteria = Field1 & " = '" & Code & "'"
    
    Result = Not IsNull(DLookup(Field1, TableName, Criteria))
    
    IsCurrencyCode = Result

End Function

