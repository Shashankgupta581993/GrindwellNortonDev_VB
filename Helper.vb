Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports System.Data


Public Class Helper

    '===========================================================
    ' ProcessMappingMatrix
    '
    ' Loads Process_Mapping.csv as a matrix:
    '   Rows   = operations (Operation_number and/or Operation_name)
    '   Cols   = cycle types (e.g., "150 VT", "PTK", etc.)
    '   Cells  = numeric value (generic "days")
    '
    ' Primary intent:
    '   GetDays(cycleType, operationNumber) -> days
    '
    ' Practical reality:
    '   Your current file has Operation_number header but values may be blank.
    '   So we ALSO store a fallback index by Operation_name.
    '
    ' IMPORTANT:
    ' - This class DOES NOT apply working-day calendars.
    ' - Values are returned as generic days (Decimal).
    '===========================================================


    Public ReadOnly Property RawTable As DataTable

        ' Preferred index:
        '   operationNumber -> (cycle -> days)
        Private ReadOnly _opNoToCycleDays As Dictionary(Of String, Dictionary(Of String, Decimal))

        ' Fallback index:
        '   operationName -> (cycle -> days)
        Private ReadOnly _opNameToCycleDays As Dictionary(Of String, Dictionary(Of String, Decimal))

        ' Optional aliases for cycle names:
        '   requestCycleNorm -> actualCycleNormInFile
        Private ReadOnly _cycleAliases As Dictionary(Of String, String)

        Private ReadOnly _strict As Boolean

        Private Sub New(dt As DataTable,
                    opNoToCycleDays As Dictionary(Of String, Dictionary(Of String, Decimal)),
                    opNameToCycleDays As Dictionary(Of String, Dictionary(Of String, Decimal)),
                    cycleAliases As Dictionary(Of String, String),
                    strict As Boolean)
            RawTable = dt
            _opNoToCycleDays = opNoToCycleDays
            _opNameToCycleDays = opNameToCycleDays
            _cycleAliases = cycleAliases
            _strict = strict
        End Sub

        '-----------------------------------------------------------
        ' Preferred lookup: by Operation Number (string or numeric text)
        ' Throws if missing when strict=True.
        '-----------------------------------------------------------
        Public Function GetDays(ByVal cycleType As String, ByVal operationNumber As String) As Decimal
            Dim days As Decimal
            If TryGetDays(cycleType, operationNumber, days) Then
                Return days
            End If

            If _strict Then
                Throw New Exception("Process mapping missing for Operation_number='" & operationNumber &
                                "', Cycle='" & cycleType & "'.")
            End If

            Return 0D
        End Function

        Public Function TryGetDays(ByVal cycleType As String,
                               ByVal operationNumber As String,
                               ByRef days As Decimal) As Boolean

            days = 0D

            Dim opNoNorm As String = NormalizeOpNumber(operationNumber)
            If String.IsNullOrWhiteSpace(opNoNorm) Then Return False

            Dim cycleNorm As String = ResolveCycleKey(cycleType)

            If Not _opNoToCycleDays.ContainsKey(opNoNorm) Then Return False
            Dim inner = _opNoToCycleDays(opNoNorm)

            If Not inner.ContainsKey(cycleNorm) Then Return False
            days = inner(cycleNorm)
            Return True
        End Function

        '-----------------------------------------------------------
        ' Fallback lookup: by Operation Name (useful while op numbers are blank)
        ' Throws if missing when strict=True.
        '-----------------------------------------------------------
        Public Function GetDaysByName(ByVal cycleType As String, ByVal operationName As String) As Decimal
            Dim days As Decimal
            If TryGetDaysByName(cycleType, operationName, days) Then
                Return days
            End If

            If _strict Then
                Throw New Exception("Process mapping missing for Operation_name='" & operationName &
                                "', Cycle='" & cycleType & "'.")
            End If

            Return 0D
        End Function

        Public Function TryGetDaysByName(ByVal cycleType As String,
                                     ByVal operationName As String,
                                     ByRef days As Decimal) As Boolean

            days = 0D

            Dim opNameNorm As String = NormalizeKey(operationName)
            If String.IsNullOrWhiteSpace(opNameNorm) Then Return False

            Dim cycleNorm As String = ResolveCycleKey(cycleType)

            If Not _opNameToCycleDays.ContainsKey(opNameNorm) Then Return False
            Dim inner = _opNameToCycleDays(opNameNorm)

            If Not inner.ContainsKey(cycleNorm) Then Return False
            days = inner(cycleNorm)
            Return True
        End Function

    '===========================================================
    ' Loader
    '===========================================================
    Public Shared Function LoadFromCsv(ByVal filePath As String,
                                       Optional ByVal strict As Boolean = True,
                                       Optional ByVal idColumnName As String = "Id",
                                       Optional ByVal operationNameColumnName As String = "Operation_name",
                                       Optional ByVal operationNumberColumnName As String = "Operation_number",
                                       Optional ByVal headerAliases As Dictionary(Of String, String) = Nothing) As Helper
        If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
        If Not File.Exists(filePath) Then Throw New FileNotFoundException("Process mapping file not found: " & filePath)

        ' Build alias map (normalized)
        Dim aliases As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If headerAliases IsNot Nothing Then
            For Each kvp In headerAliases
                Dim k As String = NormalizeKey(kvp.Key)
                Dim v As String = NormalizeKey(kvp.Value)
                If k.Length > 0 AndAlso v.Length > 0 Then
                    aliases(k) = v
                End If
            Next
        End If

        ' Read CSV into DataTable (string typed; we parse as needed)
        Dim dt As New DataTable("Process_Mapping")

        Using parser As New TextFieldParser(filePath)
            parser.TextFieldType = FieldType.Delimited
            parser.SetDelimiters(","c)
            parser.HasFieldsEnclosedInQuotes = True
            parser.TrimWhiteSpace = False

            If parser.EndOfData Then Throw New Exception("Process mapping CSV is empty: " & filePath)

            Dim headers As String() = parser.ReadFields()
            If headers Is Nothing OrElse headers.Length = 0 Then
                Throw New Exception("Process mapping CSV has no header row: " & filePath)
            End If

            For Each h In headers
                Dim colName As String = If(h, String.Empty).Trim()
                If String.IsNullOrWhiteSpace(colName) Then Throw New Exception("Blank header found in process mapping CSV.")
                dt.Columns.Add(colName, GetType(String))
            Next

            While Not parser.EndOfData
                Dim fields As String() = parser.ReadFields()
                If fields Is Nothing Then Continue While
                If fields.Length <> dt.Columns.Count Then
                    Throw New Exception("Row has " & fields.Length & " columns but header has " & dt.Columns.Count & ". File=" & filePath)
                End If

                Dim dr As DataRow = dt.NewRow()
                For i As Integer = 0 To fields.Length - 1
                    dr(i) = If(fields(i), String.Empty)
                Next
                dt.Rows.Add(dr)
            End While
        End Using

        ' Validate key columns
        If Not dt.Columns.Contains(operationNameColumnName) Then
            Throw New Exception("Missing required column: " & operationNameColumnName)
        End If
        If Not dt.Columns.Contains(operationNumberColumnName) Then
            Throw New Exception("Missing required column: " & operationNumberColumnName)
        End If

        ' Identify cycle columns: everything except Id/op-name/op-number
        Dim cycleCols As New List(Of String)()
        For Each col As DataColumn In dt.Columns
            Dim name As String = col.ColumnName
            If name.Equals(idColumnName, StringComparison.OrdinalIgnoreCase) Then Continue For
            If name.Equals(operationNameColumnName, StringComparison.OrdinalIgnoreCase) Then Continue For
            If name.Equals(operationNumberColumnName, StringComparison.OrdinalIgnoreCase) Then Continue For
            cycleCols.Add(name)
        Next

        If cycleCols.Count = 0 Then
            Throw New Exception("No cycle columns detected in process mapping CSV. Check headers.")
        End If

        Dim opNoToCycleDays As New Dictionary(Of String, Dictionary(Of String, Decimal))(StringComparer.OrdinalIgnoreCase)
        Dim opNameToCycleDays As New Dictionary(Of String, Dictionary(Of String, Decimal))(StringComparer.OrdinalIgnoreCase)

        ' Load each row
        For Each row As DataRow In dt.Rows

            Dim opNameRaw As String = row(operationNameColumnName).ToString()
            Dim opNameNorm As String = NormalizeKey(opNameRaw)

            Dim opNoRaw As String = row(operationNumberColumnName).ToString()
            Dim opNoNorm As String = NormalizeOpNumber(opNoRaw)

            ' Prepare dictionaries if keys are present
            If opNameNorm.Length > 0 AndAlso Not opNameToCycleDays.ContainsKey(opNameNorm) Then
                opNameToCycleDays(opNameNorm) = New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
            End If

            If opNoNorm.Length > 0 AndAlso Not opNoToCycleDays.ContainsKey(opNoNorm) Then
                opNoToCycleDays(opNoNorm) = New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
            End If

            ' Parse numeric values for each cycle column
            For Each cc In cycleCols
                Dim cellRaw As String = row(cc).ToString().Trim()
                If String.IsNullOrWhiteSpace(cellRaw) Then Continue For ' missing value -> ignore for now

                Dim v As Decimal
                If Not Decimal.TryParse(cellRaw, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then
                    Throw New Exception("Non-numeric value '" & cellRaw & "' in process mapping. Operation_name='" &
                                        opNameRaw & "', CycleColumn='" & cc & "'.")
                End If

                Dim cycleNorm As String = NormalizeKey(cc)

                ' Store by op-name (fallback index)
                If opNameNorm.Length > 0 Then
                    opNameToCycleDays(opNameNorm)(cycleNorm) = v
                End If

                ' Store by op-number (preferred index; only if provided)
                If opNoNorm.Length > 0 Then
                    opNoToCycleDays(opNoNorm)(cycleNorm) = v
                End If
            Next

        Next

        Return New Helper(dt, opNoToCycleDays, opNameToCycleDays, aliases, strict)

    End Function

    '===========================================================
    ' Cycle key resolution:
    ' 1) normalize requested cycle
    ' 2) if alias exists, redirect to the aliased header
    '===========================================================
    Private Function ResolveCycleKey(ByVal cycleType As String) As String
            Dim c As String = NormalizeKey(cycleType)
            If c.Length = 0 Then Return c

            If _cycleAliases IsNot Nothing AndAlso _cycleAliases.ContainsKey(c) Then
                Return _cycleAliases(c)
            End If

            Return c
        End Function

        '===========================================================
        ' Normalize generic keys (trim + collapse whitespace)
        '===========================================================
        Private Shared Function NormalizeKey(ByVal raw As String) As String
            If raw Is Nothing Then Return String.Empty
            Dim t As String = raw.Trim()
            If t.Length = 0 Then Return String.Empty
            Return Regex.Replace(t, "\s+", " ")
        End Function

        '===========================================================
        ' Normalize operation numbers:
        ' - Accept numeric strings like "300", "300.0"
        ' - Return integer-like string (e.g., "300")
        '===========================================================
        Private Shared Function NormalizeOpNumber(ByVal raw As String) As String
            Dim t As String = NormalizeKey(raw)
            If t.Length = 0 Then Return String.Empty

            ' Try integer first
            Dim i As Integer
            If Integer.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, i) Then
                Return i.ToString(CultureInfo.InvariantCulture)
            End If

            ' Try decimal (e.g., "300.0") and coerce to int if whole
            Dim d As Decimal
            If Decimal.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, d) Then
                Dim whole As Decimal = Decimal.Truncate(d)
                If d = whole Then
                    Return CInt(whole).ToString(CultureInfo.InvariantCulture)
                End If
            End If

            ' If it isn't numeric, keep the normalized string (still usable, but not ideal)
            Return t
        End Function

End Class