Option Strict On
Option Explicit On

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Text

Public Module SharedHelpers

    Public Sub RequireColumn(dt As DataTable, name As String)
        If Not dt.Columns.Contains(name) Then Throw New ArgumentException($"Missing required column: '{name}'")
    End Sub

    Public Function SafeInt(o As Object) As Integer
        If o Is Nothing Then Return 0
        If TypeOf o Is Integer Then Return CInt(o)
        Dim s As String = o.ToString().Trim()
        Dim v As Integer
        If Integer.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, v) Then Return v
        Return 0
    End Function

    Public Function SafeDbl(o As Object) As Double
        If o Is Nothing Then Return 0
        Dim v As Double
        If Double.TryParse(o.ToString().Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, v) Then Return v
        Return 0
    End Function

    Public Function SafeDate(o As Object) As DateTime
        If o Is Nothing Then Return DateTime.MinValue
        If TypeOf o Is DateTime Then Return CType(o, DateTime)
        Dim d As DateTime
        If DateTime.TryParse(o.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then Return d
        Return DateTime.MinValue
    End Function

    Public Function SafeStr(o As Object) As String
        If o Is Nothing Then Return ""
        Return o.ToString()
    End Function

    Public Function SafeBool(o As Object) As Boolean
        If o Is Nothing Then Return False
        Dim s As String = o.ToString().Trim().ToUpperInvariant()
        Return s = "TRUE" OrElse s = "T" OrElse s = "1" OrElse s = "YES" OrElse s = "Y"
    End Function

    Public Function SafeArray(arr As String(), idx As Integer) As String
        If arr Is Nothing Then Return ""
        If idx < 0 OrElse idx >= arr.Length Then Return ""
        Return If(arr(idx), "")
    End Function

    Public Function IsTruthy(s As String) As Boolean
        If s Is Nothing Then Return False
        Dim u As String = s.Trim().ToUpperInvariant()
        Return u = "1" OrElse u = "TRUE" OrElse u = "T" OrElse u = "YES" OrElse u = "Y"
    End Function

    Public Function Csv(value As String) As String
        If value Is Nothing Then value = ""
        Dim mustQuote As Boolean = value.Contains(","c) OrElse value.Contains(""""c) OrElse value.Contains(ControlChars.Cr) OrElse value.Contains(ControlChars.Lf)
        If value.Contains(""""c) Then value = value.Replace("""", """""")
        If mustQuote Then Return """" & value & """"
        Return value
    End Function

    Public Function FormatDateOrBlank(d As DateTime) As String
        If d = DateTime.MinValue Then Return ""
        Return d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
    End Function

    Public Function GetOrDefault(Of TKey, TValue)(dict As Dictionary(Of TKey, TValue), key As TKey, defaultValue As TValue) As TValue
        If dict Is Nothing Then Return defaultValue
        Dim v As TValue = defaultValue
        If dict.TryGetValue(key, v) Then Return v
        Return defaultValue
    End Function

    Public Function GetOrEmpty(Of TKey)(dict As Dictionary(Of TKey, String), key As TKey) As String
        If dict Is Nothing Then Return ""
        Dim v As String = ""
        If dict.TryGetValue(key, v) Then Return If(v, "")
        Return ""
    End Function

    Public Function GetOrEmptyDate(Of TKey)(dict As Dictionary(Of TKey, DateTime), key As TKey) As String
        If dict Is Nothing Then Return ""
        Dim v As DateTime
        If dict.TryGetValue(key, v) Then Return v.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
        Return ""
    End Function

    Public Function ParseDateDdMmYyyy(s As String) As DateTime
        If String.IsNullOrWhiteSpace(s) Then Return DateTime.MinValue
        Dim formats As String() = {"dd-MM-yyyy", "d-M-yyyy", "dd-M-yyyy", "d-MM-yyyy"}
        Dim dt As DateTime
        If DateTime.TryParseExact(s.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, dt) Then
            Return dt.Date
        End If
        Throw New FormatException("Invalid date: " & s)
    End Function

    Public Function ParseDueAsEndOfDay(o As Object) As DateTime
        Dim s As String = SafeStr(o).Trim()
        If s = "" Then Return DateTime.MinValue

        Dim d As DateTime
        If DateTime.TryParseExact(s,
                                  "dd-MM-yyyy",
                                  CultureInfo.InvariantCulture,
                                  DateTimeStyles.None,
                                  d) Then
            Return d.Date.AddDays(1).AddTicks(-1) ' end of day
        End If

        If DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then
            Return d.Date.AddDays(1).AddTicks(-1)
        End If

        Return DateTime.MinValue
    End Function

End Module