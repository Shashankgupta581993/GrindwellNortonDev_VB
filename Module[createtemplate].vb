Imports System
Imports System.IO
Imports System.Text

' =====================================================================================
'  CSV TEMPLATE GENERATOR (for testing Cycle Builder + Day Assignment)
'
'  What this does:
'   - Creates empty CSV files with ONLY the headers needed by the planner logic.
'   - You can paste data rows later.
'
'  Files created:
'   1) Orders.csv
'   2) MatchingCycle.csv
'   3) Capacity.csv
'   4) DayCapacity.csv  (optional but useful for testing day-cap logic)
'
'  Notes:
'   - Headers MUST match what the planner reads (case-sensitive recommended).
'   - All columns are written as plain CSV headers; data types are not enforced here.
' =====================================================================================

Public Module CsvTemplateGenerator

    Public Sub CreateCyclePlannerCsvTemplates(outputFolder As String)
        If String.IsNullOrWhiteSpace(outputFolder) Then Throw New ArgumentNullException(NameOf(outputFolder))

        ' Ensure folder exists
        Directory.CreateDirectory(outputFolder)

        ' 1) Orders template
        ' Required by logic:
        '   OrderId, WeekKey, DueDate, Priority, EquipType, FiringCycleCode, Qty, UnitTonnage, UnitVolume
        Dim ordersHeaders As String() = {
            "OrderId",
            "WeekKey",
            "DueDate",
            "Priority",
            "EquipType",
            "FiringCycleCode",
            "Qty",
            "UnitTonnage",
            "UnitVolume"
        }
        WriteCsvHeader(Path.Combine(outputFolder, "Orders.csv"), ordersHeaders)

        ' 2) Matching cycle template
        ' Required by logic:
        '   CycleA, CycleB, IsAllowed
        Dim matchingHeaders As String() = {
            "CycleA",
            "CycleB",
            "IsAllowed"
        }
        WriteCsvHeader(Path.Combine(outputFolder, "MatchingCycle.csv"), matchingHeaders)

        ' 3) Capacity template
        ' Required by logic:
        '   EquipType, MaxTonnage, MaxVolume
        ' Optional by logic:
        '   WeekKey (only if capacity varies by week)
        '
        ' IMPORTANT:
        '  - The planner checks if "WeekKey" column exists; if it exists, it will match on weekKey too.
        '  - For simplest testing, you can keep WeekKey present and fill same capacities for each week.
        Dim capacityHeaders As String() = {
            "EquipType",
            "WeekKey",
            "MaxTonnage",
            "MaxVolume"
        }
        WriteCsvHeader(Path.Combine(outputFolder, "Capacity.csv"), capacityHeaders)

        ' 4) Day capacity template (optional)
        ' Used only if you pass DayCapacity.csv path; otherwise greedy load-balance ignores caps.
        ' Required by logic:
        '   DayIndex, MaxTonnage, MaxVolume
        Dim dayCapacityHeaders As String() = {
            "DayIndex",
            "MaxTonnage",
            "MaxVolume"
        }
        WriteCsvHeader(Path.Combine(outputFolder, "DayCapacity.csv"), dayCapacityHeaders)

        ' Optional: create a "README" quick note
        Dim readmePath = Path.Combine(outputFolder, "README_Templates.txt")
        File.WriteAllText(readmePath,
$"CSV Templates created on {DateTime.Now:yyyy-MM-dd HH:mm:ss}
Files:
 - Orders.csv
 - MatchingCycle.csv
 - Capacity.csv
 - DayCapacity.csv

Tips:
 - WeekKey: use a consistent format (e.g., 2025-12-29) and use the SAME value you pass as weekKey in code.
 - DueDate: use ISO format (yyyy-MM-dd) to avoid parsing surprises.
 - EquipType: use TUNNEL or BATCH.
 - IsAllowed: use True/False or 1/0.
 - DayIndex: 1=Mon ... 7=Sun (based on planner mapping).")

    End Sub

    ' Writes a single header row into a CSV file.
    ' If file exists, it will be overwritten to ensure a clean template.
    Private Sub WriteCsvHeader(filePath As String, headers As String())
        If headers Is Nothing OrElse headers.Length = 0 Then Throw New ArgumentException("headers must not be empty.")

        ' Basic CSV escaping for header fields (not strictly needed for these names, but safe)
        Dim escaped = headers.Select(Function(h) CsvEscape(h)).ToArray()
        Dim line = String.Join(",", escaped)

        Using sw As New StreamWriter(filePath, append:=False, encoding:=New UTF8Encoding(encoderShouldEmitUTF8Identifier:=False))
            sw.WriteLine(line)
        End Using
    End Sub

    ' Minimal CSV escape: wrap in quotes if it contains comma or quote; double quotes inside.
    Private Function CsvEscape(value As String) As String
        If value Is Nothing Then Return ""
        Dim mustQuote = value.Contains(",") OrElse value.Contains("""") OrElse value.Contains(vbCr) OrElse value.Contains(vbLf)
        If value.Contains("""") Then value = value.Replace("""", """""")
        If mustQuote Then Return $"""{value}"""
        Return value
    End Function

End Module

' Example call:
