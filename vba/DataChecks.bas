Attribute VB_Name = "DataChecks"
Option Explicit

' Data validation checks for the Returns sheet.
' Rules: missing values, duplicate ReturnID, invalid date, future date,
' value not in the approved lists, Qty must be > 0, StockUnits must be >= 0.

Private Function IsPositiveNumber(ByVal v As Variant) As Boolean
    If IsNumeric(v) Then IsPositiveNumber = (CDbl(v) > 0)
End Function

Private Function IsNonNegativeNumber(ByVal v As Variant) As Boolean
    If IsNumeric(v) Then IsNonNegativeNumber = (CDbl(v) >= 0)
End Function

Private Function IsBlank(ByVal v As Variant) As Boolean
    IsBlank = IsEmpty(v) Or (Trim(CStr(v)) = "")
End Function

Private Sub AddIssue(wi As Worksheet, ByRef n As Long, srcRow As Long, _
                     idText As String, fieldName As String, ruleText As String, _
                     v As Variant, target As Range)
    n = n + 1
    wi.Cells(n, 1).Value = srcRow
    wi.Cells(n, 2).Value = idText
    wi.Cells(n, 3).Value = fieldName
    wi.Cells(n, 4).Value = ruleText
    wi.Cells(n, 5).Value = "'" & CStr(v)
    target.Interior.Color = RGB(255, 199, 206)
End Sub

Public Sub RunDataChecks()
    Dim ws As Worksheet, wl As Worksheet, wi As Worksheet
    Dim lastRow As Long, r As Long, c As Long, n As Long, i As Long
    Dim v As Variant, idText As String
    Dim idRange As Range, regList As Range, prodList As Range, defList As Range
    Dim rules As Variant

    Set ws = ThisWorkbook.Worksheets("Returns")
    Set wl = ThisWorkbook.Worksheets("Lists")
    Set wi = ThisWorkbook.Worksheets("Issues")

    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Set idRange = ws.Range("A2:A" & lastRow)
    Set regList = wl.Range("A2:A" & wl.Cells(wl.Rows.Count, 1).End(xlUp).Row)
    Set prodList = wl.Range("B2:B" & wl.Cells(wl.Rows.Count, 2).End(xlUp).Row)
    Set defList = wl.Range("C2:C" & wl.Cells(wl.Rows.Count, 3).End(xlUp).Row)

    Application.ScreenUpdating = False
    wi.Range("A2:H1000").ClearContents
    ws.Range("A2:G" & lastRow).Interior.Pattern = xlNone
    n = 1

    For r = 2 To lastRow
        idText = Trim(CStr(ws.Cells(r, 1).Value))

        ' 1. Missing values
        For c = 1 To 7
            v = ws.Cells(r, c).Value
            If IsBlank(v) Then
                AddIssue wi, n, r, idText, CStr(ws.Cells(1, c).Value), "Missing value", "", ws.Cells(r, c)
            End If
        Next c

        ' 2. Duplicate ReturnID
        If idText <> "" Then
            If Application.WorksheetFunction.CountIf(idRange, idText) > 1 Then
                AddIssue wi, n, r, idText, "ReturnID", "Duplicate ReturnID", idText, ws.Cells(r, 1)
            End If
        End If

        ' 3. Dates
        v = ws.Cells(r, 2).Value
        If Not IsBlank(v) Then
            If Not IsDate(v) Then
                AddIssue wi, n, r, idText, "ReturnDate", "Invalid date", v, ws.Cells(r, 2)
            ElseIf CDate(v) > Date Then
                AddIssue wi, n, r, idText, "ReturnDate", "Future date", Format(CDate(v), "dd/mm/yyyy"), ws.Cells(r, 2)
            End If
        End If

        ' 4. Approved lists (Region, ProductType, DefectType)
        v = ws.Cells(r, 3).Value
        If Not IsBlank(v) Then
            If Application.WorksheetFunction.CountIf(regList, CStr(v)) = 0 Then _
                AddIssue wi, n, r, idText, "Region", "Not in approved list", v, ws.Cells(r, 3)
        End If
        v = ws.Cells(r, 4).Value
        If Not IsBlank(v) Then
            If Application.WorksheetFunction.CountIf(prodList, CStr(v)) = 0 Then _
                AddIssue wi, n, r, idText, "ProductType", "Not in approved list", v, ws.Cells(r, 4)
        End If
        v = ws.Cells(r, 5).Value
        If Not IsBlank(v) Then
            If Application.WorksheetFunction.CountIf(defList, CStr(v)) = 0 Then _
                AddIssue wi, n, r, idText, "DefectType", "Not in approved list", v, ws.Cells(r, 5)
        End If

        ' 5. Quantity must be a positive number
        v = ws.Cells(r, 6).Value
        If Not IsBlank(v) Then
            If Not IsPositiveNumber(v) Then _
                AddIssue wi, n, r, idText, "Qty", "Quantity must be a positive number", v, ws.Cells(r, 6)
        End If

        ' 6. Stock units must be a number >= 0
        v = ws.Cells(r, 7).Value
        If Not IsBlank(v) Then
            If Not IsNonNegativeNumber(v) Then _
                AddIssue wi, n, r, idText, "StockUnits", "StockUnits must be a number >= 0", v, ws.Cells(r, 7)
        End If
    Next r

    ' Summary by rule
    rules = Array("Missing value", "Duplicate ReturnID", "Invalid date", "Future date", _
                  "Not in approved list", "Quantity must be a positive number", _
                  "StockUnits must be a number >= 0")
    wi.Cells(1, 7).Value = "Rule"
    wi.Cells(1, 8).Value = "Issues"
    wi.Range("G1:H1").Font.Bold = True
    For i = 0 To UBound(rules)
        wi.Cells(i + 2, 7).Value = rules(i)
        If n >= 2 Then
            wi.Cells(i + 2, 8).Value = Application.WorksheetFunction.CountIf(wi.Range("D2:D" & n), rules(i))
        Else
            wi.Cells(i + 2, 8).Value = 0
        End If
    Next i
    wi.Cells(UBound(rules) + 3, 7).Value = "Total issues"
    wi.Cells(UBound(rules) + 3, 8).Value = n - 1
    wi.Cells(UBound(rules) + 3, 7).Font.Bold = True
    wi.Cells(UBound(rules) + 3, 8).Font.Bold = True
    wi.Columns("A:H").AutoFit

    Application.ScreenUpdating = True
    wi.Activate
    MsgBox (n - 1) & " issues found in " & (lastRow - 1) & " rows. See the Issues sheet.", vbInformation, "Data checks"
End Sub

Public Sub ClearHighlights()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Returns")
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ws.Range("A2:G" & lastRow).Interior.Pattern = xlNone
End Sub
