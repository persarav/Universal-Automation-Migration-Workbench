"""Generate a compact, modular UiPath Windows project for the SYNC VBA process.

The generated project uses native workbook activities for file IO and a small
number of Invoke Code blocks for dense, deterministic DataTable business rules.
SAP GUI transactions remain explicit capture-required adapters because their
selectors are environment-specific and must be captured in the target Studio.
"""
from __future__ import annotations
import csv, io, json, re, html, textwrap
from pathlib import Path
from typing import Any

PROJECT_JSON = {
    "name": "UniversalMigration_SYNC_ClubCheck",
    "projectVersion": "1.0.0",
    "description": "Modular UiPath migration of the SYNC New Club STO/PO workflow",
    "main": "Main.xaml",
    "dependencies": {
        "UiPath.System.Activities": "[24.10.8]",
        "UiPath.Excel.Activities": "[2.24.5]"
    },
    "schemaVersion": "4.0",
    "studioVersion": "24.10.8",
    "projectType": "Workflow",
    "runtimeOptions": {
        "autoDispose": False,
        "isPausable": True,
        "isAttended": False,
        "requiresUserInteraction": False,
        "supportsPersistence": False,
        "excludedLoggedData": ["Private:*"],
        "executionType": "Workflow",
        "readyForPiP": False,
        "startsInPiP": False
    },
    "designOptions": {
        "projectProfile": "Development",
        "outputType": "Process",
        "processOptions": {"ignoredFiles": []},
        "libraryOptions": {"includeOriginalXaml": False, "privateWorkflows": []},
        "fileInfoCollection": []
    },
    "expressionLanguage": "VisualBasic",
    "targetFramework": "Windows",
    "isTemplate": False,
    "templateProjectData": {},
    "publishData": {}
}

COMMON_VB = r'''Dim normalizeText As System.Func(Of Object, String) = Function(v)
    If v Is Nothing OrElse v Is System.DBNull.Value Then Return ""
    Return System.Text.RegularExpressions.Regex.Replace(Convert.ToString(v).Replace(ChrW(160), " ").Trim(), "\s+", " ")
End Function
Dim normalizeHeader As System.Func(Of Object, String) = Function(v)
    Return System.Text.RegularExpressions.Regex.Replace(normalizeText(v).ToUpperInvariant(), "[^A-Z0-9]", "")
End Function
Dim articleKey As System.Func(Of Object, String) = Function(v)
    Dim s As String = normalizeText(v)
    If s.Length = 0 Then Return ""
    If System.Text.RegularExpressions.Regex.IsMatch(s, "^\d+$") Then
        s = s.TrimStart("0"c)
        If s.Length = 0 Then s = "0"
    End If
    Return s.ToUpperInvariant()
End Function
Dim findColumn As System.Func(Of System.Data.DataTable, String(), Integer) = Function(dt, aliases)
    If dt Is Nothing Then Return -1
    For Each aliasName As String In aliases
        Dim wanted As String = normalizeHeader(aliasName)
        For c As Integer = 0 To dt.Columns.Count - 1
            If normalizeHeader(dt.Columns(c).ColumnName) = wanted Then Return c
        Next
    Next
    Return -1
End Function
Dim numberValue As System.Func(Of Object, Decimal) = Function(v)
    If v Is Nothing OrElse v Is System.DBNull.Value Then Return 0D
    Dim n As Decimal = 0D
    Dim s As String = normalizeText(v).Replace(",", "")
    If Decimal.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, n) Then Return n
    Return 0D
End Function
Dim dateValue As System.Func(Of Object, DateTime) = Function(v)
    If v Is Nothing OrElse v Is System.DBNull.Value Then Return DateTime.MinValue
    If TypeOf v Is DateTime Then Return CDate(v)
    Dim serial As Double = 0
    Dim s As String = normalizeText(v)
    If Double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, serial) AndAlso serial > 20000 AndAlso serial < 80000 Then
        Try
            Return DateTime.FromOADate(serial)
        Catch
        End Try
    End If
    Dim parsed As DateTime
    If DateTime.TryParse(s, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.AssumeLocal, parsed) Then Return parsed
    Return DateTime.MinValue
End Function
Dim formatQty As System.Func(Of Decimal, String) = Function(v) v.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)
Dim formatDate As System.Func(Of DateTime, String) = Function(v) If(v = DateTime.MinValue, "", v.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture))
Dim fourDigit As System.Func(Of Object, String) = Function(v)
    If v Is Nothing OrElse v Is System.DBNull.Value Then Return ""
    ' Match dynSiteNo: read only the leading numeric site from values such as "0276 LECANTO".
    Dim m As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(normalizeText(v), "^\d+")
    If Not m.Success Then Return ""
    Dim n As Integer = 0
    If Integer.TryParse(m.Value, n) Then Return n.ToString("0000", System.Globalization.CultureInfo.InvariantCulture)
    Return m.Value
End Function
Dim isDeleted As System.Func(Of Object, Boolean) = Function(v)
    Dim s As String = normalizeText(v).ToUpperInvariant()
    Return Not (s = "" OrElse s = "#" OrElse s = "-" OrElse s = "--" OrElse s = "." OrElse s = "0" OrElse s = "N" OrElse s = "NO" OrElse s = "FALSE" OrElse s = "NULL" OrElse s = "N/A" OrElse s = "NA" OrElse s = "NOT ASSIGNED")
End Function
Dim inScope As System.Func(Of Object, String, Boolean) = Function(v, selection)
    Dim scopeText As String = normalizeText(selection)
    If scopeText.Length = 0 Then Return True
    Dim current As String = normalizeText(v)
    If current.Length = 0 Then Return False
    Dim parts = scopeText.Split(New Char() {"|"c, ";"c, ","c, ChrW(10), ChrW(13)}, StringSplitOptions.RemoveEmptyEntries)
    For Each part As String In parts
        If String.Equals(current, part.Trim(), StringComparison.OrdinalIgnoreCase) Then Return True
    Next
    Return False
End Function
'''

STO_VB = COMMON_VB + r'''
If dtClub Is Nothing OrElse dtSto Is Nothing Then Throw New Exception("New Club data and STO Export DataTables are required.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article", "Article Number", "Material", "SKU"})
Dim cQty As Integer = findColumn(dtClub, New String() {"Initial Set Quantity", "Club Qty", "Set Quantity"})
Dim cDiv As Integer = findColumn(dtClub, New String() {"Division", "Art_AH2", "AH2"})
Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review", "STO Check Date", "STO Date"})
Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments", "STO Status", "STO Comment"})
If cArt < 0 OrElse cQty < 0 OrElse cDate < 0 OrElse cComment < 0 Then Throw New Exception("Required New Club columns missing: Article, Initial Set Quantity, Date of Last Review or CSC2 Comments.")
Dim eArt As Integer = findColumn(dtSto, New String() {"Article", "Article Number", "Material", "SKU"})
Dim eSto As Integer = findColumn(dtSto, New String() {"STO", "STO Number", "Purchasing Document", "PO Number", "Document Number"})
Dim eQty As Integer = findColumn(dtSto, New String() {"Quantity in SKU", "Qty in SKU", "Quantity"})
Dim eDue As Integer = findColumn(dtSto, New String() {"Due Date", "Delivery Date", "Scheduled Delivery Date", "Requested Delivery Date", "Stat.-Rel. Del. Date"})
Dim eDel As Integer = findColumn(dtSto, New String() {"Deletion Indicator", "Deletion Ind", "Del Ind", "Delete Flag"})
Dim eSite As Integer = findColumn(dtSto, New String() {"Site", "Plant", "Receiving Site"})
Dim eMark As Integer = findColumn(dtSto, New String() {"PO history/release documentation", "PO history release documentation"})
If eSite < 0 Then eSite = 0
If eArt < 0 OrElse eSto < 0 OrElse eQty < 0 OrElse eDue < 0 Then Throw New Exception("STO Export is missing Article, Purchasing Document, Quantity in SKU or delivery date.")
If eDel < 0 Then Throw New Exception("STO Export is missing the Deletion Indicator column; safety check cannot continue.")
If eMark < 0 AndAlso dtSto.Columns.Count >= 17 Then eMark = 16
If eMark < 0 Then Throw New Exception("STO Export release-documentation marker column was not found.")

Dim clubRows As Integer = 0, siteRows As Integer = 0, deletedSiteRows As Integer = 0
For Each er As System.Data.DataRow In dtSto.Rows
    If er.RowState = System.Data.DataRowState.Deleted Then Continue For
    Dim hasAny As Boolean = False
    For Each cell As Object In er.ItemArray
        If normalizeText(cell).Length > 0 Then hasAny = True : Exit For
    Next
    If Not hasAny Then Continue For
    clubRows += 1
    If fourDigit(er(eSite)) = fourDigit(clubNo) Then
        siteRows += 1
        If isDeleted(er(eDel)) Then deletedSiteRows += 1
    End If
Next
If clubRows > 0 AndAlso siteRows = 0 Then Throw New Exception("STO Export Site does not match the Action Sheet New Club Number. No data was changed.")
If siteRows > 0 AndAlso deletedSiteRows >= siteRows Then Throw New Exception("All STO rows for this club have a deletion indicator. No data was changed.")

Dim totalByArt As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
Dim deliveredByArt As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
Dim openByArt As New Dictionary(Of String, Dictionary(Of String, Object()))(StringComparer.OrdinalIgnoreCase)
Dim deliveredLines As New Dictionary(Of String, Dictionary(Of String, Decimal))(StringComparer.OrdinalIgnoreCase)
Dim latestDeliveredDue As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
For Each er As System.Data.DataRow In dtSto.Rows
    If fourDigit(er(eSite)) <> fourDigit(clubNo) OrElse isDeleted(er(eDel)) Then Continue For
    Dim key As String = articleKey(er(eArt))
    If key.Length = 0 Then Continue For
    Dim q As Decimal = numberValue(er(eQty))
    Dim due As DateTime = dateValue(er(eDue))
    Dim stoNo As String = normalizeText(er(eSto))
    Dim marked As Boolean = normalizeText(er(eMark)).IndexOf("PO history/release documentation", StringComparison.OrdinalIgnoreCase) >= 0
    If Not totalByArt.ContainsKey(key) Then totalByArt(key) = 0D
    totalByArt(key) += q
    If marked Then
        If Not deliveredByArt.ContainsKey(key) Then deliveredByArt(key) = 0D
        deliveredByArt(key) += q
        If Not deliveredLines.ContainsKey(key) Then deliveredLines(key) = New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
        If Not deliveredLines(key).ContainsKey(stoNo) Then deliveredLines(key)(stoNo) = 0D
        deliveredLines(key)(stoNo) += q
        If due <> DateTime.MinValue AndAlso (Not latestDeliveredDue.ContainsKey(key) OrElse due > latestDeliveredDue(key)) Then latestDeliveredDue(key) = due
    Else
        If Not openByArt.ContainsKey(key) Then openByArt(key) = New Dictionary(Of String, Object())(StringComparer.OrdinalIgnoreCase)
        Dim lineKey As String = stoNo & ChrW(1) & If(due = DateTime.MinValue, "0", due.ToString("yyyyMMdd"))
        If Not openByArt(key).ContainsKey(lineKey) Then
            openByArt(key)(lineKey) = New Object() {stoNo, q, due}
        Else
            openByArt(key)(lineKey)(1) = CDec(openByArt(key)(lineKey)(1)) + q
        End If
    End If
Next

outUpdated = 0
outOverdue = 0
outMissing = 0
For Each cr As System.Data.DataRow In dtClub.Rows
    Dim key As String = articleKey(cr(cArt))
    If key.Length = 0 Then Continue For
    If Not inScope(If(cDiv >= 0, cr(cDiv), ""), divisionScope) Then Continue For
    If Not totalByArt.ContainsKey(key) Then
        outMissing += 1
        Continue For
    End If
    Dim need As Decimal = numberValue(cr(cQty))
    Dim totalQty As Decimal = totalByArt(key)
    Dim doneQty As Decimal = If(deliveredByArt.ContainsKey(key), deliveredByArt(key), 0D)
    Dim doneDue As DateTime = If(latestDeliveredDue.ContainsKey(key), latestDeliveredDue(key), DateTime.MinValue)
    Dim openLines As New System.Collections.Generic.List(Of Object())
    If openByArt.ContainsKey(key) Then
        For Each entry As KeyValuePair(Of String, Object()) In openByArt(key)
            openLines.Add(entry.Value)
        Next
    End If
    openLines.Sort(Function(a, b)
        Dim da As DateTime = CDate(a(2)), db As DateTime = CDate(b(2))
        If da = DateTime.MinValue AndAlso db <> DateTime.MinValue Then Return 1
        If db = DateTime.MinValue AndAlso da <> DateTime.MinValue Then Return -1
        Return DateTime.Compare(da, db)
    End Function)
    Dim comment As String = "STO Manifested"
    Dim hasOverdue As Boolean = False
    If need > 0D AndAlso doneQty >= need Then
        comment = "STO Manifested"
    ElseIf need > 0D AndAlso doneQty > 0D Then
        Dim deliveredParts As New System.Collections.Generic.List(Of String)
        If deliveredLines.ContainsKey(key) Then
            For Each de As KeyValuePair(Of String, Decimal) In deliveredLines(key)
                deliveredParts.Add("STO " & de.Key & " qty " & formatQty(de.Value))
            Next
        End If
        Dim openParts As New System.Collections.Generic.List(Of String)
        Dim windowQty As Decimal = 0D, windowCount As Integer = 0
        Dim nextSto As String = "", nextDue As DateTime = DateTime.MaxValue
        For Each ln As Object() In openLines
            Dim due As DateTime = CDate(ln(2)), q As Decimal = CDec(ln(1)), stoNo As String = CStr(ln(0))
            Dim part As String = "STO " & stoNo & " qty " & formatQty(q)
            If due <> DateTime.MinValue Then
                part &= " due on " & formatDate(due)
                If due.Date < todayDate.Date AndAlso (doneDue = DateTime.MinValue OrElse due > doneDue) Then part &= " (passed, not delivered)" : hasOverdue = True
                If due.Date >= todayDate.Date AndAlso due.Date <= truckDate.Date Then
                    windowQty += q : windowCount += 1
                    If due < nextDue Then nextDue = due : nextSto = stoNo
                End If
            End If
            openParts.Add(part)
        Next
        If windowCount > 0 AndAlso doneQty + windowQty >= need Then
            comment = "Partially manifested - next incoming STO " & nextSto & " due " & formatDate(nextDue) & ", will manifest"
            If windowCount > 1 Then comment &= " (+" & (windowCount - 1).ToString() & " more STO)"
        Else
            comment = "Partial set quantity delivered " & formatQty(doneQty) & " out of " & formatQty(need)
            If deliveredParts.Count > 0 Then comment &= "; delivered " & String.Join(", ", deliveredParts)
            If openParts.Count > 0 Then comment &= "; open " & String.Join(", ", openParts)
            Dim manualQty As Decimal = need - totalQty
            If manualQty > 0D Then comment &= "; need manual STO for " & formatQty(manualQty)
        End If
    ElseIf totalQty < need Then
        comment = "Need manual STO"
        If need - totalQty > 0D Then comment &= " for remaining " & formatQty(need - totalQty)
        Dim openParts As New System.Collections.Generic.List(Of String)
        For Each ln As Object() In openLines
            Dim due As DateTime = CDate(ln(2)), q As Decimal = CDec(ln(1)), stoNo As String = CStr(ln(0))
            Dim part As String = "STO " & stoNo & " qty " & formatQty(q)
            If due <> DateTime.MinValue Then
                part &= " due on " & formatDate(due)
                If due.Date < todayDate.Date AndAlso (doneDue = DateTime.MinValue OrElse due > doneDue) Then part &= " (passed, not delivered)" : hasOverdue = True
            End If
            openParts.Add(part)
        Next
        If openParts.Count > 0 Then comment &= "; open " & String.Join(", ", openParts)
    ElseIf openLines.Count > 0 Then
        Dim overdueParts As New System.Collections.Generic.List(Of String)
        Dim upcomingParts As New System.Collections.Generic.List(Of String)
        For Each ln As Object() In openLines
            Dim due As DateTime = CDate(ln(2)), stoNo As String = CStr(ln(0))
            If due <> DateTime.MinValue AndAlso due.Date < todayDate.Date Then
                If doneDue = DateTime.MinValue OrElse due > doneDue Then
                    overdueParts.Add("STO " & stoNo & " Due date " & formatDate(due) & " is passed STO not delivered")
                    hasOverdue = True
                End If
            Else
                upcomingParts.Add(If(due = DateTime.MinValue, "STO " & stoNo, "STO " & stoNo & " Due " & formatDate(due)))
            End If
        Next
        If overdueParts.Count = 0 Then
            comment = If(upcomingParts.Count = 0, "STO Manifested", upcomingParts(0) & If(upcomingParts.Count > 1, " (+" & (upcomingParts.Count - 1).ToString() & " more)", ""))
        Else
            comment = String.Join("; ", overdueParts)
            If upcomingParts.Count > 0 Then comment &= "; Next: " & upcomingParts(0) & If(upcomingParts.Count > 1, " (+" & (upcomingParts.Count - 1).ToString() & " more)", "")
        End If
    End If
    cr(cComment) = comment
    cr(cDate) = todayDate.Date
    outUpdated += 1
    If hasOverdue Then outOverdue += 1
Next
'''

PO_VB = COMMON_VB + r'''
If dtClub Is Nothing OrElse dtPo Is Nothing Then Throw New Exception("New Club data and PO Export DataTables are required.")
If truckDate = DateTime.MinValue Then Throw New Exception("Last truck date is required and must be valid before PO Check.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article", "Article Number", "Material"})
Dim cGen As Integer = findColumn(dtClub, New String() {"Generic Article", "Generic Material"})
Dim cDiv As Integer = findColumn(dtClub, New String() {"Division", "Art_AH2", "AH2"})
Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review", "STO Check Date", "STO Date"})
Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments", "STO Status", "STO Comment"})
If cArt < 0 OrElse cComment < 0 OrElse cDate < 0 Then Throw New Exception("New Club data is missing Article, CSC2 Comments or Date of Last Review.")
Dim eArt As Integer = findColumn(dtPo, New String() {"Article", "Material"})
Dim ePo As Integer = findColumn(dtPo, New String() {"Purchasing Document"})
Dim eSite As Integer = findColumn(dtPo, New String() {"Site", "Plant", "Receiving Site"})
Dim eDel As Integer = findColumn(dtPo, New String() {"Deletion Indicator", "Deletion Ind", "Del Ind", "Delete Flag"})
Dim eDue As Integer = findColumn(dtPo, New String() {"Delivery Date", "Due Date", "Scheduled Delivery Date", "Promise Date"})
If eArt < 0 OrElse ePo < 0 OrElse eDel < 0 Then Throw New Exception("PO Export is missing Article, Purchasing Document or Deletion Indicator.")
If eSite < 0 Then eSite = 0
If eDue < 0 Then
    For c As Integer = 0 To dtPo.Columns.Count - 1
        Dim h As String = normalizeHeader(dtPo.Columns(c).ColumnName)
        If h.Contains("DELIVER") OrElse h.Contains("PROMISE") OrElse h.Contains("DUE") OrElse h = "DATE" Then eDue = c : Exit For
    Next
End If
If eDue < 0 Then Throw New Exception("PO delivery date column could not be resolved; configure the correct export header before running PO Check.")
Dim dcRows As Integer = 0
For Each er As System.Data.DataRow In dtPo.Rows
    If fourDigit(er(eSite)) = fourDigit(dcNo) Then dcRows += 1
Next
If dtPo.Rows.Count > 0 AndAlso dcRows = 0 Then Throw New Exception("PO Export Site does not match the Action Sheet DC Number. No New Club comments were changed.")

Dim poByArticle As New Dictionary(Of String, System.Collections.Generic.List(Of Object()))(StringComparer.OrdinalIgnoreCase)
For Each er As System.Data.DataRow In dtPo.Rows
    If isDeleted(er(eDel)) Then Continue For
    If fourDigit(er(eSite)) <> fourDigit(dcNo) Then Continue For
    Dim due As DateTime = dateValue(er(eDue))
    If due = DateTime.MinValue OrElse due.Date < todayDate.Date OrElse due.Date > truckDate.Date Then Continue For
    Dim key As String = articleKey(er(eArt))
    Dim poNo As String = normalizeText(er(ePo))
    If key.Length = 0 OrElse poNo.Length = 0 Then Continue For
    If Not poByArticle.ContainsKey(key) Then poByArticle(key) = New System.Collections.Generic.List(Of Object())
    poByArticle(key).Add(New Object() {poNo, due})
Next
For Each kv As KeyValuePair(Of String, System.Collections.Generic.List(Of Object())) In poByArticle
    kv.Value.Sort(Function(a,b) DateTime.Compare(CDate(a(1)), CDate(b(1))))
Next
outUpdated = 0
outNoMatch = 0
outPreserved = 0
For Each cr As System.Data.DataRow In dtClub.Rows
    If Not inScope(If(cDiv >= 0, cr(cDiv), ""), divisionScope) Then
        outPreserved += 1 : Continue For
    End If
    Dim oldComment As String = normalizeText(cr(cComment))
    Dim eligible As Boolean = oldComment.Length = 0 OrElse oldComment.StartsWith("PO ", StringComparison.OrdinalIgnoreCase)
    If Not eligible Then outPreserved += 1 : Continue For
    Dim key As String = articleKey(cr(cArt))
    If key.Length = 0 Then Continue For
    If Not poByArticle.ContainsKey(key) AndAlso cGen >= 0 Then
        Dim genericKey As String = articleKey(cr(cGen))
        If poByArticle.ContainsKey(genericKey) Then key = genericKey
    End If
    If poByArticle.ContainsKey(key) Then
        Dim parts As New System.Collections.Generic.List(Of String)
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each p As Object() In poByArticle(key)
            Dim one As String = "PO " & CStr(p(0)) & " Due On " & formatDate(CDate(p(1)))
            If Not seen.Contains(one) Then parts.Add(one) : seen.Add(one)
        Next
        cr(cComment) = String.Join(" | ", parts)
        cr(cDate) = todayDate.Date
        outUpdated += 1
    ElseIf oldComment.Length = 0 Then
        outNoMatch += 1
    End If
Next
'''

OTHER_VB = COMMON_VB + r'''
If dtClub Is Nothing Then Throw New Exception("New Club data DataTable is required.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article", "Article Number"})
Dim cGen As Integer = findColumn(dtClub, New String() {"Generic Article", "Generic Material"})
Dim cDiv As Integer = findColumn(dtClub, New String() {"Division", "Art_AH2", "AH2"})
Dim cPlacement As Integer = findColumn(dtClub, New String() {"Placement"})
Dim cListing As Integer = findColumn(dtClub, New String() {"Listing Status"})
Dim cQty As Integer = findColumn(dtClub, New String() {"Initial Set Quantity"})
Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"})
Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
If cArt < 0 OrElse cDate < 0 OrElse cComment < 0 Then Throw New Exception("New Club data is missing required Article/comment/review-date columns.")
outHvs = 0 : outDelisted = 0 : outNoQty = 0
For Each cr As System.Data.DataRow In dtClub.Rows
    If Not inScope(If(cDiv >= 0, cr(cDiv), ""), divisionScope) Then Continue For
    If normalizeText(cr(cComment)).Length > 0 Then Continue For
    Dim hasArticle As Boolean = articleKey(cr(cArt)).Length > 0 OrElse (cGen >= 0 AndAlso articleKey(cr(cGen)).Length > 0)
    Dim newComment As String = ""
    If cPlacement >= 0 AndAlso normalizeText(cr(cPlacement)).Equals("HVS", StringComparison.OrdinalIgnoreCase) Then
        newComment = "HVS article" : outHvs += 1
    ElseIf cListing >= 0 AndAlso normalizeText(cr(cListing)).Equals("Delisted", StringComparison.OrdinalIgnoreCase) Then
        newComment = "Delisted article" : outDelisted += 1
    ElseIf hasArticle AndAlso cQty >= 0 Then
        Dim rawQty As String = normalizeText(cr(cQty))
        If rawQty.Length = 0 OrElse numberValue(cr(cQty)) = 0D Then
            newComment = "No initial set quantity for STO creation" : outNoQty += 1
        End If
    End If
    If newComment.Length > 0 Then
        cr(cComment) = newComment
        cr(cDate) = todayDate.Date
    End If
Next
'''

STORAGE_VB = COMMON_VB + r'''
If dtClub Is Nothing OrElse dtStorage Is Nothing Then Throw New Exception("New Club data and Storage Detail DataTables are required.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article", "Article Number"})
Dim cGen As Integer = findColumn(dtClub, New String() {"Generic Article", "Generic Material"})
Dim cDiv As Integer = findColumn(dtClub, New String() {"Division", "Art_AH2", "AH2"})
Dim cQty As Integer = findColumn(dtClub, New String() {"Initial Set Quantity"})
Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"})
Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
If cArt < 0 OrElse cDate < 0 OrElse cComment < 0 Then Throw New Exception("New Club data is missing Article/comment/review-date columns.")

' The Storage Report has a two-row grouped header. Preserve all rows by reading with AddHeaders=False.
Dim headerRow As Integer = -1, sArt As Integer = -1
For r As Integer = 0 To Math.Min(19, dtStorage.Rows.Count - 1)
    For c As Integer = 0 To dtStorage.Columns.Count - 1
        If normalizeHeader(dtStorage.Rows(r)(c)) = "ARTICLENUMBER" Then headerRow = r : sArt = c : Exit For
    Next
    If headerRow >= 0 Then Exit For
Next
If headerRow < 1 OrElse sArt < 0 Then Throw New Exception("Storage Detail: Article Number header was not found within the first 20 rows.")
Dim groupRow As Integer = headerRow - 1
Dim findDcColumn As System.Func(Of String, Integer) = Function(groupName)
    Dim groupCol As Integer = -1
    For c As Integer = 0 To dtStorage.Columns.Count - 1
        If normalizeHeader(dtStorage.Rows(groupRow)(c)) = normalizeHeader(groupName) Then groupCol = c : Exit For
    Next
    If groupCol < 0 Then Return -1
    For c As Integer = groupCol To dtStorage.Columns.Count - 1
        If c > groupCol AndAlso normalizeText(dtStorage.Rows(groupRow)(c)).Length > 0 Then Exit For
        Dim dcDigits As String = System.Text.RegularExpressions.Regex.Replace(normalizeText(dtStorage.Rows(headerRow)(c)), "[^0-9]", "")
        If dcDigits.Length > 0 AndAlso fourDigit(dcDigits) = fourDigit(dcNo) Then Return c
    Next
    Return -1
End Function
Dim sUnits As Integer = findDcColumn("On Site Units")
Dim sCases As Integer = findDcColumn("On Site Cases")
If sUnits < 0 Then Throw New Exception("Storage Detail: DC " & dcNo & " was not found under the On Site Units group heading.")
Dim stock As New Dictionary(Of String, Object())(StringComparer.OrdinalIgnoreCase)
For r As Integer = headerRow + 1 To dtStorage.Rows.Count - 1
    Dim key As String = articleKey(dtStorage.Rows(r)(sArt))
    If key.Length = 0 Then Continue For
    Dim units As Decimal = numberValue(dtStorage.Rows(r)(sUnits))
    Dim cases As Decimal = If(sCases >= 0, numberValue(dtStorage.Rows(r)(sCases)), 0D)
    If Not stock.ContainsKey(key) Then stock(key) = New Object() {0D, 0D}
    stock(key)(0) = CDec(stock(key)(0)) + units
    stock(key)(1) = CDec(stock(key)(1)) + cases
Next
outUpdated = 0 : outBelowSet = 0
For Each cr As System.Data.DataRow In dtClub.Rows
    If Not inScope(If(cDiv >= 0, cr(cDiv), ""), divisionScope) Then Continue For
    If normalizeText(cr(cComment)).Length > 0 Then Continue For
    Dim key As String = articleKey(cr(cArt))
    If key.Length = 0 OrElse Not stock.ContainsKey(key) Then
        If cGen >= 0 Then key = articleKey(cr(cGen))
    End If
    If key.Length = 0 OrElse Not stock.ContainsKey(key) Then Continue For
    Dim units As Decimal = CDec(stock(key)(0)), cases As Decimal = CDec(stock(key)(1))
    If units <= 0D Then Continue For
    Dim comment As String = "Storage DC " & dcNo & ": " & units.ToString("#,##0", System.Globalization.CultureInfo.InvariantCulture) & " units"
    If cases > 0D Then comment &= " (" & cases.ToString("#,##0", System.Globalization.CultureInfo.InvariantCulture) & " cs)"
    comment &= " on site - available to release"
    cr(cComment) = comment
    cr(cDate) = todayDate.Date
    outUpdated += 1
    If cQty >= 0 AndAlso numberValue(cr(cQty)) > 0D AndAlso units < numberValue(cr(cQty)) Then outBelowSet += 1
Next
'''

TRANSFER_VB = COMMON_VB + r'''
If dtClub Is Nothing OrElse dtInvStat Is Nothing Then Throw New Exception("New Club data and ZPRR_INVSTAT export DataTables are required.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article"})
Dim cGen As Integer = findColumn(dtClub, New String() {"Generic Article"})
Dim cDiv As Integer = findColumn(dtClub, New String() {"Division"})
Dim cQty As Integer = findColumn(dtClub, New String() {"Initial Set Quantity"})
Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"})
Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
Dim eClub As Integer = findColumn(dtInvStat, New String() {"Club", "Site"})
Dim eArt As Integer = findColumn(dtInvStat, New String() {"Article", "Material"})
Dim eOH As Integer = findColumn(dtInvStat, New String() {"OH", "On Hand"})
Dim eWos As Integer = findColumn(dtInvStat, New String() {"WOS", "Weeks of Supply"})
If cArt < 0 OrElse cComment < 0 OrElse eClub < 0 OrElse eArt < 0 OrElse eOH < 0 Then Throw New Exception("Club transfer input is missing Article, Club/Site or OH columns.")
Dim allowed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
For Each part As String In normalizeText(allowedClubs).Split(New Char() {"|"c, ";"c, ","c, " "c}, StringSplitOptions.RemoveEmptyEntries)
    allowed.Add(fourDigit(part))
Next
Dim candidates As New Dictionary(Of String, System.Collections.Generic.List(Of Object()))(StringComparer.OrdinalIgnoreCase)
For Each er As System.Data.DataRow In dtInvStat.Rows
    Dim key As String = articleKey(er(eArt)), site As String = fourDigit(er(eClub))
    Dim oh As Decimal = numberValue(er(eOH)), wos As Decimal = If(eWos >= 0, numberValue(er(eWos)), 0D)
    If key.Length = 0 OrElse site.Length = 0 OrElse oh <= 0D Then Continue For
    If site = fourDigit(newClubNo) OrElse site = fourDigit(dcNo) Then Continue For
    If allowed.Count > 0 AndAlso Not allowed.Contains(site) Then Continue For
    If Not candidates.ContainsKey(key) Then candidates(key) = New System.Collections.Generic.List(Of Object())
    candidates(key).Add(New Object() {site, oh, wos})
Next
outSuggested = 0 : outNone = 0
For Each cr As System.Data.DataRow In dtClub.Rows
    If Not inScope(If(cDiv >= 0, cr(cDiv), ""), divisionScope) Then Continue For
    Dim comment As String = normalizeText(cr(cComment))
    If Not comment.StartsWith("Need Club transfer", StringComparison.OrdinalIgnoreCase) OrElse comment.IndexOf("no club in DC", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
    Dim key As String = articleKey(cr(cArt))
    If (key.Length = 0 OrElse Not candidates.ContainsKey(key)) AndAlso cGen >= 0 Then key = articleKey(cr(cGen))
    Dim isq As Decimal = If(cQty >= 0, numberValue(cr(cQty)), 0D)
    Dim text As String = ""
    If candidates.ContainsKey(key) Then
        Dim list = candidates(key)
        list.Sort(Function(a,b)
            Dim ac As Boolean = (isq > 0D AndAlso CDec(a(1)) >= isq), bc As Boolean = (isq > 0D AndAlso CDec(b(1)) >= isq)
            If ac <> bc Then Return If(ac, -1, 1)
            If ac Then
                If CDec(a(2)) <> CDec(b(2)) Then Return Decimal.Compare(CDec(b(2)), CDec(a(2)))
                Return Decimal.Compare(CDec(b(1)), CDec(a(1)))
            End If
            If CDec(a(1)) <> CDec(b(1)) Then Return Decimal.Compare(CDec(b(1)), CDec(a(1)))
            Return Decimal.Compare(CDec(b(2)), CDec(a(2)))
        End Function)
        Dim picks As New System.Collections.Generic.List(Of String)
        For i As Integer = 0 To Math.Min(2, list.Count - 1)
            picks.Add(CStr(list(i)(0)) & " (OH " & formatQty(CDec(list(i)(1))) & ", WOS " & formatQty(CDec(list(i)(2))) & ")")
        Next
        text = "Transfer from: " & String.Join("; ", picks)
    End If
    If text.Length = 0 Then
        cr(cComment) = "Need Club transfer - no club in DC " & dcNo & " has stock"
        outNone += 1
    Else
        cr(cComment) = text : outSuggested += 1
    End If
    If cDate >= 0 Then cr(cDate) = todayDate.Date
Next
'''

APPT_VB = COMMON_VB + r'''
If dtClub Is Nothing OrElse dtApptConfirmations Is Nothing Then Throw New Exception("New Club data and ME23N confirmation export DataTables are required.")
Dim cDiv As Integer = findColumn(dtClub, New String() {"Division"})
Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"})
Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
Dim eDoc As Integer = findColumn(dtApptConfirmations, New String() {"Document", "Purchasing Document", "PO", "STO", "Document Number"})
Dim eDate As Integer = findColumn(dtApptConfirmations, New String() {"Delivery Date", "Confirmation Date", "Appointment Date"})
If cComment < 0 OrElse eDoc < 0 OrElse eDate < 0 Then Throw New Exception("Appointment lookup requires CSC2 Comments and confirmation Document/Delivery Date columns.")
Dim confirmations As New Dictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
For Each er As System.Data.DataRow In dtApptConfirmations.Rows
    Dim k As String = articleKey(er(eDoc)), d As DateTime = dateValue(er(eDate))
    If k.Length > 0 AndAlso d <> DateTime.MinValue AndAlso Not confirmations.ContainsKey(k) Then confirmations(k) = d
Next
outFound = 0 : outNoDate = 0 : outSkipped = 0
For Each cr As System.Data.DataRow In dtClub.Rows
    If Not inScope(If(cDiv >= 0, cr(cDiv), ""), divisionScope) Then Continue For
    Dim original As String = normalizeText(cr(cComment))
    If original.Length = 0 OrElse (original.IndexOf("STO", StringComparison.OrdinalIgnoreCase) < 0 AndAlso original.IndexOf("PO", StringComparison.OrdinalIgnoreCase) < 0) Then Continue For
    Dim clean As String = System.Text.RegularExpressions.Regex.Replace(original, "\s*\|\s*(Appt date.*|no Appt date.*)$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim()
    Dim matches = System.Text.RegularExpressions.Regex.Matches(clean, "\d+")
    Dim doc As String = ""
    For Each m As System.Text.RegularExpressions.Match In matches
        If m.Value.Length >= 8 AndAlso m.Value.Length >= doc.Length Then doc = m.Value
    Next
    If doc.Length < 8 Then outSkipped += 1 : Continue For
    If confirmations.ContainsKey(doc) Then
        cr(cComment) = clean & " | Appt date " & formatDate(confirmations(doc)) : outFound += 1
    Else
        cr(cComment) = clean & " | no Appt date" : outNoDate += 1
    End If
    If cDate >= 0 Then cr(cDate) = todayDate.Date
Next
'''

STF_VB = COMMON_VB + r'''
' Similarity used by the source VBA: hierarchy match plus Dice word overlap.
Dim scoreArticles As System.Func(Of Object(), Object(), Double) = Function(a,b)
    Dim score As Double = 0D
    If CStr(a(4)).Length > 0 AndAlso String.Equals(CStr(a(4)), CStr(b(4)), StringComparison.OrdinalIgnoreCase) Then score += 30D
    If CStr(a(3)).Length > 0 AndAlso String.Equals(CStr(a(3)), CStr(b(3)), StringComparison.OrdinalIgnoreCase) Then score += 20D
    If CStr(a(2)).Length > 0 AndAlso String.Equals(CStr(a(2)), CStr(b(2)), StringComparison.OrdinalIgnoreCase) Then score += 10D
    If CStr(a(1)).Length > 0 AndAlso String.Equals(CStr(a(1)), CStr(b(1)), StringComparison.OrdinalIgnoreCase) Then score += 10D
    If String.Equals(CStr(b(5)), "Live", StringComparison.OrdinalIgnoreCase) Then score += 5D
    Dim wa As New HashSet(Of String)(System.Text.RegularExpressions.Regex.Split(System.Text.RegularExpressions.Regex.Replace(CStr(a(0)).ToUpperInvariant(), "[^A-Z0-9]", " "), "\s+"), StringComparer.OrdinalIgnoreCase)
    Dim wb As New HashSet(Of String)(System.Text.RegularExpressions.Regex.Split(System.Text.RegularExpressions.Regex.Replace(CStr(b(0)).ToUpperInvariant(), "[^A-Z0-9]", " "), "\s+"), StringComparer.OrdinalIgnoreCase)
    wa.Remove("") : wb.Remove("")
    Dim hit As Integer = wa.Intersect(wb).Count()
    If wa.Count + wb.Count > 0 Then score += 60D * (2D * hit / (wa.Count + wb.Count))
    Return score
End Function

If dtClub Is Nothing OrElse dtAllocation Is Nothing Then Throw New Exception("New Club data and BJs Allocation Report DataTable are required for STF suggestions.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article"})
Dim cGen As Integer = findColumn(dtClub, New String() {"Generic Article"})
Dim cDiv As Integer = findColumn(dtClub, New String() {"Division"})
Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"})
Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
Dim rClub As Integer = findColumn(dtAllocation, New String() {"Club"})
Dim rPog As Integer = findColumn(dtAllocation, New String() {"POG_Name"})
Dim rArt As Integer = findColumn(dtAllocation, New String() {"Article"})
Dim rName As Integer = findColumn(dtAllocation, New String() {"Article_Name"})
Dim rMc As Integer = findColumn(dtAllocation, New String() {"MerchCat"})
Dim r4 As Integer = findColumn(dtAllocation, New String() {"AH4"})
Dim r5 As Integer = findColumn(dtAllocation, New String() {"AH5"})
Dim r6 As Integer = findColumn(dtAllocation, New String() {"AH6"})
Dim rStatus As Integer = findColumn(dtAllocation, New String() {"POG_STATUS"})
If cArt < 0 OrElse cComment < 0 OrElse rClub < 0 OrElse rPog < 0 OrElse rArt < 0 Then Throw New Exception("STF workflow requires Club, POG_Name and Article columns in the allocation report.")
Dim articleInfo As New Dictionary(Of String, Object())(StringComparer.OrdinalIgnoreCase)
Dim articlePogs As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
Dim pogArticles As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
Dim pogNames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
For Each rr As System.Data.DataRow In dtAllocation.Rows
    If articleKey(rr(rClub)) <> articleKey(clubNo) Then Continue For
    Dim a As String = articleKey(rr(rArt)), p As String = normalizeText(rr(rPog)).ToUpperInvariant()
    If a.Length = 0 OrElse p.Length = 0 Then Continue For
    If Not articleInfo.ContainsKey(a) Then articleInfo(a) = New Object() {If(rName >= 0, normalizeText(rr(rName)), ""), If(rMc >= 0, normalizeText(rr(rMc)), ""), If(r4 >= 0, normalizeText(rr(r4)), ""), If(r5 >= 0, normalizeText(rr(r5)), ""), If(r6 >= 0, normalizeText(rr(r6)), ""), If(rStatus >= 0, normalizeText(rr(rStatus)), "")}
    If Not articlePogs.ContainsKey(a) Then articlePogs(a) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    articlePogs(a).Add(p)
    If Not pogArticles.ContainsKey(p) Then pogArticles(p) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    pogArticles(p).Add(a) : pogNames(p) = normalizeText(rr(rPog))
Next
Dim stfKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
For Each cr As System.Data.DataRow In dtClub.Rows
    If normalizeText(cr(cComment)).IndexOf("need stf", StringComparison.OrdinalIgnoreCase) >= 0 Then
        Dim a As String = articleKey(cr(cArt))
        If a.Length = 0 AndAlso cGen >= 0 Then a = articleKey(cr(cGen))
        If a.Length > 0 Then stfKeys.Add(a)
    End If
Next
outSuggestions = 0 : outNoPog = 0 : outNoCandidates = 0
For Each cr As System.Data.DataRow In dtClub.Rows
    If Not inScope(If(cDiv >= 0, cr(cDiv), ""), divisionScope) Then Continue For
    Dim original As String = normalizeText(cr(cComment))
    If original.IndexOf("need stf", StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
    Dim sourceArt As String = articleKey(cr(cArt))
    If sourceArt.Length = 0 AndAlso cGen >= 0 Then sourceArt = articleKey(cr(cGen))
    If sourceArt.Length = 0 OrElse Not articleInfo.ContainsKey(sourceArt) Then
        cr(cComment) = original & " | STF options: article " & sourceArt & " is not on any POG for club " & clubNo
        outNoPog += 1 : Continue For
    End If
    Dim base As String = System.Text.RegularExpressions.Regex.Replace(original, "\s*\|\s*STF options.*$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim()
    Dim cand As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
    Dim meInfo As Object() = articleInfo(sourceArt)
    For Each p As String In articlePogs(sourceArt)
        For Each a As String In pogArticles(p)
            If a = sourceArt OrElse stfKeys.Contains(a) OrElse Not articleInfo.ContainsKey(a) OrElse cand.ContainsKey(a) Then Continue For
            cand(a) = 100D + scoreArticles(meInfo, articleInfo(a))
        Next
    Next
    If cand.Count < 3 Then
        For Each a As String In articleInfo.Keys
            If a = sourceArt OrElse stfKeys.Contains(a) OrElse cand.ContainsKey(a) Then Continue For
            Dim other As Object() = articleInfo(a)
            If (CStr(meInfo(4)).Length > 0 AndAlso String.Equals(CStr(meInfo(4)), CStr(other(4)), StringComparison.OrdinalIgnoreCase)) OrElse (CStr(meInfo(1)).Length > 0 AndAlso String.Equals(CStr(meInfo(1)), CStr(other(1)), StringComparison.OrdinalIgnoreCase)) Then cand(a) = scoreArticles(meInfo, other)
        Next
    End If
    Dim picks As New System.Collections.Generic.List(Of String)
    For j As Integer = 1 To 3
        If cand.Count = 0 Then Exit For
        Dim bestKey As String = "", bestScore As Double = Double.MinValue
        For Each kv As KeyValuePair(Of String, Double) In cand
            If kv.Value > bestScore Then bestKey = kv.Key : bestScore = kv.Value
        Next
        If bestKey.Length = 0 Then Exit For
        picks.Add(j.ToString() & ") " & bestKey & " " & CStr(articleInfo(bestKey)(0)))
        cand.Remove(bestKey)
    Next
    If picks.Count = 0 Then
        Dim pogNamesForArticle As New System.Collections.Generic.List(Of String)
        For Each p As String In articlePogs(sourceArt)
            pogNamesForArticle.Add(If(pogNames.ContainsKey(p), pogNames(p), p))
        Next
        cr(cComment) = base & " | STF options: no other article on POG " & String.Join(", ", pogNamesForArticle) & " for club " & clubNo
        outNoCandidates += 1
    Else
        Dim pogNamesForArticle As New System.Collections.Generic.List(Of String)
        For Each p As String In articlePogs(sourceArt)
            pogNamesForArticle.Add(If(pogNames.ContainsKey(p), pogNames(p), p))
        Next
        cr(cComment) = base & " | STF options are here, please check and update: " & String.Join("; ", picks) & " (POG " & String.Join(", ", pogNamesForArticle) & ")"
        outSuggestions += 1
    End If
    If cDate >= 0 Then cr(cDate) = todayDate.Date
Next
'''
# Additional VB scoring helpers embedded at end of STF procedure code.


REPLACEMENT_VB = COMMON_VB + r'''
If dtClub Is Nothing OrElse dtSto Is Nothing OrElse dtPo Is Nothing Then Throw New Exception("Replacement Article workflow requires Club, STO Export and PO Export DataTables.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article"}), cDiv As Integer = findColumn(dtClub, New String() {"Division"}), cQty As Integer = findColumn(dtClub, New String() {"Initial Set Quantity"}), cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"}), cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
If cComment < 0 OrElse cDate < 0 Then Throw New Exception("Replacement Article requires CSC2 Comments and Date of Last Review.")

' This reproduces RepStoInfo: delivered = Order Quantity - Still to be delivered;
' earliest remaining open STO is selected using Stat.-Rel. Del. Date then Delivery Date.
Dim sArt As Integer = findColumn(dtSto, New String() {"Article"}), sDoc As Integer = findColumn(dtSto, New String() {"Purchasing Document"}), sOrder As Integer = findColumn(dtSto, New String() {"Order Quantity"}), sStill As Integer = findColumn(dtSto, New String() {"Still to be delivered (qty)"}), sStat As Integer = findColumn(dtSto, New String() {"Stat.-Rel. Del. Date"}), sDelivery As Integer = findColumn(dtSto, New String() {"Delivery Date"}), sDel As Integer = findColumn(dtSto, New String() {"Deletion Indicator"}), sSite As Integer = findColumn(dtSto, New String() {"Site", "Plant"})
If sArt < 0 OrElse sDoc < 0 OrElse sOrder < 0 OrElse sStill < 0 Then Throw New Exception("Replacement STO export requires Article, Purchasing Document, Order Quantity and Still to be delivered (qty).")
If sSite < 0 Then sSite = 0
Dim stoFacts As New Dictionary(Of String, Object())(StringComparer.OrdinalIgnoreCase)
For Each sr As System.Data.DataRow In dtSto.Rows
    If fourDigit(sr(sSite)) <> fourDigit(newClubNo) Then Continue For
    If sDel >= 0 AndAlso isDeleted(sr(sDel)) Then Continue For
    Dim k As String = articleKey(sr(sArt))
    If k.Length = 0 Then Continue For
    Dim ordered As Decimal = numberValue(sr(sOrder)), still As Decimal = numberValue(sr(sStill))
    If Not stoFacts.ContainsKey(k) Then stoFacts(k) = New Object() {0D, "", DateTime.MinValue}
    Dim f As Object() = stoFacts(k)
    If ordered > still Then f(0) = CDec(f(0)) + (ordered - still)
    If still > 0D Then
        Dim due As DateTime = DateTime.MinValue
        If sStat >= 0 Then due = dateValue(sr(sStat))
        If due = DateTime.MinValue AndAlso sDelivery >= 0 Then due = dateValue(sr(sDelivery))
        If CStr(f(1)).Length = 0 OrElse (due <> DateTime.MinValue AndAlso (CDate(f(2)) = DateTime.MinValue OrElse due < CDate(f(2)))) Then
            f(1) = normalizeText(sr(sDoc)) : f(2) = due
        End If
    End If
    stoFacts(k) = f
Next

' This reproduces BuildBestPO: open only (Still to be delivered > 0), not deleted,
' within Today..Last Truck Date; retain the earliest due date per Article.
Dim pArt As Integer = findColumn(dtPo, New String() {"Article"}), pPo As Integer = findColumn(dtPo, New String() {"Purchasing Document"}), pStat As Integer = findColumn(dtPo, New String() {"Stat.-Rel. Del. Date"}), pDue As Integer = findColumn(dtPo, New String() {"Delivery Date", "Due Date"}), pStill As Integer = findColumn(dtPo, New String() {"Still to be delivered (qty)"}), pDel As Integer = findColumn(dtPo, New String() {"Deletion Indicator"}), pSite As Integer = findColumn(dtPo, New String() {"Site", "Plant"})
If pArt < 0 OrElse pPo < 0 Then Throw New Exception("Replacement PO export requires Article and Purchasing Document.")
If pSite < 0 Then pSite = 0
Dim poBest As New Dictionary(Of String, Object())(StringComparer.OrdinalIgnoreCase)
For Each pr As System.Data.DataRow In dtPo.Rows
    If fourDigit(pr(pSite)) <> fourDigit(dcNo) Then Continue For
    If pDel >= 0 AndAlso isDeleted(pr(pDel)) Then Continue For
    If pStill >= 0 AndAlso numberValue(pr(pStill)) <= 0D Then Continue For
    Dim due As DateTime = DateTime.MinValue
    If pStat >= 0 Then due = dateValue(pr(pStat))
    If due = DateTime.MinValue AndAlso pDue >= 0 Then due = dateValue(pr(pDue))
    If due = DateTime.MinValue OrElse due.Date < todayDate.Date OrElse due.Date > truckDate.Date Then Continue For
    Dim k As String = articleKey(pr(pArt)), doc As String = normalizeText(pr(pPo))
    If k.Length = 0 OrElse doc.Length = 0 Then Continue For
    If Not poBest.ContainsKey(k) OrElse due < CDate(poBest(k)(0)) Then poBest(k) = New Object() {due, doc}
Next

' Build optional Storage lookup once, retaining the original two-row grouped header structure.
Dim storageUnits As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
If dtStorage IsNot Nothing AndAlso dtStorage.Rows.Count > 0 Then
    Dim sh As Integer = -1, sa As Integer = -1
    For rr As Integer = 0 To Math.Min(19, dtStorage.Rows.Count - 1)
        For cc As Integer = 0 To dtStorage.Columns.Count - 1
            If normalizeHeader(dtStorage.Rows(rr)(cc)) = "ARTICLENUMBER" Then sh = rr : sa = cc : Exit For
        Next
        If sh >= 1 Then Exit For
    Next
    If sh >= 1 Then
        Dim groupRow As Integer = sh - 1, groupCol As Integer = -1, su As Integer = -1
        For cc As Integer = 0 To dtStorage.Columns.Count - 1
            If normalizeHeader(dtStorage.Rows(groupRow)(cc)) = "ONSITEUNITS" Then groupCol = cc : Exit For
        Next
        If groupCol >= 0 Then
            For cc As Integer = groupCol To dtStorage.Columns.Count - 1
                If cc > groupCol AndAlso normalizeText(dtStorage.Rows(groupRow)(cc)).Length > 0 Then Exit For
                Dim digits As String = System.Text.RegularExpressions.Regex.Replace(normalizeText(dtStorage.Rows(sh)(cc)), "[^0-9]", "")
                If digits.Length > 0 AndAlso fourDigit(digits) = fourDigit(dcNo) Then su = cc : Exit For
            Next
        End If
        If sa >= 0 AndAlso su >= 0 Then
            For rr As Integer = sh + 1 To dtStorage.Rows.Count - 1
                Dim k As String = articleKey(dtStorage.Rows(rr)(sa))
                If k.Length = 0 Then Continue For
                If Not storageUnits.ContainsKey(k) Then storageUnits(k) = 0D
                storageUnits(k) += numberValue(dtStorage.Rows(rr)(su))
            Next
        End If
    End If
End If

outUpdated = 0
For Each cr As System.Data.DataRow In dtClub.Rows
    If Not inScope(If(cDiv >= 0, cr(cDiv), ""), divisionScope) Then Continue For
    Dim original As String = normalizeText(cr(cComment))
    Dim m As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(original, "replac\w*[^0-9]*(\d{3,})", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
    If Not m.Success Then Continue For
    Dim replacementArt As String = articleKey(m.Groups(1).Value), need As Decimal = If(cQty >= 0, numberValue(cr(cQty)), 0D)
    Dim result As String = ""
    If stoFacts.ContainsKey(replacementArt) Then
        Dim f As Object() = stoFacts(replacementArt), delivered As Decimal = CDec(f(0)), openDoc As String = CStr(f(1)), openDue As DateTime = CDate(f(2))
        If openDoc.Length = 0 Then
            result = "STO manifested"
        ElseIf delivered > 0D AndAlso (need = 0D OrElse delivered < need) Then
            result = "partially manifested, next STO " & openDoc & " due " & formatDate(openDue)
        ElseIf delivered > 0D Then
            result = "STO manifested"
        Else
            result = "STO " & openDoc & " due " & formatDate(openDue)
        End If
    ElseIf poBest.ContainsKey(replacementArt) Then
        result = "PO " & CStr(poBest(replacementArt)(1)) & " due " & formatDate(CDate(poBest(replacementArt)(0)))
    ElseIf storageUnits.ContainsKey(replacementArt) AndAlso storageUnits(replacementArt) > 0D Then
        result = "Storage DC " & dcNo & ": " & storageUnits(replacementArt).ToString("#,##0", System.Globalization.CultureInfo.InvariantCulture) & " units"
    Else
        result = "no STO or PO"
    End If
    cr(cComment) = "Replacement article " & replacementArt & " - " & result
    cr(cDate) = todayDate.Date : outUpdated += 1
Next
'''

TRACKING_VB = COMMON_VB + r'''
If dtClub Is Nothing OrElse dtTracking Is Nothing Then Throw New Exception("Source New Club data and Tracking Report DataTables are required.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article"}), cDesc As Integer = findColumn(dtClub, New String() {"Article Description", "Description"}), cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"}), cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
Dim tArt As Integer = findColumn(dtTracking, New String() {"Article"}), tDesc As Integer = findColumn(dtTracking, New String() {"Article Description", "Description"}), tDate As Integer = findColumn(dtTracking, New String() {"Date of Last Review"}), tComment As Integer = findColumn(dtTracking, New String() {"CSC2 Comments"})
If cArt < 0 OrElse cDesc < 0 OrElse cDate < 0 OrElse cComment < 0 OrElse tArt < 0 OrElse tDesc < 0 OrElse tDate < 0 OrElse tComment < 0 Then Throw New Exception("Tracking Report must include Article, Article Description, Date of Last Review and CSC2 Comments.")
Dim source As New Dictionary(Of String, Object())(StringComparer.OrdinalIgnoreCase)
For Each r As System.Data.DataRow In dtClub.Rows
    If normalizeText(r(cComment)).Length = 0 Then Continue For
    Dim key As String = articleKey(r(cArt)) & ChrW(1) & normalizeText(r(cDesc)).ToUpperInvariant()
    source(key) = New Object() {r(cDate), r(cComment)}
Next
outMatched = 0 : outNotFound = 0
For Each tr As System.Data.DataRow In dtTracking.Rows
    Dim key As String = articleKey(tr(tArt)) & ChrW(1) & normalizeText(tr(tDesc)).ToUpperInvariant()
    If source.ContainsKey(key) Then
        tr(tDate) = source(key)(0) : tr(tComment) = source(key)(1) : outMatched += 1
    Else
        outNotFound += 1
    End If
Next
'''

SUMMARY_VB = COMMON_VB + r'''
If dtClub Is Nothing Then Throw New Exception("New Club data DataTable is required.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article"}), cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"}), cDiv As Integer = findColumn(dtClub, New String() {"Division", "Art_AH2", "AH2"}), cMerch As Integer = findColumn(dtClub, New String() {"Merch Cat", "Merchandise Category", "Art_AH4 - Merch Cat", "MerchCat"}), cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"}), cDesc As Integer = findColumn(dtClub, New String() {"Article Description", "Description", "Short Text"})
If cArt < 0 OrElse cComment < 0 Then Throw New Exception("Summary needs Article and CSC2 Comments columns.")
Dim category As System.Func(Of Object, Integer) = Function(v)
    Dim t As String = normalizeText(v).ToLowerInvariant()
    If t.Length = 0 Then Return 9
    If t.StartsWith("sto manifested") Then Return 0
    If t.StartsWith("partial") Then Return 1
    If t.StartsWith("need manual sto") Then Return 2
    If t.Contains("stf") Then Return 7
    If t.StartsWith("sto ") Then
        If t.Contains("is passed sto not delivered") Then Return 3
        Return 4
    End If
    If t.StartsWith("po ") AndAlso t.Contains(" due on ") Then Return 5
    If t.StartsWith("storage dc") Then Return 6
    Return 8
End Function
Dim sumNorm As System.Func(Of Object, String) = Function(v) normalizeText(v).ToUpperInvariant()
Dim sumDivKey As System.Func(Of Object, String) = Function(v)
    Dim t As String = System.Text.RegularExpressions.Regex.Replace(sumNorm(v), "[^A-Z0-9]", "")
    If t = "AHGM" OrElse t = "GM" Then t = "AHGENERALMERCHANDISE"
    Return t
End Function
Dim sumLetters As System.Func(Of String, String) = Function(v) System.Text.RegularExpressions.Regex.Replace(If(v, "").ToUpperInvariant(), "[^A-Z]", "")
Dim levDist As System.Func(Of String, String, Integer) = Function(a, b)
    If Math.Abs(a.Length - b.Length) > 2 Then Return 99
    Dim d(a.Length, b.Length) As Integer
    For i As Integer = 0 To a.Length : d(i, 0) = i : Next
    For j As Integer = 0 To b.Length : d(0, j) = j : Next
    For i As Integer = 1 To a.Length
        For j As Integer = 1 To b.Length
            Dim cost As Integer = If(a(i - 1) = b(j - 1), 0, 1)
            d(i, j) = Math.Min(Math.Min(d(i - 1, j) + 1, d(i, j - 1) + 1), d(i - 1, j - 1) + cost)
        Next
    Next
    Return d(a.Length, b.Length)
End Function
Dim person As System.Func(Of Object, String) = Function(v)
    Dim t As String = normalizeText(v)
    If t = "" OrElse t.ToUpperInvariant() = "NA" OrElse t.ToUpperInvariant() = "N/A" OrElse t.ToUpperInvariant() = "#N/A" OrElse t = "0" Then Return "Not owned by CSC2"
    Return t
End Function
Dim mapCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
Dim mapAnalyst As System.Func(Of String, String, String) = Function(divText, merchText)
    Dim dk As String = sumDivKey(divText), mk As String = sumNorm(merchText), cacheKey As String = dk & "|" & mk
    If mapCache.ContainsKey(cacheKey) Then Return mapCache(cacheKey)
    Dim best As Integer = -1
    If dtCategoryMap IsNot Nothing AndAlso dtCategoryMap.Columns.Count >= 7 AndAlso mk.Length > 0 Then
        For i As Integer = 0 To dtCategoryMap.Rows.Count - 1
            If sumDivKey(dtCategoryMap.Rows(i)(0)) = dk AndAlso sumNorm(dtCategoryMap.Rows(i)(2)) = mk Then best = i : Exit For
        Next
        If best < 0 Then
            For i As Integer = 0 To dtCategoryMap.Rows.Count - 1
                If sumDivKey(dtCategoryMap.Rows(i)(0)) = dk AndAlso sumNorm(dtCategoryMap.Rows(i)(1)) = mk Then best = i : Exit For
            Next
        End If
        If best < 0 Then
            Dim wanted As String = sumLetters(mk)
            For i As Integer = 0 To dtCategoryMap.Rows.Count - 1
                If sumDivKey(dtCategoryMap.Rows(i)(0)) <> dk Then Continue For
                Dim candidate As String = sumLetters(sumNorm(dtCategoryMap.Rows(i)(2)))
                If candidate = "NA" OrElse candidate.Length = 0 Then candidate = sumLetters(sumNorm(dtCategoryMap.Rows(i)(1)))
                If candidate.Length >= 5 AndAlso wanted.Length >= 5 AndAlso (candidate = wanted OrElse wanted.Contains(candidate) OrElse candidate.Contains(wanted) OrElse levDist(candidate, wanted) <= 2) Then best = i : Exit For
            Next
        End If
    End If
    Dim owner As String = If(best < 0, "Not owned by CSC2", person(dtCategoryMap.Rows(best)(6)))
    mapCache(cacheKey) = owner
    Return owner
End Function
Dim total As Integer = 0
Dim names As String() = New String() {"STO Manifested", "Partial STO", "Manual STO", "STO Past Due", "Open STO", "PO Due", "Storage DC", "STF", "Other", "Blank"}
Dim counts(9) As Integer
Dim divBuckets As New Dictionary(Of String, Integer())(StringComparer.OrdinalIgnoreCase)
Dim analystBuckets As New Dictionary(Of String, Integer())(StringComparer.OrdinalIgnoreCase)
outCommentSummary = New System.Data.DataTable("CommentSummary")
outCommentSummary.Columns.Add("Category", GetType(String))
outCommentSummary.Columns.Add("CategoryNo", GetType(Integer))
outCommentSummary.Columns.Add("Lines", GetType(Integer))
outCommentSummary.Columns.Add("Percent", GetType(Decimal))
outDivisionSummary = New System.Data.DataTable("DivisionSummary")
For Each col As String In New String() {"Division", "TotalLines", "Manifested", "ToAction", "Other"} : outDivisionSummary.Columns.Add(col, If(col = "Division", GetType(String), GetType(Integer))) : Next
outAnalystSummary = New System.Data.DataTable("AnalystSummary")
For Each col As String In New String() {"CSC2 Analyst", "ToAction", "AllLines", "BlankLines", "NotReviewed3PlusDays"} : outAnalystSummary.Columns.Add(col, If(col = "CSC2 Analyst", GetType(String), GetType(Integer))) : Next
outActionDetails = New System.Data.DataTable("ActionDetails")
For Each col As String In New String() {"Article", "Description", "Division", "Merch Cat", "Category", "Comment", "ReviewDate", "CSC2 Analyst"} : outActionDetails.Columns.Add(col, GetType(String)) : Next
Dim today As DateTime = todayDate.Date
For Each r As System.Data.DataRow In dtClub.Rows
    Dim art As String = articleKey(r(cArt))
    If art.Length = 0 Then Continue For
    Dim divText As String = If(cDiv >= 0, normalizeText(r(cDiv)), "")
    If Not inScope(divText, divisionFilter) Then Continue For
    total += 1
    Dim cat As Integer = category(r(cComment))
    counts(cat) += 1
    Dim divKey As String = If(divText.Length = 0, "(no division)", divText)
    If Not divBuckets.ContainsKey(divKey) Then divBuckets(divKey) = New Integer() {0, 0, 0, 0}
    Dim dv As Integer() = divBuckets(divKey) : dv(0) += 1
    If cat = 0 Then
        dv(1) += 1
    ElseIf cat = 1 OrElse cat = 2 OrElse cat = 9 Then
        dv(2) += 1
    Else
        dv(3) += 1
    End If
    Dim mc As String = If(cMerch >= 0, normalizeText(r(cMerch)), "")
    Dim analyst As String = mapAnalyst(divText, mc)
    If Not analystBuckets.ContainsKey(analyst) Then analystBuckets(analyst) = New Integer() {0, 0, 0, 0}
    Dim av As Integer() = analystBuckets(analyst) : av(1) += 1
    If cat = 1 OrElse cat = 2 OrElse cat = 9 Then av(0) += 1
    If cat = 9 Then av(2) += 1
    Dim old As Boolean = False
    If cat <> 0 AndAlso cat <> 6 AndAlso cat <> 9 AndAlso cDate >= 0 Then
        Dim review As DateTime = dateValue(r(cDate))
        old = (review <> DateTime.MinValue AndAlso review.Date <= today.AddDays(-3))
    End If
    If old Then av(3) += 1
    If cat = 1 OrElse cat = 2 OrElse cat = 9 Then
        Dim desc As String = If(cDesc >= 0, normalizeText(r(cDesc)), "")
        Dim reviewText As String = If(cDate >= 0, normalizeText(r(cDate)), "")
        outActionDetails.Rows.Add(art, desc, divText, mc, names(cat), normalizeText(r(cComment)), reviewText, analyst)
    End If
Next
For i As Integer = 0 To 9
    outCommentSummary.Rows.Add(names(i), i, counts(i), If(total = 0, 0D, Math.Round(CDec(counts(i)) / CDec(total) * 100D, 2)))
Next
For Each kv As KeyValuePair(Of String, Integer()) In divBuckets
    outDivisionSummary.Rows.Add(kv.Key, kv.Value(0), kv.Value(1), kv.Value(2), kv.Value(3))
Next
For Each kv As KeyValuePair(Of String, Integer()) In analystBuckets
    outAnalystSummary.Rows.Add(kv.Key, kv.Value(0), kv.Value(1), kv.Value(2), kv.Value(3))
Next
'''

SNAPSHOT_VB = COMMON_VB + r'''
If dtClub Is Nothing Then Throw New Exception("New Club data DataTable is required.")
If dtSnapshot Is Nothing Then dtSnapshot = New System.Data.DataTable("CommentSnapshot")
If dtSnapshot.Columns.Count = 0 Then
    dtSnapshot.Columns.Add("Snapshot Date", GetType(DateTime))
    dtSnapshot.Columns.Add("Article", GetType(String))
    dtSnapshot.Columns.Add("Category no.", GetType(Integer))
    dtSnapshot.Columns.Add("Comment", GetType(String))
End If
Dim cArt As Integer = findColumn(dtClub, New String() {"Article"}), cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
If cArt < 0 OrElse cComment < 0 Then Throw New Exception("Snapshot requires Article and CSC2 Comments columns.")
' Keep at most one snapshot per article per day; refresh today's values on repeat runs.
For i As Integer = dtSnapshot.Rows.Count - 1 To 0 Step -1
    If Not dtSnapshot.Rows(i).IsNull(0) AndAlso dateValue(dtSnapshot.Rows(i)(0)).Date = todayDate.Date Then dtSnapshot.Rows.RemoveAt(i)
Next
For Each r As System.Data.DataRow In dtClub.Rows
    Dim a As String = articleKey(r(cArt))
    If a.Length = 0 Then Continue For
    Dim c As String = normalizeText(r(cComment)), cat As Integer = 8
    Dim low As String = c.ToLowerInvariant()
    If c.Length = 0 Then
        cat = 9
    ElseIf low.StartsWith("sto manifested") Then
        cat = 0
    ElseIf low.StartsWith("partial") Then
        cat = 1
    ElseIf low.StartsWith("need manual sto") Then
        cat = 2
    ElseIf low.StartsWith("sto ") AndAlso low.Contains("is passed sto not delivered") Then
        cat = 3
    ElseIf low.StartsWith("sto ") Then
        cat = 4
    ElseIf low.StartsWith("po ") AndAlso low.Contains(" due on ") Then
        cat = 5
    ElseIf low.StartsWith("storage dc") Then
        cat = 6
    ElseIf low.Contains("stf") Then
        cat = 7
    End If
    dtSnapshot.Rows.Add(todayDate.Date, a, cat, c)
Next
'''

LOG_VB = COMMON_VB + r'''
If dtLog Is Nothing Then dtLog = New System.Data.DataTable("RunLog")
If dtLog.Columns.Count = 0 Then
    dtLog.Columns.Add("Date / time", GetType(DateTime))
    dtLog.Columns.Add("User", GetType(String))
    dtLog.Columns.Add("Club", GetType(String))
    dtLog.Columns.Add("Action", GetType(String))
    dtLog.Columns.Add("Result", GetType(String))
End If
dtLog.Rows.Add(DateTime.Now, Environment.UserName, clubNo, actionName, resultText)
'''

PREPARE_WRITE_VB = COMMON_VB + r'''
Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"})
Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
If cDate < 0 OrElse cComment < 0 Then Throw New Exception("Cannot write back: Date of Last Review or CSC2 Comments header missing.")
outDateColumn = New System.Data.DataTable("DateColumn")
outCommentColumn = New System.Data.DataTable("CommentColumn")
outDateColumn.Columns.Add("Date of Last Review", GetType(Object))
outCommentColumn.Columns.Add("CSC2 Comments", GetType(Object))
For Each r As System.Data.DataRow In dtClub.Rows
    outDateColumn.Rows.Add(If(r.IsNull(cDate), System.DBNull.Value, r(cDate)))
    outCommentColumn.Rows.Add(If(r.IsNull(cComment), System.DBNull.Value, r(cComment)))
Next
Dim toColumnLetter As System.Func(Of Integer, String) = Function(colNo)
    Dim n As Integer = colNo, s As String = ""
    While n > 0
        n -= 1 : s = ChrW(65 + (n Mod 26)) & s : n = n \ 26
    End While
    Return s
End Function
outDateStartCell = toColumnLetter(cDate + 1) & "2"
outCommentStartCell = toColumnLetter(cComment + 1) & "2"
'''

VALIDATE_VB = COMMON_VB + r'''
If dtAction Is Nothing OrElse dtClub Is Nothing Then Throw New Exception("Action Sheet and New Club data are required.")
Dim findLabelValue As System.Func(Of String, Object) = Function(label)
    Dim wanted As String = normalizeHeader(label)
    For r As Integer = 0 To dtAction.Rows.Count - 2
        For c As Integer = 0 To dtAction.Columns.Count - 1
            If normalizeHeader(dtAction.Rows(r)(c)) = wanted Then
                Return dtAction.Rows(r + 1)(c)
            End If
        Next
    Next
    Return Nothing
End Function
Dim rawClub As Object = findLabelValue("New Club Number")
Dim rawDc As Object = findLabelValue("DC Number")
Dim rawTruck As Object = findLabelValue("Last truck date")
outClubNo = fourDigit(rawClub)
outDcNo = fourDigit(rawDc)
outTruckDate = dateValue(rawTruck)
If outClubNo.Length = 0 Then Throw New Exception("New Club Number is missing from the Action Sheet.")
If outDcNo.Length = 0 Then Throw New Exception("DC Number is missing from the Action Sheet.")
If outTruckDate = DateTime.MinValue Then Throw New Exception("Last truck date is missing or invalid on the Action Sheet.")
If dtClub.Rows.Count = 0 Then Throw New Exception("New Club data contains no rows.")
Dim required As New String() {"Article", "Initial Set Quantity", "Date of Last Review", "CSC2 Comments"}
For Each h As String In required
    If findColumn(dtClub, New String() {h}) < 0 Then Throw New Exception("New Club data is missing required column: " & h)
Next
outValidationMessage = "Validated Action Sheet: Club " & outClubNo & ", DC " & outDcNo & ", Last truck date " & formatDate(outTruckDate)
'''


def _arg_xml(name: str, direction: str, xaml_type: str, expression: str | None = None) -> str:
    arg_tag = {"In":"InArgument", "Out":"OutArgument", "InOut":"InOutArgument"}[direction]
    body = "" if expression is None else f"[{html.escape(expression, quote=False)}]"
    return f'<{arg_tag} x:TypeArguments="{xaml_type}" x:Key="{name}">{body}</{arg_tag}>'


def _xaml_escape_attr(s: str) -> str:
    return html.escape(s, quote=True).replace("\n", "&#xA;").replace("\r", "")


def _member(name: str, direction: str, xaml_type: str) -> str:
    return f'<x:Property Name="{name}" Type="{direction}Argument({xaml_type})" />'


def workflow_xaml(class_name: str, display_name: str, source_map: str, args: list[tuple[str,str,str]], code: str, imports: list[str] | None = None) -> str:
    members = "\n".join(_member(*a) for a in args)
    code_args = []
    for name, direction, typ in args:
        if name in {"in_"}:
            continue
        code_name = name
        if name.startswith("in_"):
            # Invoke Code identifiers follow the source block's local parameter names.
            code_name = name[3:]
        elif name.startswith("out_"):
            # Output arguments map to the local variable names used in VB code.
            # The `dt` prefix is a type hint in argument names, not part of the
            # shorter local names used by the table-preparation/report routines.
            special_outputs = {
                "out_dtDateColumn": "outDateColumn",
                "out_dtCommentColumn": "outCommentColumn",
                "out_dtSummary": "outSummary",
            }
            code_name = special_outputs.get(name, "out" + name[4:5].upper() + name[5:])
        elif name.startswith("io_"):
            code_name = name[3:]
        expr = name
        code_args.append(_arg_xml(code_name, direction, typ, expr))
    args_xml = "\n".join(code_args)
    imp_xml = ""
    if imports:
        imp_xml = "\n".join(f'<TextExpression.Import><Import>"{html.escape(i)}"</Import></TextExpression.Import>' for i in imports)
    code_attr = _xaml_escape_attr(code)
    return f'''<Activity mc:Ignorable="sap sap2010" x:Class="{class_name}" sap2010:WorkflowViewState.IdRef="ActivityBuilder_1"
 xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities"
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:sap="http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation"
 xmlns:sap2010="http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation"
 xmlns:ui="http://schemas.uipath.com/workflow/activities"
 xmlns:scg="clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib"
 xmlns:sco="clr-namespace:System.Collections.ObjectModel;assembly=System.Private.CoreLib"
 xmlns:s="clr-namespace:System.Data;assembly=System.Data.Common"
 xmlns:sys="clr-namespace:System;assembly=System.Private.CoreLib">
  <TextExpression.NamespacesForImplementation>
    <sco:Collection x:TypeArguments="x:String">
      <x:String>System</x:String><x:String>System.Collections.Generic</x:String><x:String>System.Data</x:String><x:String>System.Linq</x:String><x:String>System.Text</x:String><x:String>System.Text.RegularExpressions</x:String><x:String>System.Globalization</x:String>
    </sco:Collection>
  </TextExpression.NamespacesForImplementation>
  <TextExpression.ReferencesForImplementation>
    <sco:Collection x:TypeArguments="AssemblyReference">
      <AssemblyReference>System</AssemblyReference><AssemblyReference>System.Data.Common</AssemblyReference><AssemblyReference>UiPath.System.Activities</AssemblyReference>
    </sco:Collection>
  </TextExpression.ReferencesForImplementation>
  <x:Members>
    {members}
  </x:Members>
  <!-- Source-to-workflow mapping: {html.escape(source_map, quote=False)} -->
  <Sequence DisplayName="{html.escape(display_name, quote=True)}" sap2010:WorkflowViewState.IdRef="Sequence_1">
    <ui:InvokeCode Code="{code_attr}" Language="VBNet" Private="True" DisplayName="Apply {html.escape(display_name, quote=True)} logic" sap2010:WorkflowViewState.IdRef="InvokeCode_1">
      <ui:InvokeCode.Arguments>
        <scg:Dictionary x:TypeArguments="x:String, Argument">
          {args_xml}
        </scg:Dictionary>
      </ui:InvokeCode.Arguments>
    </ui:InvokeCode>
  </Sequence>
</Activity>
'''


def _invoke_workflow(path: str, mappings: list[tuple[str,str,str,str]], idx: int, display: str) -> str:
    args = []
    for child_name, direction, typ, parent_expr in mappings:
        args.append(_arg_xml(child_name, direction, typ, parent_expr))
    return f'''<ui:InvokeWorkflowFile WorkflowFileName="{path}" DisplayName="{html.escape(display)}" sap2010:WorkflowViewState.IdRef="InvokeWorkflowFile_{idx}">
  <ui:InvokeWorkflowFile.Arguments>
    <scg:Dictionary x:TypeArguments="x:String, Argument">{''.join(args)}</scg:Dictionary>
  </ui:InvokeWorkflowFile.Arguments>
</ui:InvokeWorkflowFile>'''


def _main_xaml() -> str:
    variables = [
        ("dtAction", "s:DataTable", ""), ("dtClub", "s:DataTable", ""), ("dtSto", "s:DataTable", ""), ("dtPo", "s:DataTable", ""),
        ("dtStorage", "s:DataTable", ""), ("dtSnapshot", "s:DataTable", ""), ("dtLog", "s:DataTable", ""),
        ("dtWriteDate", "s:DataTable", ""), ("dtWriteComment", "s:DataTable", ""),
        ("clubNo", "x:String", ""), ("dcNo", "x:String", ""), ("truckDate", "sys:DateTime", ""),
        ("reviewStartCell", "x:String", ""), ("commentStartCell", "x:String", ""),
        ("updatedSto", "x:Int32", ""), ("updatedPo", "x:Int32", ""), ("updatedStorage", "x:Int32", ""),
        ("overdueSto", "x:Int32", ""), ("missingSto", "x:Int32", ""), ("noMatchPo", "x:Int32", ""), ("preservedPo", "x:Int32", ""),
        ("belowSet", "x:Int32", ""), ("msg", "x:String", ""), ("runResult", "x:String", "")
    ]
    var_xml = "\n".join(f'<Variable x:TypeArguments="{typ}" Name="{name}" />' for name, typ, _ in variables)
    acts=[]
    acts.append('<ui:ReadRange WorkbookPath="[in_WorkbookPath]" SheetName="Action Sheet " Range="A1:AZ100" AddHeaders="False" DataTable="[dtAction]" DisplayName="Read Action Sheet" sap2010:WorkflowViewState.IdRef="Read_Action" />')
    acts.append('<ui:ReadRange WorkbookPath="[in_WorkbookPath]" SheetName="New Club data" AddHeaders="True" DataTable="[dtClub]" DisplayName="Read New Club data" sap2010:WorkflowViewState.IdRef="Read_Club" />')
    acts.append('<ui:ReadRange WorkbookPath="[in_WorkbookPath]" SheetName="STO Export" AddHeaders="True" DataTable="[dtSto]" DisplayName="Read STO Export" sap2010:WorkflowViewState.IdRef="Read_STO" />')
    acts.append('<ui:ReadRange WorkbookPath="[in_WorkbookPath]" SheetName="PO Export" AddHeaders="True" DataTable="[dtPo]" DisplayName="Read PO Export" sap2010:WorkflowViewState.IdRef="Read_PO" />')
    acts.append('<ui:ReadRange WorkbookPath="[in_WorkbookPath]" SheetName="Comment Snapshot" AddHeaders="True" DataTable="[dtSnapshot]" DisplayName="Read Comment Snapshot" sap2010:WorkflowViewState.IdRef="Read_Snapshot" />')
    acts.append('<ui:ReadRange WorkbookPath="[in_WorkbookPath]" SheetName="Run Log" AddHeaders="True" DataTable="[dtLog]" DisplayName="Read Run Log" sap2010:WorkflowViewState.IdRef="Read_RunLog" />')
    validate_call=_invoke_workflow("Workflows/Validate_Run_Inputs.xaml", [
        ("in_dtAction","In","s:DataTable","dtAction"),("in_dtClub","In","s:DataTable","dtClub"),
        ("out_clubNo","Out","x:String","clubNo"),("out_dcNo","Out","x:String","dcNo"),("out_truckDate","Out","sys:DateTime","truckDate"),("out_validationMessage","Out","x:String","msg")], 1, "Validate inputs before processing")
    snapshot_call=_invoke_workflow("Workflows/Take_Daily_Snapshot.xaml", [
        ("in_dtClub","In","s:DataTable","dtClub"),("io_dtSnapshot","InOut","s:DataTable","dtSnapshot"),("in_todayDate","In","sys:DateTime","Now.Date")], 2, "Take the daily comment snapshot")
    sto_call=_invoke_workflow("Workflows/STO_Check.xaml", [
        ("io_dtClub","InOut","s:DataTable","dtClub"),("in_dtSto","In","s:DataTable","dtSto"),("in_clubNo","In","x:String","clubNo"),("in_divisionScope","In","x:String","in_DivisionScope"),("in_todayDate","In","sys:DateTime","Now.Date"),("in_truckDate","In","sys:DateTime","truckDate"),
        ("out_updated","Out","x:Int32","updatedSto"),("out_overdue","Out","x:Int32","overdueSto"),("out_missing","Out","x:Int32","missingSto")], 3, "Run STO Check when the STO Export contains data")
    po_call=_invoke_workflow("Workflows/PO_Check.xaml", [
        ("io_dtClub","InOut","s:DataTable","dtClub"),("in_dtPo","In","s:DataTable","dtPo"),("in_dcNo","In","x:String","dcNo"),("in_divisionScope","In","x:String","in_DivisionScope"),("in_todayDate","In","sys:DateTime","Now.Date"),("in_truckDate","In","sys:DateTime","truckDate"),
        ("out_updated","Out","x:Int32","updatedPo"),("out_noMatch","Out","x:Int32","noMatchPo"),("out_preserved","Out","x:Int32","preservedPo")], 4, "Run PO Check when the PO Export contains data")
    storage_call=_invoke_workflow("Workflows/Storage_Release.xaml", [
        ("io_dtClub","InOut","s:DataTable","dtClub"),("in_dtStorage","In","s:DataTable","dtStorage"),("in_dcNo","In","x:String","dcNo"),("in_divisionScope","In","x:String","in_DivisionScope"),("in_todayDate","In","sys:DateTime","Now.Date"),("out_updated","Out","x:Int32","updatedStorage"),("out_belowSet","Out","x:Int32","belowSet")], 5, "Run optional Storage Release check")
    log_result='[String.Format("Run All: STO={0}; PO={1}; Storage={2}; overdue STO={3}; missing STO={4}; no PO match={5}; preserved comments={6}", updatedSto, updatedPo, updatedStorage, overdueSto, missingSto, noMatchPo, preservedPo)]'
    log_call=_invoke_workflow("Workflows/Write_Run_Log.xaml", [
        ("io_dtLog","InOut","s:DataTable","dtLog"),("in_clubNo","In","x:String","clubNo"),("in_actionName","In","x:String",'"Run All"'),("in_resultText","In","x:String","runResult")], 6, "Append Run All audit entry")
    prep_call=_invoke_workflow("Workflows/Prepare_Column_Writes.xaml", [
        ("in_dtClub","In","s:DataTable","dtClub"),("out_dtDateColumn","Out","s:DataTable","dtWriteDate"),("out_dtCommentColumn","Out","s:DataTable","dtWriteComment"),("out_dateStartCell","Out","x:String","reviewStartCell"),("out_commentStartCell","Out","x:String","commentStartCell")], 7, "Prepare targeted review-date/comment write-back")
    optional_storage=f'''<If Condition="[in_RunStorageCheck]" sap2010:WorkflowViewState.IdRef="If_Storage">
      <If.Then><Sequence DisplayName="Storage check selected" sap2010:WorkflowViewState.IdRef="Seq_Storage">
        <If Condition="[String.IsNullOrWhiteSpace(in_StorageWorkbookPath)]" sap2010:WorkflowViewState.IdRef="If_StoragePath">
          <If.Then><Throw Exception="[New System.ArgumentException(&quot;Storage check was selected but in_StorageWorkbookPath is blank.&quot;)]" sap2010:WorkflowViewState.IdRef="Throw_StoragePath" /></If.Then>
          <If.Else><Sequence DisplayName="Read Storage data and run storage rules" sap2010:WorkflowViewState.IdRef="Seq_StorageRead">
            <ui:ReadRange WorkbookPath="[in_StorageWorkbookPath]" SheetName="Detail" AddHeaders="False" DataTable="[dtStorage]" DisplayName="Read Storage Report Detail" sap2010:WorkflowViewState.IdRef="Read_Storage" />
            {storage_call}
          </Sequence></If.Else>
        </If>
      </Sequence></If.Then>
      <If.Else><Sequence DisplayName="Storage check declined" sap2010:WorkflowViewState.IdRef="Seq_NoStorage"><WriteLine Text="Storage Release skipped because in_RunStorageCheck is False." sap2010:WorkflowViewState.IdRef="No_Storage_Line" /></Sequence></If.Else>
    </If>'''
    core=f'''<If Condition="[in_ConfirmRun]" sap2010:WorkflowViewState.IdRef="If_ConfirmRun">
      <If.Then><Sequence DisplayName="Run_All source sequence" sap2010:WorkflowViewState.IdRef="Seq_Run_All">
        {''.join(acts)}
        {validate_call}
        <WriteLine Text="[msg]" sap2010:WorkflowViewState.IdRef="Write_Validated" />
        {snapshot_call}
        <If Condition="[dtSto.Rows.Count > 0]" sap2010:WorkflowViewState.IdRef="If_STO">
          <If.Then><Sequence DisplayName="STO export has records" sap2010:WorkflowViewState.IdRef="Seq_STO">{sto_call}</Sequence></If.Then>
          <If.Else><Sequence DisplayName="STO export empty" sap2010:WorkflowViewState.IdRef="Seq_STO_Empty"><WriteLine Text="STO check: STO Export is empty." sap2010:WorkflowViewState.IdRef="STO_Empty_Line" /></Sequence></If.Else>
        </If>
        <If Condition="[dtPo.Rows.Count > 0]" sap2010:WorkflowViewState.IdRef="If_PO">
          <If.Then><Sequence DisplayName="PO export has records" sap2010:WorkflowViewState.IdRef="Seq_PO">{po_call}</Sequence></If.Then>
          <If.Else><Sequence DisplayName="PO export empty" sap2010:WorkflowViewState.IdRef="Seq_PO_Empty"><WriteLine Text="PO check: PO Export is empty." sap2010:WorkflowViewState.IdRef="PO_Empty_Line" /></Sequence></If.Else>
        </If>
        {optional_storage}
        <Assign sap2010:WorkflowViewState.IdRef="Assign_RunResult"><Assign.To><OutArgument x:TypeArguments="x:String">[runResult]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:String">{log_result}</InArgument></Assign.Value></Assign>
        {log_call}
        {prep_call}
        <If Condition="[in_WriteBack]" sap2010:WorkflowViewState.IdRef="If_WriteBack">
          <If.Then><Sequence DisplayName="Persist business outputs and audit tables" sap2010:WorkflowViewState.IdRef="Seq_WriteBack">
            <ui:WriteRange WorkbookPath="[in_WorkbookPath]" SheetName="New Club data" StartingCell="[reviewStartCell]" DataTable="[dtWriteDate]" AddHeaders="False" DisplayName="Write review dates" sap2010:WorkflowViewState.IdRef="Write_ReviewDate" />
            <ui:WriteRange WorkbookPath="[in_WorkbookPath]" SheetName="New Club data" StartingCell="[commentStartCell]" DataTable="[dtWriteComment]" AddHeaders="False" DisplayName="Write CSC2 comments" sap2010:WorkflowViewState.IdRef="Write_Comments" />
            <ui:WriteRange WorkbookPath="[in_WorkbookPath]" SheetName="Comment Snapshot" StartingCell="A1" DataTable="[dtSnapshot]" AddHeaders="True" DisplayName="Persist Comment Snapshot" sap2010:WorkflowViewState.IdRef="Write_Snapshot" />
            <ui:WriteRange WorkbookPath="[in_WorkbookPath]" SheetName="Run Log" StartingCell="A1" DataTable="[dtLog]" AddHeaders="True" DisplayName="Persist Run Log" sap2010:WorkflowViewState.IdRef="Write_RunLog" />
          </Sequence></If.Then>
          <If.Else><Sequence DisplayName="Dry run - no workbook writes" sap2010:WorkflowViewState.IdRef="Seq_DryRun"><WriteLine Text="Write-back disabled. DataTables were evaluated in memory; workbook remains unchanged." sap2010:WorkflowViewState.IdRef="Dry_Run_Line" /></Sequence></If.Else>
        </If>
        <WriteLine Text="[runResult]" sap2010:WorkflowViewState.IdRef="Write_Completion" />
      </Sequence></If.Then>
      <If.Else><Sequence DisplayName="Run cancelled by caller" sap2010:WorkflowViewState.IdRef="Seq_Cancelled"><WriteLine Text="Run_All cancelled. in_ConfirmRun was False; no processing or writes occurred." sap2010:WorkflowViewState.IdRef="Cancelled_Line" /></Sequence></If.Else>
    </If>'''
    return f'''<Activity mc:Ignorable="sap sap2010" x:Class="Main" sap2010:WorkflowViewState.IdRef="ActivityBuilder_1"
 xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities"
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:sap="http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation"
 xmlns:sap2010="http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation"
 xmlns:ui="http://schemas.uipath.com/workflow/activities"
 xmlns:scg="clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib"
 xmlns:sco="clr-namespace:System.Collections.ObjectModel;assembly=System.Private.CoreLib"
 xmlns:s="clr-namespace:System.Data;assembly=System.Data.Common"
 xmlns:sys="clr-namespace:System;assembly=System.Private.CoreLib">
  <TextExpression.NamespacesForImplementation><sco:Collection x:TypeArguments="x:String"><x:String>System</x:String><x:String>System.Collections.Generic</x:String><x:String>System.Data</x:String><x:String>System.Linq</x:String><x:String>System.Text</x:String><x:String>System.Text.RegularExpressions</x:String><x:String>System.Globalization</x:String></sco:Collection></TextExpression.NamespacesForImplementation>
  <TextExpression.ReferencesForImplementation><sco:Collection x:TypeArguments="AssemblyReference"><AssemblyReference>System</AssemblyReference><AssemblyReference>System.Data.Common</AssemblyReference><AssemblyReference>UiPath.System.Activities</AssemblyReference><AssemblyReference>UiPath.Excel.Activities</AssemblyReference></sco:Collection></TextExpression.ReferencesForImplementation>
  <x:Members>
    <x:Property Name="in_WorkbookPath" Type="InArgument(x:String)" />
    <x:Property Name="in_DivisionScope" Type="InArgument(x:String)" />
    <x:Property Name="in_ConfirmRun" Type="InArgument(x:Boolean)" />
    <x:Property Name="in_RunStorageCheck" Type="InArgument(x:Boolean)" />
    <x:Property Name="in_StorageWorkbookPath" Type="InArgument(x:String)" />
    <x:Property Name="in_WriteBack" Type="InArgument(x:Boolean)" />
  </x:Members>
  <Sequence DisplayName="Run_All — SYNC New Club supply workflow" sap2010:WorkflowViewState.IdRef="Sequence_Main">
    <Sequence.Variables>{var_xml}</Sequence.Variables>
    {core}
  </Sequence>
</Activity>'''


CLEAR_DATA_VB = COMMON_VB + r'''
If dtClub Is Nothing OrElse dtSto Is Nothing OrElse dtPo Is Nothing Then Throw New Exception("New Club, STO Export and PO Export DataTables are required.")
Dim selection As String = normalizeText(clearSelection)
If Not confirmed Then
    outStoRowsCleared = 0 : outPoRowsCleared = 0 : outClubRowsCleared = 0 : outCommentCellsCleared = 0
    outMessage = "Clear Data cancelled. No rows or comments were changed."
Else
    Dim doSto As Boolean = selection.Contains("1"), doPo As Boolean = selection.Contains("2"), doClub As Boolean = selection.Contains("3"), doComments As Boolean = selection.Contains("4")
    If Not (doSto OrElse doPo OrElse doClub OrElse doComments) Then Throw New Exception("Select one or more valid Clear Data options: 1=STO, 2=PO, 3=New Club data, 4=New Club review dates/comments.")
    If doClub Then doComments = False
    outStoRowsCleared = If(doSto, dtSto.Rows.Count, 0)
    outPoRowsCleared = If(doPo, dtPo.Rows.Count, 0)
    outClubRowsCleared = If(doClub, dtClub.Rows.Count, 0)
    outCommentCellsCleared = 0
    If doSto Then dtSto.Clear()
    If doPo Then dtPo.Clear()
    If doClub Then
        dtClub.Clear()
    ElseIf doComments Then
        Dim cDate As Integer = findColumn(dtClub, New String() {"Date of Last Review"})
        Dim cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"})
        If cDate < 0 OrElse cComment < 0 Then Throw New Exception("New Club data is missing Date of Last Review or CSC2 Comments.")
        For Each r As System.Data.DataRow In dtClub.Rows
            r(cDate) = System.DBNull.Value : r(cComment) = System.DBNull.Value
            outCommentCellsCleared += 2
        Next
    End If
    outMessage = "Clear Data completed. Headers/schema preserved; selected row data or review/comment values cleared. Formatting is not altered by this DataTable workflow."
End If
'''

EMAIL_ANALYSTS_VB = COMMON_VB + r'''
If dtClub Is Nothing Then Throw New Exception("New Club data DataTable is required.")
If dtCategoryMap Is Nothing OrElse dtCategoryMap.Columns.Count < 7 Then Throw New Exception("Category Map must contain the seven source columns: Art_AH2, Art_AH3, Art_AH4, CSC1 Manager, CSC1 Planner, CSC1 IA and CSC2 Analyst.")
Dim cArt As Integer = findColumn(dtClub, New String() {"Article"}), cComment As Integer = findColumn(dtClub, New String() {"CSC2 Comments"}), cDiv As Integer = findColumn(dtClub, New String() {"Division", "Art_AH2", "AH2"}), cMc As Integer = findColumn(dtClub, New String() {"Merch Cat", "Merchandise Category", "Art_AH4 - Merch Cat", "MerchCat"})
If cArt < 0 OrElse cComment < 0 Then Throw New Exception("New Club data must contain Article and CSC2 Comments.")
Dim sumNorm As System.Func(Of Object, String) = Function(v) normalizeText(v).ToUpperInvariant()
Dim sumDivKey As System.Func(Of Object, String) = Function(v)
    Dim t As String = System.Text.RegularExpressions.Regex.Replace(sumNorm(v), "[^A-Z0-9]", "")
    If t = "AHGM" OrElse t = "GM" Then t = "AHGENERALMERCHANDISE"
    Return t
End Function
Dim sumLetters As System.Func(Of String, String) = Function(v) System.Text.RegularExpressions.Regex.Replace(If(v, "").ToUpperInvariant(), "[^A-Z]", "")
Dim levDist As System.Func(Of String, String, Integer) = Function(a, b)
    If Math.Abs(a.Length - b.Length) > 2 Then Return 99
    Dim d(a.Length, b.Length) As Integer
    For i As Integer = 0 To a.Length : d(i, 0) = i : Next
    For j As Integer = 0 To b.Length : d(0, j) = j : Next
    For i As Integer = 1 To a.Length
        For j As Integer = 1 To b.Length
            Dim cost As Integer = If(a(i - 1) = b(j - 1), 0, 1)
            d(i, j) = Math.Min(Math.Min(d(i - 1, j) + 1, d(i, j - 1) + 1), d(i - 1, j - 1) + cost)
        Next
    Next
    Return d(a.Length, b.Length)
End Function
Dim person As System.Func(Of Object, String) = Function(v)
    Dim t As String = normalizeText(v)
    If t = "" OrElse t.ToUpperInvariant() = "NA" OrElse t.ToUpperInvariant() = "N/A" OrElse t.ToUpperInvariant() = "#N/A" OrElse t = "0" Then Return "Not owned by CSC2"
    Return t
End Function
Dim cache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
Dim mapAnalyst As System.Func(Of String, String, String) = Function(divText, merchText)
    Dim dk As String = sumDivKey(divText), mk As String = sumNorm(merchText), key As String = dk & "|" & mk
    If cache.ContainsKey(key) Then Return cache(key)
    Dim best As Integer = -1
    If mk.Length > 0 Then
        For i As Integer = 0 To dtCategoryMap.Rows.Count - 1
            If sumDivKey(dtCategoryMap.Rows(i)(0)) = dk AndAlso sumNorm(dtCategoryMap.Rows(i)(2)) = mk Then best = i : Exit For
        Next
        If best < 0 Then
            For i As Integer = 0 To dtCategoryMap.Rows.Count - 1
                If sumDivKey(dtCategoryMap.Rows(i)(0)) = dk AndAlso sumNorm(dtCategoryMap.Rows(i)(1)) = mk Then best = i : Exit For
            Next
        End If
        If best < 0 Then
            Dim wanted As String = sumLetters(mk)
            For i As Integer = 0 To dtCategoryMap.Rows.Count - 1
                If sumDivKey(dtCategoryMap.Rows(i)(0)) <> dk Then Continue For
                Dim candidate As String = sumLetters(sumNorm(dtCategoryMap.Rows(i)(2)))
                If candidate = "NA" OrElse candidate.Length = 0 Then candidate = sumLetters(sumNorm(dtCategoryMap.Rows(i)(1)))
                If candidate.Length >= 5 AndAlso wanted.Length >= 5 AndAlso (candidate = wanted OrElse wanted.Contains(candidate) OrElse candidate.Contains(wanted) OrElse levDist(candidate, wanted) <= 2) Then best = i : Exit For
            Next
        End If
    End If
    Dim owner As String = If(best < 0, "Not owned by CSC2", person(dtCategoryMap.Rows(best)(6)))
    cache(key) = owner
    Return owner
End Function
Dim counts As New Dictionary(Of String, Integer())(StringComparer.OrdinalIgnoreCase)
Dim ccByAnalyst As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
outEmailSummary = New System.Data.DataTable("AnalystEmailQueue")
For Each col As String In New String() {"Analyst", "To", "CC", "Partial STO", "Manual STO", "Blank", "Total", "Subject", "HTMLBody", "SourceRowIndexes", "DeliveryState"} : outEmailSummary.Columns.Add(col, GetType(String)) : Next
outEmailRows = dtClub.Clone()
If Not outEmailRows.Columns.Contains("Assigned CSC2 Analyst") Then outEmailRows.Columns.Add("Assigned CSC2 Analyst", GetType(String))
If Not outEmailRows.Columns.Contains("Source Excel Row") Then outEmailRows.Columns.Add("Source Excel Row", GetType(Integer))
Dim filterText As String = normalizeText(analystFilter)
If filterText.Length = 0 Then filterText = "ALL"
For rowIndex As Integer = 0 To dtClub.Rows.Count - 1
    Dim r As System.Data.DataRow = dtClub.Rows(rowIndex)
    Dim art As String = articleKey(r(cArt))
    If art.Length = 0 Then Continue For
    Dim divText As String = If(cDiv >= 0, normalizeText(r(cDiv)), "")
    If Not inScope(divText, divisionScope) Then Continue For
    Dim comment As String = normalizeText(r(cComment)).ToLowerInvariant()
    Dim cat As Integer = 8
    If comment.Length = 0 Then
        cat = 9
    ElseIf comment.StartsWith("sto manifested") Then
        cat = 0
    ElseIf comment.StartsWith("partial") Then
        cat = 1
    ElseIf comment.StartsWith("need manual sto") Then
        cat = 2
    ElseIf comment.Contains("stf") Then
        cat = 7
    ElseIf comment.StartsWith("sto ") Then
        cat = If(comment.Contains("is passed sto not delivered"), 3, 4)
    ElseIf comment.StartsWith("po ") AndAlso comment.Contains(" due on ") Then
        cat = 5
    ElseIf comment.StartsWith("storage dc") Then
        cat = 6
    End If
    If cat <> 1 AndAlso cat <> 2 AndAlso cat <> 9 Then Continue For
    Dim merchText As String = If(cMc >= 0, normalizeText(r(cMc)), "")
    Dim analyst As String = mapAnalyst(divText, merchText)
    If analyst = "Not owned by CSC2" Then Continue For
    If Not String.Equals(filterText, "ALL", StringComparison.OrdinalIgnoreCase) AndAlso analyst.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
    If Not counts.ContainsKey(analyst) Then counts(analyst) = New Integer() {0, 0, 0}
    If cat = 1 Then counts(analyst)(0) += 1
    If cat = 2 Then counts(analyst)(1) += 1
    If cat = 9 Then counts(analyst)(2) += 1
    If Not ccByAnalyst.ContainsKey(analyst) Then ccByAnalyst(analyst) = ""
    Dim divUpper As String = divText.ToUpperInvariant(), cc As String = ""
    If divUpper.Contains("SUNDRIES") OrElse divUpper.Contains("GROCERY") Then
        cc = "banjan@bjs.com"
    ElseIf divUpper.Contains("GENERAL MERCHANDISE") OrElse sumDivKey(divText) = "AHGENERALMERCHANDISE" Then
        cc = "Vivek Udayshankar"
    End If
    If cc.Length > 0 AndAlso Not ccByAnalyst(analyst).ToLowerInvariant().Contains(cc.ToLowerInvariant()) Then ccByAnalyst(analyst) = ccByAnalyst(analyst) & If(ccByAnalyst(analyst).Length > 0, "; ", "") & cc
    Dim nr As System.Data.DataRow = outEmailRows.NewRow()
    For c As Integer = 0 To dtClub.Columns.Count - 1 : nr(c) = r(c) : Next
    nr("Assigned CSC2 Analyst") = analyst : nr("Source Excel Row") = rowIndex + 2
    outEmailRows.Rows.Add(nr)
Next
For Each kv As KeyValuePair(Of String, Integer()) In counts
    Dim total As Integer = kv.Value(0) + kv.Value(1) + kv.Value(2)
    Dim firstName As String = kv.Key.Split(" "c)(0)
    Dim subject As String = "Action Needed - New Club " & clubNo & " - " & total.ToString() & " lines to action - " & todayDate.ToString("MM/dd/yyyy")
    Dim body As String = "<div style=""font-family:Calibri,Arial,sans-serif;font-size:11pt"">Hi " & System.Net.WebUtility.HtmlEncode(firstName) & ",<br><br>Please review the attached New Club " & System.Net.WebUtility.HtmlEncode(clubNo) & " lines that need action and update the comments.<br><br>Partial STO: <b>" & kv.Value(0).ToString() & "</b> &nbsp;|&nbsp; Manual STO: <b>" & kv.Value(1).ToString() & "</b> &nbsp;|&nbsp; Blank: <b>" & kv.Value(2).ToString() & "</b> &nbsp;|&nbsp; Total: <b>" & total.ToString() & "</b><br><br>Thanks,</div>"
    Dim rowIndexes As New System.Text.StringBuilder()
    For Each rr As System.Data.DataRow In outEmailRows.Rows
        If normalizeText(rr("Assigned CSC2 Analyst")) = kv.Key Then
            If rowIndexes.Length > 0 Then rowIndexes.Append(";")
            rowIndexes.Append(normalizeText(rr("Source Excel Row")))
        End If
    Next
    outEmailSummary.Rows.Add(kv.Key, kv.Key, ccByAnalyst(kv.Key), kv.Value(0).ToString(), kv.Value(1).ToString(), kv.Value(2).ToString(), total.ToString(), subject, body, rowIndexes.ToString(), "Ready for Outlook draft/send configuration")
Next
'''

def _procedure_workflows() -> dict[str, tuple[str, list[tuple[str,str,str]], str, str]]:
    """Workflow path -> (display, args, VB code, purpose)."""
    return {
        "Workflows/Validate_Run_Inputs.xaml": ("Validate Action Sheet and input workbook", [("in_dtAction","In","s:DataTable"),("in_dtClub","In","s:DataTable"),("out_clubNo","Out","x:String"),("out_dcNo","Out","x:String"),("out_truckDate","Out","sys:DateTime"),("out_validationMessage","Out","x:String")], VALIDATE_VB, "Reads New Club Number, DC Number and Last truck date from Action Sheet; validates required New Club headers before any updates."),
        "Workflows/STO_Check.xaml": ("STO aggregation and row decision tree", [("io_dtClub","InOut","s:DataTable"),("in_dtSto","In","s:DataTable"),("in_clubNo","In","x:String"),("in_divisionScope","In","x:String"),("in_todayDate","In","sys:DateTime"),("in_truckDate","In","sys:DateTime"),("out_updated","Out","x:Int32"),("out_overdue","Out","x:Int32"),("out_missing","Out","x:Int32")], STO_VB, "Site/deletion safety checks; Article aggregation; delivered/open quantities; partial/manual/open/manifested comments; scope-aware review dates."),
        "Workflows/PO_Check.xaml": ("PO filter, grouping and comment precedence", [("io_dtClub","InOut","s:DataTable"),("in_dtPo","In","s:DataTable"),("in_dcNo","In","x:String"),("in_divisionScope","In","x:String"),("in_todayDate","In","sys:DateTime"),("in_truckDate","In","sys:DateTime"),("out_updated","Out","x:Int32"),("out_noMatch","Out","x:Int32"),("out_preserved","Out","x:Int32")], PO_VB, "Validates DC, filters deletion/site/date rows, orders PO lines by due date, groups by Article and writes only blank/old PO comments with Generic Article fallback."),
        "Workflows/Storage_Release.xaml": ("Storage DC available-to-release comments", [("io_dtClub","InOut","s:DataTable"),("in_dtStorage","In","s:DataTable"),("in_dcNo","In","x:String"),("in_divisionScope","In","x:String"),("in_todayDate","In","sys:DateTime"),("out_updated","Out","x:Int32"),("out_belowSet","Out","x:Int32")], STORAGE_VB, "Reads DC-specific on-site units/cases and fills only blank comments; counts rows where on-site units are below Initial Set Quantity."),
        "Workflows/Other_Checks.xaml": ("HVS, Delisted and Initial Set Quantity checks", [("io_dtClub","InOut","s:DataTable"),("in_divisionScope","In","x:String"),("in_todayDate","In","sys:DateTime"),("out_hvs","Out","x:Int32"),("out_delisted","Out","x:Int32"),("out_noQty","Out","x:Int32")], OTHER_VB, "Only blank comments; precedence Placement=HVS, Listing Status=Delisted, then missing/zero Initial Set Quantity."),
        "Workflows/Club_Transfer.xaml": ("Club transfer stock ranking", [("io_dtClub","InOut","s:DataTable"),("in_dtInvStat","In","s:DataTable"),("in_dcNo","In","x:String"),("in_newClubNo","In","x:String"),("in_divisionScope","In","x:String"),("in_allowedClubs","In","x:String"),("in_todayDate","In","sys:DateTime"),("out_suggested","Out","x:Int32"),("out_none","Out","x:Int32")], TRANSFER_VB, "Consumes a pre-exported ZPRR_INVSTAT table; excludes New Club/DC; optionally filters clubs; ranks up to three candidate source clubs by set-quantity coverage, WOS and OH."),
        "Workflows/Appointment_Date.xaml": ("ME23N appointment date merge", [("io_dtClub","InOut","s:DataTable"),("in_dtApptConfirmations","In","s:DataTable"),("in_divisionScope","In","x:String"),("in_todayDate","In","sys:DateTime"),("out_found","Out","x:Int32"),("out_noDate","Out","x:Int32"),("out_skipped","Out","x:Int32")], APPT_VB, "Extracts the longest document-number digit run (>=8) from STO/PO comments, looks it up in a confirmation export, replaces previous appointment suffix, and dates the review."),
        "Workflows/Replacement_Article.xaml": ("Replacement article STO to PO to Storage", [("io_dtClub","InOut","s:DataTable"),("in_dtSto","In","s:DataTable"),("in_dtPo","In","s:DataTable"),("in_dtStorage","In","s:DataTable"),("in_dcNo","In","x:String"),("in_newClubNo","In","x:String"),("in_divisionScope","In","x:String"),("in_todayDate","In","sys:DateTime"),("in_truckDate","In","sys:DateTime"),("out_updated","Out","x:Int32")], REPLACEMENT_VB, "Processes comments containing Replacement article N; checks replacement Article supply in STO, then valid PO window, then storage stock, else no STO or PO."),
        "Workflows/Fetch_STF_Article.xaml": ("STF replacement recommendation ranking", [("io_dtClub","InOut","s:DataTable"),("in_dtAllocation","In","s:DataTable"),("in_clubNo","In","x:String"),("in_divisionScope","In","x:String"),("in_todayDate","In","sys:DateTime"),("out_suggestions","Out","x:Int32"),("out_noPog","Out","x:Int32"),("out_noCandidates","Out","x:Int32")], STF_VB, "Reads the BJs Allocation Report; uses same POG candidates first, then same AH6/MerchCat; excludes other STF articles and ranks three candidates using hierarchy and description-word similarity."),
        "Workflows/Update_Tracking_Report.xaml": ("Tracking Report composite-key update", [("in_dtClub","In","s:DataTable"),("io_dtTracking","InOut","s:DataTable"),("out_matched","Out","x:Int32"),("out_notFound","Out","x:Int32")], TRACKING_VB, "Matches normalized Article plus Article Description; updates all matching tracker rows with review date/comment; blank source comments are ignored."),
        "Workflows/Refresh_Summary.xaml": ("Full Summary data model", [("in_dtClub","In","s:DataTable"),("in_dtCategoryMap","In","s:DataTable"),("in_divisionFilter","In","x:String"),("in_todayDate","In","sys:DateTime"),("out_commentSummary","Out","s:DataTable"),("out_divisionSummary","Out","s:DataTable"),("out_analystSummary","Out","s:DataTable"),("out_actionDetails","Out","s:DataTable")], SUMMARY_VB, "Implements source comment categories 0-9, category totals/percentages, division totals/manifested/to-action/other counts, analyst action/all/blank/stale counts and actionable line details using Category Map matching."),
        "Workflows/Take_Daily_Snapshot.xaml": ("Daily comment snapshot", [("in_dtClub","In","s:DataTable"),("io_dtSnapshot","InOut","s:DataTable"),("in_todayDate","In","sys:DateTime")], SNAPSHOT_VB, "Stores one daily Article/category/comment snapshot and refreshes today's snapshot when rerun."),
        "Workflows/Write_Run_Log.xaml": ("Append run audit entry", [("io_dtLog","InOut","s:DataTable"),("in_clubNo","In","x:String"),("in_actionName","In","x:String"),("in_resultText","In","x:String")], LOG_VB, "Appends date/time, Windows user, club, action and result into a Run Log DataTable."),
        "Workflows/UiOnly_Replacement.xaml": ("Replace Excel-only progress and dashboard presentation", [("in_sourceUiAction","In","x:String"),("out_replacementNote","Out","x:String")], 'outReplacementNote = "Excel-only decorative action: " & sourceUiAction & ". Replaced by UiPath logs, counters and workflow status; business data is unchanged."', "Maps Excel-only centering, protection, color, progress animation, toast and dashboard drawing procedures to an explicit, safe status/log boundary."),
        "Workflows/Prepare_Column_Writes.xaml": ("Prepare minimal, non-destructive workbook write-back", [("in_dtClub","In","s:DataTable"),("out_dtDateColumn","Out","s:DataTable"),("out_dtCommentColumn","Out","s:DataTable"),("out_dateStartCell","Out","x:String"),("out_commentStartCell","Out","x:String")], PREPARE_WRITE_VB, "Builds one-column DataTables and resolves column letters so only Date of Last Review and CSC2 Comments are written back."),
        "Workflows/Clear_Data.xaml": ("Clear selected workbook data safely", [("io_dtClub","InOut","s:DataTable"),("io_dtSto","InOut","s:DataTable"),("io_dtPo","InOut","s:DataTable"),("in_clearSelection","In","x:String"),("in_confirmed","In","x:Boolean"),("out_stoRowsCleared","Out","x:Int32"),("out_poRowsCleared","Out","x:Int32"),("out_clubRowsCleared","Out","x:Int32"),("out_commentCellsCleared","Out","x:Int32"),("out_message","Out","x:String")], CLEAR_DATA_VB, "Implements Clear_Data options 1-4: clear STO rows, PO rows, all New Club rows, or only review-date/comment fields. Requires explicit confirmation and preserves DataTable schemas."),
        "Workflows/Email_Analysts.xaml": ("Build analyst email queue and attachment rows", [("in_dtClub","In","s:DataTable"),("in_dtCategoryMap","In","s:DataTable"),("in_clubNo","In","x:String"),("in_divisionScope","In","x:String"),("in_analystFilter","In","x:String"),("in_todayDate","In","sys:DateTime"),("out_emailSummary","Out","s:DataTable"),("out_emailRows","Out","s:DataTable")], EMAIL_ANALYSTS_VB, "Reproduces SumCat/MapPeople selection for Partial STO, Manual STO and blank rows; groups by CSC2 Analyst, applies the source CC rules, and builds HTML email envelopes plus exact attachment rows. Outlook sending/drafting is a downstream configured UiPath activity."),
    }


def _sap_adapter_xaml(class_name: str, display_name: str, transaction: str) -> str:
    code = f'''Throw New System.InvalidOperationException("SAP adapter {transaction} is a capture-required integration boundary. Capture the approved SAP GUI window and selectors in UiPath Studio first; no SAP selector or credential is fabricated by this generator.")'''
    return workflow_xaml("SAP_Adapters_" + class_name, display_name, "SAP UI automation procedures", [], code)


def _classify_procedure(symbol: dict[str, Any]) -> tuple[str,str,str]:
    name = str(symbol.get("name", ""))
    n = re.sub(r"[^a-z0-9]", "", name.lower())
    src = str(symbol.get("source_file", ""))
    mod = re.sub(r"[^a-z0-9]", "", Path(src).stem.lower())

    # Explicit macro entry points and user actions.
    if n in {"cleardata", "clearbelowheader", "lastdatarow"}:
        return "Workflows/Clear_Data.xaml", "Business data reset / helper", "Implements Clear Data options 1–4 in a guarded DataTable workflow; confirmation is required."
    if n in {"emailanalysts", "safename", "htmlesc"}:
        return "Workflows/Email_Analysts.xaml", "Business reporting / communications", "Builds assigned analyst groups, action counts, source CC rules, HTML body and row sets for an attachment; actual Outlook draft/send is configured separately."
    if n == "runfullclubcheck":
        return "SAP_Adapters/Pull_From_SAP.xaml", "SAP orchestration / selector capture required", "Maps ME2M New Club STO → STO Check → remaining blank Articles → ME2M DC PO → PO Check → optional Storage. Live GUI fields are environment-specific."
    if n in {"runall", "hasdata"}:
        return "Main.xaml", "Run_All orchestration", "Main follows source Run_All: daily snapshot → conditional STO → conditional PO → optional Storage → run log → targeted write-back."

    # Module 8 presentation helpers (keep calculations separate from drawing).
    summary_visual = {
        "removesheetprotection", "styleactionsheet", "applytabcolours", "colourtab", "wbstructlock", "wbstructunlock",
        "protectactionsheet", "centeractionsheet", "sumcleardash", "sumdashband", "summakebtn", "sumtile", "sumchart",
        "scol", "stext", "sumband", "sumsection", "sumheader", "sumalign", "sumbody", "sumtotal", "sumdivrow", "sumactrow",
    }
    if mod == "module8" and n in summary_visual:
        return "Workflows/UiOnly_Replacement.xaml", "Excel dashboard / presentation", "Replaces Excel styling, protection, chart/shape drawing and layout with structured output and log/status text; core summary numbers remain in Refresh_Summary.xaml."
    if mod == "module8" and n == "takedailysnapshot":
        return "Workflows/Take_Daily_Snapshot.xaml", "Audit snapshot", "Stores today's Article/category/comment snapshot and refreshes today's records on rerun."
    if mod == "module8" and n in {"refreshsummary", "sumcat", "mappeople", "sumperson", "sumnorm", "sumdivkey", "sumletters", "sumlev", "sortedkeys", "newcounts", "addcounts", "sumpickdiv", "dyncommentsummary"}:
        return "Workflows/Refresh_Summary.xaml", "Reporting / business calculation", "Builds source 0–9 comment categories, division/analyst ownership and workload aggregates; visual rendering is replaced by output DataTables."
    if mod == "module8" and n == "dynlog":
        return "Workflows/Write_Run_Log.xaml", "Audit logging", "Appends date/time, user, club, action and result into the run audit table."

    # SAP screen and grid interaction: selectors are not inferred from source code.
    sap_ui = {
        "getsapsession", "waitsap", "sapvkey", "gotcode", "loadvariant", "setclipboard", "setmultiselect",
        "sapfind", "sapdate", "closepopups", "sbartext", "findgrid", "readme2m", "readinvstat",
        "olewaitpromptoff", "olewaitprompton", "sapvkeyw", "sbartype", "sapfindbytext", "sapfindbytip", "sapscan",
        "sapfieldbylabel", "saptry", "findanygrid", "apptdatefromme23n", "readconfdate", "pressotherpo",
        "findeditableinwindow", "expanditemdetail", "collecttabs", "pickconftab",
    }
    if n in sap_ui:
        return "SAP_Adapters/Capture_SAP_UI.xaml", "Environment-specific SAP UI", "SAP GUI screen/grid integration; capture the real control targets in UiPath Studio and do not guess selectors or credentials."

    # SAP export preparation (non-UI) procedures belong to the pull adapter contract.
    if n in {"writeexport", "collectarticles", "collectblankcount", "clearexporttab", "filterblanks", "findsheet", "fourdigit", "celltext", "normart", "normhdr", "hdridx", "colkind", "convertcell", "parsesapdate", "sapnum", "fourdigitorblank", "lastrowincol"} and mod == "autoclubcheck":
        return "SAP_Adapters/Pull_From_SAP.xaml", "SAP export preparation helper", "Prepares/normalizes SAP result tables and Article filters for the source ME2M orchestration; it does not interact with a screen by itself."

    # Former partial-STO PO enrichment code is retained in the source but its caller is commented out/removed.
    if mod == "autoclubcheck" and n in {"buildbestpo", "writepartialcomments"}:
        return "Workflows/PO_Check.xaml", "Retired legacy helper", "Legacy partial-STO next-PO enrichment routine. Source Pull from SAP has this call branch explicitly removed; retained here for traceability but not invoked by Main."

    # Standalone business workflows.
    if mod == "autoclubcheck" and n in {"otherchecks", "ocwrite"}:
        return "Workflows/Other_Checks.xaml", "Business exception rules", "Blank-comment-only decision order: HVS, then Delisted, then missing/zero Initial Set Quantity."
    if mod == "autoclubcheck" and n in {"clubtransfer", "isneedtransfer", "pickclub", "picktopclubs", "transferless", "parseclublist"}:
        return "Workflows/Club_Transfer.xaml", "Business transfer selection / ranking", "Selects Need Club transfer rows and ranks candidate source clubs from on-hand stock/coverage; raw stock extraction is in the SAP adapter."
    if mod == "autoclubcheck" and n in {"appointmentdate", "isapptflag", "extractdoc", "stripappt", "hasstoorpo", "doctype", "collectbykind"}:
        return "Workflows/Appointment_Date.xaml", "Appointment-date business logic", "Parses STO/PO document IDs in comments and merges confirmation dates from a ME23N result table; screen navigation is a separate SAP adapter."
    if mod == "autoclubcheck" and n in {"replacementarticle", "replacementof", "repstoinfo", "isopenpartial"}:
        return "Workflows/Replacement_Article.xaml", "Replacement supply resolution", "Extracts replacement Article and checks its supply in order STO → eligible PO → Storage → no supply."
    if mod == "module6" and n in {"fetchstfarticle", "restoreclub", "findreport", "newestreport", "readreporttab", "filterreporttab", "bestreplacements", "score", "wordsim", "words", "poglist", "sameclub", "leadingdigits", "txtof", "hdrcol", "readcol", "findbyname", "resolvesheet"}:
        return "Workflows/Fetch_STF_Article.xaml", "STF recommendation logic / helper", "Finds Allocation Report data, prefers same POG then same AH6/MerchCat and ranks up to three replacement suggestions."
    if mod == "autoclubcheck" and n == "hasstoorpo":
        return "Workflows/Appointment_Date.xaml", "Appointment-date helper", "Confirms whether the row comment contains a qualifying STO/PO reference before appointment-date lookup."

    if mod == "autoclubcheck" and n in {"readinvstat"}:
        return "SAP_Adapters/Capture_SAP_UI.xaml", "SAP inventory query", "ZPRR_INVSTAT screen/grid extraction; output is consumed by Club_Transfer.xaml."

    # Core supply check workflows and their same-module helpers.
    if mod == "module3":
        if n == "hideworksheet1":
            return "Workflows/UiOnly_Replacement.xaml", "Excel presentation helper", "Hides the temporary work sheet; no business result is altered."
        return "Workflows/STO_Check.xaml", "STO business logic / helper", "Article-level STO totals, delivered/open quantities, due-date classification, comment decision tree and its header/date/text helpers."
    if mod == "module6":
        if n == "clearfilters":
            return "Workflows/PO_Check.xaml", "PO data preparation helper", "Source clears worksheet filter state; target filters DataTables in memory."
        return "Workflows/PO_Check.xaml", "PO business logic / helper", "DC/site/deletion/date validation, valid PO aggregation, due-date ordering, Article/Generic Article matching and safe comment preservation."
    if mod == "storagerelease":
        return "Workflows/Storage_Release.xaml", "Storage business logic / helper", "Finds DC-specific grouped On Site Units/cases columns, aggregates stock by Article and updates blank comments."
    if mod == "module7":
        return "Workflows/Update_Tracking_Report.xaml", "Tracking report matching / helper", "Normalizes Article + Description composite key, selects matching report sheet, and updates every matching target row."
    if mod == "module1":
        if n == "dynlog":
            return "Workflows/Write_Run_Log.xaml", "Audit logging", "Adds a run audit entry."
        if n in {"dyninscope", "dynscopeflags", "dynscopelabel", "dynpickdivisions"}:
            return "Workflows/Validate_Run_Inputs.xaml", "Division scope / input helper", "Resolves selected division scope and exact row membership."
        if n in {"dynshowall", "dynprogressstart", "dynprogressstep", "dynprogressdetail", "dynprogressend", "prgdelete", "prgcreate", "prgtext", "prgdraw", "prgspinchar", "prgbuildindicator", "prganimateto", "dynquip", "dyndonequip", "dyndonetoast", "dyntoastreset", "dyntoaststop", "dyntoastpump", "prgpause", "prglighten", "prgrefresh"}:
            return "Workflows/UiOnly_Replacement.xaml", "Excel presentation / status helper", "Progress drawings, toast messages, centering and UI-only filter state are replaced with UiPath status/log output."
        return "Workflows/Prepare_Column_Writes.xaml", "Shared workbook / normalization helper", "Dynamic sheet/header/cell discovery, Article/site normalization, deletion-flag parsing, safe text conversion and minimal write-back helpers are consolidated."
    if mod == "thisworkbook" or mod.startswith("sheet"):
        if n == "workbooksheetactivate":
            return "Workflows/Refresh_Summary.xaml", "Workbook event / Summary refresh", "The source event refreshes Summary when the Summary sheet is activated; target refresh is an explicit workflow invocation."
        return "Workflows/UiOnly_Replacement.xaml", "Workbook event / presentation", "Workbook-open/window resize/worksheet code-behind is replaced by explicit initialization or status behavior."
    if mod == "autoclubcheck":
        return "SAP_Adapters/Capture_SAP_UI.xaml", "SAP UI / workflow helper", "SAP transaction interaction or data-specific helper; see mapping note and SAP capture checklist."

    return "Workflows/Prepare_Column_Writes.xaml", "Shared utility / helper", "Consolidated into reusable normalization, header-resolution or safe-write helpers in the owning business workflow."

def _mapping_csv(symbols: list[dict[str, Any]]) -> str:
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Source module/file", "Procedure", "Kind", "Source line", "End line", "Signature", "Detected calls", "Target XAML / artifact", "Migration classification", "Mapping note"])
    seen = set()
    for s in symbols:
        key = (s.get("source_file"), s.get("name"), s.get("line"), s.get("signature"))
        if key in seen: continue
        seen.add(key)
        path, cls, note = _classify_procedure(s)
        w.writerow([s.get("source_file",""), s.get("name",""), s.get("kind",""), s.get("line",0), s.get("end_line",0), s.get("signature",""), "; ".join(s.get("calls",[]) or []), path, cls, note])
    return buf.getvalue()


def _procedure_catalog_md(symbols: list[dict[str, Any]]) -> str:
    rows = ["# Extracted VBA procedure catalogue", "", f"Total detected procedures/functions: **{len(symbols)}**", "", "Each source procedure is mapped to a logical UiPath workflow rather than generating one XAML file per VBA helper. The CSV is the complete traceability register.", "", "| Module | Procedure | Lines | Target workflow | Classification |", "|---|---|---:|---|---|"]
    for s in symbols:
        path, cls, _ = _classify_procedure(s)
        rows.append(f"| {Path(s.get('source_file','')).stem} | `{s.get('name','')}` | {s.get('line',0)}–{s.get('end_line',0)} | `{path}` | {cls} |")
    return "\n".join(rows) + "\n"


def _design_md(symbols: list[dict[str, Any]], is_sync: bool) -> str:
    workflow_count = len(_procedure_workflows()) + 3
    return f'''# SYNC VBA → UiPath migration design

## Migration summary

- Detected source procedures/functions: **{len(symbols)}**
- Generated XAML count: **{workflow_count}**
- Target: UiPath Windows project using VB expressions
- Generation: deterministic and offline; no model API key is required
- Design: one orchestrator plus cohesive capability workflows. Small VBA helpers are consolidated into their owning workflow and trace-mapped in `Migration/Function_Mapping.csv`.

## Why this is not 249 XAML files

The VBA project contains 249 procedures, including small formatting helpers, normalization functions, event handlers and private functions called by larger business procedures. A separate XAML for every helper would create a fragile chain of tiny workflows and excessive argument plumbing. A single large XAML would be difficult to review and maintain.

The project groups related procedures around business capabilities. `Main.xaml` orchestrates the source `Run_All` path. Each business workflow owns the procedures and helpers needed for that capability. DataTable-heavy aggregation and scoring are encapsulated in bounded VB `Invoke Code` activities where a native-activity implementation would create hundreds of low-level activities; workbook I/O and workflow orchestration remain native UiPath activities. This keeps the visual workflow structure limited while retaining readable, testable rule blocks.

## Main workflow — exact source `Run_All` order

1. Read `Action Sheet `, `New Club data`, `STO Export`, `PO Export`, `Comment Snapshot` and `Run Log` into DataTables.
2. Validate New Club Number, DC Number, Last Truck Date and required New Club columns. Stop before any write if validation fails.
3. Take/refresh the daily comment snapshot.
4. Run `STO_Check.xaml` only when STO Export contains rows. Validate requested club/site and deletion state before Article aggregation and comments.
5. Run `PO_Check.xaml` only when PO Export contains rows. Validate DC/site, deletion indicator and the inclusive Today-to-Last-Truck-Date window before updating blank or prior-PO comments.
6. Run `Storage_Release.xaml` only when `in_RunStorageCheck=True`; read the selected report's `Detail` sheet and apply its on-site-units lookup to comments still blank.
7. Append the run audit entry and prepare two targeted write-back DataTables.
8. When `in_WriteBack=True`, write only `Date of Last Review` and `CSC2 Comments`, then persist Comment Snapshot and Run Log. With write-back disabled, row changes remain in memory for review.

`Run_All` does **not** call Other Checks, Replacement Article, Club Transfer, Appointment Date, STF or Email Analysts in the original VBA. Those remain separate callable workflows, matching their separate source buttons/actions. Do not add them to the main run without a business-owner decision.

## Business capability workflows

| XAML | Source capability | Design responsibility |
|---|---|---|
| `Validate_Run_Inputs.xaml` | Action Sheet and column validation | Read Club/DC/truck date and fail before changes on missing inputs or columns. |
| `STO_Check.xaml` | `STO_Check`, `BuildExportInfo`, `OpenStoText`, `PartialStoText`, `ManualStoText` and aliases | Site/deletion safety; Article key aggregation; delivered/open quantities; STO document/due-date details; manifested, partial, manual and overdue outcomes. |
| `PO_Check.xaml` | `PO_Check`, comment builder, filtering, date sort, deletion and key helpers | Validate DC; exclude deleted, wrong-site or invalid/out-of-window dates; aggregate/sort PO comments; Article and Generic Article fallback; preserve non-PO comments. |
| `Storage_Release.xaml` | Storage Release and `sr*` helpers | Resolve grouped DC On Site Units header, aggregate Article stock and update blank comments. |
| `Other_Checks.xaml` | HVS / Delisted / Initial Set Quantity rules | Apply HVS first, then Delisted, then missing/zero Initial Set Quantity, only to eligible blank comments. |
| `Club_Transfer.xaml` | `Club_Transfer` and ranking helpers | Consume ZPRR_INVSTAT result data; rank source clubs from OH/WOS relative to Initial Set Quantity; exclude the new club and apply any allowed club list. |
| `Appointment_Date.xaml` | `Appointment_Date` and document/date helpers | Extract PO/STO document key from the existing comment and merge the first applicable confirmation date from ME23N data. |
| `Replacement_Article.xaml` | `Replacement_Article`, `ReplacementOf`, `RepStoInfo`, `IsOpenPartial` | Extract replacement Article; evaluate replacement supply in order STO → eligible PO → Storage → no supply. |
| `Fetch_STF_Article.xaml` | `Fetch_STF_Article` and candidate scoring helpers | Prefer same-POG candidates, then same AH6/MerchCat candidates; exclude STF source Articles; rank up to three suggestions. |
| `Update_Tracking_Report.xaml` | Module7 report-copy routines | Match normalized Article + Article Description and update every matching target row; optional STO/PO-only mode follows the source filter. |
| `Refresh_Summary.xaml` | `SumCat`, `MapPeople`, `Refresh_Summary` and dashboard data builders | Produce structured comment, division, analyst and action-detail DataTables. Excel drawing/charts are not required for business calculations. |
| `Take_Daily_Snapshot.xaml` | Daily snapshot routines | Store today's Article/category/comment snapshot, replacing today's prior snapshot on rerun. |
| `Write_Run_Log.xaml` | Run Log helpers | Append date/time, Windows user, club, action and result. |
| `Clear_Data.xaml` | `Clear_Data` and clear-range helpers | Implement source selection options 1–4 with confirmation; clear selected DataTables/cells while retaining schemas. |
| `Email_Analysts.xaml` | `Email_Analysts`, owner and HTML helpers | Reproduce analyst ownership/category selection, group rows and produce subject/body/CC/attachment row sets. Outlook draft/send must be configured separately. |
| `Prepare_Column_Writes.xaml` | Dynamic column and safe-write helpers | Build only the Review Date and Comments write-back DataTables and resolve target cells. |
| `UiOnly_Replacement.xaml` | Excel-only UI/presentation helpers | Replace progress animation, color, centering, toast, protection and dashboard drawing with logs/status text; no business data is altered. |

## Comment and comparison rules preserved

### STO

- Normalize Article keys; numeric Article values are compared without leading zeroes.
- Validate that remaining export rows belong to the requested New Club and do not silently clear the entire export when site/deletion checks fail.
- Aggregate Article quantity into total, delivered/marked and open quantities.
- When delivered quantity covers a positive Initial Set Quantity, use `STO Manifested`.
- When some quantity is delivered but the requirement is not yet covered, calculate incoming open STO supply through Last Truck Date and distinguish will-manifest from manual-shortfall text.
- When no delivered quantity is available and total STO quantity is insufficient, produce `Need manual STO` including remaining quantity and open STO references.
- Preserve due dates and mark overdue open STOs as passed/not delivered. Future open STO rows are retained because source constant `STO_PAST_ONLY` is `False`.

### PO

- A PO row is valid only when it is not deleted, its Site equals the configured DC, it has a valid due date, and the date is within the inclusive Today-to-Last-Truck-Date window.
- Group valid rows by normalized Article; sort by due date and produce `PO <document> Due On <m/d/yyyy>`; multiple rows are joined with ` | `.
- Match Article first and Generic Article second. Only blank comments or an existing `PO ` comment may be updated; other comments are preserved.

### Storage and exception rules

- Storage uses the selected DC's On Site Units column and updates blank comments only. Cases text is optional when no cases quantity exists.
- Other Checks precedence is HVS → Delisted → missing/zero Initial Set Quantity and only runs on blank comments.
- Replacement Article uses replacement Article supply precedence STO → PO → Storage; it replaces only rows carrying the replacement trigger.

## SAP boundary — intentionally not fabricated

The source VBA automates SAP GUI transactions ME2M, ZPRR_INVSTAT and ME23N. Target GUI selectors cannot be derived safely from VBA source because they depend on actual SAP GUI version, screen layout, control IDs, user role and runtime state. `SAP_Adapters/Capture_SAP_UI.xaml` is an explicit fail-closed adapter contract; `SAP_Adapters/Pull_From_SAP.xaml` documents the `Run_Full_Club_Check` order and throws until selectors are recorded in UiPath Studio. The generated `Main.xaml` currently accepts already-exported STO and PO DataTables via workbook tabs.

Capture approved selectors in Studio UI Explorer / Object Repository, use stable anchors where possible, add timeout/retry and SAP status-bar checks, and run in a non-production environment first. Do not replace the fail-closed adapter with guessed selectors.

## Outlook boundary

`Email_Analysts.xaml` builds deterministic email groups, CC fields, subjects, HTML bodies and row sets for attachments. Actual Outlook delivery/draft creation is a separate external side effect and requires a configured UiPath Outlook/Microsoft 365 activity, connection and mailbox policy. The migration does not silently send messages.

## Skill Markdown policy

The project includes curated offline Markdown playbooks and a full skill snapshot of any repositories the user imported. Skill Markdown is guidance/reference material; imported scripts and executable binaries are not run. The app can fetch public `.md`/`.mdx` files from the supplied GitHub skill repositories without a model API key, subject to GitHub anonymous API rate limits. GitHub topic pages (Python Skills and Agentic Skills) are indexes rather than canonical repositories; select concrete repositories from those topics before import.

## UiPath Studio validation required

1. Open `project.json` in the intended UiPath Studio release and restore the declared packages.
2. Validate every XAML. Confirm activity property schemas against the exact installed package versions.
3. Build the whole project in Studio; static XML parsing does not compile VB.NET or prove UiPath can load each activity.
4. Run `Main.xaml` on a backup workbook with `in_WriteBack=False` first. Review comments and row-preservation results.
5. Compare against original VBA with `Migration/Regression_Test_Plan.csv`.
6. Capture and test SAP adapters and configure Outlook only after approved access is available.

This is a deterministic migration implementation and auditable XAML project draft. It must pass Studio Validate/Build and environment-specific regression tests before production use.
'''

def generate_uipath_project(analysis: dict[str, Any], ir: dict[str, Any] | None = None, relevant_skills: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    symbols = analysis.get("symbols", []) or []
    names = {str(s.get("name", "")).lower() for s in symbols}
    is_sync = {"run_full_club_check", "sto_check", "po_check", "storage_release_check"}.issubset(names)
    if not is_sync:
        return generate_generic_uipath_project(analysis, ir or {})

    files: list[dict[str, str]] = []
    def add(path: str, content: str, purpose: str):
        files.append({"path": path, "content": content, "purpose": purpose})

    add("project.json", json.dumps(PROJECT_JSON, indent=2), "UiPath Windows project manifest; VB expressions; System and Excel activities dependencies.")
    # Preserve the exact knowledge references selected by the offline skill matcher.
    skills = relevant_skills or []
    skill_doc = ["# Selected Skill References", "", "These Markdown references were loaded by the local migration engine and used as authoring guidance. They are included for auditability; imported scripts are not executed.", ""]
    for skill in skills:
        skill_doc.extend([f"## {skill.get('name', 'Skill')}", f"Source: {skill.get('source', 'bundled')}", "", str(skill.get('description', '')), ""])
        content = str(skill.get("content", ""))
        if content:
            skill_doc.extend([content[:18000], ""])
    if not skills:
        skill_doc.extend(["No additional imported skill documents were selected. The generator uses its embedded UiPath XAML authoring and migration rules."])
    add("Skills/Applied_Skill_References.md", "\n".join(skill_doc), "Skill evidence loaded by the offline deterministic generator.")
    add("Main.xaml", _main_xaml(), "Main orchestrator: reads workbook exports, validates inputs, runs supply checks and writes only the two target output columns.")
    add("README_FIRST.md", _design_md(symbols, True), "Implementation notes, source rules, target mapping, SAP capture boundary and Studio validation instructions.")
    for path, (display, args, code, purpose) in _procedure_workflows().items():
        class_name = path.replace("\\", "/").removesuffix(".xaml").replace("/", "_")
        # Build each child workflow root with typed arguments and one bounded code activity.
        add(path, workflow_xaml(class_name, display, purpose, args, code), purpose)
    add("SAP_Adapters/Capture_SAP_UI.xaml", _sap_adapter_xaml("Capture_SAP_UI", "SAP UI capture-required adapter", "ME2M / ZPRR_INVSTAT / ME23N"), "Explicit failure boundary to prevent fabricated SAP selectors; replace with recorded selectors after Studio capture.")
    pull_code = '''Dim articleRows As Integer = If(dtClub Is Nothing, 0, dtClub.Rows.Count)
outMessage = "Source sequence: clear STO/PO export tabs; connect to SAP; run ME2M for New Club " & clubNo & " and " & articleRows.ToString() & " New Club rows; run STO Check; collect still-blank Articles; run ME2M for DC " & dcNo & " through Last Truck Date " & truckDate.ToString("yyyy-MM-dd") & "; run PO Check; optionally load Storage Report. The legacy partial-STO PO enrichment branch is explicitly removed in the source."
Throw New System.InvalidOperationException("Live SAP pull is capture-required. Record SAP session, ME2M selection fields, Article multi-select popup, ALV Delivery Schedule grid and export column mapping in SAP_Adapters/Capture_SAP_UI.xaml before enabling this workflow.")'''
    add("SAP_Adapters/Pull_From_SAP.xaml", workflow_xaml("SAP_Adapters_Pull_From_SAP", "Pull from SAP — ME2M source sequence", "Run_Full_Club_Check orchestration from AutoClubCheck", [("in_clubNo","In","x:String"),("in_dcNo","In","x:String"),("in_truckDate","In","sys:DateTime"),("in_dtClub","In","s:DataTable"),("out_message","Out","x:String")], pull_code), "Documents and gates the exact SAP Pull from SAP sequence; fails closed until target SAP controls and selectors are captured.")
    add("SAP_Adapters/README_SAP_CAPTURE_REQUIRED.md", '''# SAP GUI adapter capture required\n\nThe source macro uses SAP transactions ME2M, ZPRR_INVSTAT and ME23N. Their data mapping and business rules are translated into the business workflows, but the target UI selectors are not guessed. Capture the target SAP screens in UiPath Studio using the installed UIAutomation package and the approved SAP environment. Implement these adapters:\n\n- **ME2M**: Site, Article multi-select, optional delivery date range, ALV result grid and export schema mapping.\n- **ZPRR_INVSTAT**: DC, Article multi-select, Club/Site, Article, OH, WOS; source VBA batches at most 5 articles and suggests up to 3 source clubs.\n- **ME23N**: PO/STO document entry, item detail, Confirmations tab and first Delivery Date.\n\nNever hand-author fragile selectors from source code. Use Studio Capture / UI Explorer, store selectors in Object Repository where appropriate, add timeout/retry and SAP status-bar handling, and validate against the exact SAP GUI version and user role. Until adapters are captured, supply the exported SAP DataTables to the included workflows.\n''', "Environment-specific SAP GUI integration requirements.")
    add("Migration/Function_Mapping.csv", _mapping_csv(symbols), "Complete per-procedure traceability from all detected VBA procedures/functions to logical target XAML or explicit capture/omission boundary.")
    add("Migration/Procedure_Catalog.md", _procedure_catalog_md(symbols), "Readable complete procedure catalogue.")
    add("Migration/Architecture.md", _design_md(symbols, True), "Detailed workflow design and business decision order.")
    add("Migration/Regression_Test_Plan.csv", _test_plan_csv(), "Row-level behavior tests to compare VBA and UiPath outputs.")
    add("Config/BusinessRules.json", json.dumps({
        "dateWindow": {"poStartInclusive": "Today", "poEndInclusive": "Last Truck Date"},
        "stoPastOnly": False,
        "stoCommentTemplates": {"manifested": "STO Manifested", "manualPrefix": "Need manual STO", "partialPrefix": "Partial set quantity delivered", "willManifestPrefix": "Partially manifested - next incoming STO", "overduePhrase": "is passed STO not delivered"},
        "poCommentTemplate": "PO {document} Due On {M/d/yyyy}",
        "poJoinSeparator": " | ",
        "poEligibleExistingComments": ["blank", "starts with PO "],
        "storageCommentTemplate": "Storage DC {dc}: {units} units ({cases} cs) on site - available to release",
        "otherCheckPrecedence": ["Placement=HVS", "Listing Status=Delisted", "Initial Set Quantity blank or zero"],
        "divisionScope": "Empty means all divisions; otherwise pipe/comma/semicolon separated exact division labels.",
        "deletionIndicatorWhitelistNotDeleted": ["", "#", "-", "--", ".", "0", "N", "NO", "FALSE", "NULL", "N/A", "NA", "NOT ASSIGNED"],
        "writeBackColumnsOnly": ["Date of Last Review", "CSC2 Comments"],
        "sapCaptureRequired": ["ME2M", "ZPRR_INVSTAT", "ME23N"]
    }, indent=2), "Configurable business constants extracted from the source VBA.")
    return {
        "executive_summary": f"Generated a modular, offline UiPath Windows project for the SYNC VBA process. {len(symbols)} source procedures/functions are mapped to {len([f for f in files if f['path'].endswith('.xaml')])} XAML workflows/adapters. Core STO, PO, Storage, Other Checks, Replacement Article, Clear Data, Email Analyst selection/grouping, transfer ranking, appointment-date merge, STF recommendations, tracking matching and full Summary data tables are implemented as deterministic VB.NET blocks inside a compact set of cohesive XAML workflows. All {len(symbols)} procedures are trace-mapped; SAP UI selectors are explicitly marked capture-required, and Outlook send/draft actions require configured mail activities.",
        "source_logic": {"detected_procedures": [s.get("name") for s in symbols], "patterns": analysis.get("patterns", {}), "external_touchpoints": analysis.get("external_touchpoints", [])},
        "target_design": {"target": "UiPath Windows XAML project", "status": "modular business-logic implementation with SAP and Outlook capture/configuration boundaries", "workflow_count": len([f for f in files if f['path'].endswith('.xaml')]), "procedure_count": len(symbols), "uses_api_key": False},
        "generated_files": files,
        "mapping_table": [{"source": s.get("name", ""), "module": s.get("source_file", ""), "target": _classify_procedure(s)[0], "status": _classify_procedure(s)[1], "mapping": _classify_procedure(s)[2]} for s in symbols],
        "unresolved_items": ["SAP ME2M, ZPRR_INVSTAT and ME23N UI selectors must be captured in the target SAP GUI with UiPath Studio; adapter XAML intentionally fails closed until configured.", "Confirm the exact Storage Report Detail grouped-header layout and DC columns with a representative report.", "Validate the project against the organization's UiPath Studio/runtime version and package feed; run a whole-project build.", "Regression-test against the VBA output for comments, row preservation, date values and scope filtering."],
        "test_plan": ["Use test rows for delivered STO quantity equal to, greater than and less than Initial Set Quantity.", "Test partial delivered + open STO due by Last Truck Date sufficient to manifest; test insufficient incoming quantity; test overdue open STO.", "Test STO Export wrong Site and all rows deleted; ensure failure occurs before any write.", "Test PO deletion indicators, wrong DC, missing/invalid dates, dates before Today and after Last Truck Date.", "Test PO grouping by earliest date, multiple PO separator, Generic Article fallback, and preservation of existing non-PO comments.", "Test Storage comment on blank rows and preserve nonblank rows.", "Test HVS / Delisted / missing quantity precedence.", "Test Replacement Article STO→PO→Storage precedence, club transfer ranking, ME23N appointment suffix, STF same-POG preference, and tracker Article+Description matching."],
        "risks": ["SAP GUI automation is not executable until target selectors are captured.", "DataTable header reading depends on export reports preserving header rows.", "The workbook uses Excel-specific formatting and dashboard objects which are not replicated as XAML UI."],
        "assumptions": ["UiPath Windows project and VB expression language are acceptable.", "SAP STO/PO exports are provided by Main workflow inputs until SAP adapters are captured.", "The source code was extracted from the supplied workbook; no VBA macros were executed.", "No external AI/model API key is used by this deterministic generation."],
        "confidence": "medium",
        "generation_mode": "Deterministic offline UiPath XAML generation",
    }


def generate_generic_uipath_project(analysis: dict[str, Any], ir: dict[str, Any]) -> dict[str, Any]:
    """For unknown source systems, emit a truthful typed workflow map, not fake code translation."""
    symbols = analysis.get("symbols", []) or []
    names = [s.get("name", "") for s in symbols]
    files = []
    files.append({"path":"project.json", "content":json.dumps(PROJECT_JSON, indent=2), "purpose":"UiPath Windows workflow project manifest."})
    files.append({"path":"Main.xaml", "content":_generic_main_xaml(symbols), "purpose":"Entry workflow with structured phase boundaries and a safe review gate."})
    files.append({"path":"Migration/Function_Mapping.csv", "content":_mapping_csv(symbols), "purpose":"Complete mapping of discovered source functions and procedures."})
    files.append({"path":"Migration/Source_Analysis.md", "content":"# Source-to-UiPath mapping\n\nThis input did not match the known SYNC workbook process. All detected procedures are catalogued in Function_Mapping.csv. A deterministic translator cannot infer arbitrary language semantics, API contracts, selectors, transactions and side effects from names/excerpts alone. The XAML entry workflow therefore exposes explicit implementation boundaries rather than generating false executable logic. Supply full source files and representative input/output examples for a source-specific rule implementation.\n", "purpose":"Honest conversion boundary."})
    return {"executive_summary":f"Created a UiPath project map for {len(symbols)} discovered procedures. The known SYNC business-rule converter was not activated because this source does not contain its orchestration signature. No API key was used.", "source_logic":{"detected_procedures":names}, "target_design":{"target":"UiPath Windows XAML","status":"workflow map; arbitrary semantics require source-specific conversion"}, "generated_files":files,"mapping_table":[{"source":n,"target":_classify_procedure(s)[0],"status":"mapped for review"} for n,s in zip(names,symbols)],"unresolved_items":["Source-specific input/output contracts and side effects require semantic analysis before implementing generic procedures."],"test_plan":["Provide representative source input and expected output fixtures.","Build and run generated project in UiPath Studio."],"risks":["Do not treat the generic map as an executable conversion."],"assumptions":[],"confidence":"low","generation_mode":"Deterministic source map; no API key"}


def _generic_main_xaml(symbols: list[dict[str, Any]]) -> str:
    note = f"Source project contains {len(symbols)} detected functions or procedures. Review Migration/Function_Mapping.csv and implement target-specific subworkflows before production."
    # LogMessage is a standard built-in activity; no fake conversion is represented as executable behavior.
    return f'''<Activity mc:Ignorable="sap sap2010" x:Class="Main" sap2010:WorkflowViewState.IdRef="ActivityBuilder_1" xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:sap="http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation" xmlns:sap2010="http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation">
  <Sequence DisplayName="Migration map - review and implement target logic" sap2010:WorkflowViewState.IdRef="Sequence_1">
    <WriteLine Text="{html.escape(note)}" sap2010:WorkflowViewState.IdRef="WriteLine_1" />
    <Throw Exception="[New System.InvalidOperationException(&quot;This generic map is not executable until source-specific rules and target activities are implemented. See Migration/Function_Mapping.csv.&quot;)]" sap2010:WorkflowViewState.IdRef="Throw_1" />
  </Sequence>
</Activity>'''


def _test_plan_csv() -> str:
    rows = [
        ("STO", "Delivered qty = required qty", "Comment exactly STO Manifested; review date today"),
        ("STO", "Delivered qty > required qty", "Comment exactly STO Manifested"),
        ("STO", "Partial delivered + open supply by truck date covers set", "Partially manifested - next incoming STO ... will manifest"),
        ("STO", "Partial delivered + open supply insufficient", "Partial set quantity delivered...; lists delivered/open; manual gap"),
        ("STO", "No delivered and total STO < required", "Need manual STO for remaining..."),
        ("STO", "Overdue open STO not covered by delivered due date", "Contains is passed STO not delivered"),
        ("STO", "Export Site does not equal Action Sheet Club", "Business exception before row write"),
        ("STO", "All selected club rows are deletion-marked", "Business exception; do not wipe export"),
        ("PO", "PO due today", "Valid PO comment"),
        ("PO", "PO due on Last Truck Date", "Valid PO comment"),
        ("PO", "PO due before today or after truck date", "Excluded"),
        ("PO", "PO deletion marker true or wrong DC", "Excluded"),
        ("PO", "Existing STO/non-PO comment with matching PO", "Existing comment preserved"),
        ("PO", "Article not matched but Generic Article is matched", "PO comment via fallback"),
        ("Storage", "Blank comment with units and cases", "Storage DC comment includes units and cases"),
        ("Storage", "Existing nonblank comment", "Preserved"),
        ("Other", "HVS + Delisted + no qty", "HVS article - first rule wins"),
        ("Other", "Delisted + no qty", "Delisted article"),
        ("Other", "Blank/zero Initial Set Quantity", "No initial set quantity for STO creation"),
        ("Replacement", "Replacement article in STO, then PO, then Storage", "STO wins; else PO; else Storage"),
        ("Transfer", "Some club OH >= initial set qty", "Coverage-qualified candidates rank before non-covering; then WOS/OH"),
        ("STF", "Candidates on same POG", "Same-POG candidates are preferred"),
        ("Tracking", "Duplicate Article + Description in report", "Every matching tracker row updated"),
        ("Scope", "Division not selected", "Comment and review date preserved"),
    ]
    buf=io.StringIO(); w=csv.writer(buf); w.writerow(["Workflow","Test condition","Expected result"]); w.writerows(rows); return buf.getvalue()
