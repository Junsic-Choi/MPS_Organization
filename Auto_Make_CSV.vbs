On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("_FinalList.csv", True)
' Write CSV Header
outFile.WriteLine "Site,Group,Model,RPM,Month,Code,Product"

Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

Dim dir
dir = fso.GetAbsolutePathName(".")
Dim path
path = dir & "\data_working.xlsx"

Set workbook = excel.Workbooks.Open(path)
If Err.Number <> 0 Then
    WScript.Echo "Error opening: " & Err.Description
    WScript.Quit 1
End If

' 1. Build Meta Map (Dictionary)
Set metaMap = CreateObject("Scripting.Dictionary")
Set wsMeta = workbook.Sheets(2)
lastRowMeta = wsMeta.UsedRange.Rows.Count
lastS = "": lastG = ""

For r = 7 To lastRowMeta
    s = Trim(wsMeta.Cells(r, 1).Value)
    g = Trim(wsMeta.Cells(r, 2).Value)
    m = Trim(wsMeta.Cells(r, 3).Value)
    rpm = Trim(wsMeta.Cells(r, 4).Value)
    
    If s <> "" Then lastS = s Else s = lastS
    If g <> "" Then lastG = g Else g = lastG
    
    If m <> "" Then
        key = UCase(Replace(m, "LYNX ", ""))
        If Not metaMap.Exists(key) Then
            metaMap.Add key, Array(s, g, m, rpm)
        End If
    End If
Next

' 2. Month Headers from Sheet 4
Set wsMps = workbook.Sheets(4)
targetCols = Array(9, 13, 18, 23, 29, 35)
Dim months(5)
For i = 0 To 5
    h = wsMps.Cells(3, targetCols(i)).Value
    ' Simple extract digit from "26.2월 실적" -> "2월"
    vMonth = ""
    For k = 1 To Len(h)
        ch = Mid(h, k, 1)
        If IsNumeric(ch) Then vMonth = vMonth & ch
    Next
    If vMonth <> "" Then months(i) = vMonth & ChrW(50900) Else months(i) = h
Next

' 3. Extract Loop
lastC = "": lastP = "": curMeta = Empty
count = 0
For r = 7 To 10000
    c = Trim(wsMps.Cells(r, 4).Value)
    p = Trim(wsMps.Cells(r, 5).Value)
    
    If c <> "" Then lastC = c
    If p <> "" Then
        lastP = p
        ' Extract model part (e.g., XG800-...)
        kPart = UCase(Split(p, "-")(0))
        curMeta = Empty
        If metaMap.Exists(kPart) Then
            curMeta = metaMap(kPart)
        Else
            ' Fuzzy search
            keys = metaMap.Keys
            For Each mk In keys
                If InStr(mk, kPart) > 0 Or InStr(kPart, mk) > 0 Then
                    curMeta = metaMap(mk)
                    Exit For
                End If
            Next
        End If
    End If
    
    ' Month loop should run for every row
    For i = 0 To 5
        vQty = wsMps.Cells(r, targetCols(i)).Value
        If IsNumeric(vQty) Then
            If vQty > 0 Then
                ' Combined Metadata
                mS = "": mG = "": mM = "": mR = ""
                If Not IsEmpty(curMeta) Then
                    mS = curMeta(0): mG = curMeta(1): mM = curMeta(2): mR = curMeta(3)
                End If
                For q = 1 To vQty
                    ' Escape commas for CSV
                    line = """" & mS & """,""" & mG & """,""" & mM & """,""" & mR & """,""" & months(i) & """,""" & lastC & """,""" & lastP & """"
                    outFile.WriteLine line
                    count = count + 1
                Next
            End If
        End If
    Next
    
    If r Mod 100 = 0 Then ' Log progress occasionally if needed
    End If
    
    ' Safeguard: If both code and product are empty for this row, and we've passed a certain row, exit.
    ' This prevents processing thousands of empty rows at the end of the sheet.
    If c = "" And p = "" And r > 2000 Then Exit For
Next

outFile.Close
workbook.Close False
excel.Quit
WScript.Echo "Success: " & count & " rows."
