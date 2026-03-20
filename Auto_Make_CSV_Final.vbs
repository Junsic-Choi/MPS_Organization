On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile("FinalList_UTF16.csv", True, True)
outFile.WriteLine "Site,Group,Model,RPM,Month,Code,Product"

Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(fso.GetAbsolutePathName(".") & "\data_working.xlsx")

' 1. Meta (Sheet 2)
Set metaMap = CreateObject("Scripting.Dictionary")
Set wsMeta = workbook.Sheets(2)
lastRowMeta = 1000 ' Fixed or detect
lastS = "": lastG = ""
For r = 7 To 1000
    s = Trim(wsMeta.Cells(r, 1).Value)
    g = Trim(wsMeta.Cells(r, 2).Value)
    m = Trim(wsMeta.Cells(r, 3).Value)
    rpm = Trim(wsMeta.Cells(r, 4).Value)
    If s <> "" Then lastS = s Else s = lastS
    If g <> "" Then lastG = g Else g = lastG
    If m <> "" Then
        key = UCase(Replace(m, "LYNX ", ""))
        mMeta = Array(s, g, m, rpm)
        If Not metaMap.Exists(key) Then metaMap.Add key, mMeta
    End If
Next

' 2. Months (Sheet 4)
Set wsMps = workbook.Sheets(4)
tCols = Array(9, 13, 18, 23, 29, 35)
Dim ms(5)
For i = 0 To 5
    h = wsMps.Cells(3, tCols(i)).Value
    mNum = ""
    For k = 1 To Len(h)
        If IsNumeric(Mid(h, k, 1)) Then mNum = mNum & Mid(h, k, 1)
    Next
    If mNum <> "" Then ms(i) = mNum & ChrW(50900) Else ms(i) = h
Next

' 3. Loop
count = 0
lastC = "" : lastP = "" : curM = Empty
For r = 7 To 2000
    c = Trim(wsMps.Cells(r, 4).Value)
    p = Trim(wsMps.Cells(r, 5).Value)
    If c <> "" Then lastC = c
    If p <> "" Then
        lastP = p
        kP = UCase(Split(p, "-")(0))
        curM = Empty
        If metaMap.Exists(kP) Then
            curM = metaMap(kP)
        Else
            ks = metaMap.Keys
            For Each mk In ks
                If InStr(mk, kP) > 0 Or InStr(kP, mk) > 0 Then
                    curM = metaMap(mk)
                    Exit For
                End If
            Next
        End If
    End If
    
    ' Check if current row has ANY data in target columns
    hasAny = False
    For i = 0 To 5
        v = wsMps.Cells(r, tCols(i)).Value
        If IsNumeric(v) And v > 0 Then
            hasAny = True
            mS = "" : mG = "" : mM = "" : mR = ""
            If Not IsEmpty(curM) Then
                mS = curM(0): mG = curM(1): mM = curM(2): mR = curM(3)
            End If
            For q = 1 To v
                line = """" & mS & """,""" & mG & """,""" & mM & """,""" & mR & """,""" & ms(i) & """,""" & lastC & """,""" & lastP & """"
                outFile.WriteLine line
                count = count + 1
            Next
        End If
    Next
    ' If we reached far end and still seeing blanks, exit
    If c = "" And p = "" And r > 1600 Then Exit For
Next

outFile.Close
workbook.Close False
excel.Quit
WScript.Echo "Rows: " & count
