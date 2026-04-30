' final_extraction.vbs
On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set logFile = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\mapping_log.txt", True)
Set csvFile = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\Final_Mapped_Data.csv", True)

Sub WriteLog(msg)
    logFile.WriteLine "[" & Time & "] " & msg
End Sub

WriteLog "Starting Final Mapping"

Set xl = CreateObject("Excel.Application")
Set wb = Nothing
For Each w In xl.Workbooks
    If InStr(w.Name, "MPS2603") > 0 Or InStr(w.Name, "생산배포용") > 0 Then 
        Set wb = w
        Exit For
    End If
Next

If wb Is Nothing Then
    WriteLog "Workbook NOT FOUND. Please ensure the file is open in Excel."
    logFile.Close
    WScript.Quit
End If

Set wsP = wb.Worksheets(2) ' 생산배포용
Set wsM = wb.Worksheets(4) ' MPS

WriteLog "Using MPS WB: " & wb.Name
WriteLog "S2: " & wsP.Name & ", S4: " & wsM.Name

' 1. Extract Production List from 생산배포용
' Col A: 생산처, B: 기종분류, C: 기종(Model), D: RPM
' Qty cols: E(2월), H(3월), I(4월), J(5월), K(6월), M(7월)
' 월 이름 맵핑
Dim monthsP(5), colP(5)
monthsP(0) = "2월": colP(0) = 5   ' E
monthsP(1) = "3월": colP(1) = 8   ' H
monthsP(2) = "4월": colP(2) = 9   ' I
monthsP(3) = "5월": colP(3) = 10  ' J
monthsP(4) = "6월": colP(4) = 11  ' K
monthsP(5) = "7월": colP(5) = 13  ' M

' 유닛 리스트 저장 (Dictionary 사용 어려우므로 Array 활용)
Dim units()
ReDim units(10000, 5) ' (Index, {Site, Cat, Model, RPM, Month})
uCount = 0

WriteLog "Scanning Sheet 2 for units..."
For r = 1 To 1000 ' 데이터 예상 범위
    site = wsP.Cells(r, 1).Value
    If site <> "" And site <> "생산처" And site <> "합계" Then
        cat = wsP.Cells(r, 2).Value
        mdl = wsP.Cells(r, 3).Value
        rpm = wsP.Cells(r, 4).Value
        
        For m = 0 To 5
            qty = wsP.Cells(r, colP(m)).Value
            If IsNumeric(qty) Then
                For q = 1 To CInt(qty)
                    units(uCount, 0) = site
                    units(uCount, 1) = cat
                    units(uCount, 2) = mdl
                    units(uCount, 3) = rpm
                    units(uCount, 4) = monthsP(m)
                    units(uCount, 5) = "Unused"
                    uCount = uCount + 1
                Next
            End If
        Next
    End If
Next
WriteLog "Total Units collected from S2: " & uCount

' 2. Map to MPS Tab and Output
' CSV Header
csvFile.WriteLine """Site"",""Category"",""Model"",""RPM"",""Month"",""MPS_Model"",""MPS_Product"",""MPS_Site"",""MPS_Ver"""

WriteLog "Mapping units to Sheet 4 (MPS)..."
Dim monthsM(5), colM(5)
' I(9), M(13), R(18), W(23), AC(29), AI(35)
colM(0) = 9: colM(1) = 13: colM(2) = 18: colM(3) = 23: colM(4) = 29: colM(5) = 35

mappedCount = 0

' MPS 데이터 시작: R6
For r = 6 To 3000 ' MPS 예상 범위
    mpsModel = wsM.Cells(r, 4).Value   ' D
    mpsProd = wsM.Cells(r, 5).Value    ' E
    mpsSite = wsM.Cells(r, 7).Value    ' G
    mpsVer = wsM.Cells(r, 8).Value     ' H
    
    If mpsModel = "" And mpsProd = "" Then Exit For
    
    For m = 0 To 5
        mQty = wsM.Cells(r, colM(m)).Value
        If IsNumeric(mQty) Then
            For q = 1 To CInt(mQty)
                ' Find matching unit in the collected list
                foundIdx = -1
                For i = 0 To uCount - 1
                    If units(i, 5) = "Unused" And units(i, 4) = monthsP(m) Then
                        ' basic matching by Model (Fuzzy)
                        s1 = Replace(UCase(units(i, 2)), " ", "")
                        s2 = Replace(UCase(mpsModel), " ", "")
                        If InStr(s1, s2) > 0 Or InStr(s2, s1) > 0 Then
                            foundIdx = i
                            Exit For
                        End If
                    End If
                Next
                
                If foundIdx <> -1 Then
                    units(foundIdx, 5) = "Used"
                    csvFile.WriteLine """" & units(foundIdx, 0) & """,""" & units(foundIdx, 1) & """,""" & units(foundIdx, 2) & """,""" & units(foundIdx, 3) & """,""" & units(foundIdx, 4) & """,""" & mpsModel & """,""" & mpsProd & """,""" & mpsSite & """,""" & mpsVer & """"
                    mappedCount = mappedCount + 1
                Else
                    ' Fallback: If no fuzzy match, just take any unused for that month (or mark as MISSING)
                    csvFile.WriteLine """MISSING"",""MISSING"",""MISSING"",""MISSING"",""" & monthsP(m) & """,""" & mpsModel & """,""" & mpsProd & """,""" & mpsSite & """,""" & mpsVer & """"
                End If
            Next
        End If
    Next
Next

WriteLog "Total units mapped: " & mappedCount
WriteLog "Done."

csvFile.Close
logFile.Close
