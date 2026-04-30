' robust_dump_v96.vbs
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Excel Not Found"
    WScript.Quit 1
End If

Set wb = xl.ActiveWorkbook
If wb Is Nothing Then
    WScript.Echo "Error: No Active Workbook"
    WScript.Quit 1
End IF
WScript.Echo "Active: " & wb.Name

' Find Sheets
Set wsMPS = Nothing: Set wsProd = Nothing
For Each sh In wb.Sheets
    If InStr(sh.Name, "MPS") > 0 Then Set wsMPS = sh
    If InStr(sh.Name, "생산") > 0 Then Set wsProd = sh
Next
If wsMPS Is Nothing Then Set wsMPS = wb.Sheets(4)
If wsProd Is Nothing Then Set wsProd = wb.Sheets(2)

WScript.Echo "MPS: " & wsMPS.Name & " | Prod: " & wsProd.Name

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\data_dump.json", True)

f.Write "{""mps"":["
arr = wsMPS.Range("A1:AD1500").Value
first = True
For r = 1 To 1500
    c = Trim(arr(r, 4)): p = Trim(arr(r, 5)): n = Trim(arr(r, 7))
    If c <> "" And c <> "Model" Then
        If Not first Then f.Write ","
        f.Write "{""c"":""" & c & """,""n"":""" & n & """,""p"":""" & p & """}"
        first = False
    End If
Next
f.Write "],""prod"":["

arrP = wsProd.Range("A1:N3000").Value
first = True
For r = 6 To 3000
    site = Trim(arrP(r, 1)): model = Trim(arrP(r, 3))
    If model <> "" And site <> "" And model <> "기종" Then
        If Not first Then f.Write ","
        f.Write "{""s"":""" & site & """,""m"":""" & model & """,""r"":""" & arrP(r, 4) & """,""v"":["
        f.Write arrP(r, 5) & "," & arrP(r, 8) & "," & arrP(r, 9) & "," & arrP(r, 10) & "," & arrP(r, 11) & "," & arrP(r, 13)
        f.Write "]}"
        first = False
    End If
Next
f.Write "]}"
f.Close
WScript.Echo "SUCCESS: Data Dumped"
