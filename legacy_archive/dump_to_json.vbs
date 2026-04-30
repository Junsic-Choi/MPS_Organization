' dump_to_json.vbs
On Error Resume Next
Set xl = GetObject(, "Excel.Application")
Set wb = xl.ActiveWorkbook
If wb Is Nothing Then WScript.Quit 1

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\data_dump.json", True)

' Dump MPS (Sheet 4)
Set wsMPS = wb.Sheets(4)
f.Write "{""mps"":["
arrMPS = wsMPS.Range("A1:H1500").Value
first = True
For r = 1 To 1500
    c = Trim(arrMPS(r, 4))
    n = Trim(arrMPS(r, 7))
    p = Trim(arrMPS(r, 5))
    If c <> "" And c <> "Model" Then
        If Not first Then f.Write ","
        f.Write "{""c"":""" & c & """,""n"":""" & n & """,""p"":""" & p & """}"
        first = False
    End If
Next
f.Write "],""prod"":["

' Dump Prod (Sheet 2)
Set wsProd = wb.Sheets(2)
arrProd = wsProd.Range("A1:N3000").Value
first = True
For r = 6 To 3000
    site = Trim(arrProd(r, 1))
    model = Trim(arrProd(r, 3))
    If model <> "" And site <> "" And model <> "기종" Then
        If Not first Then f.Write ","
        f.Write "{""s"":""" & site & """,""m"":""" & model & """,""r"":""" & Trim(arrProd(r, 4)) & """,""v"":["
        ' Monthly values (Feb-Jul)
        f.Write Trim(arrProd(r, 5)) & "," & Trim(arrProd(r, 8)) & "," & Trim(arrProd(r, 9)) & "," & Trim(arrProd(r, 10)) & "," & Trim(arrProd(r, 11)) & "," & Trim(arrProd(r, 13))
        f.Write "]}"
        first = False
    End If
Next
f.Write "]}"
f.Close
