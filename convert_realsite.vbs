' convert_realsite.vbs
Set xl = CreateObject("Excel.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
dest = "c:\Users\i0215099\Desktop\MPS_UPDATE\realsite_simple.txt"
Set out = fso.CreateTextFile(dest, True)

Set wb = xl.Workbooks.Open("c:\Users\i0215099\Desktop\MPS_UPDATE\Real site.xlsx")
Set ws = wb.Sheets(1)
For r = 1 To 50
    line = ""
    For c = 1 To 10
        val = ws.Cells(r, c).Text
        line = line & val & "||"
    Next
    out.WriteLine line
Next

wb.Close False
xl.Quit
out.Close
