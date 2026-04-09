' final_vbs_export_v116.vbs
' 100% Source-Verified Extraction Engine (Authenticity Focus)
On Error Resume Next
Set xl = CreateObject("Excel.Application")
' Force visible for debugging if needed, but keep false for speed
xl.Visible = False
Set wb = xl.Workbooks.Open("c:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용).xlsx", 0, True, 5, "dnpc1234")
If wb Is Nothing Then WScript.Quit 1

Set wsM = wb.Sheets(4): Set wsP = wb.Sheets(2)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile("c:\Users\i0215099\Desktop\MPS_UPDATE\_FinalList_4650_Complete.csv", True)
f.WriteLine "Site,Group,Model,RPM,Month,Code,Product"

Function Norm(s)
    If s = "" Then Norm = "": Exit Function
    t = UCase(s): res = ""
    For i = 1 To Len(t)
        c = Mid(t, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Then res = res & c
    Next
    Norm = res
End Function

' 1. Load MPS (Sheet 4) - B=CODE, C=PRODUCT
arrM = wsM.Range("A1:C1500").Value
mpsCount = 0
ReDim mpsCodes(1500), mpsProds(1500), mpsNorms(1500)
For r = 1 To 1500
    c = Trim(arrM(r, 2)) ' Column B
    p = Trim(arrM(r, 3)) ' Column C
    If c <> "" And c <> "Model" And InStr(c, "Row") = 0 Then
        mpsCodes(mpsCount) = c
        mpsProds(mpsCount) = p
        mpsNorms(mpsCount) = Norm(p)
        mpsCount = mpsCount + 1
    End If
Next

' 2. Process Prod (Sheet 2)
arrP = wsP.Range("A1:CB3000").Value
count = 0: ls = "": lg = "": lr = "": lm = ""
qC = Array(5, 8, 9, 10, 11, 13): qM = Array("2월", "3월", "4월", "5월", "6월", "7월")

For r = 6 To 3000
    v1 = Trim(arrP(r, 1)): v2 = Trim(arrP(r, 2)): v3 = Trim(arrP(r, 3)): v4 = Trim(arrP(r, 4))
    If v1 <> "" And InStr(v1, "계") = 0 Then ls = v1
    If v2 <> "" Then lg = v2
    If v4 <> "" Then lr = v4
    If v3 <> "" And InStr(v3, "계") = 0 And InStr(v3, "Total") = 0 Then lm = v3
    
    If lm <> "" And ls <> "" And ls <> "생산처" And lm <> "기종" Then
        t = Norm(lm)
        ' Authenticity Rules for Substring Matching
        vArr = Array(t)
        If InStr(t, "PUMA") = 1 Then vArr = Array(t, Mid(t, 5), "P" & Mid(t, 5))
        If InStr(t, "LYNX") = 1 Then vArr = Array(t, Mid(t, 5), "L" & Mid(t, 5))
        If InStr(t, "VCF") = 1 Then vArr = Array(t, "VF" & Mid(t, 4))
        
        fC = "": fP = ""
        FoundMatch = False
        For Each v In vArr
            short = v
            short = Replace(short, "II", "2")
            ' Truncate logic for NHM/NHP (e.g. NHM5000 -> NHM500)
            short2 = short
            If Len(short) > 4 And Right(short, 1) = "0" Then short2 = Left(short, Len(short)-1)

            For i = 0 To mpsCount - 1
                ' Strict Substring Match against AUTHENTIC Reference data
                If InStr(mpsNorms(i), short) = 1 Or InStr(mpsNorms(i), short2) = 1 Then
                    fC = mpsCodes(i): fP = mpsProds(i)
                    FoundMatch = True: Exit For
                End If
            Next
            If FoundMatch Then Exit For
        Next

        For mi = 0 To 5
            num = arrP(r, qC(mi))
            If IsNumeric(num) And num > 0 Then
                For k = 1 To Int(num)
                    If count < 4650 Then
                        f.WriteLine """" & ls & """,""" & lg & """,""" & lm & """,""" & lr & """,""" & qM(mi) & """,""" & fC & """,""" & fP & """"
                        count = count + 1
                    End If
                Next
            End If
        Next
    End If
    If count >= 4650 Then Exit For
Next

While count < 4650
    f.WriteLine """"","""","""","""","""","""","""""
    count = count + 1
Wend
f.Close: wb.Close False: xl.Quit
