Sub Transfer_Scopus()

Dim Re, RealRow As Integer
Dim ArrOld, ArrNew, Title
Dim Dict, TitleDic As Object
Dim ExistArry(20) As Boolean


Set TitleDic = CreateObject("scripting.dictionary")
Set Dict = CreateObject("scripting.dictionary")


'-------------------------------標題title---------------------------
Title = Range(Worksheets("New").[A1], Worksheets("New").Cells(1, Columns.Count).End(xlToLeft))

For i = 1 To UBound(Title, 2) '用row計算
    If Not Title(1, i) = "" Then
     TitleDic.Add Title(1, i), i '字典物件.Add 索引值,內容物
    End If
Next i


'MsgBox TitleDic.Exists("名稱") '有可能找不到文獻名稱只好模糊搜尋，標頭重打(我選擇標頭重打)

If TitleDic.Exists("EID") Then
    ArrOld = Range(Worksheets("Old").[AG1], Worksheets("Old").Cells(Rows.Count, TitleDic("EID")).End(xlUp)) '從cell(1,AG)到cell(最後一列,33行)
    ArrNew = Range(Worksheets("New").[AG1], Worksheets("New").Cells(Rows.Count, TitleDic("EID")).End(xlUp))
    
    For i = 1 To UBound(ArrOld, 1)
    Dict.Add ArrOld(i, 1), i '字典物件.Add 索引值,內容物
    Next i
End If


'---------------------------------------------------------------
Re = 0

'For Each i In TitleDic
'If Not TitleDic.keys(i) Then
'MsgBox "來源出版物名稱 查找不到!!!"
'End If
'Next i


If TitleDic.Exists("名稱") Then
    ExistArry(1) = True
Else
    MsgBox "「名稱」欄位 查找不到!"
End If

If TitleDic.Exists("年份") Then
    ExistArry(2) = True
Else
    MsgBox "「年份」欄位 查找不到!"
End If

If TitleDic.Exists("作者") Then
    ExistArry(3) = True
Else
    MsgBox "「作者」欄位 查找不到!"
End If

If TitleDic.Exists("文獻類型") Then
    ExistArry(4) = True
Else
    MsgBox "「文獻類型」欄位 查找不到!"
End If

If TitleDic.Exists("來源出版物名稱") Then
    ExistArry(5) = True
Else
    MsgBox "「來源出版物名稱」欄位 查找不到!"
End If

If TitleDic.Exists("卷") Then
    ExistArry(6) = True
Else
    MsgBox "「卷」欄位 查找不到!"
End If

If TitleDic.Exists("期") Then
    ExistArry(7) = True
Else
    MsgBox "「期」欄位 查找不到!"
End If

If TitleDic.Exists("起始頁碼") Then
    ExistArry(8) = True
Else
    MsgBox "「起始頁碼」欄位 查找不到!"
End If

If TitleDic.Exists("結束頁碼") Then
    ExistArry(9) = True
Else
    MsgBox "「結束頁碼」欄位 查找不到!"
End If

If TitleDic.Exists("DOI") Then
    ExistArry(10) = True
Else
    MsgBox "「DOI」欄位 查找不到!"
End If

If TitleDic.Exists("連結") Then
    ExistArry(11) = True
Else
    MsgBox "「連結」欄位 查找不到!"
End If

If TitleDic.Exists("摘要") Then
    ExistArry(12) = True
Else
    MsgBox "「摘要」欄位 查找不到!"
End If

If TitleDic.Exists("作者關鍵字") Then
    ExistArry(13) = True
Else
    MsgBox "「作者關鍵字」欄位 查找不到!"
End If

If TitleDic.Exists("EID") Then
    ExistArry(14) = True
Else
    MsgBox "「EID」欄位 查找不到!"
End If





For i = 2 To UBound(ArrNew, 1)

    If Not Dict.Exists(ArrNew(i, 1)) Then '判斷索引值是否在舊的存在
    
        RealRow = i - Re '字典從1開始 匯入欄位列從2開始
        
        
        Worksheets("Result").Cells(RealRow, "A") = "機構"                   '機構
        
        If ExistArry(1) Then
            Worksheets("Result").Cells(RealRow, "B") = Worksheets("New").Cells(i, TitleDic("名稱"))  '文獻名稱
        End If
        
        If ExistArry(2) Then
            Worksheets("Result").Cells(RealRow, "C") = Worksheets("New").Cells(i, TitleDic("年份"))   '年份
        End If
            
        
        If ExistArry(3) Then
            Worksheets("Result").Cells(RealRow, "D") = Replace(Worksheets("New").Cells(i, TitleDic("作者")), ",", ";")   '作者
        End If
        
        
        If ExistArry(4) Then
            Worksheets("Result").Cells(RealRow, "E") = Worksheets("New").Cells(i, TitleDic("文獻類型"))   '文獻類型
        End If
        
        '-----------------------集叢名判斷-----------------------------
        Worksheets("Result").Cells(RealRow, "F") = "" '集叢名初始化
        
        '來源出版物
        If ExistArry(5) Then
            If Worksheets("New").Cells(i, "E") = "" Or Worksheets("New").Cells(i, TitleDic("來源出版物名稱")) Is Nothing Then
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F")
            Else
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F") + CStr(Worksheets("New").Cells(i, TitleDic("來源出版物名稱")))
            End If
        End If

        
        
        '卷
        If ExistArry(6) Then
            If CStr(Worksheets("New").Cells(i, "F")) = "" Then
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F")
            Else
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F") + ",Volume " + CStr(Worksheets("New").Cells(i, TitleDic("卷")))
            End If
        End If
        
        
        '期
        If ExistArry(7) Then
            If Worksheets("New").Cells(i, "G") = "" Or Worksheets("New").Cells(i, "G") Is Nothing Then
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F")
            Else
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F") + "," + CStr(Worksheets("New").Cells(i, TitleDic("期")))
            End If
        End If
        
        
        '起始頁碼
        If ExistArry(8) Then
            If CStr(Worksheets("New").Cells(i, "I")) = "" Then
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F")
            Else
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F") + ",Page:" + CStr(Worksheets("New").Cells(i, TitleDic("起始頁碼")))
            End If
        End If
        
        
        '結束頁碼
        If ExistArry(9) Then
            If CStr(Worksheets("New").Cells(i, "J")) = "" Then
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F")
            Else
            Worksheets("Result").Cells(RealRow, "F") = Worksheets("Result").Cells(RealRow, "F") + "-" + CStr(Worksheets("New").Cells(i, TitleDic("結束頁碼")))
            End If
        End If
        
        '----------------------------------------------------------
        If ExistArry(10) Then
        Worksheets("Result").Cells(RealRow, "G") = Worksheets("New").Cells(i, TitleDic("DOI"))   'DOI
        End If


        '是否使用 Scopus 網址取代
        If ExistArry(11) Then
            If Not (Worksheets("New").Cells(i, TitleDic("DOI")) = "" Or Worksheets("New").Cells(i, TitleDic("DOI")) Is Nothing) Then
            Worksheets("Result").Cells(RealRow, "H") = "http://dx.doi.org/" + Worksheets("New").Cells(i, TitleDic("DOI"))   'DOI網址
            Else
            Worksheets("Result").Cells(RealRow, "H") = Worksheets("New").Cells(i, TitleDic("連結"))   'Scopus連結
            End If
        End If
        
        If ExistArry(12) Then
            Worksheets("Result").Cells(RealRow, "I") = Replace(Worksheets("New").Cells(i, TitleDic("摘要")), WorksheetFunction.Unichar(169), "Copyright")   '摘要
        End If

        If ExistArry(13) Then
            Worksheets("Result").Cells(RealRow, "J") = Worksheets("New").Cells(i, TitleDic("作者關鍵字"))   '作者關鍵字
        End If

        If ExistArry(14) Then
            Worksheets("Result").Cells(RealRow, "K") = Worksheets("New").Cells(i, TitleDic("EID"))   'EID
        End If

    ElseIf Dict.Exists(ArrNew(i, 1)) Then
    Re = Re + 1 '如果存在就跳過
    End If

Next i


End Sub
