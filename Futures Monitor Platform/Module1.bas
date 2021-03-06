Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Public msg As String
Public ErrNumber As Integer

Sub 產生資訊()

    Dim i, j As Integer
    Dim dummy(11) As String
    Dim WSN As String
    
    WSN = ActiveSheet.Name
    
    For i = 3 To Cells(2, 1).End(xlDown).Row
        If Cells(i, 1).Value = "TSE" Then
            WSN = "TW"
             For j = 2 To 12
                dummy(j - 1) = Cells(1, j).Value
                Cells(i, j).Value = "=XQFAP|Quote!'" & Cells(i, 1).Value & "." & WSN & "-" & dummy(j - 1) & "'"
            Next j
            WSN = ActiveSheet.Name
        Else
            For j = 2 To 12
                dummy(j - 1) = Cells(1, j).Value
                Cells(i, j).Value = "=XQFAP|Quote!'" & Cells(i, 1).Value & "." & WSN & "-" & dummy(j - 1) & "'"
            Next j
        End If
    Next i
End Sub

Sub update_TF()
    Dim i As Integer
    Dim ID As String
    Dim dummy As Date
    Dim WBN As Object, 主表 As Object, TF As Object
    Dim AdoConn As New ADODB.Connection
    Dim strConn As String
    Dim strSQL As String
    Dim output As Variant
    
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & ThisWorkbook.Path & Application.PathSeparator & "Monitor platform .accdb;" '連接database和source
    AdoConn.Open strConn  'excel連接database
    
    Set WBN = Workbooks("Monitor Platform .xlsm")
    Set 主表 = Workbooks("Monitor Platform .xlsm").Sheets("主表")
    Set TF = Workbooks("Monitor Platform .xlsm").Sheets("TF")
    
    WBN.Activate
    
    With 主表
        '=========================================================
        '-------------
        '更新原本資料
        '-------------
            
        ErrNumber = 0
        
        On Error GoTo 100
        
        For i = 3 To 14
        
            ID = .Cells(i, 2).Value
            
            If .Cells(i, 6).Value = "--" Or .Cells(i, 7).Value = "--" Then
                ErrNumber = 100
            Else
                dummy = .Cells(i, 3).Value
                strSQL = "SELECT TOP 1 時間,最高價,最低價 FROM " & ID & " ORDER BY 時間 DESC"      'select欄位 'from什麼table 'orderby排序 'desn升冪降冪排列
                WBN.Sheets("temp").Cells(1, 1).CopyFromRecordset AdoConn.Execute(strSQL)   '在日期把select出的東西貼在1,1
                
                If dummy <> Sheets("temp").Cells(1, 1).Value And Day(.Cells(1, 3).Value) <> Day(dummy) Then '現在日期和db與過去worksheet存在的日期不一樣
                    ErrNumber = 200
                    strSQL = "INSERT INTO " & ID & "(時間,開盤價,最高價,最低價,收盤價) VALUES(#" & TF.Cells(i, 3).Value & "#," & _
                    TF.Cells(i, 5).Value & "," & _
                    TF.Cells(i, 6).Value & "," & _
                    TF.Cells(i, 7).Value & "," & _
                    TF.Cells(i, 8).Value & ")"
                    AdoConn.Execute strSQL
                    
                    'insertinto為sql語法塞ws的data到db裡
                    
                ElseIf dummy = Sheets("temp").Cells(1, 1).Value Then
                    ErrNumber = 210
                    If .Cells(i, 6).Value > WBN.Sheets("temp").Cells(1, 2).Value Then
                        ErrNumber = 211
                        strSQL = "UPDATE " & ID & " SET 最高價=" & .Cells(i, 6).Value & " WHERE 時間=#" & .Cells(i, 3).Value & "#"
                        AdoConn.Execute strSQL
                        WBN.Sheets("temp").Cells(1, 2).Value = .Cells(i, 6).Value
                    End If
                    If .Cells(i, 7).Value < WBN.Sheets("temp").Cells(1, 3).Value Then
                        ErrNumber = 212
                        strSQL = "UPDATE " & ID & " SET 最低價=" & .Cells(i, 7).Value & " WHERE 時間=#" & .Cells(i, 3).Value & "#"
                        AdoConn.Execute strSQL
                        WBN.Sheets("temp").Cells(1, 3).Value = .Cells(i, 7).Value
                    End If
                End If
            End If
            
'            '3日內最高最低   'db選出來到ws
'            ErrNumber = 300
'            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC)"
'            AdoConn.Execute strSQL
'            .Cells(i, 8).CopyFromRecordset AdoConn.Execute(strSQL)
'            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
'            AdoConn.Execute strSQL
'            .Cells(i, 9).CopyFromRecordset AdoConn.Execute(strSQL)
'            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
'            AdoConn.Execute strSQL
'            .Cells(i, 11).CopyFromRecordset AdoConn.Execute(strSQL)

            '5個營業日內最高最低(1-week)
            ErrNumber = 301
            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 13).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 14).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 16).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '10個營業日內最高最低(2-weeks)
            ErrNumber = 302
            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 18).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 19).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 21).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '1-Month
            ErrNumber = 303
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 23).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 24).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 26).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '2個月內最高最低
            ErrNumber = 304
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 28).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 29).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 31).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '1季內最高最低
            ErrNumber = 305
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 33).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 34).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 36).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '2季內最高最低
            ErrNumber = 306
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 38).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 39).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 41).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '1年內最高最低
            ErrNumber = 307
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 43).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 44).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 46).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '2年內最高最低
            ErrNumber = 308
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 48).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 49).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 51).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '歷史最高與最低
            ErrNumber = 309
            strSQL = "SELECT TOP 1 最高價,時間 FROM " & ID & " ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 53).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM " & ID & " ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 55).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '資料起始日
            ErrNumber = 310
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 57).CopyFromRecordset AdoConn.Execute(strSQL)
   
        Next i
        '=========================================================

    End With
    Set WBN = Nothing
    Set 主表 = Nothing
    Set TF = Nothing
    
    AdoConn.Close
    ErrNumber = 400
    ThisWorkbook.Save
Exit Sub
    
100:
    Set WBN = Nothing
    Set 主表 = Nothing
    Set TF = Nothing
    
    '遇錯誤excel自動關閉
    
    AdoConn.Close
    ThisWorkbook.Save
    msg = Err.Description & " ErrNumber is " & ErrNumber & " i=" & i & "請重開此檔案"
    Hotmail_err
    ThisWorkbook.Close
End Sub

Sub update_yeswinRTD()  '測試元大RTD之用途

    Dim i As Integer
    Dim ID As String
    Dim dummy As Date
    Dim WBN As Object, 主表 As Object, yeswinRTD As Object
    Dim AdoConn As New ADODB.Connection
    Dim strConn As String
    Dim strSQL As String
    Dim output As Variant
    
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & ThisWorkbook.Path & Application.PathSeparator & "Monitor platform.accdb;"
    AdoConn.Open strConn
    
    Set WBN = Workbooks("Monitor Platform.xlsm")
    Set 主表 = Workbooks("Monitor Platform.xlsm").Sheets("主表")
    Set yeswinRTD = Workbooks("Monitor Platform.xlsm").Sheets("yeswinRTD")
    
    WBN.Activate
    
    With 主表
        '=========================================================
        '-------------
        '更新原本資料
        '-------------
            
        ErrNumber = 1000
        
        On Error GoTo 100
        
        For i = 3 To 4
            .Cells(i, 3).Value = Year(Now()) & "/" & .Cells(i, 58).Value
            ID = .Cells(i, 2).Value
            
            If .Cells(i, 6).Value = "--" Or .Cells(i, 7).Value = "--" Then
                ErrNumber = 1100
            Else
                dummy = .Cells(i, 3).Value
                strSQL = "SELECT TOP 1 時間,最高價,最低價 FROM " & ID & " ORDER BY 時間 DESC"
                WBN.Sheets("temp").Cells(1, 1).CopyFromRecordset AdoConn.Execute(strSQL)
                
                If dummy <> Sheets("temp").Cells(1, 1).Value And Day(.Cells(1, 3).Value) <> Day(dummy) Then
                    ErrNumber = 1200
                    strSQL = "INSERT INTO " & ID & "(時間,開盤價,最高價,最低價,收盤價) VALUES(#" & yeswinRTD.Cells(i, 3).Value & "#," & _
                    yeswinRTD.Cells(i, 5).Value & "," & _
                    yeswinRTD.Cells(i, 6).Value & "," & _
                    yeswinRTD.Cells(i, 7).Value & "," & _
                    yeswinRTD.Cells(i, 8).Value & ")"
                    AdoConn.Execute strSQL

                ElseIf dummy = Sheets("temp").Cells(1, 1).Value Then
                    ErrNumber = 1210
                    If .Cells(i, 6).Value > WBN.Sheets("temp").Cells(1, 2).Value Then
                        ErrNumber = 1211
                        strSQL = "UPDATE " & ID & " SET 最高價=" & .Cells(i, 6).Value & " WHERE 時間=#" & .Cells(i, 3).Value & "#"
                        AdoConn.Execute strSQL
                        WBN.Sheets("temp").Cells(1, 2).Value = .Cells(i, 6).Value
                    End If
                    If .Cells(i, 7).Value < WBN.Sheets("temp").Cells(1, 3).Value Then
                        ErrNumber = 1212
                        strSQL = "UPDATE " & ID & " SET 最低價=" & .Cells(i, 7).Value & " WHERE 時間=#" & .Cells(i, 3).Value & "#"
                        AdoConn.Execute strSQL
                        WBN.Sheets("temp").Cells(1, 3).Value = .Cells(i, 7).Value
                    End If
                End If
            End If
            
            '3日內最高最低
            ErrNumber = 1300
            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 8).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 9).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 11).CopyFromRecordset AdoConn.Execute(strSQL)

            '5個營業日內最高最低(1-week)
            ErrNumber = 1301
            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 13).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 14).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 16).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '10個營業日內最高最低(2-weeks)
            ErrNumber = 1302
            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 18).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 19).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 21).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '1-Month
            ErrNumber = 1303
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 23).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 24).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 26).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '2個月內最高最低
            ErrNumber = 1304
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 28).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 29).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 31).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '1季內最高最低
            ErrNumber = 1305
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 33).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 34).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 36).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '2季內最高最低
            ErrNumber = 1306
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 38).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 39).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 41).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '1年內最高最低
            ErrNumber = 1307
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 43).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 44).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 46).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '2年內最高最低
            ErrNumber = 1308
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 48).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 49).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 51).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '歷史最高與最低
            ErrNumber = 1309
            strSQL = "SELECT TOP 1 最高價,時間 FROM " & ID & " ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 53).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM " & ID & " ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 55).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '資料起始日
            ErrNumber = 1310
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 57).CopyFromRecordset AdoConn.Execute(strSQL)
   
        Next i
        '=========================================================

    End With
    Set WBN = Nothing
    Set 主表 = Nothing
    Set yeswinRTD = Nothing
    
    AdoConn.Close
    ErrNumber = 1400
    ThisWorkbook.Save
Exit Sub
    
100:
    Set WBN = Nothing
    Set 主表 = Nothing
    Set yeswinRTD = Nothing
    
    AdoConn.Close
    ThisWorkbook.Save
    msg = Err.Description & " ErrNumber is " & ErrNumber & " i=" & i & "請重開此檔案"
    'email_err
    'Hotmail_err
    ThisWorkbook.Close
End Sub

Sub update_Ryan()
    Dim i As Integer
    Dim ID As String
    Dim dummy As Date
    Dim WBN As Object, 主表 As Object, 犀利環球 As Object
    Dim AdoConn As New ADODB.Connection
    Dim strConn As String
    Dim strSQL As String
    Dim output As Variant
    
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & ThisWorkbook.Path & Application.PathSeparator & "Monitor platform 201503 for APP.accdb;"
    AdoConn.Open strConn
    
    Set WBN = Workbooks("Monitor Platform 201503 for APP.xlsm")
    Set 主表 = Workbooks("Monitor Platform 201503 for APP.xlsm").Sheets("主表")
    Set 犀利環球 = Workbooks("Monitor Platform 201503 for APP.xlsm").Sheets("犀利環球")
    
    WBN.Activate
    
    With 主表
        '=========================================================
        '-------------
        '更新原本資料
        '-------------
            
        ErrNumber = 2000
        
        On Error GoTo 100
        
        For i = 280 To 288
        
            ID = .Cells(i, 2).Value
            
            If .Cells(i, 6).Value = "--" Or .Cells(i, 7).Value = "--" Then
                ErrNumber = 2100
            Else
                dummy = .Cells(i, 3).Value - 1
                strSQL = "SELECT TOP 1 時間,最高價,最低價 FROM " & ID & " ORDER BY 時間 DESC"
                WBN.Sheets("temp").Cells(1, 1).CopyFromRecordset AdoConn.Execute(strSQL)
                
                If dummy <> Sheets("temp").Cells(1, 1).Value And Day(.Cells(1, 3).Value) <> Day(dummy) Then
                    ErrNumber = 2200
                    strSQL = "INSERT INTO " & ID & "(時間,開盤價,最高價,最低價,收盤價) VALUES(#" & 犀利環球.Cells(i - 277, 3).Value & "#," & _
                    犀利環球.Cells(i - 277, 5).Value & "," & _
                    犀利環球.Cells(i - 277, 6).Value & "," & _
                    犀利環球.Cells(i - 277, 7).Value & "," & _
                    犀利環球.Cells(i - 277, 8).Value & ")"
                    AdoConn.Execute strSQL

                ElseIf dummy = Sheets("temp").Cells(1, 1).Value Then
                    ErrNumber = 2210
                    If .Cells(i, 6).Value > WBN.Sheets("temp").Cells(1, 2).Value Then
                        ErrNumber = 1211
                        strSQL = "UPDATE " & ID & " SET 最高價=" & .Cells(i, 6).Value & " WHERE 時間=#" & .Cells(i, 3).Value & "#"
                        AdoConn.Execute strSQL
                        WBN.Sheets("temp").Cells(1, 2).Value = .Cells(i, 6).Value
                    End If
                    If .Cells(i, 7).Value < WBN.Sheets("temp").Cells(1, 3).Value Then
                        ErrNumber = 1212
                        strSQL = "UPDATE " & ID & " SET 最低價=" & .Cells(i, 7).Value & " WHERE 時間=#" & .Cells(i, 3).Value & "#"
                        AdoConn.Execute strSQL
                        WBN.Sheets("temp").Cells(1, 3).Value = .Cells(i, 7).Value
                    End If
                End If
            End If
            
            '3日內最高最低
            ErrNumber = 2300
            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 8).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 9).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 3 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 11).CopyFromRecordset AdoConn.Execute(strSQL)

            '5個營業日內最高最低(1-week)
            ErrNumber = 2301
            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 13).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 14).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 5 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 16).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '10個營業日內最高最低(2-weeks)
            ErrNumber = 2302
            strSQL = "SELECT LAST(時間) FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 18).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 19).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT TOP 10 * FROM " & ID & " ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 21).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '1-Month
            ErrNumber = 2303
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 23).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 24).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 26).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '2個月內最高最低
            ErrNumber = 2304
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 28).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 29).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""m"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 31).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '1季內最高最低
            ErrNumber = 2305
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 33).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 34).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 36).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '2季內最高最低
            ErrNumber = 2306
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 38).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 39).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""q"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 41).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '1年內最高最低
            ErrNumber = 2307
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 43).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 44).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-1,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 46).CopyFromRecordset AdoConn.Execute(strSQL)
                    
            '2年內最高最低
            ErrNumber = 2308
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 48).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最高價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 49).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM (SELECT * FROM " & ID & " WHERE 時間>DATEADD(""yyyy"",-2,#" & .Cells(i, 3).Value & "#) ORDER BY 時間 DESC) ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 51).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '歷史最高與最低
            ErrNumber = 2309
            strSQL = "SELECT TOP 1 最高價,時間 FROM " & ID & " ORDER BY 最高價 DESC"
            AdoConn.Execute strSQL
            .Cells(i, 53).CopyFromRecordset AdoConn.Execute(strSQL)
            strSQL = "SELECT TOP 1 最低價,時間 FROM " & ID & " ORDER BY 最低價 ASC"
            AdoConn.Execute strSQL
            .Cells(i, 55).CopyFromRecordset AdoConn.Execute(strSQL)
            
            '資料起始日
            ErrNumber = 2310
            strSQL = "SELECT FIRST(時間) FROM (SELECT * FROM " & ID & " ORDER BY 時間 DESC)"
            AdoConn.Execute strSQL
            .Cells(i, 57).CopyFromRecordset AdoConn.Execute(strSQL)
   
        Next i
        '=========================================================

    End With
    Set WBN = Nothing
    Set 主表 = Nothing
    Set 犀利環球 = Nothing
    
    AdoConn.Close
    ErrNumber = 2400
    ThisWorkbook.Save
Exit Sub
    
100:
    Set WBN = Nothing
    Set 主表 = Nothing
    Set 犀利環球 = Nothing
    
    AdoConn.Close
    ThisWorkbook.Save
    msg = Err.Description & " ErrNumber is " & ErrNumber & " i=" & i & "請重開此檔案"
    'Hotmail_err
    ThisWorkbook.Close
End Sub


Sub notify()

    Dim i As Integer
    Dim flat As Integer
    Dim WBN As Object
    flat = 0
    ErrNumber = 0
    
    Set WBN = Workbooks("Monitor Platform.xlsm")
    WBN.Activate
    
    On Error GoTo 100
    If flat = 0 Then
        With WBN.Sheets("主表")
            For i = 3 To 4
                .Cells(i, 3).Value = Year(Now()) & "/" & .Cells(i, 58).Value
                If .Cells(i, 6).Value = "--" Or .Cells(i, 7).Value = "--" Then
                    ErrNumber = 500
                Else
                    If .Cells(i, 6).Value > .Cells(i, 53).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 54).Value) Then  '創資料庫新高
                        msg = "[New High After " & .Cells(i, 57).Value & "] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 54).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 53).Value, 2))
                        ErrNumber = 505
                        .Cells(i, 54).Value = .Cells(i, 3).Value    'DB的最高日期更新
                        .Cells(i, 53).Value = .Cells(i, 6).Value    'DB的最高價格更新
                        .Cells(i, 50).Value = .Cells(i, 3).Value    '2 Years的最高日期更新
                        .Cells(i, 49).Value = .Cells(i, 6).Value    '2 Years的最高價格更新
                        .Cells(i, 45).Value = .Cells(i, 3).Value    '1 Years的最高日期更新
                        .Cells(i, 44).Value = .Cells(i, 6).Value    '1 Years的最高價格更新
                        .Cells(i, 40).Value = .Cells(i, 3).Value    '2 Quarters的最高日期更新
                        .Cells(i, 39).Value = .Cells(i, 6).Value    '2 Quarters的最高價格更新
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
                        
                    ElseIf .Cells(i, 6).Value > .Cells(i, 49).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 50).Value) Then  '創2 Years新高
                        msg = "[2Y-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 50).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 49).Value, 2))
                        ErrNumber = 510
                        .Cells(i, 50).Value = .Cells(i, 3).Value    '2 Years的最高日期更新
                        .Cells(i, 49).Value = .Cells(i, 6).Value    '2 Years的最高價格更新
                        .Cells(i, 45).Value = .Cells(i, 3).Value    '1 Years的最高日期更新
                        .Cells(i, 44).Value = .Cells(i, 6).Value    '1 Years的最高價格更新
                        .Cells(i, 40).Value = .Cells(i, 3).Value    '2 Quarters的最高日期更新
                        .Cells(i, 39).Value = .Cells(i, 6).Value    '2 Quarters的最高價格更新
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
                   
                    ElseIf .Cells(i, 6).Value > .Cells(i, 44).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 45).Value) Then  '創1 Years新高
                        msg = "[1Y-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 45).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 44).Value, 2))
                        ErrNumber = 511
                        .Cells(i, 45).Value = .Cells(i, 3).Value    '1 Years的最高日期更新
                        .Cells(i, 44).Value = .Cells(i, 6).Value    '1 Years的最高價格更新
                        .Cells(i, 40).Value = .Cells(i, 3).Value    '2 Quarters的最高日期更新
                        .Cells(i, 39).Value = .Cells(i, 6).Value    '2 Quarters的最高價格更新
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
                   
                    ElseIf .Cells(i, 6).Value > .Cells(i, 39).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 40).Value) Then  '創2 Quarters新高
                        msg = "[2Q-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 40).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 39).Value, 2))
                        ErrNumber = 512
                        .Cells(i, 40).Value = .Cells(i, 3).Value    '2 Quarters的最高日期更新
                        .Cells(i, 39).Value = .Cells(i, 6).Value    '2 Quarters的最高價格更新
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
                  
                    ElseIf .Cells(i, 6).Value > .Cells(i, 34).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 35).Value) Then  '創1 Quarter新高
                        msg = "[1Q-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 35).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 34).Value, 2))
                        ErrNumber = 513
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
             
                    ElseIf .Cells(i, 6).Value > .Cells(i, 29).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 30).Value) Then  '創2 Months新高
                        msg = "[2M-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 30).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 29).Value, 2))
                        ErrNumber = 514
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
            
                    ElseIf .Cells(i, 6).Value > .Cells(i, 24).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 25).Value) Then  '創1 Month新高
                        msg = "[1M-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 25).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 24).Value, 2))
                        ErrNumber = 515
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
             
                    ElseIf .Cells(i, 6).Value > .Cells(i, 19).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 20).Value) Then  '創2 Weeks新高
                        msg = "[2W-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 20).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 19).Value, 2))
                        ErrNumber = 516
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
                 
                    ElseIf .Cells(i, 6).Value > .Cells(i, 14).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 15).Value) Then  '創1 Week新高
                        msg = "[1W-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 15).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 14).Value, 2))
                        ErrNumber = 517
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        'email_short
                    End If
                End If
            
                If .Cells(i, 7).Value = "--" Then
                    ErrNumber = 600
                Else
                    If .Cells(i, 7).Value < .Cells(i, 55).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 56).Value) Then  '創DB新低
                        msg = "[New Low After " & .Cells(i, 57).Value & "] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 56).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 55).Value, 2))
                        ErrNumber = 605
                        .Cells(i, 56).Value = .Cells(i, 3).Value    'DB的最低日期更新
                        .Cells(i, 55).Value = .Cells(i, 7).Value    'DB的最低價格更新
                        .Cells(i, 52).Value = .Cells(i, 3).Value    '2 Years的最低日期更新
                        .Cells(i, 51).Value = .Cells(i, 7).Value    '2 Years的最低價格更新
                        .Cells(i, 47).Value = .Cells(i, 3).Value    '1 Years的最低日期更新
                        .Cells(i, 46).Value = .Cells(i, 7).Value    '1 Years的最低價格更新
                        .Cells(i, 42).Value = .Cells(i, 3).Value    '2 Quarters的最低日期更新
                        .Cells(i, 41).Value = .Cells(i, 7).Value    '2 Quarters的最低價格更新
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
                        
                    ElseIf .Cells(i, 7).Value < .Cells(i, 51).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 52).Value) Then  '創2 Years新低
                        msg = "[2Y-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 52).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 51).Value, 2))
                        ErrNumber = 610
                        .Cells(i, 52).Value = .Cells(i, 3).Value    '2 Years的最低日期更新
                        .Cells(i, 51).Value = .Cells(i, 7).Value    '2 Years的最低價格更新
                        .Cells(i, 47).Value = .Cells(i, 3).Value    '1 Years的最低日期更新
                        .Cells(i, 46).Value = .Cells(i, 7).Value    '1 Years的最低價格更新
                        .Cells(i, 42).Value = .Cells(i, 3).Value    '2 Quarters的最低日期更新
                        .Cells(i, 41).Value = .Cells(i, 7).Value    '2 Quarters的最低價格更新
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
               
                    ElseIf .Cells(i, 7).Value < .Cells(i, 46).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 47).Value) Then '創1 Years新低
                        msg = "[1Y-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 47).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 46).Value, 2))
                        ErrNumber = 611
                        .Cells(i, 47).Value = .Cells(i, 3).Value    '1 Years的最低日期更新
                        .Cells(i, 46).Value = .Cells(i, 7).Value    '1 Years的最低價格更新
                        .Cells(i, 42).Value = .Cells(i, 3).Value    '2 Quarters的最低日期更新
                        .Cells(i, 41).Value = .Cells(i, 7).Value    '2 Quarters的最低價格更新
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
               
                    ElseIf .Cells(i, 7).Value < .Cells(i, 41).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 42).Value) Then  '創2 Quarters新低
                        msg = "[2Q-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 42).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 41).Value, 2))
                        ErrNumber = 612
                        .Cells(i, 42).Value = .Cells(i, 3).Value    '2 Quarters的最低日期更新
                        .Cells(i, 41).Value = .Cells(i, 7).Value    '2 Quarters的最低價格更新
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
                
                    ElseIf .Cells(i, 7).Value < .Cells(i, 36).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 37).Value) Then  '創1 Quarter新低
                        msg = "[1Q-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 37).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 36).Value, 2))
                        ErrNumber = 613
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
            
                    ElseIf .Cells(i, 7).Value < .Cells(i, 31).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 32).Value) Then  '創2 Months新低
                        msg = "[2M-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 32).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 31).Value, 2))
                        ErrNumber = 614
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
              
                    ElseIf .Cells(i, 7).Value < .Cells(i, 26).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 27).Value) Then  '創1 Month新低
                        msg = "[1M-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 27).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 26).Value, 2))
                        ErrNumber = 615
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
               
                    ElseIf .Cells(i, 7).Value < .Cells(i, 21).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 22).Value) Then  '創2 Weeks新低
                        msg = "[2W-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 22).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 21).Value, 2))
                        ErrNumber = 616
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
              
                    ElseIf .Cells(i, 7).Value < .Cells(i, 16).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 17).Value) Then '創1 Week新低
                        msg = "[1W-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 17).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 16).Value, 2))
                        ErrNumber = 617
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        'email_short
                    End If
                End If
               
            Next i
        End With
        flat = flat + 1
    End If
    Application.OnTime CDate(Now) + TimeValue("00:30:00"), "notify", , True '每30分中檢查此邏輯
    'Hotmail_normal    '若符合此邏輯emailnormail
    'email_normal
    Set WBN = Nothing
Exit Sub
    
100:
    msg = Err.Description & " " & CStr(Now) & " ErrNumber is " & ErrNumber & " i=" & i
    'Hotmail_err
    'email_err
    Set WBN = Nothing
End Sub

Sub notify_Ryan()

    Dim i As Integer
    Dim flat As Integer
    Dim WBN As Object
    flat = 0
    ErrNumber = 0
    
    Set WBN = Workbooks("Monitor Platform 201503 for APP.xlsm")
    WBN.Activate
    
    On Error GoTo 100
    If flat = 0 Then
        With WBN.Sheets("主表")
            For i = 280 To 288
                
                If .Cells(i, 6).Value = "--" Or .Cells(i, 7).Value = "--" Then
                    ErrNumber = 1500
                Else
                    If .Cells(i, 6).Value > .Cells(i, 53).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 54).Value) Then  '創資料庫新高
                        msg = "[New High After " & .Cells(i, 57).Value & "] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 54).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 53).Value, 2))
                        ErrNumber = 1505
                        .Cells(i, 54).Value = .Cells(i, 3).Value    'DB的最高日期更新
                        .Cells(i, 53).Value = .Cells(i, 6).Value    'DB的最高價格更新
                        .Cells(i, 50).Value = .Cells(i, 3).Value    '2 Years的最高日期更新
                        .Cells(i, 49).Value = .Cells(i, 6).Value    '2 Years的最高價格更新
                        .Cells(i, 45).Value = .Cells(i, 3).Value    '1 Years的最高日期更新
                        .Cells(i, 44).Value = .Cells(i, 6).Value    '1 Years的最高價格更新
                        .Cells(i, 40).Value = .Cells(i, 3).Value    '2 Quarters的最高日期更新
                        .Cells(i, 39).Value = .Cells(i, 6).Value    '2 Quarters的最高價格更新
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
                        
                    ElseIf .Cells(i, 6).Value > .Cells(i, 49).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 50).Value) Then  '創2 Years新高
                        msg = "[2Y-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 50).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 49).Value, 2))
                        ErrNumber = 1510
                        .Cells(i, 50).Value = .Cells(i, 3).Value    '2 Years的最高日期更新
                        .Cells(i, 49).Value = .Cells(i, 6).Value    '2 Years的最高價格更新
                        .Cells(i, 45).Value = .Cells(i, 3).Value    '1 Years的最高日期更新
                        .Cells(i, 44).Value = .Cells(i, 6).Value    '1 Years的最高價格更新
                        .Cells(i, 40).Value = .Cells(i, 3).Value    '2 Quarters的最高日期更新
                        .Cells(i, 39).Value = .Cells(i, 6).Value    '2 Quarters的最高價格更新
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
                   
                    ElseIf .Cells(i, 6).Value > .Cells(i, 44).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 45).Value) Then  '創1 Years新高
                        msg = "[1Y-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 45).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 44).Value, 2))
                        ErrNumber = 1511
                        .Cells(i, 45).Value = .Cells(i, 3).Value    '1 Years的最高日期更新
                        .Cells(i, 44).Value = .Cells(i, 6).Value    '1 Years的最高價格更新
                        .Cells(i, 40).Value = .Cells(i, 3).Value    '2 Quarters的最高日期更新
                        .Cells(i, 39).Value = .Cells(i, 6).Value    '2 Quarters的最高價格更新
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
                   
                    ElseIf .Cells(i, 6).Value > .Cells(i, 39).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 40).Value) Then  '創2 Quarters新高
                        msg = "[2Q-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 40).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 39).Value, 2))
                        ErrNumber = 1512
                        .Cells(i, 40).Value = .Cells(i, 3).Value    '2 Quarters的最高日期更新
                        .Cells(i, 39).Value = .Cells(i, 6).Value    '2 Quarters的最高價格更新
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
                  
                    ElseIf .Cells(i, 6).Value > .Cells(i, 34).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 35).Value) Then  '創1 Quarter新高
                        msg = "[1Q-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 35).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 34).Value, 2))
                        ErrNumber = 1513
                        .Cells(i, 35).Value = .Cells(i, 3).Value    '1 Quarter的最高日期更新
                        .Cells(i, 34).Value = .Cells(i, 6).Value    '1 Quarter的最高價格更新
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
             
                    ElseIf .Cells(i, 6).Value > .Cells(i, 29).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 30).Value) Then  '創2 Months新高
                        msg = "[2M-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 30).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 29).Value, 2))
                        ErrNumber = 1514
                        .Cells(i, 30).Value = .Cells(i, 3).Value    '2 Months的最高日期更新
                        .Cells(i, 29).Value = .Cells(i, 6).Value    '2 Months的最高價格更新
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
            
                    ElseIf .Cells(i, 6).Value > .Cells(i, 24).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 25).Value) Then  '創1 Month新高
                        msg = "[1M-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 25).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 24).Value, 2))
                        ErrNumber = 1515
                        .Cells(i, 25).Value = .Cells(i, 3).Value    '1 Month的最高日期更新
                        .Cells(i, 24).Value = .Cells(i, 6).Value    '1 Month的最高價格更新
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
             
                    ElseIf .Cells(i, 6).Value > .Cells(i, 19).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 20).Value) Then  '創2 Weeks新高
                        msg = "[2W-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 20).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 19).Value, 2))
                        ErrNumber = 1516
                        .Cells(i, 20).Value = .Cells(i, 3).Value    '2 Weeks的最高日期更新
                        .Cells(i, 19).Value = .Cells(i, 6).Value    '2 Weeks的最高價格更新
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
                 
                    ElseIf .Cells(i, 6).Value > .Cells(i, 14).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 15).Value) Then  '創1 Week新高
                        msg = "[1W-High] " & CStr(.Cells(i, 2).Value) & " 站上 " & CStr(.Cells(i, 15).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 14).Value, 2))
                        ErrNumber = 1517
                        .Cells(i, 15).Value = .Cells(i, 3).Value    '1 Week的最高日期更新
                        .Cells(i, 14).Value = .Cells(i, 6).Value    '1 Week的最高價格更新
                        email_Ryan
                    End If
                End If
            
                If .Cells(i, 7).Value = "--" Then
                    ErrNumber = 1600
                Else
                    If .Cells(i, 7).Value < .Cells(i, 55).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 56).Value) Then  '創DB新低
                        msg = "[New Low After " & .Cells(i, 57).Value & "] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 56).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 55).Value, 2))
                        ErrNumber = 1605
                        .Cells(i, 56).Value = .Cells(i, 3).Value    'DB的最低日期更新
                        .Cells(i, 55).Value = .Cells(i, 7).Value    'DB的最低價格更新
                        .Cells(i, 52).Value = .Cells(i, 3).Value    '2 Years的最低日期更新
                        .Cells(i, 51).Value = .Cells(i, 7).Value    '2 Years的最低價格更新
                        .Cells(i, 47).Value = .Cells(i, 3).Value    '1 Years的最低日期更新
                        .Cells(i, 46).Value = .Cells(i, 7).Value    '1 Years的最低價格更新
                        .Cells(i, 42).Value = .Cells(i, 3).Value    '2 Quarters的最低日期更新
                        .Cells(i, 41).Value = .Cells(i, 7).Value    '2 Quarters的最低價格更新
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
                        
                    ElseIf .Cells(i, 7).Value < .Cells(i, 51).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 52).Value) Then  '創2 Years新低
                        msg = "[2Y-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 52).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 51).Value, 2))
                        ErrNumber = 1610
                        .Cells(i, 52).Value = .Cells(i, 3).Value    '2 Years的最低日期更新
                        .Cells(i, 51).Value = .Cells(i, 7).Value    '2 Years的最低價格更新
                        .Cells(i, 47).Value = .Cells(i, 3).Value    '1 Years的最低日期更新
                        .Cells(i, 46).Value = .Cells(i, 7).Value    '1 Years的最低價格更新
                        .Cells(i, 42).Value = .Cells(i, 3).Value    '2 Quarters的最低日期更新
                        .Cells(i, 41).Value = .Cells(i, 7).Value    '2 Quarters的最低價格更新
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
               
                    ElseIf .Cells(i, 7).Value < .Cells(i, 46).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 47).Value) Then '創1 Years新低
                        msg = "[1Y-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 47).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 46).Value, 2))
                        ErrNumber = 1611
                        .Cells(i, 47).Value = .Cells(i, 3).Value    '1 Years的最低日期更新
                        .Cells(i, 46).Value = .Cells(i, 7).Value    '1 Years的最低價格更新
                        .Cells(i, 42).Value = .Cells(i, 3).Value    '2 Quarters的最低日期更新
                        .Cells(i, 41).Value = .Cells(i, 7).Value    '2 Quarters的最低價格更新
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
               
                    ElseIf .Cells(i, 7).Value < .Cells(i, 41).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 42).Value) Then  '創2 Quarters新低
                        msg = "[2Q-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 42).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 41).Value, 2))
                        ErrNumber = 1612
                        .Cells(i, 42).Value = .Cells(i, 3).Value    '2 Quarters的最低日期更新
                        .Cells(i, 41).Value = .Cells(i, 7).Value    '2 Quarters的最低價格更新
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
                
                    ElseIf .Cells(i, 7).Value < .Cells(i, 36).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 37).Value) Then  '創1 Quarter新低
                        msg = "[1Q-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 37).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 36).Value, 2))
                        ErrNumber = 1613
                        .Cells(i, 37).Value = .Cells(i, 3).Value    '1 Quarter的最低日期更新
                        .Cells(i, 36).Value = .Cells(i, 7).Value    '1 Quarter的最低價格更新
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
            
                    ElseIf .Cells(i, 7).Value < .Cells(i, 31).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 32).Value) Then  '創2 Months新低
                        msg = "[2M-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 32).Value) & " 日最高點 " & CStr(WorksheetFunction.Round(.Cells(i, 31).Value, 2))
                        ErrNumber = 1614
                        .Cells(i, 32).Value = .Cells(i, 3).Value    '2 Months的最低日期更新
                        .Cells(i, 31).Value = .Cells(i, 7).Value    '2 Months的最低價格更新
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
              
                    ElseIf .Cells(i, 7).Value < .Cells(i, 26).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 27).Value) Then  '創1 Month新低
                        msg = "[1M-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 27).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 26).Value, 2))
                        ErrNumber = 1615
                        .Cells(i, 27).Value = .Cells(i, 3).Value    '1 Month的最低日期更新
                        .Cells(i, 26).Value = .Cells(i, 7).Value    '1 Month的最低價格更新
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
               
                    ElseIf .Cells(i, 7).Value < .Cells(i, 21).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 22).Value) Then  '創2 Weeks新低
                        msg = "[2W-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 22).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 21).Value, 2))
                        ErrNumber = 1616
                        .Cells(i, 22).Value = .Cells(i, 3).Value    '2 Weeks的最低日期更新
                        .Cells(i, 21).Value = .Cells(i, 7).Value    '2 Weeks的最低價格更新
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
              
                    ElseIf .Cells(i, 7).Value < .Cells(i, 16).Value And CDate(.Cells(i, 3).Value) <> CDate(.Cells(i, 17).Value) Then '創1 Week新低
                        msg = "[1W-Low] " & CStr(.Cells(i, 2).Value) & " 跌破 " & CStr(.Cells(i, 17).Value) & " 日最低點 " & CStr(WorksheetFunction.Round(.Cells(i, 16).Value, 2))
                        ErrNumber = 1617
                        .Cells(i, 17).Value = .Cells(i, 3).Value    '1 Week的最低日期更新
                        .Cells(i, 16).Value = .Cells(i, 7).Value    '1 Week的最低價格更新
                        email_Ryan
                    End If
                End If
               
            Next i
        End With
        flat = flat + 1
    End If
    Application.OnTime CDate(Now) + TimeValue("00:05:00"), "notify_Ryan", , True
    Hotmail_normal
    Set WBN = Nothing
Exit Sub
    
100:
    msg = Err.Description & " " & CStr(Now) & " ErrNumber is " & ErrNumber & " i=" & i
    Hotmail_err
    Set WBN = Nothing
End Sub


'Private Sub email_short()
'    'Outlook Objects
'    Dim objOutlook As Object
'    Dim objOutlookMsg As Object
'
'    'Excel Objects
'    Set objOutlook = CreateObject("outlook.application")
'    Set objOutlookMsg = objOutlook.CreateItem(0)
'
'    With objOutlookMsg
'        .Display
'        '.To = "yj-chen@outlook.com;02153440@scu.edu.tw;yungrrrr@gmail.com;"
'        .CC = "yungrrrr@gmail.com;jay.cc.hsieh@gmail.com"
'
'        .Subject = msg
'        .Body = "[Auto Message]"
'        .Body = .Body & Chr(10) & Chr(10) & _
'                  "This is auto e-mail testing" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "Sincerely yours" & Chr(10) & _
'                  "-------------------------------------------------" & Chr(10) & _
'                  "Peter YJ" & Chr(10) & _
'                  "Email： yj-chen@outlook.com.com" & Chr(10) & _
'                  "-------------------------------------------------"
'        .Send
'    End With
'
'    msg = ""   '清空
'    Set objOutlookMsg = Nothing  '指標還在但不指定,釋放記憶體
'    Set objOutlook = Nothing
'End Sub

'Private Sub email_long()
'    'Outlook Objects
'    Dim objOutlook As Object
'    Dim objOutlookMsg As Object
'
'    'Excel Objects
'    Set objOutlook = CreateObject("outlook.application")
'    Set objOutlookMsg = objOutlook.CreateItem(0)
'
'    With objOutlookMsg
'        '.Display
'
'        .To = "yungrrrr@gmail.com;jay.cc.hsieh@gmail.com"
'
'        .Subject = msg
'        .Body = "[Auto Message]"
'        .Body = .Body & Chr(10) & Chr(10) & _
'                  "This is auto e-mail testing" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "Sincerely yours" & Chr(10) & _
'                  "-------------------------------------------------" & Chr(10) & _
'                  "Peter YJ" & Chr(10) & _
'                  "Email： yj-chen@outlook.com.com" & Chr(10) & _
'                  "-------------------------------------------------"
'        .Send
'    End With
'
'    msg = ""
'    Set objOutlookMsg = Nothing
'    Set objOutlook = Nothing
'End Sub
'
'Private Sub email_err()
'    'Outlook Objects
'    Dim objOutlook As Object
'    Dim objOutlookMsg As Object
'
'    'Excel Objects
'    Set objOutlook = CreateObject("outlook.application")
'    Set objOutlookMsg = objOutlook.CreateItem(0)
'
'    With objOutlookMsg
'        '.Display
'        .To = "02153440@scu.edu.tw;"
'        .CC = "jay.cc.hsieh@gmail.com"
'        .Subject = "err" & ErrNumber
'        .Body = "[Auto Message]"
'        .Body = .Body & Chr(10) & Chr(10) & _
'                  "This is auto e-mail testing" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "Sincerely yours" & Chr(10) & _
'                  "-------------------------------------------------" & Chr(10) & _
'                  "Peter YJ" & Chr(10) & _
'                  "Email： yj-chen@outlook.com.com" & Chr(10) & _
'                  "-------------------------------------------------"
'        .Send
'    End With
'
'    msg = ""
'    Set objOutlookMsg = Nothing
'    Set objOutlook = Nothing
'End Sub
'
'Private Sub email_normal()
'    'Outlook Objects
'    Dim objOutlook As Object
'    Dim objOutlookMsg As Object
'
'    'Excel Objects
'    Set objOutlook = CreateObject("outlook.application")
'    Set objOutlookMsg = objOutlook.CreateItem(0)
'
'    With objOutlookMsg
'        .Display
'        '.To = "02153440@scu.edu.tw;"
'        .To = "yungrrrr@gmail.com;jay.cc.hsieh@gmail.com"
'
'        .Subject = "auto_ok"
'        .Body = "[Auto Message]"
'        .Body = .Body & Chr(10) & Chr(10) & _
'                  "This is auto e-mail testing" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "" & Chr(10) & _
'                  "Sincerely yours" & Chr(10) & _
'                  "-------------------------------------------------" & Chr(10) & _
'                  "Peter YJ" & Chr(10) & _
'                  "Email： yj-chen@outlook.com.com" & Chr(10) & _
'                  "-------------------------------------------------"
'        .Send
'    End With
'
'    msg = ""
'    Set objOutlookMsg = Nothing
'    Set objOutlook = Nothing
'End Sub
'
'
'Private Sub Hotmail_short()
'    '============================================================================================
'    '由於需使用到CDO物件，在編寫VBA程式碼前，須先設定引用"Microsoft CDO for Windows 2000 Library"
'    '============================================================================================
'    Dim Mail As CDO.Message
'    Set Mail = New CDO.Message
'    With Mail.Configuration.Fields
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-mail.outlook.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yungrrrr@hotmail.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "dtshomg100"
'        .Update
'    End With
'
'    With Mail
'        .Subject = "Hotmail_short" & msg
'        .From = "okihuvguyy@gmail.com"
'        .To = "yungrrrr@gmail.com;jay.cc.hsieh@gmail.com"
'        .CC = ""
'        .HTMLBody = msg
'        .BodyPart.Charset = "utf-8"
'        .HTMLBodyPart.Charset = "utf-8"
'        .Send
'    End With
'    'MsgBox "信件已寄出", vbInformation, "寄出"
'
'    Set Mail = Nothing
'
'    End Sub
'    Private Sub Hotmail_normal()
'    '============================================================================================
'    '由於需使用到CDO物件，在編寫VBA程式碼前，須先設定引用"Microsoft CDO for Windows 2000 Library"
'    '============================================================================================
'    Dim Mail As CDO.Message
'    Set Mail = New CDO.Message
'    With Mail.Configuration.Fields
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-mail.outlook.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yungrrrr@hotmail.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "dtshomg100"
'        .Update
'    End With
'
'    With Mail
'        .Subject = "ok"
'        .From = "yungrrrr@hotmail.com"
'        .To = "yungrrrr@gmail.com;jay.cc.hsieh@gmail.com"
'        .CC = ""
'        .HTMLBody = ""
'        .BodyPart.Charset = "utf-8"
'        .HTMLBodyPart.Charset = "utf-8"
'        .Send
'    End With
'    'MsgBox "信件已寄出", vbInformation, "寄出"
'
'    Set Mail = Nothing
'End Sub
'
'Private Sub Hotmail_long()
'    '============================================================================================
'    '由於需使用到CDO物件，在編寫VBA程式碼前，須先設定引用"Microsoft CDO for Windows 2000 Library"
'    '============================================================================================
'    Dim Mail As CDO.Message
'    Set Mail = New CDO.Message
'    With Mail.Configuration.Fields
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-mail.outlook.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yungrrrr@hotmail.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "dtshomg100"
'        .Update
'    End With
'
'    With Mail
'        .Subject = "Hotmail_long" & msg
'        .From = "okihuvguyy@gmail.com"
'        .To = "yungrrrr@gmail.com;jay.cc.hsieh@gmail.com"
'        .CC = ""
'        .HTMLBody = msg
'        .BodyPart.Charset = "utf-8"
'        .HTMLBodyPart.Charset = "utf-8"
'        .Send
'    End With
'    'MsgBox "信件已寄出", vbInformation, "寄出"
'
'    Set Mail = Nothing
'End Sub


'Private Sub Hotmail_err()
'    '============================================================================================
'    '由於需使用到CDO物件，在編寫VBA程式碼前，須先設定引用"Microsoft CDO for Windows 2000 Library"
'    '============================================================================================
'    Dim Mail As CDO.Message
'    Set Mail = New CDO.Message
'    With Mail.Configuration.Fields
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-mail.outlook.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "yungrrrr@hotmail.com"
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "dtshomg100"
'        .Update
'    End With
'
'    With Mail
'        .Subject = "Hotmail_err" & msg
'        .From = "okihuvguyy@gmail.com"
'        .To = "yungrrrr@gmail.com;jay.cc.hsieh@gmail.com"
'        .CC = ""
'        .HTMLBody = msg
'        .BodyPart.Charset = "utf-8"
'        .HTMLBodyPart.Charset = "utf-8"
'        .Send
'    End With
'    'MsgBox "信件已寄出", vbInformation, "寄出"
'
'    Set Mail = Nothing
'End Sub
'
