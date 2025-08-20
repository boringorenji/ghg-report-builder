Attribute VB_Name = "Module1"
Sub export_value_as_formatted()
    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim r As Long, c As Long
    Dim rowCount As Long, colCount As Long
    Dim displayText As String
    Dim tempVal As Variant
    Dim cellFormat As String

    ' 建立新檔案
    Set newWb = Workbooks.Add

    ' 移除多餘分頁
    Application.DisplayAlerts = False
    Do While newWb.Sheets.Count > 1
        newWb.Sheets(2).Delete
    Loop
    Application.DisplayAlerts = True

    ' 處理所有工作表
    For Each ws In ThisWorkbook.Worksheets
        ' 新增對應分頁
        If newWb.Sheets.Count = 1 And (newWb.Sheets(1).Name = "工作表1" Or newWb.Sheets(1).Name = "Sheet1") Then
            Set newWs = newWb.Sheets(1)
        Else
            Set newWs = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.Count))
        End If
        newWs.Name = ws.Name

        ' 找出範圍
        On Error Resume Next
        rowCount = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        colCount = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        On Error GoTo 0

        ' 複製資料
        If rowCount > 0 And colCount > 0 Then
            For r = 1 To rowCount
                For c = 1 To colCount
                    With ws.Cells(r, c)
                        displayText = .Text
                        cellFormat = .NumberFormat

                        ' 嘗試將顯示值轉為數值（若適用）
                        On Error Resume Next
                        If IsNumeric(displayText) Then
                            tempVal = CDbl(displayText)
                        Else
                            tempVal = displayText
                        End If
                        On Error GoTo 0

                        ' 寫入值
                        newWs.Cells(r, c).Value = tempVal

                        ' ? 防止科學記號、處理 General 格式
                        If cellFormat = "General" Then
                            If IsNumeric(.Value) Then
                                If Abs(CDbl(.Value)) < 0.0001 Then
                                    newWs.Cells(r, c).NumberFormat = "0.0000000000"
                                Else
                                    newWs.Cells(r, c).NumberFormat = "0.00"
                                End If
                            Else
                                newWs.Cells(r, c).NumberFormat = "@" ' 非數字：設成文字格式
                            End If
                        Else
                            newWs.Cells(r, c).NumberFormat = cellFormat
                        End If
                    End With
                Next c
            Next r

            ' 複製欄寬
            For c = 1 To colCount
                newWs.Columns(c).ColumnWidth = ws.Columns(c).ColumnWidth
            Next c
        End If
    Next ws

    MsgBox "Done! All pages have been copied as pure values, retaining the display format, no formulas, no scientific notation, and decimals and percentages are correctly retained."
End Sub
