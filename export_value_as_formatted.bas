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

    ' �إ߷s�ɮ�
    Set newWb = Workbooks.Add

    ' �����h�l����
    Application.DisplayAlerts = False
    Do While newWb.Sheets.Count > 1
        newWb.Sheets(2).Delete
    Loop
    Application.DisplayAlerts = True

    ' �B�z�Ҧ��u�@��
    For Each ws In ThisWorkbook.Worksheets
        ' �s�W��������
        If newWb.Sheets.Count = 1 And (newWb.Sheets(1).Name = "�u�@��1" Or newWb.Sheets(1).Name = "Sheet1") Then
            Set newWs = newWb.Sheets(1)
        Else
            Set newWs = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.Count))
        End If
        newWs.Name = ws.Name

        ' ��X�d��
        On Error Resume Next
        rowCount = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        colCount = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        On Error GoTo 0

        ' �ƻs���
        If rowCount > 0 And colCount > 0 Then
            For r = 1 To rowCount
                For c = 1 To colCount
                    With ws.Cells(r, c)
                        displayText = .Text
                        cellFormat = .NumberFormat

                        ' ���ձN��ܭ��ର�ƭȡ]�Y�A�Ρ^
                        On Error Resume Next
                        If IsNumeric(displayText) Then
                            tempVal = CDbl(displayText)
                        Else
                            tempVal = displayText
                        End If
                        On Error GoTo 0

                        ' �g�J��
                        newWs.Cells(r, c).Value = tempVal

                        ' ? �����ǰO���B�B�z General �榡
                        If cellFormat = "General" Then
                            If IsNumeric(.Value) Then
                                If Abs(CDbl(.Value)) < 0.0001 Then
                                    newWs.Cells(r, c).NumberFormat = "0.0000000000"
                                Else
                                    newWs.Cells(r, c).NumberFormat = "0.00"
                                End If
                            Else
                                newWs.Cells(r, c).NumberFormat = "@" ' �D�Ʀr�G�]����r�榡
                            End If
                        Else
                            newWs.Cells(r, c).NumberFormat = cellFormat
                        End If
                    End With
                Next c
            Next r

            ' �ƻs��e
            For c = 1 To colCount
                newWs.Columns(c).ColumnWidth = ws.Columns(c).ColumnWidth
            Next c
        End If
    Next ws

    MsgBox "Done! All pages have been copied as pure values, retaining the display format, no formulas, no scientific notation, and decimals and percentages are correctly retained."
End Sub
