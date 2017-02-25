Attribute VB_Name = "md_table"
'版本号
Public Const TOOL_VERSION As String = "6.0"

'Excel查询表格中，第一个待查单词的行数
Public Const HEAD_ROW As Integer = 5

'获取Excel纵向单词的range值
Public Function getWordList() As Range
    Dim r As Integer
    
    r = Range("Sheet1!A65536").End(xlUp).Row
    If r >= HEAD_ROW Then
        Set getWordList = Range("Sheet1!A" & HEAD_ROW & ":A" & r)
    Else
        Set getwordlits = Nothing
    End If
End Function
'获取Excel横向表头的range值
Public Function getColList() As Range
    Dim c As Integer
    
    c = Range("Sheet1!A" & HEAD_ROW).End(xlToRight).Column
    If c >= 1 Then
        Set getColList = Sheets("Sheet1").Range(Cells(HEAD_ROW, 1).Address, Cells(HEAD_ROW, c).Address)
    Else
        Set getColList = Nothing
    End If

End Function

Sub WriteOut()

    
    Dim fsT, cText, tFilePath As String
    Dim x, y, columnlist As Range
    
    
    tFilePath = Application.ActiveWorkbook.Path + "\reading_log.txt"
    'Create Stream object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2
    fsT.Charset = "utf-8"

    'Open the stream And write binary data To the object
    fsT.Open
    
    If Not getWordList() Is Nothing Then
        
        Set columnlist = getColList()
        
        '输出表头
        cText = "![](reading_log.png)" & vbCrLf & vbCrLf
        fsT.writetext cText
        
        '按单词列表输出的主循环
        For Each x In getWordList()
            cText = ""
            For Each y In columnlist
                cText = cText & "|" & x.Offset(0, y.Column() - 1).Value
            Next y
            cText = cText & "|" & vbCrLf '将最后一个tab换成回车符
            fsT.writetext cText
        Next x

            
        '保存文件
        fsT.SaveToFile tFilePath, 2
        MsgBox "文件已经生成！" & vbCrLf & vbCrLf & "目录: " & tFilePath

    Else
        MsgBox "无记录可保存！"
    End If
    fsT.Close
End Sub

