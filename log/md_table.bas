Attribute VB_Name = "md_table"
'�汾��
Public Const TOOL_VERSION As String = "6.0"

'Excel��ѯ����У���һ�����鵥�ʵ�����
Public Const HEAD_ROW As Integer = 5

'��ȡExcel���򵥴ʵ�rangeֵ
Public Function getWordList() As Range
    Dim r As Integer
    
    r = Range("Sheet1!A65536").End(xlUp).Row
    If r >= HEAD_ROW Then
        Set getWordList = Range("Sheet1!A" & HEAD_ROW & ":A" & r)
    Else
        Set getwordlits = Nothing
    End If
End Function
'��ȡExcel�����ͷ��rangeֵ
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
        
        '�����ͷ
        cText = "![](reading_log.png)" & vbCrLf & vbCrLf
        fsT.writetext cText
        
        '�������б��������ѭ��
        For Each x In getWordList()
            cText = ""
            For Each y In columnlist
                cText = cText & "|" & x.Offset(0, y.Column() - 1).Value
            Next y
            cText = cText & "|" & vbCrLf '�����һ��tab���ɻس���
            fsT.writetext cText
        Next x

            
        '�����ļ�
        fsT.SaveToFile tFilePath, 2
        MsgBox "�ļ��Ѿ����ɣ�" & vbCrLf & vbCrLf & "Ŀ¼: " & tFilePath

    Else
        MsgBox "�޼�¼�ɱ��棡"
    End If
    fsT.Close
End Sub

