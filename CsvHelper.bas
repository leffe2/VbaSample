Attribute VB_Name = "CsvHelper"
Option Explicit

Public Function GetCsvPath() As String

    Dim filePath As String
    filePath = Application.GetOpenFilename(FileFilter:="csv�t�@�C��(*.csv;),")
    
    If filePath = "False" Then
    
        GetCsvPath = ""
        
        Exit Function
    
    End If
    
    GetCsvPath = filePath

End Function


Public Function ReadCsv(csvPath As String) As Variant
        
    With CreateObject("ADODB.Stream")
    
        .Charset = "UTF-8"      ' �����R�[�h�w��i"UTF-8"�A"shift_jis"�j
        .LineSeparator = 10     ' ���s�R�[�h�w��i-1�FCRLF�A10�FLF�ACR�F13�j
        .Open
        .LoadFromFile csvPath
        
        Dim lines() As Variant
        
        Dim i As Long
        
        Do Until .EOS
    
            Dim line As String
            line = .ReadText(-2) ' �ǂݍ��ݕ��@�i-1�F�ꊇ�A-2�F1�s���j
            
            Dim values As Variant
            values = Split(line, ",")
            
            ReDim Preserve lines(UBound(values), i)
    
            Dim j As Long
            For j = LBound(values) To UBound(values)
            
                On Error GoTo Catch
                lines(j, i) = values(j)
            
            Next
            
            i = i + 1
            
        Loop
    
        .Close
    
    End With
    
    ReadCsv = Transpose(lines)
    
Catch:
    If Err.Number = 9 Then
    
        MsgBox "CSV�̃w�b�_�̗񐔂ƃf�[�^�̗񐔂���v���܂���B" & vbCrLf & Err.Number & ":" & Err.Description
    
    ElseIf Err.Number <> 0 Then
    
        MsgBox "�\�����ʃG���[���������܂���" & vbCrLf & Err.Number & ":" & Err.Description
    
    End If

End Function


Private Function Transpose(lines As Variant) As Variant

    Dim tmp As Variant
    ReDim tmp(LBound(lines, 2) To UBound(lines, 2), LBound(lines, 1) To UBound(lines, 1))
    
    Dim i As Long
    
    Dim j As Long
    
    For i = LBound(lines, 1) To UBound(lines, 1)
    
        For j = LBound(lines, 2) To UBound(lines, 2)
        
            tmp(j, i) = lines(i, j)
        
        Next
        
    Next

    Transpose = tmp

End Function
