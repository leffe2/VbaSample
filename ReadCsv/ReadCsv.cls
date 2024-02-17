VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReadCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ReadCsv(csvPath As String) As Variant
        
    With CreateObject("ADODB.Stream")
    
        .Charset = "UTF-8"      ' 文字コード指定（"UTF-8"、"shift_jis"）
        .LineSeparator = 10     ' 改行コード指定（-1：CRLF、10：LF、CR：13）
        .Open
        .LoadFromFile csvPath
        
        Dim lines() As Variant
        Dim i As Long
        Do Until .EOS
    
            Dim line As String
            line = .ReadText(-2) ' 読み込み方法（-1：一括、-2：1行ずつ）
            
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
    
        MsgBox "CSVのヘッダの列数とデータの列数が一致しません。" & vbCrLf & Err.Number & ":" & Err.Description
    
    ElseIf Err.Number <> 0 Then
    
        MsgBox "予期せぬエラーが発生しました" & vbCrLf & Err.Number & ":" & Err.Description
    
    End If

End Function

Public Function GetCsvPath() As String

    Dim filePath As String
    filePath = Application.GetOpenFilename(FileFilter:="csvファイル(*.csv;),")
    
    If filePath = "False" Then
    
        GetCsvPath = ""
        
        Exit Function
    
    End If
    
    GetCsvPath = filePath

End Function

Public Sub WriteWorkSheet(lines As Variant, wsName As String)

    Dim ws As Worksheet
    Set ws = Worksheets().Add(After:=Worksheets(Worksheets.Count))
    ws.Name = wsName

    With ws
    
        Dim rowLength As Long
        rowLength = UBound(lines, 1) - LBound(lines, 1) + 1
        
        Dim columnLength As Long
        columnLength = UBound(lines, 2) - LBound(lines, 2) + 1

        .Range(.Cells(1, 1), .Cells(1, 1)).Resize(rowLength, columnLength).NumberFormatLocal = "@"
        .Range(.Cells(1, 1), .Cells(1, 1)).Resize(rowLength, columnLength).Value = lines
    
    End With

End Sub

Private Function Transpose(lines As Variant) As Variant

    Dim tmp As Variant
    ReDim tmp(LBound(lines, 2) To UBound(lines, 2), LBound(lines, 1) To UBound(lines, 1))
        
    Dim i As Long, j As Long
    For i = LBound(lines, 1) To UBound(lines, 1)
    
        For j = LBound(lines, 2) To UBound(lines, 2)
        
            tmp(j, i) = lines(i, j)
        
        Next
        
    Next

    Transpose = tmp

End Function
