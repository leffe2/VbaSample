VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WriteCsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub WriteCsv()

    MsgBox "処理を開始します。"

    Dim nowDate As String
    nowDate = Format(Now, "yyyymmddhhnnss")

    Dim folderPath As String
    folderPath = ThisWorkbook.path & "\CSV_" & nowDate
    
    Call CreateFolder(folderPath)

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        Call WriteCsvByWs(ws, folderPath)
    
    Next
    
    MsgBox "処理が完了しました。" & folderPath & "を確認下さい。"

End Sub

Private Sub CreateFolder(path As String)

    With CreateObject("Scripting.FileSystemObject")
    
        .CreateFolder path
    
    End With

End Sub

Private Sub WriteCsvByWs(ws As Worksheet, folderPath As String)

    Dim arrays As Variant
    arrays = GetArrays(ws)
    
    With CreateObject("ADODB.Stream")
    
        .Charset = "UTF-8"
        .LineSeparator = 10 '10:LF
        .Open
        
        Dim i As Long
        For i = LBound(arrays, 1) To UBound(arrays, 1)
        
            Dim line As String
            line = Join(WorksheetFunction.Index(arrays, i, 0), ",") '0:その行全体を取得する。
        
            .WriteText line, 1 '1:文字列の末尾に改行コードを付けて書き込む。
        
        Next
        
        Dim csvPath As String
        csvPath = folderPath & "\" & ws.Name & ".csv"
    
        .SaveToFile csvPath, 1 '1:ファイルが存在する場合エラー。無い場合は、新規作成する。
        .Close
    
    End With

End Sub

Private Function GetArrays(ws As Worksheet) As Variant

    With ws
    
        Dim lastRow As Long
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim lastColumn As Long
        lastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
        
        Dim arrays As Variant
        arrays = .Range(.Cells(1, 1), .Cells(lastRow, lastColumn)).Value
    
    End With
    
    If UBound(arrays, 1) < 2 Then
    
        Err.Raise Number:="40000", Description:="エラーが発生しました。:" & ws.Name & "をCSVにできません。"
    
    End If
    
    GetArrays = arrays

End Function



