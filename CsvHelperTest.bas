Attribute VB_Name = "CsvHelperTest"
Option Explicit


Public Sub Csvのファイルパスが取得できる()

    Dim csvPath As String
    csvPath = GetCsvPath()
    
    Debug.Print (csvPath)

End Sub

Public Sub Csvが2次元配列として取得できる()

    Dim csvPath As String
    csvPath = GetCsvPath()
    
    Debug.Print (csvPath)
    
    Dim lines As Variant
    lines = ReadCsv(csvPath)

    With ThisWorkbook.Sheets(1)
    
    
        Dim rowLength As Long
        rowLength = UBound(lines, 1) - LBound(lines, 1) + 1
        
        Dim columnLength As Long
        columnLength = UBound(lines, 2) - LBound(lines, 2) + 1
        
        .Range(.Cells(1, 1), .Cells(1, 1)).Resize(rowLength, columnLength).NumberFormatLocal = "@"
        
        .Range(.Cells(1, 1), .Cells(1, 1)).Resize(rowLength, columnLength).Value = lines

    End With

End Sub
