Option Explicit

' =====================================================
'  ConcatListBuilder
'  Author: Carlos Pulido Rosas
'  Purpose: Concatenate values from one or two columns into a single string.
'  Usage:
'       Select a cell within the target column, then run the macro.
'
'       Parameters:
'           - quoteMode: 0=None, 1=Single quotes, 2=Double quotes
'           - separator: string (default ", ")
'           - pairMode: True to concatenate value pairs from two columns
' =====================================================
Public Sub ConcatListBuilder(Optional ByVal quoteMode As Integer = 0, _
                             Optional ByVal separator As String = ", ", _
                             Optional ByVal pairMode As Boolean = False)

    Dim ws As Worksheet
    Dim sourceCol As Long
    Dim lastRow As Long
    Dim valuesArr As Variant
    Dim resultStr As String
    Dim quoteChar As String
    Dim i As Long
    Dim val1 As String
    Dim val2 As String

    On Error GoTo ErrHandler

    Set ws = ActiveSheet
    sourceCol = ActiveCell.Column
    lastRow = ws.Cells(ws.Rows.Count, sourceCol).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found below header.", vbExclamation, "ConcatListBuilder"
        Exit Sub
    End If

    ' --- Select quote character ---
    Select Case quoteMode
        Case 1: quoteChar = "'"     ' Single quote
        Case 2: quoteChar = Chr(34) ' Double quote
        Case Else: quoteChar = ""
    End Select

    ' --- Load data into memory (faster than reading cell-by-cell) ---
    valuesArr = ws.Range(ws.Cells(2, sourceCol), ws.Cells(lastRow, sourceCol)).Value
    resultStr = ""

    ' --- Build concatenation string ---
    If pairMode Then
        For i = 1 To UBound(valuesArr, 1)
            val1 = Trim(valuesArr(i, 1))
            val2 = Trim(ws.Cells(i + 1, sourceCol + 1).Value)
            If val1 <> "" And val1 <> "-1" Then
                resultStr = resultStr & "(" & val1 & separator & val2 & "), "
            End If
        Next i
    Else
        For i = 1 To UBound(valuesArr, 1)
            val1 = Trim(valuesArr(i, 1))
            If val1 <> "" And val1 <> "-1" Then
                resultStr = resultStr & quoteChar & val1 & quoteChar & separator
            End If
        Next i
    End If

    ' --- Trim trailing separator ---
    If Len(resultStr) >= Len(separator) Then
        resultStr = Left(resultStr, Len(resultStr) - Len(separator))
    End If

    ' --- Output beside source column ---
    ws.Cells(1, sourceCol + 1).Value = resultStr

    MsgBox "Concatenation complete.", vbInformation, "ConcatListBuilder"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "ConcatListBuilder"

End Sub
