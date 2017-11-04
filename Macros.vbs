Function ConcatenateIfs(ConcatenateRange As Range, ParamArray Criteria() As Variant) As Variant
        ' Source: EileensLounge.com, August 2014
        Dim i As Long
        Dim c As Long
        Dim n As Long
        Dim f As Boolean
        Dim Separator As String
        Dim strResult As String
        Dim col As Collection
        On Error GoTo ErrHandler
        n = UBound(Criteria)
        If n < 3 Then
            ' Too few arguments
            GoTo ErrHandler
        End If
        If n Mod 3 = 0 Then
            ' Separator specified explicitly
            Separator = Criteria(n)
        Else
            ' Use default separator
            Separator = ","
        End If
        ' Initialize collection of unique items
        Set col = New Collection
        ' Loop through the cells of the concatenate range
        For i = 1 To ConcatenateRange.Count
            ' Start by assuming that we have a match
            f = True
            ' Loop through the conditions
            For c = 0 To n - 1 Step 3
                ' Does cell in criteria range match the condition?
                Select Case Criteria(c + 1)
                    Case "<="
                        If Criteria(c).Cells(i).Value > Criteria(c + 2) Then
                            f = False
                            Exit For
                        End If
                    Case "<"
                        If Criteria(c).Cells(i).Value >= Criteria(c + 2) Then
                            f = False
                            Exit For
                        End If
                    Case ">="
                        If Criteria(c).Cells(i).Value < Criteria(c + 2) Then
                            f = False
                            Exit For
                        End If
                    Case ">"
                        If Criteria(c).Cells(i).Value <= Criteria(c + 2) Then
                            f = False
                            Exit For
                        End If
                    Case "<>"
                        If Criteria(c).Cells(i).Value = Criteria(c + 2) Then
                            f = False
                            Exit For
                        End If
                    Case Else
                        If Criteria(c).Cells(i).Value <> Criteria(c + 2) Then
                            f = False
                            Exit For
                        End If
                End Select
            Next c
            ' Were all criteria satisfied?
            If f Then
                ' If so, add value to collection, if it has not been added yet
                On Error Resume Next
                col.Add Item:=ConcatenateRange.Cells(i).Value, _
                    Key:=CStr(ConcatenateRange.Cells(i).Value)
                On Error GoTo ErrHandler
            End If
        Next i
        If col.Count > 0 Then
            ' Sort the results
            SortCollection col
            ' Concatenate them
            For i = 1 To col.Count
                strResult = strResult & Separator & col(i)
            Next i
            ' Remove first separator
            strResult = Mid(strResult, Len(Separator) + 1)
        End If
        ConcatenateIfs = strResult
        Exit Function
        ErrHandler:
            ConcatenateIfs = CVErr(xlErrValue)
End Function

Sub SortCollection(col As Collection)
            Dim i As Long
            Dim j As Long
            Dim tmp As Variant
            For i = 1 To col.Count - 1
                For j = i + 1 To col.Count
                    If col(j) < col(i) Then
                        tmp = col(j)
                        col.Remove Index:=j
                        col.Add Item:=tmp, Key:=CStr(tmp), Before:=i
                    End If
                Next j
            Next i
End Sub

Function GetDiffs(Cell1 As Range, Cell2 As Range) As String
    Dim Array1, Array2, lLoop As Long
    Dim strDiff As String, strDiffs As String
    Dim lCheck As Long
     
     
    Array1 = Split(Replace(Cell1, " ", ""), ",")
    Array2 = Split(Replace(Cell2, " ", ""), ",")
    On Error Resume Next
    With WorksheetFunction
        For lLoop = 0 To UBound(Array1)
            strDiff = vbNullString
            strDiff = .Index(Array2, 1, .Match(Array1(lLoop), Array2, 0))
            If strDiff = vbNullString Then
                lCheck = 0
                lCheck = .Match(Array1(lLoop), Array2, 0)
                 
                If lCheck = 0 Then
                    strDiffs = strDiffs & "," & Array1(lLoop)
                End If
            End If
             
        Next lLoop
    End With
     
    GetDiffs = Trim(Right(strDiffs, Len(strDiffs) - 1))
End Function

Sub SustituirUnidad()
    ' Replaces all empty cell with 0s and all numerical cells with 1.
    ' This macro copies the selected range values to an array and turns off
    ' external events to reduce time consumtion.

    Dim arr() As Variant
    Dim rng As Range
    Dim i As Long, _
        j As Long
    Dim floor As Long
    Dim ceiling As Long
    Dim replacement_value

    'assign the worksheet range to a variable
    Set rng = Selection
    floor = 0
    'ceiling = 500
    replacement_value = 1

    ' copy the values in the worksheet range to the array
    arr = rng

    ' turn off time-consuming external operations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    'loop through each element in the array
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            'do the comparison of the value in an array element
            'with the criteria for replacing the value
            If arr(i, j) > floor Then 'And arr(i, j) < ceiling Then
                arr(i, j) = replacement_value
            ElseIf arr(i, j) = "" Then
                arr(i, j) = 0
            End If
        Next j
    Next i

    'copy array back to worksheet range
    rng = arr

    'turn events back on
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Sub InsSepColumn()
    ' InsSepColumn Macro
    ' Inserta columna, ancho 2 y en blanco, como separador.

    Selection.EntireColumn.Insert
    ActiveCell.Columns("A:A").EntireColumn.Select
    ActiveCell.Activate
    Selection.Clear
    Selection.ColumnWidth = 2
End Sub

Sub Multi_FindReplace()
    ' Find & Replace a list of text/values throughout a selection from a table

    Dim fndList As Integer
    Dim rplcList As Integer
    Dim tbl As ListObject
    Dim myArray As Variant

    'Create variable to point to your table
    Set tbl = Worksheets("Reemplazar").ListObjects("tblReemplazar")

    'Create an Array out of the Table's Data
    Set TempArray = tbl.DataBodyRange
    myArray = Application.Transpose(TempArray)
  
    'Designate Columns for Find/Replace data
    fndList = 1
    rplcList = 2

    'Loop through each item in Array lists
    For x = LBound(myArray, 1) To UBound(myArray, 2)
        Selection.Cells.Replace What:=myArray(fndList, x), Replacement:=myArray(rplcList, x), _
        LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False
    Next x
End Sub

Sub GroupCols()
    ' GroupCols Macro
    ' Agrupa n columnas a la derecha.

    ActiveCell.Columns("A:C").EntireColumn.Select
    ActiveCell.Activate
    Selection.Columns.Group
End Sub

Sub GroupReplace()
    ' GroupReplace Macro

    Application.Run "PERSONAL.XLSB!GroupCols"
    ActiveCell.Offset(0, 2).Columns("A:A").EntireColumn.Select
    Application.Run "PERSONAL.XLSB!Multi_FindReplace"
End Sub

Sub uniqueValues()
    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        Dim d
        Set d = CreateObject("Scripting.Dictionary")

        Dim rng As Range
        Dim arr() as Variant

        Set rng = Selection
        arr = rng
        ary = rng

        For i = LBound(arr, 1) To UBound(arr, 1)
            If d.Exists(arr(i, 1)) = true Then
                ary(i, 1) = 0
            Else
                ary(i, 1) = 1
                d.Add arr(i, 1), i
            End If
        Next
        rng.Offset(0, 1) = ary
        ' Turn events back on
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub

Sub multiplyRange()
    ' Takes a range as input and multiplies it by a factor.
        
    Set updateRng = Application.InputBox(prompt:="Select a range", Type:=8)
    Dim factor As Double
    factor = 1.045
    
    With Application
        ' Turn off screen updating to increase performance
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        
        Dim rng As Range
        Dim arr() As Variant

        Set rng = updateRng
        arr = rng

        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                arr(i, j) = arr(i, j) * factor
            Next
        Next
        rng = arr
        ' Turn events back on
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
