Attribute VB_Name = "Select_and_Fill"


' -------------------------------------------------------------

Sub Fill_yellow_from_selected_cell_to_last_non_empty_cell_in_column()
Attribute Fill_yellow_from_selected_cell_to_last_non_empty_cell_in_column.VB_Description = "Выделяет столбец желтым цветом от выделенной ячейки до последней заполненной ячейки в столбце"
Attribute Fill_yellow_from_selected_cell_to_last_non_empty_cell_in_column.VB_ProcData.VB_Invoke_Func = " \n14"
' Выделяет столбец желтым цветом от выделенной ячейки 
' до последней заполненной ячейки в столбце

    ' Выделение требуемого диапазона:
    Range(Selection, Selection.End(xlDown)).Select
    
    ' Назначение для выделенного диапазона
    ' заливки фона ячеек желтым цветом:
    Selection.Interior.Color = vbYellow
    
End Sub


' -------------------------------------------------------------

Sub Fill_yellow_from_top_cell_to_last_non_empty_cell_in_column()
' Выделяет столбец желтым цветом от первой ячейки данного столбца
' до последней заполненной ячейки в столбце

    ' Объявление переменной, в которой будет храниться
    ' номер активного столбца:
    Dim activeColumn As Integer
    

    ' Сохранение в переменную номера активного столбца:
    activeColumn = ActiveCell.Column
    
    ' Выделение требуемого диапазона:
    Range(Cells(1, activeColumn), Selection.End(xlDown)).Select
    
    ' Назначение для выделенного диапазона
    ' заливки фона ячеек желтым цветом:
    Selection.Interior.Color = vbYellow

End Sub


' -------------------------------------------------------------

Sub Fill_yellow_entire_column_of_selected_cell()
' Выделяет желтым цветом весь столбец,
' в котором располагается выделенная ячейка

    ' Объявление переменной, в которой будет храниться
    ' номер активного столбца:
    Dim activeColumn As Integer
    
    
    ' Сохранение в переменную номера активного столбца:
    activeColumn = ActiveCell.Column
    
    ' Выделение требуемого диапазона:
    Range(Columns(activeColumn), Columns(activeColumn)).Select
    
    ' Назначение для выделенного диапазона
    ' заливки фона ячеек желтым цветом:
    Selection.Interior.Color = vbYellow

End Sub

' -------------------------------------------------------------
