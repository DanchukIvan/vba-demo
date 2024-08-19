Attribute VB_Name = "DataTransformer"
'@Folder "S3FileDownloader.DataProcessing"
'@ModuleDescription("Модуль содержит функции для преобразования данных в соответствии с _
                     настройками скрипта.")
                        
Option Explicit
Option Private Module

'Процедура выполняет роль точки входа для всех функций преобразования данных.
Public Sub Pipeline(ByRef Config As ConfigClass)
        If Config.FieldsWithDates(0) <> False Then CleareDates Config
        If Config.FieldsWithTags(0) <> False Then CleareTags Config
End Sub

'Процедура очищает определенные пользователем поля таблицы от тегов, заменяя их на пустоту.
Public Sub CleareTags(ByRef Config As ConfigClass)
    Dim Filter As New RegExp, Field As Variant, ColumnIdx As Long, ResultArr() As Variant
    Dim DataBody As Variant, CellValue As String
    Dim WorkRange As ListObject, idx As Variant, OldText As String
    
    Filter.Pattern = "</?(?:\w+)?/?>"
    Filter.Global = True
    
    Set WorkRange = Sheets("Лист1").ListObjects("result")
    
    For Each Field In Config.FieldsWithTags
        ColumnIdx = WorkRange.HeaderRowRange.Find(Field).Column
        DataBody = Application.Transpose(WorkRange.ListColumns(ColumnIdx).DataBodyRange.Value)
        For idx = LBound(DataBody) To UBound(DataBody)
            CellValue = DataBody(idx)
            If Filter.Test(CellValue) Then
                OldText = Filter.Replace(CellValue, "")
                DataBody(idx) = OldText
            End If
        Next idx
        WorkRange.ListColumns(ColumnIdx).DataBodyRange = Application.Transpose(DataBody)
    Next Field
    
    ActiveWorkbook.Save
    
End Sub


'В отличие от предыдущей процедуры, эта извлекает текст из ячеек нужного столбца, значения которых _
содержат не поддерживаемые VBA и Excel строки дат, преобразуя в валидные даты
Public Sub CleareDates(ByRef Config As ConfigClass)
    Dim Filter As New RegExp, Matches As Object, ColumnIdx As Long, WorkRange As ListObject
    Dim DateString As String, idx As Variant, ResultArr() As Variant, DataBody As Variant, Field As Variant
    
    Filter.Pattern = "(?=\d{2})\d{2}.\d{2}.\d{4}|(?=\d{4})\d{4}.\d{2}.\d{2}"
    Filter.Global = True

    Set WorkRange = Sheets("Лист1").ListObjects("result")
    For Each Field In Config.FieldsWithDates
        ColumnIdx = WorkRange.HeaderRowRange.Find(Field).Column
        DataBody = Application.Transpose(WorkRange.ListColumns(ColumnIdx).DataBodyRange.Value)
        For idx = LBound(DataBody) To UBound(DataBody)
            If Filter.Test(DataBody(idx)) Then
                Set Matches = Filter.Execute(DataBody(idx))
                ResultArr = Utils.CollectionToArray(Matches)
                DateString = WorksheetFunction.Index(ResultArr, 1)
                DataBody(idx) = CDate(DateString)
            End If
        Next idx
        WorkRange.ListColumns(ColumnIdx).DataBodyRange = Application.Transpose(DataBody)
    Next Field
    
    ActiveWorkbook.Save
    
End Sub
