Attribute VB_Name = "FormEntrypoint"
'@Folder "S3FileDownloader.FormEntrypoint"
'@ModuleDescription("Модуль содержит различные утилиты, которые используются только при работе с _
                    формами.")

Option Explicit

'Используются списки, так как с ними гораздо удобнее работать чем с массивами
Public MacroVarsList As New mscorlib.ArrayList
Public DPVarsList As New mscorlib.ArrayList
Public FolderPickersList As New mscorlib.ArrayList

Public MacroRng As Range
Public DataProcRng As Range

Public DescDict As New Dictionary
Public Memento As New Dictionary

'@VariableDescription("Данная переменная - константа. Определяется в момент начала работы формы. _
                    Используется для манипуляции значениями настроек.")
Public ConfWorkingWsh As Worksheet

'Точка входа для отображения окна настройки скрипта
Public Sub StartVacancyParser()
    Load MainWindows
    MainWindows.Show
End Sub

'Процедура инициализирует дефолтный стейт - собирает все переменные, которые можно изменить _
в форме настройки и кэширует их в словаре на случай, если понадобится восстановить их.

Public Sub SetInitVars()

    Dim vFolderPickers As Variant
    Dim MacroVars() As String
    Dim DPVars() As String
    Dim FolderPickers() As String
    Dim var As Variant
    Dim DescRng As Range
    Dim Row As Variant
    
    'Очищаем кэш если в нем что-то есть на момент инициализации (навряд ли что-то будет конечно).
    If Memento.Count <> 0 Then
        Memento.RemoveAll
    End If
    
    'Активируем лист с бэкендом формы
    Set ConfWorkingWsh = ThisWorkbook.Sheets("Конфигуратор")
    ConfWorkingWsh.Activate
    
    'Рэндж с настройками макроса, которые являются обязательными (не могут быть пустыми)
    Set MacroRng = ConfWorkingWsh.Range("required_params").Offset(0, -1)
    'Рэндж с настройками обработки (могут быть пустыми)
    Set DataProcRng = ConfWorkingWsh.Range("postproc_settings").Columns(1)
    'Собираем настройки, которые требуют выбора папок, в массив. Он нужен для разграничения настроек, _
    которым нужна директория и которым нужен файл.
    vFolderPickers = Application.Transpose(ConfWorkingWsh.Range("folder_picker").Offset(0, -1))
    
    'Обходим все настройки и кэшируем дефолтные значения в словарь
    MacroVars = Split(Join(Application.Transpose(MacroRng.Value), ","), ",")
    For Each var In MacroRng.Rows
        Debug.Print var.Columns(1).Value
        MacroVarsList.Add var.Columns(1).Value
        If Not Memento.Exists(var.Columns(1).Value) Then
            Memento.Add var.Columns(1).Value, var.Columns(2).Value
        End If
    Next var
    
    DPVars = Split(Join(Application.Transpose(DataProcRng.Value), ","), ",")
    For Each var In DataProcRng.Rows
        DPVarsList.Add var.Columns(1).Value
        If Not Memento.Exists(var.Columns(1).Value) Then
            Memento.Add var.Columns(1).Value, var.Columns(2).Value
        End If
    Next var
    
    FolderPickers = Split(Join(vFolderPickers, ","), ",")
    For Each var In FolderPickers
        FolderPickersList.Add var
    Next var
    
    'Кэшируем описания полей в словаре, из которого их потом удобно и просто доставать
    Set DescRng = ThisWorkbook.Sheets("Описание полей").Range("param_list")
    
    For Each Row In DescRng.Rows
        If Not DescDict.Exists(Row.Columns(1).Value) Then
            DescDict.Add Row.Columns(1).Value, Row.Columns(2).Value
        End If
    Next Row
 
End Sub



