Attribute VB_Name = "CreateDataset"
'@Folder "S3FileDownloader.DataProcessing"
'@ModuleDescription("Модуль используется для подготовки датасета из скаченных файлов с помощью _
                        готовых файлов запроса Power Query.")

Option Explicit
Option Private Module

'Процедура создает датасет, прогнав его через пайплайн

Public Sub CreateQueryTable(ByVal Config As ConfigClass)
    Dim i As Variant
    Dim ResultBookPath As String, FileODC As String
    Dim ResultWb As Workbook, InsertPath As New RegExp
    
    'Так как у скрипта есть стандартная поставка odc файлов и их названия фиксированы через _
    приватные константы класса конфига, то и конфиг сразу нам возвращает полный путь файла подключения, _
    в котором зашиты связанные скрипты Power Query
    
    FileODC = Config.ConnectionPath
    ResultBookPath = ActiveWorkbook.Path + "\" + Config.ResultFilename
    If Dir(FileODC) = vbNullString Then
        MsgBox "Отсутствует файл подключения к запросу " + FileODC + "." + "Повторно скачайте файл из репозитория. Выходим из программы.", _
               vbCritical, "Критическая ошибка"
        Exit Sub
    End If
    
    'Открываем или создаем файл для сохранения датасета
    Application.StatusBar = "Открываем файл для сохранения результатов"
    If Dir(ResultBookPath) = vbNullString Then
        Set ResultWb = Workbooks.Add
        ResultWb.SaveAs ResultBookPath
    Else
        Set ResultWb = Workbooks.Open(ResultBookPath)
    End If
    
    ResultWb.Activate
    
    'Проверяем есть ли какие-то подключения в файле-носителе результата. _
    Если нет, то присоединяем к нему подключение. Если есть, то полагаем, что это нужное подключение _
    и продолжаем работу.
    
    If ResultWb.Connections.Count = 0 Then
        Application.StatusBar = "Присоединяем файл запроса к файлу результатов"
        With ResultWb
            .Connections.AddFromFile FileODC, False, False
        End With
    End If
    
    Dim FormulaM As String
    
    'Так как местоположение скрипта может меняться, нам нужно динамически менять папку, из которой _
    скрипты Power Query забирают контент. Для этого мы меняем в формуле содержимое первой строки, _
    предварительно достав её текст из запроса.
    
    FormulaM = ResultWb.Queries(1).Formula
    InsertPath.Global = True
    InsertPath.MultiLine = True
'    InsertPath.Pattern = "(?:\" & Chr$(34) & "\w:.*\" & Chr$(34) & ")"
    InsertPath.Pattern = "(" & Chr$(34) & "\w:.*(?=[\)" & Chr$(34) & "])" & ")"
'    FormulaM = InsertPath.Replace(FormulaM, Chr$(34) & Left(Config.DownloadPath, Len(Config.DownloadPath) - 1) & Chr(34))
    FormulaM = InsertPath.Replace(FormulaM, Chr$(34) & Config.DownloadPath & Chr(34))
    ResultWb.Queries(1).Formula = FormulaM
    
    'В файле результатов создаем умную таблицу, с которой гораздо легче работать в плане использования _
    подключений к данным. Если умная таблица уже есть, а значит и запрос к данным, то просто обновляем его.
    
    If ActiveSheet.ListObjects.Count <> 0 Then
        Application.StatusBar = "Получаем новые данные"
        
        ResultWb.Sheets("Лист1").ListObjects("result").QueryTable.Refresh
    Else
        Application.StatusBar = "Создаем таблицу результатов"
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=ResultWb.Connections("Запрос — result_query"), Destination:=Range("$A$1")).QueryTable
            .Refresh
            .ListObject.Name = "result"
            .SourceConnectionFile = FileODC
        End With
    End If

    ResultWb.Save
    
    'Вызываем скрипты обработки тэгов и дат если их формат в исходных файлах был некорректный.
    Call DataTransformer.Pipeline(Config)
    
    'Проверяем диапазон умной таблицы на наличие дубликатов. Здесь в цикле мы собираем массив номеров _
    колонок диапазона умной таблицы. Проверка проводится помимо той, что есть в Power Query.
    
    With ResultWb.Sheets("Лист1").ListObjects("result")
        Application.StatusBar = "Проверяем данные на наличие дубликатов"
        
        Dim ColCnt As Long, arr() As Variant
        
        ColCnt = .ListColumns.Count
        ReDim arr(0 To ColCnt - 1)
        
        For i = 1 To ColCnt
            arr(i - 1) = i
        Next i
        
        Dim r As Range
        Set r = .Range
        
        r.RemoveDuplicates Columns:=(arr), Header:=xlYes
        
        ResultWb.Save
    End With

    ResultWb.Close

End Sub


