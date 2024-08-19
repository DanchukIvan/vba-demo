Attribute VB_Name = "GetS3Item"
'@Folder "S3FileDownloader.S3Handling"
'@ModuleDescription("Модуль содержит функции для загрузки и сохранения файлов из S3 бакета")

Option Explicit
Option Private Module

'Процедура обходит url адреса файлов и сохраняет их в директории, путь к которой указан в конфиге

Public Sub SendGetRequests(ByRef Config As ConfigClass, ByRef aUrls As Variant)

    Dim Request As WinHttpRequest, Stream As ADODB.Stream, SaveFolder As String
    Dim Url As Variant, NewPath As String
    
    'Проверяем что директория, на которую указывает путь, существует и если нет, то создаем новую
    SaveFolder = Config.DownloadPath
    If Dir(SaveFolder, vbDirectory) = vbNullString Then
        MkDir SaveFolder
    End If
    
    'Создаем http клиента и объект потока, который будет записывать данные в память, а потом их сбрасывать на диск
    Set Request = New WinHttpRequest
    Set Stream = New ADODB.Stream
    
    For Each Url In aUrls
        Application.StatusBar = "Отправляем запрос в S3"
        If Not CheckUrl(Config, Url) Then
            GoTo SkipBadUrl
        End If
        Request.Open "GET", Url, False
        Request.Send
        'Любой иной статус запроса, кроме 200, отбрасываем и прерываем выполнение
        If Request.Status = 200 Then
            NewPath = Utils.MakePathConsist(SaveFolder) + FileFromItemUrl(Url)
            'Если таких файлов нет в директории назначения, то сохраняем файл. Иначе - пропускаем.
            'Не обрабатывается случай, когда название файла осталось прежним, но изменились метаданные и контент.
            If Len(Dir(NewPath)) = 0 Then
                Application.StatusBar = "Записываем файл в папку"
                Stream.Open
                Stream.Type = adTypeBinary
                Stream.Write Request.ResponseBody
                Stream.SaveToFile NewPath, adSaveCreateNotExist
                Stream.Close
            End If
        Else
            MsgBox "Что-то пошло не так - запрос прошел неуспешно." & vbNewLine & "Тело ответа:" & vbNewLine & _
            Request.ResponseText, vbCritical, "Статус запроса - " & Request.Status
            Exit Sub
        End If

'Если url невалидный (не соответствует заданной пользователем маске), то мы просто переходим к другому
SkipBadUrl:
    Next Url
    
    'Завершаем вызов процедуры подготовкой датасета
    Call CreateDataset.CreateQueryTable(Config)
    
End Sub

'Функция проводит проверку url на соответствие заданной пользователем маске
Private Function CheckUrl(ByRef Config As ConfigClass, ByRef Url As Variant) As Boolean
    Dim CheckRegex As New RegExp, Matches As Object
    
    CheckUrl = False
    
    CheckRegex.Pattern = Config.UrlFilterPattern
    CheckRegex.MultiLine = True
    CheckRegex.Global = True
    
    Set Matches = CheckRegex.Execute(Url)
    
    'Функция возвращает True только если url не пустой и мы нашли все заданные маски в url
    If Matches.Count - 1 = Config.PatternKWCount And Url <> vbNullString Then
        CheckUrl = True
    End If

End Function

'Функция возвращает имя файла из url. Специфична для формата ссылки на объект S3.
Private Function FileFromItemUrl(ByRef Url As Variant) As String
    Dim UrlPart As String
    
    UrlPart = Split(Url, "?")(0)
    
    FileFromItemUrl = Right$(UrlPart, Len(UrlPart) - InStrRev(UrlPart, "/"))
End Function

