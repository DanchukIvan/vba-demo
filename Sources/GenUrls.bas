Attribute VB_Name = "GenUrls"
'@Folder "S3FileDownloader.S3Handling"
'@ModuleDescription("Модуль используется для работы со скриптом Python, генерирующим ссылки. _
                     Скрипт под капотом использует boto3. Скрипт скомпилирован в exe архив.")

Option Explicit
Option Private Module

'Процедура осуществляет вызов скрипта получения url из S3 и проверяет, что есть что обрабатывать

Public Sub GetUrlFromPython(ByRef Config As ConfigClass)
    Dim vUrls As Variant
    
    vUrls = PythonCaller(Config.PyUrlGeneratorPath)
    If UBound(vUrls) = 0 Then
        MsgBox "Нет ссылок для обработки. Проверьте корректность скрипта Python и содержимое хранилища.", vbOKOnly, _
        "Получен пустой массив"
        Exit Sub
    End If

    Call GetS3Item.SendGetRequests(Config, vUrls)
    
End Sub

'Функция вызывает скрипт Python через exe и возвращает список url в S3, которые нужно обойти
'http клиентом

Private Function PythonCaller(ByVal PyUrlGeneratorPath As String) As String():
    Dim oSh As Object, oExec As Object
    Dim var As Variant
    Dim aOutput() As String
    
    Application.StatusBar = "Вызываем скрипт генерации ссылок на Python"
    
    'Вызываем шелл (комнадный терминал)
    Set oSh = New WshShell
    
    'Делаем так, чтобы текущей рабочей директорией стало расположения скрипта
    oSh.CurrentDirectory = PyUrlGeneratorPath
    
    'Выполняем экзешник и ждем пока он завершится
    Set oExec = oSh.Exec(".\get_url.exe")
    Application.Wait (Now + TimeValue("0:00:03"))
    
    'Если код возврата ненулевой, то отображаем текст ошибки из стандартного вывода
    'в пушапе и уведомляем пользователя. Данная ошибка обработается выше по стеку вызовов.
    If oExec.ExitCode <> 0 Then
        MsgBox "Текст ошибки:" & vbCrLf & oExec.StdOut, vbCritical, "В процессе выполнения скрипта Python возникла ошибка."
        aOutput = Array()
    Else
        Application.StatusBar = "Генерируем ссылки"
        aOutput = Split(oExec.StdOut.ReadAll, vbCrLf)
    End If
    
    
    
    Set oExec = Nothing: Set oSh = Nothing
    PythonCaller = aOutput
    
End Function

