VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "S3FileDownloader.CommonUtils"

'@ModuleDescription("Модуль класса для аккумуляции ключей конфигурации. Используется всеми модулями за исключением модуля утилит.")

Option Explicit

'@VariableDescription("Переменная-плейсхолдер для строковых данных. Используется по мере необходимости.")
Private ResultString As String

'@VariableDescription("Количество ключевых слов, введенных пользователем для фильтрации ссылок")
Private KeyWordCount As Long

'@VariableDescription("Имя файла подключения, входящего в стандартную поставку. Не должно переопределяться")
Private Const PrepareDatasetConn As String = "prepare_dataset.odc"


'@Description("Ключевые слова для фильтрации ссылок на объекты в S3")
Private Property Get ResString() As String
    ResString = ResultString
End Property

Private Property Let ResString(ByVal RHS As String)
    ResultString = RHS
End Property

'@Description("Необходимое количество совпадений с паттерном регулярного выражения при фильтрации ссылок")
Public Property Get PatternKWCount() As Long
    PatternKWCount = KeyWordCount
End Property

Public Property Let PatternKWCount(ByVal RHS As Long)
    KeyWordCount = RHS
End Property

'@Description("Путь для поиска скрипта Python")
Public Property Get PyUrlGeneratorPath() As String
    PyUrlGeneratorPath = Лист2.Range("C7").Value
End Property

'@Description("Путь для поиска файла подключения к скрипту Power Query")
Public Property Get ConnectionPath() As String

    ConnectionPath = Лист2.Range("C6").Value
End Property

'@Description("Путь для сохранения скаченных файлов")
Public Property Get DownloadPath() As String

    DownloadPath = Лист2.Range("C8").Value
End Property

'@Description("Название файла для сохранения обработанного результата")
Public Property Get ResultFilename() As String

    ResultFilename = Лист2.Range("C9").Value
End Property


'@Description("Вспомогательная функция, преобразующая строку в список по заданному разделителю")
Private Function StringToArray(ByVal sFieldString As String, Optional ByVal sDelimPattern As String = "[, ]+") As Variant
    Dim aFields() As String, oDelimFinder As New RegExp
    Dim Choose As Boolean, i As Variant
    
    oDelimFinder.Pattern = sDelimPattern
    oDelimFinder.Global = True
     
    If oDelimFinder.Test(sFieldString) Then
        sFieldString = oDelimFinder.Replace(sFieldString, ",")
        aFields = Split(sFieldString, ",")
    Else
        If sFieldString <> vbNullString Then
            Dim arr(0) As String
            arr(0) = sFieldString
            aFields = arr
        Else
            Err.Raise 31004, "ConfigClass.StringToArray", _
                      "Указана неверная или пустая строка, нужно отредактировать конфигурацию"
        End If
    End If
    
    StringToArray = aFields

End Function

Public Property Get CreateDownloadsList() As Boolean
    
    CreateDownloadsList = False
    
    If Лист2.Range("C13").Value = "Да" Then
        CreateDownloadsList = True
    End If
          
End Property



'@Description("Список названий колонок таблицы, которые нужно очистить от тэгов")
'@Enumerator
Public Property Get FieldsWithTags() As Variant
    Dim FalsyArr(0) As Variant: FalsyArr(0) = False
    
    Dim sFields As String: sFields = Лист2.Range("C14").Value
    
    If sFields = vbNullString Then
        FieldsWithTags = FalsyArr
    Else
        FieldsWithTags = StringToArray(sFields)
    End If
End Property

'@Description("Список названий колонок таблицы, даты которых нужно привести в корректный вид")
'@Enumerator
Public Property Get FieldsWithDates() As Variant
    Dim FalsyArr(0) As Variant: FalsyArr(0) = False
    
    Dim sFields As String: sFields = Лист2.Range("C15").Value
    
    If sFields = vbNullString Then
        FieldsWithDates = FalsyArr
    Else
        FieldsWithDates = StringToArray(sFields)
    End If
End Property


'@Description("Паттерн регулярного выражения для фильтрации ссылок на объекты в S3.")
Public Property Get UrlFilterPattern() As String
    Dim idx As Variant
    
    ResString = Лист2.Range("C16").Value
    Dim aMasks() As String
    
    If ResString = vbNullString Then
        GoTo NullStringHandler
    End If
    
    aMasks = StringToArray(ResString)
    For idx = LBound(aMasks) To UBound(aMasks)
        aMasks(idx) = "\b" & aMasks(idx) & "\b"
    Next idx
    
    'Устанавливается значение свойства класса
    PatternKWCount = UBound(aMasks)
    Debug.Print PatternKWCount
    
    Dim sPattern As String
    sPattern = vbNullString
    
    
    sPattern = Join(aMasks, "|")
    Debug.Print sPattern
    
    UrlFilterPattern = sPattern
    
NullStringHandler:
    
    If ResString = vbNullString Then
        'Поиск по regex находит минимум 2 совпадения - слово и, почему-то, пустую строку
        PatternKWCount = 1
        'Если не задан паттерн, то мы любую ссылку воспринимаем как валидную
        UrlFilterPattern = ".*"
    End If

End Property


