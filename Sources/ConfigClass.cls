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

'@ModuleDescription("������ ������ ��� ����������� ������ ������������. ������������ ����� �������� �� ����������� ������ ������.")

Option Explicit

'@VariableDescription("����������-����������� ��� ��������� ������. ������������ �� ���� �������������.")
Private ResultString As String

'@VariableDescription("���������� �������� ����, ��������� ������������� ��� ���������� ������")
Private KeyWordCount As Long

'@VariableDescription("��� ����� �����������, ��������� � ����������� ��������. �� ������ ����������������")
Private Const PrepareDatasetConn As String = "prepare_dataset.odc"


'@Description("�������� ����� ��� ���������� ������ �� ������� � S3")
Private Property Get ResString() As String
    ResString = ResultString
End Property

Private Property Let ResString(ByVal RHS As String)
    ResultString = RHS
End Property

'@Description("����������� ���������� ���������� � ��������� ����������� ��������� ��� ���������� ������")
Public Property Get PatternKWCount() As Long
    PatternKWCount = KeyWordCount
End Property

Public Property Let PatternKWCount(ByVal RHS As Long)
    KeyWordCount = RHS
End Property

'@Description("���� ��� ������ ������� Python")
Public Property Get PyUrlGeneratorPath() As String
    PyUrlGeneratorPath = ����2.Range("C7").Value
End Property

'@Description("���� ��� ������ ����� ����������� � ������� Power Query")
Public Property Get ConnectionPath() As String

    ConnectionPath = ����2.Range("C6").Value
End Property

'@Description("���� ��� ���������� ��������� ������")
Public Property Get DownloadPath() As String

    DownloadPath = ����2.Range("C8").Value
End Property

'@Description("�������� ����� ��� ���������� ������������� ����������")
Public Property Get ResultFilename() As String

    ResultFilename = ����2.Range("C9").Value
End Property


'@Description("��������������� �������, ������������� ������ � ������ �� ��������� �����������")
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
                      "������� �������� ��� ������ ������, ����� ��������������� ������������"
        End If
    End If
    
    StringToArray = aFields

End Function

Public Property Get CreateDownloadsList() As Boolean
    
    CreateDownloadsList = False
    
    If ����2.Range("C13").Value = "��" Then
        CreateDownloadsList = True
    End If
          
End Property



'@Description("������ �������� ������� �������, ������� ����� �������� �� �����")
'@Enumerator
Public Property Get FieldsWithTags() As Variant
    Dim FalsyArr(0) As Variant: FalsyArr(0) = False
    
    Dim sFields As String: sFields = ����2.Range("C14").Value
    
    If sFields = vbNullString Then
        FieldsWithTags = FalsyArr
    Else
        FieldsWithTags = StringToArray(sFields)
    End If
End Property

'@Description("������ �������� ������� �������, ���� ������� ����� �������� � ���������� ���")
'@Enumerator
Public Property Get FieldsWithDates() As Variant
    Dim FalsyArr(0) As Variant: FalsyArr(0) = False
    
    Dim sFields As String: sFields = ����2.Range("C15").Value
    
    If sFields = vbNullString Then
        FieldsWithDates = FalsyArr
    Else
        FieldsWithDates = StringToArray(sFields)
    End If
End Property


'@Description("������� ����������� ��������� ��� ���������� ������ �� ������� � S3.")
Public Property Get UrlFilterPattern() As String
    Dim idx As Variant
    
    ResString = ����2.Range("C16").Value
    Dim aMasks() As String
    
    If ResString = vbNullString Then
        GoTo NullStringHandler
    End If
    
    aMasks = StringToArray(ResString)
    For idx = LBound(aMasks) To UBound(aMasks)
        aMasks(idx) = "\b" & aMasks(idx) & "\b"
    Next idx
    
    '��������������� �������� �������� ������
    PatternKWCount = UBound(aMasks)
    Debug.Print PatternKWCount
    
    Dim sPattern As String
    sPattern = vbNullString
    
    
    sPattern = Join(aMasks, "|")
    Debug.Print sPattern
    
    UrlFilterPattern = sPattern
    
NullStringHandler:
    
    If ResString = vbNullString Then
        '����� �� regex ������� ������� 2 ���������� - ����� �, ������-��, ������ ������
        PatternKWCount = 1
        '���� �� ����� �������, �� �� ����� ������ ������������ ��� ��������
        UrlFilterPattern = ".*"
    End If

End Property


