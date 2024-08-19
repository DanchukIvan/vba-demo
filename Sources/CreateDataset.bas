Attribute VB_Name = "CreateDataset"
'@Folder "S3FileDownloader.DataProcessing"
'@ModuleDescription("������ ������������ ��� ���������� �������� �� ��������� ������ � ������� _
                        ������� ������ ������� Power Query.")

Option Explicit
Option Private Module

'��������� ������� �������, ������� ��� ����� ��������

Public Sub CreateQueryTable(ByVal Config As ConfigClass)
    Dim i As Variant
    Dim ResultBookPath As String, FileODC As String
    Dim ResultWb As Workbook, InsertPath As New RegExp
    
    '��� ��� � ������� ���� ����������� �������� odc ������ � �� �������� ����������� ����� _
    ��������� ��������� ������ �������, �� � ������ ����� ��� ���������� ������ ���� ����� �����������, _
    � ������� ������ ��������� ������� Power Query
    
    FileODC = Config.ConnectionPath
    ResultBookPath = ActiveWorkbook.Path + "\" + Config.ResultFilename
    If Dir(FileODC) = vbNullString Then
        MsgBox "����������� ���� ����������� � ������� " + FileODC + "." + "�������� �������� ���� �� �����������. ������� �� ���������.", _
               vbCritical, "����������� ������"
        Exit Sub
    End If
    
    '��������� ��� ������� ���� ��� ���������� ��������
    Application.StatusBar = "��������� ���� ��� ���������� �����������"
    If Dir(ResultBookPath) = vbNullString Then
        Set ResultWb = Workbooks.Add
        ResultWb.SaveAs ResultBookPath
    Else
        Set ResultWb = Workbooks.Open(ResultBookPath)
    End If
    
    ResultWb.Activate
    
    '��������� ���� �� �����-�� ����������� � �����-�������� ����������. _
    ���� ���, �� ������������ � ���� �����������. ���� ����, �� ��������, ��� ��� ������ ����������� _
    � ���������� ������.
    
    If ResultWb.Connections.Count = 0 Then
        Application.StatusBar = "������������ ���� ������� � ����� �����������"
        With ResultWb
            .Connections.AddFromFile FileODC, False, False
        End With
    End If
    
    Dim FormulaM As String
    
    '��� ��� �������������� ������� ����� ��������, ��� ����� ����������� ������ �����, �� ������� _
    ������� Power Query �������� �������. ��� ����� �� ������ � ������� ���������� ������ ������, _
    �������������� ������ � ����� �� �������.
    
    FormulaM = ResultWb.Queries(1).Formula
    InsertPath.Global = True
    InsertPath.MultiLine = True
'    InsertPath.Pattern = "(?:\" & Chr$(34) & "\w:.*\" & Chr$(34) & ")"
    InsertPath.Pattern = "(" & Chr$(34) & "\w:.*(?=[\)" & Chr$(34) & "])" & ")"
'    FormulaM = InsertPath.Replace(FormulaM, Chr$(34) & Left(Config.DownloadPath, Len(Config.DownloadPath) - 1) & Chr(34))
    FormulaM = InsertPath.Replace(FormulaM, Chr$(34) & Config.DownloadPath & Chr(34))
    ResultWb.Queries(1).Formula = FormulaM
    
    '� ����� ����������� ������� ����� �������, � ������� ������� ����� �������� � ����� ������������� _
    ����������� � ������. ���� ����� ������� ��� ����, � ������ � ������ � ������, �� ������ ��������� ���.
    
    If ActiveSheet.ListObjects.Count <> 0 Then
        Application.StatusBar = "�������� ����� ������"
        
        ResultWb.Sheets("����1").ListObjects("result").QueryTable.Refresh
    Else
        Application.StatusBar = "������� ������� �����������"
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=ResultWb.Connections("������ � result_query"), Destination:=Range("$A$1")).QueryTable
            .Refresh
            .ListObject.Name = "result"
            .SourceConnectionFile = FileODC
        End With
    End If

    ResultWb.Save
    
    '�������� ������� ��������� ����� � ��� ���� �� ������ � �������� ������ ��� ������������.
    Call DataTransformer.Pipeline(Config)
    
    '��������� �������� ����� ������� �� ������� ����������. ����� � ����� �� �������� ������ ������� _
    ������� ��������� ����� �������. �������� ���������� ������ ���, ��� ���� � Power Query.
    
    With ResultWb.Sheets("����1").ListObjects("result")
        Application.StatusBar = "��������� ������ �� ������� ����������"
        
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


