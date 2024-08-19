Attribute VB_Name = "GetS3Item"
'@Folder "S3FileDownloader.S3Handling"
'@ModuleDescription("������ �������� ������� ��� �������� � ���������� ������ �� S3 ������")

Option Explicit
Option Private Module

'��������� ������� url ������ ������ � ��������� �� � ����������, ���� � ������� ������ � �������

Public Sub SendGetRequests(ByRef Config As ConfigClass, ByRef aUrls As Variant)

    Dim Request As WinHttpRequest, Stream As ADODB.Stream, SaveFolder As String
    Dim Url As Variant, NewPath As String
    
    '��������� ��� ����������, �� ������� ��������� ����, ���������� � ���� ���, �� ������� �����
    SaveFolder = Config.DownloadPath
    If Dir(SaveFolder, vbDirectory) = vbNullString Then
        MkDir SaveFolder
    End If
    
    '������� http ������� � ������ ������, ������� ����� ���������� ������ � ������, � ����� �� ���������� �� ����
    Set Request = New WinHttpRequest
    Set Stream = New ADODB.Stream
    
    For Each Url In aUrls
        Application.StatusBar = "���������� ������ � S3"
        If Not CheckUrl(Config, Url) Then
            GoTo SkipBadUrl
        End If
        Request.Open "GET", Url, False
        Request.Send
        '����� ���� ������ �������, ����� 200, ����������� � ��������� ����������
        If Request.Status = 200 Then
            NewPath = Utils.MakePathConsist(SaveFolder) + FileFromItemUrl(Url)
            '���� ����� ������ ��� � ���������� ����������, �� ��������� ����. ����� - ����������.
            '�� �������������� ������, ����� �������� ����� �������� �������, �� ���������� ���������� � �������.
            If Len(Dir(NewPath)) = 0 Then
                Application.StatusBar = "���������� ���� � �����"
                Stream.Open
                Stream.Type = adTypeBinary
                Stream.Write Request.ResponseBody
                Stream.SaveToFile NewPath, adSaveCreateNotExist
                Stream.Close
            End If
        Else
            MsgBox "���-�� ����� �� ��� - ������ ������ ���������." & vbNewLine & "���� ������:" & vbNewLine & _
            Request.ResponseText, vbCritical, "������ ������� - " & Request.Status
            Exit Sub
        End If

'���� url ���������� (�� ������������� �������� ������������� �����), �� �� ������ ��������� � �������
SkipBadUrl:
    Next Url
    
    '��������� ����� ��������� ����������� ��������
    Call CreateDataset.CreateQueryTable(Config)
    
End Sub

'������� �������� �������� url �� ������������ �������� ������������� �����
Private Function CheckUrl(ByRef Config As ConfigClass, ByRef Url As Variant) As Boolean
    Dim CheckRegex As New RegExp, Matches As Object
    
    CheckUrl = False
    
    CheckRegex.Pattern = Config.UrlFilterPattern
    CheckRegex.MultiLine = True
    CheckRegex.Global = True
    
    Set Matches = CheckRegex.Execute(Url)
    
    '������� ���������� True ������ ���� url �� ������ � �� ����� ��� �������� ����� � url
    If Matches.Count - 1 = Config.PatternKWCount And Url <> vbNullString Then
        CheckUrl = True
    End If

End Function

'������� ���������� ��� ����� �� url. ���������� ��� ������� ������ �� ������ S3.
Private Function FileFromItemUrl(ByRef Url As Variant) As String
    Dim UrlPart As String
    
    UrlPart = Split(Url, "?")(0)
    
    FileFromItemUrl = Right$(UrlPart, Len(UrlPart) - InStrRev(UrlPart, "/"))
End Function

