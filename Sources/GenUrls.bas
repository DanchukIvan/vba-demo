Attribute VB_Name = "GenUrls"
'@Folder "S3FileDownloader.S3Handling"
'@ModuleDescription("������ ������������ ��� ������ �� �������� Python, ������������ ������. _
                     ������ ��� ������� ���������� boto3. ������ ������������� � exe �����.")

Option Explicit
Option Private Module

'��������� ������������ ����� ������� ��������� url �� S3 � ���������, ��� ���� ��� ������������

Public Sub GetUrlFromPython(ByRef Config As ConfigClass)
    Dim vUrls As Variant
    
    vUrls = PythonCaller(Config.PyUrlGeneratorPath)
    If UBound(vUrls) = 0 Then
        MsgBox "��� ������ ��� ���������. ��������� ������������ ������� Python � ���������� ���������.", vbOKOnly, _
        "������� ������ ������"
        Exit Sub
    End If

    Call GetS3Item.SendGetRequests(Config, vUrls)
    
End Sub

'������� �������� ������ Python ����� exe � ���������� ������ url � S3, ������� ����� ������
'http ��������

Private Function PythonCaller(ByVal PyUrlGeneratorPath As String) As String():
    Dim oSh As Object, oExec As Object
    Dim var As Variant
    Dim aOutput() As String
    
    Application.StatusBar = "�������� ������ ��������� ������ �� Python"
    
    '�������� ���� (��������� ��������)
    Set oSh = New WshShell
    
    '������ ���, ����� ������� ������� ����������� ����� ������������ �������
    oSh.CurrentDirectory = PyUrlGeneratorPath
    
    '��������� �������� � ���� ���� �� ����������
    Set oExec = oSh.Exec(".\get_url.exe")
    Application.Wait (Now + TimeValue("0:00:03"))
    
    '���� ��� �������� ���������, �� ���������� ����� ������ �� ������������ ������
    '� ������ � ���������� ������������. ������ ������ ������������ ���� �� ����� �������.
    If oExec.ExitCode <> 0 Then
        MsgBox "����� ������:" & vbCrLf & oExec.StdOut, vbCritical, "� �������� ���������� ������� Python �������� ������."
        aOutput = Array()
    Else
        Application.StatusBar = "���������� ������"
        aOutput = Split(oExec.StdOut.ReadAll, vbCrLf)
    End If
    
    
    
    Set oExec = Nothing: Set oSh = Nothing
    PythonCaller = aOutput
    
End Function

