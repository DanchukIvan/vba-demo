Attribute VB_Name = "FormEntrypoint"
'@Folder "S3FileDownloader.FormEntrypoint"
'@ModuleDescription("������ �������� ��������� �������, ������� ������������ ������ ��� ������ � _
                    �������.")

Option Explicit

'������������ ������, ��� ��� � ���� ������� ������� �������� ��� � ���������
Public MacroVarsList As New mscorlib.ArrayList
Public DPVarsList As New mscorlib.ArrayList
Public FolderPickersList As New mscorlib.ArrayList

Public MacroRng As Range
Public DataProcRng As Range

Public DescDict As New Dictionary
Public Memento As New Dictionary

'@VariableDescription("������ ���������� - ���������. ������������ � ������ ������ ������ �����. _
                    ������������ ��� ����������� ���������� ��������.")
Public ConfWorkingWsh As Worksheet

'����� ����� ��� ����������� ���� ��������� �������
Public Sub StartVacancyParser()
    Load MainWindows
    MainWindows.Show
End Sub

'��������� �������������� ��������� ����� - �������� ��� ����������, ������� ����� �������� _
� ����� ��������� � �������� �� � ������� �� ������, ���� ����������� ������������ ��.

Public Sub SetInitVars()

    Dim vFolderPickers As Variant
    Dim MacroVars() As String
    Dim DPVars() As String
    Dim FolderPickers() As String
    Dim var As Variant
    Dim DescRng As Range
    Dim Row As Variant
    
    '������� ��� ���� � ��� ���-�� ���� �� ������ ������������� (������ �� ���-�� ����� �������).
    If Memento.Count <> 0 Then
        Memento.RemoveAll
    End If
    
    '���������� ���� � �������� �����
    Set ConfWorkingWsh = ThisWorkbook.Sheets("������������")
    ConfWorkingWsh.Activate
    
    '����� � ����������� �������, ������� �������� ������������� (�� ����� ���� �������)
    Set MacroRng = ConfWorkingWsh.Range("required_params").Offset(0, -1)
    '����� � ����������� ��������� (����� ���� �������)
    Set DataProcRng = ConfWorkingWsh.Range("postproc_settings").Columns(1)
    '�������� ���������, ������� ������� ������ �����, � ������. �� ����� ��� ������������� ��������, _
    ������� ����� ���������� � ������� ����� ����.
    vFolderPickers = Application.Transpose(ConfWorkingWsh.Range("folder_picker").Offset(0, -1))
    
    '������� ��� ��������� � �������� ��������� �������� � �������
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
    
    '�������� �������� ����� � �������, �� �������� �� ����� ������ � ������ ���������
    Set DescRng = ThisWorkbook.Sheets("�������� �����").Range("param_list")
    
    For Each Row In DescRng.Rows
        If Not DescDict.Exists(Row.Columns(1).Value) Then
            DescDict.Add Row.Columns(1).Value, Row.Columns(2).Value
        End If
    Next Row
 
End Sub



