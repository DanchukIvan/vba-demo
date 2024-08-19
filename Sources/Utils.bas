Attribute VB_Name = "Utils"

'@Folder "S3FileDownloader.CommonUtils"
'@ModuleDescription("������ ������ ������. ����� ����������� ��� �������� ��������������� � ����� Excel")

Option Explicit

'������� ��������� ����� ������ ����������� � ���� ����������
Public Function MakePathConsist(ByRef OldPath As String) As String
    
    If Right$(OldPath, 1) = "\" Then
        MakePathConsist = OldPath
    Else
        MakePathConsist = OldPath + "\"
    End If
End Function

'������� ���������, ��� ��� �������� ������ � ������ ������ �� ����
Public Function IsPath(ByRef Value As Variant) As Boolean
    Dim PathRegex As New RegExp
    PathRegex.Pattern = "\w:[\\\w]+"
    PathRegex.Global = True
    
    IsPath = WorksheetFunction.IsText(Value) And PathRegex.Test(Value)

End Function

'������ ����������� ��������� � �������
Public Function CollectionToArray(ByRef myCol As Variant) As Variant
    Dim result  As Variant
    Dim cnt     As Long
    Dim Val As Variant

    ReDim result(myCol.Count - 1)
    
    cnt = 0
    For Each Val In myCol
        result(cnt) = Val
        cnt = cnt + 1
    Next Val

    CollectionToArray = result

End Function

'��������� ������� ���� �� ������� ��������� ������
Public Sub CreateDwnldsSheet(ByRef Config As ConfigClass)
    Dim Sh As Worksheet, Rng As Range, FSO As Scripting.FileSystemObject, Folder As Scripting.Folder
    Dim Files As Scripting.Files, File As Scripting.File, FilesArr() As Variant, cnt As Long
    
    Application.StatusBar = "��������� ������ ����������� ������"
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(Config.DownloadPath)
    Set Files = Folder.Files
    
    ReDim FilesArr(1 To Files.Count, 1 To 2)
    
    cnt = 1
    For Each File In Files
        FilesArr(cnt, 1) = File.Name
        FilesArr(cnt, 2) = Date
        cnt = cnt + 1
    Next File
    
    Set FSO = Nothing
    Set Folder = Nothing
    Set Files = Nothing
    
    With ThisWorkbook
        .Activate
        Set Sh = .Sheets.Add
        Sh.Name = "������ ��������� ������"
        With Sh
            '� ������ ����� ���� ��������� ���������� Rng ������ ����������������, ��� ����� ��������.
            Set Rng = .Range("A1")
            Rng = "������������ �����"
            Rng.Offset(0, 1) = "���� ����������"
            Set Rng = Rng.Resize(UBound(FilesArr, 1) + 1, UBound(FilesArr, 2))
            Rng.Offset(1, 0) = FilesArr
            Rng.Font.FontStyle = "Arial Narrow"
            Rng.Font.Size = 11
            .ListObjects.Add xlSrcRange, Source:=Rng, XlListObjectHasHeaders:=xlYes
            .ListObjects(1).DisplayName = "ExctractedFiles"
            Set Rng = .ListObjects(1).Range
            Rng.Font.FontStyle = "Arial Narrow"
            Rng.VerticalAlignment = xlCenter
            Rng.HorizontalAlignment = xlCenter
        End With
        '���� ����� �� ������� �������� ���, �� �� ������� ��� �� ����� ���������� �������� �����. _
        ���� ����, �� �� ���������� ��� � �������� ����� ������ �� ����� ������� ������ (������ ���������).
        If Dir("��������� �����.xlsx") <> "" Then
            Dim Wb As Workbook, Tbl As ListObject, Data As Range
            
            Set Wb = Workbooks.Open("��������� �����.xlsx")
            Sh.ListObjects("ExctractedFiles").DataBodyRange.Copy
            Wb.Activate
            Set Tbl = Wb.Sheets("������ ��������� ������").ListObjects("ExctractedFiles")
            Tbl.HeaderRowRange.End(xlDown).Select
            ActiveSheet.Paste
        Else
            Sh.Copy
            ActiveWorkbook.SaveAs "��������� �����.xlsx", xlWorkbookDefault
        End If
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        Sh.Delete
    End With
    
    ThisWorkbook.Save
            
End Sub

'������� ������������� ������ �������� ������� ����� � ������� �����, � ������� �������� ������ _
�� ������ ������ �������
Public Function WShExists(ByRef WshName As String) As Boolean
    Dim Sh As Worksheet
    
    WShExists = False
    
    On Error Resume Next
    Set Sh = ActiveWorkbook.Sheets(WshName)
    If Not Sh Is Nothing Then
        WShExists = True
    End If
    
End Function

'������� �������� ������ ������ ����������. ���� ������������ ������ �� ������ - �� ������ ������. _
���������� ���� ����������.
Public Function FolderPicker(Optional ByVal OldValue As String = "") As String
    Dim Folder As String, Resp As Integer
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .ButtonName = "������� �����"
        Resp = .Show
        If Resp = -1 Then
            Folder = .SelectedItems(1)
            FolderPicker = Utils.MakePathConsist(Folder)
        Else
            If Len(OldValue) <> 0 Then
                FolderPicker = OldValue
            Else
                FolderPicker = ThisWorkbook.Path
            End If
        End If
    End With
    
End Function

'������� �������� ������ ������ �����. ���� ������������ ������ �� ������ - �� ������ ������. _
���������� ������ ��� �����, � �� ������ ����.
Public Function FilePicker(Optional ByVal OldValue As String = "") As String
    Dim File As String, Resp As Integer
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .ButtonName = "������� ����"
        Resp = .Show
        If Resp = -1 Then
            File = .SelectedItems(1)
            FilePicker = Right$(File, Len(File) - InStrRev(File, "\"))
        Else
            FilePicker = OldValue
        End If
    End With
        
End Function
