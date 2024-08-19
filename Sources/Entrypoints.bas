Attribute VB_Name = "Entrypoints"
'@Folder "S3FileDownloader"
'@ModuleDescription("������� ����� ����� ��� ������� �������� � ����� ��������")

Option Explicit
Option Private Module

'��������� ��������� ������� GetUrlFromPython, ������� ��������� ��� ��������� �������.

Public Sub RunScript(ByRef Config As ConfigClass)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo ResetSettings
    Call GenUrls.GetUrlFromPython(Config)
    
    If Config.CreateDownloadsList Then
        Call Utils.CreateDwnldsSheet(Config)
    End If
    
ResetSettings:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

