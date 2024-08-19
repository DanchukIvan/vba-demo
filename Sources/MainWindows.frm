VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainWindows 
   Caption         =   "MainWindows"
   ClientHeight    =   4330
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7770
   OleObjectBlob   =   "MainWindows.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State As String
Public IsChanged As Boolean
Public Config As New ConfigClass

'@VariableDescription("Имя файла подключения, входящего в стандартную поставку. Не должно переопределяться")
Private Const PrepareDatasetConn As String = "prepare_dataset.odc"

'Процедура вызывает форму с информацией о скрипте
Private Sub About_Click()
    Load AboutForm
    AboutForm.Show
End Sub


Private Sub CancelButton_Click()
    Dim Choice As Boolean
    
    If IsChanged Then
        Choice = MsgBox("Сохранить внесенные изменения?", vbYesNoCancel, "Выход из программы")
        If Not Choice Or Choice = vbCancel Then
            Call Restore
        End If
    End If
    
    FormEntrypoint.ConfWorkingWsh.Parent.Save
    Me.Hide
    Unload Me
    
End Sub

Private Sub StartupMacro_Click()
    
    FormEntrypoint.ConfWorkingWsh.Parent.Save
    Call Entrypoints.RunScript(Config)
    If Err = 0 Then
        MsgBox "Загрузка файлов завершена успешно", vbOKOnly
    End If
    Me.Hide
    Unload Me
    
End Sub

Private Sub MacrosVariable_Click()
    
    State = "Macro"
    With MainWindows.SettingsArea
        .textAlign = fmTextAlignLeft
        .columnCount = 1
        .RowSource = FormEntrypoint.MacroRng.Address
    End With
    
End Sub

Private Sub ProcessingOpt_Click()
    
    State = "DProcess"
    With MainWindows.SettingsArea
        .textAlign = fmTextAlignLeft
        .Enabled = True
        .columnCount = 2
        .RowSource = FormEntrypoint.DataProcRng.Address
    End With
    
End Sub
'Процедура восстанавливает значения ВСЕХ полей до значений по умолчанию (в т.ч. до пустых полей)
Public Sub Restore()
    Dim Key As Variant
    
    For Each Key In FormEntrypoint.Memento.Keys
        If FormEntrypoint.MacroVarsList.contains(Key) Then
            FormEntrypoint.MacroRng.Find(Key).Offset(0, 1).Value = FormEntrypoint.Memento(Key)
        End If
        If FormEntrypoint.DPVarsList.contains(Key) Then
            FormEntrypoint.DataProcRng.Find(Key).Offset(0, 1).Value = FormEntrypoint.Memento(Key)
        End If
        If SettingsArea.Value = Key Then
            Selection.Value = FormEntrypoint.Memento(Key)
        End If
    Next Key
    
    IsChanged = False
    
    ConfWorkingWsh.Parent.Save
    
End Sub

Private Sub RestoreButton_Click()
    Call Restore
End Sub

Private Sub SettingsArea_Click()
    Dim Val As Variant
    Dim DefaultValue As Variant
    
    SettingsArea.ControlTipText = "Дважды щелкните левой кнопкой мыши по строке для вызова диалогового окна"
    If State = "Macro" Then
        DefaultValue = FormEntrypoint.MacroRng.Find(SettingsArea.Value).Offset(0, 1).Value
        Me.Selection.Value = DefaultValue
        Me.DescriptWindow.Value = FormEntrypoint.DescDict(SettingsArea.Value)
    End If
    
    If State = "DProcess" Then
        DefaultValue = FormEntrypoint.DataProcRng.Find(SettingsArea.Value).Offset(0, 1).Value
        Me.Selection.Value = DefaultValue
        Me.DescriptWindow.Value = FormEntrypoint.DescDict(SettingsArea.Value)
    End If
    
End Sub

Private Sub SettingsArea_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim Path As String, OldPath As String, Choice As Integer
    Dim TargetRng As Range
    
    If State = "Macro" Then
        If FormEntrypoint.MacroVarsList.contains(SettingsArea.Value) Then
            Set TargetRng = FormEntrypoint.MacroRng.Find(SettingsArea.Value).Offset(0, 1)
            If FormEntrypoint.FolderPickersList.contains(SettingsArea.Value) Then
                Path = Utils.FolderPicker(TargetRng.Value)
            Else
                Path = Utils.FilePicker(TargetRng.Value)
            End If
            If Path <> vbNullString Or TypeName(Path) <> "Boolean" Then
                Me.Selection.Value = Path
                IsChanged = True
                TargetRng.Value = Path
            End If
        End If
        GoTo ExitCode
    End If
    
    If State = "DProcess" Then
        Set TargetRng = FormEntrypoint.DataProcRng.Find(SettingsArea.Value).Offset(0, 1)
        If TargetRng.Address = "$C$13" Then
            Choice = MsgBox("Выберите один вариант.", vbYesNo, "Создать список загруженных файлов?")
            If Choice = 6 Then
                Path = "Да"
            Else
                Path = "Нет"
            End If
        Else
            If FormEntrypoint.DPVarsList.contains(SettingsArea.Value) Then
                Path = Application.InputBox("Введите значение параметра: ", SettingsArea.Value, Type:=2)
            End If
        End If
    End If
        
    Select Case Path
        Case "False"
            GoTo ExitCode
        Case ""
            GoTo ExitCode
        Case Else
            Me.Selection.Value = Path
            IsChanged = True
            TargetRng.Value = Path
    End Select
    
ExitCode:
    Exit Sub
    
End Sub

Private Sub UserForm_Initialize()

    If Лист2.Range("C7").Value = vbNullString Then
        Лист2.Range("C7").Value = ThisWorkbook.Path
    End If
    If Лист2.Range("C6").Value = vbNullString Then
        Лист2.Range("C6").Value = Utils.MakePathConsist(ThisWorkbook.Path) + PrepareDatasetConn
    End If
    If Лист2.Range("C9").Value = vbNullString Then
        Лист2.Range("C9").Value = "ResultDataset.xlsx"
    End If
    If Лист2.Range("C8").Value = vbNullString Then
        Лист2.Range("C8").Value = Utils.MakePathConsist(ThisWorkbook.Path) + "downloads"
    End If
    
    Call FormEntrypoint.SetInitVars
    
    State = "Macro"
    With Me.SettingsArea
        .MultiSelect = fmMultiSelectSingle
        .textAlign = fmTextAlignLeft
        .columnCount = 1
        .RowSource = FormEntrypoint.MacroRng.Address
        .SetFocus
    End With
    
    IsChanged = False
        
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Dim Choice As Boolean
    
    If IsChanged And CloseMode <> 1 Then
        Choice = MsgBox("Сохранить внесенные изменения?", vbYesNoCancel, "Выход из программы")
        If Not Choice Or Choice = vbCancel Then
            Call Restore
        End If
    End If
    
    FormEntrypoint.ConfWorkingWsh.Parent.Save
    
    Me.Hide
    Unload Me
    
End Sub
