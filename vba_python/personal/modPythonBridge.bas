Attribute VB_Name = "modPythonBridge"
Option Explicit

Private Const BOOK_KIND As String = "personal"
Private mCompletionPath As String
Private mLogPath As String
Private mOperationName As String
Private mNextCheck As Date
Private mStartedAt As Date
Private mCheckScheduled As Boolean

Public Sub PingPythonBridge()
End Sub

Public Sub PythonExportJson()
    RunPythonAsync "export-auto " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName), "JSON Export"
End Sub

Public Sub PythonImportJson()
    Dim selectedPath As String
    selectedPath = SelectJsonFile()
    If Len(selectedPath) = 0 Then
        Exit Sub
    End If
    If MsgBox("Import前検証とバックアップ後にMerge取込します。続行しますか？", vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If
    RunPythonAsync "import-json " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName) & " " & QuoteArg(selectedPath) & _
        " --backup-dir " & QuoteArg(ThisWorkbook.Path & Application.PathSeparator & "backup"), "JSON Import"
End Sub

Public Sub PythonBackup()
    RunPythonAsync "backup-auto " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName), "バックアップ"
End Sub

Public Sub PythonRefreshAll()
    RunPythonAsync "refresh-all " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName), "全体更新"
End Sub

Public Sub PythonAssignIDs()
    RunPythonAsync "assign-ids " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName) & " --backup-dir " & _
        QuoteArg(ThisWorkbook.Path & Application.PathSeparator & "backup"), "ID採番"
End Sub

Public Sub PythonValidate()
    RunPythonAsync "validate " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName), "整合性チェック"
End Sub

Public Sub PythonSyncProgress()
    RunPythonAsync "sync-progress " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName) & " --backup-dir " & _
        QuoteArg(ThisWorkbook.Path & Application.PathSeparator & "backup"), "進捗ログ同期"
End Sub

Public Sub PythonRefreshViews()
    RunPythonAsync "refresh-views " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName) & " --backup-dir " & _
        QuoteArg(ThisWorkbook.Path & Application.PathSeparator & "backup"), "表示更新"
End Sub

Public Sub PythonRefreshGantt()
    RunPythonAsync "refresh-gantt " & BOOK_KIND & " " & QuoteArg(ThisWorkbook.FullName) & " --backup-dir " & _
        QuoteArg(ThisWorkbook.Path & Application.PathSeparator & "backup"), "ガント更新"
End Sub

Public Sub PythonProcessInbox()
    RunPythonAsync "process-inbox " & QuoteArg(ThisWorkbook.FullName) & " --backup-dir " & _
        QuoteArg(ThisWorkbook.Path & Application.PathSeparator & "backup"), "インボックス処理"
End Sub

Public Sub PythonSyncCompleted()
    RunPythonAsync "sync-completed " & QuoteArg(ThisWorkbook.FullName) & " --backup-dir " & _
        QuoteArg(ThisWorkbook.Path & Application.PathSeparator & "backup"), "完了ログ同期"
End Sub

Private Sub RunPythonAsync(ByVal arguments As String, ByVal operationName As String)
    Dim runnerPath As String
    Dim logFolder As String
    Dim logPath As String
    Dim commandText As String
    Dim processId As Double
    On Error GoTo EH
    If mCheckScheduled Then
        MsgBox mOperationName & "を実行中です。完了後にもう一度実行してください。", vbExclamation
        Exit Sub
    End If
    runnerPath = ThisWorkbook.Path & Application.PathSeparator & "task_management_runner.bat"
    If Dir$(runnerPath) = vbNullString Then
        MsgBox "実行用BATが見つかりません。" & vbCrLf & runnerPath, vbExclamation
        Exit Sub
    End If
    ThisWorkbook.Save
    logFolder = ResolveBridgeLogFolder()
    logPath = logFolder & Application.PathSeparator & Format$(Now, "yyyymmdd_hhnnss") & "_" & BOOK_KIND & ".log"
    mCompletionPath = logPath & ".status"
    If Dir$(mCompletionPath) <> vbNullString Then
        Kill mCompletionPath
    End If
    commandText = "cmd.exe /D /V:ON /S /C ""cd /D " & QuoteArg(ThisWorkbook.Path) & _
        " && call " & QuoteArg(runnerPath) & " " & arguments & _
        " > " & QuoteArg(logPath) & " 2>&1 & echo !errorlevel! > " & QuoteArg(mCompletionPath) & """"
    processId = Shell(commandText, vbHide)
    mOperationName = operationName
    mLogPath = logPath
    mStartedAt = Now
    Application.StatusBar = operationName & "を実行中です..."
    SchedulePythonCompletionCheck
    Exit Sub
EH:
    Application.StatusBar = False
    mCheckScheduled = False
    MsgBox operationName & "を開始できませんでした。" & vbCrLf & _
        "エラー番号: " & CStr(Err.Number) & vbCrLf & _
        "内容: " & Err.Description & vbCrLf & _
        "BAT: " & runnerPath & vbCrLf & _
        "ログ: " & logPath, vbCritical
End Sub

Public Sub PythonBridgeCheckCompletion()
    Dim fileNumber As Integer
    Dim exitCodeText As String
    Dim exitCode As Long
    On Error GoTo RetryLater
    mCheckScheduled = False
    If Dir$(mCompletionPath) <> vbNullString Then
        fileNumber = FreeFile
        Open mCompletionPath For Input As #fileNumber
        Line Input #fileNumber, exitCodeText
        Close #fileNumber
        exitCode = CLng(Trim$(exitCodeText))
        Kill mCompletionPath
        Application.StatusBar = False
        If exitCode = 0 Then
            MsgBox mOperationName & "が完了しました。", vbInformation
        Else
            MsgBox mOperationName & "に失敗しました。" & vbCrLf & _
                "ログを確認してください。" & vbCrLf & mLogPath, vbCritical
        End If
        ClearPythonCompletionState
        Exit Sub
    End If
    If DateDiff("n", mStartedAt, Now) >= 120 Then
        Application.StatusBar = False
        MsgBox mOperationName & "の完了を確認できませんでした。" & vbCrLf & _
            "ログを確認してください。" & vbCrLf & mLogPath, vbExclamation
        ClearPythonCompletionState
        Exit Sub
    End If
RetryLater:
    SchedulePythonCompletionCheck
End Sub

Public Sub CancelPythonCompletionCheck()
    On Error Resume Next
    If mCheckScheduled Then
        Application.OnTime EarliestTime:=mNextCheck, Procedure:=CompletionProcedureName(), Schedule:=False
    End If
    Application.StatusBar = False
    ClearPythonCompletionState
    On Error GoTo 0
End Sub

Private Sub SchedulePythonCompletionCheck()
    mNextCheck = Now + TimeSerial(0, 0, 2)
    mCheckScheduled = True
    Application.OnTime EarliestTime:=mNextCheck, Procedure:=CompletionProcedureName(), Schedule:=True
End Sub

Private Function CompletionProcedureName() As String
    CompletionProcedureName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PythonBridgeCheckCompletion"
End Function

Private Sub ClearPythonCompletionState()
    mCheckScheduled = False
    mCompletionPath = vbNullString
    mLogPath = vbNullString
    mOperationName = vbNullString
    mStartedAt = 0
End Sub

Private Function SelectJsonFile() As String
    Dim dialogObject As FileDialog
    Set dialogObject = Application.FileDialog(msoFileDialogFilePicker)
    dialogObject.Title = "ImportするJSON v3を選択"
    dialogObject.Filters.Clear
    dialogObject.Filters.Add "JSON", "*.json"
    If dialogObject.Show = -1 Then
        SelectJsonFile = dialogObject.SelectedItems(1)
    End If
End Function

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr$(34) & value & Chr$(34)
End Function

Private Sub EnsureBridgeFolder(ByVal folderPath As String)
    If Dir$(folderPath, vbDirectory) = vbNullString Then
        MkDir folderPath
    End If
End Sub

Private Function ResolveBridgeLogFolder() As String
    Dim folderPath As String
    On Error GoTo UseTempFolder
    folderPath = ThisWorkbook.Path & Application.PathSeparator & "logs"
    EnsureBridgeFolder folderPath
    ResolveBridgeLogFolder = folderPath
    Exit Function
UseTempFolder:
    Err.Clear
    folderPath = Environ$("TEMP") & Application.PathSeparator & "TaskManagementLogs"
    EnsureBridgeFolder folderPath
    ResolveBridgeLogFolder = folderPath
End Function
