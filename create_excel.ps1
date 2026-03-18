# nagios-conf-to-xlsx.xlsm を生成するスクリプト
# 実行前に Excel の「トラストセンター」で
# 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。

$OutputPath = "$PSScriptRoot\nagios-conf-to-xlsx.xlsm"

# NOTE: VBA code must NOT contain Japanese characters (encoding issue with AddFromString).
#       Sheet access uses index (2) instead of name. Messages are in English.
$vbaCode = @'
Option Explicit

' ===== Windows API: get local/UTC system time for timezone detection =====
Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub GetLocalTime  Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Sub GetLocalTime  Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

' Returns the local timezone offset from UTC in hours (e.g. +9 for JST)
Function GetUTCOffsetHours() As Double
    Dim localST As SYSTEMTIME
    Dim utcST   As SYSTEMTIME
    GetLocalTime  localST
    GetSystemTime utcST
    Dim localDt As Date
    Dim utcDt   As Date
    localDt = DateSerial(localST.wYear, localST.wMonth, localST.wDay) + _
              TimeSerial(localST.wHour, localST.wMinute, localST.wSecond)
    utcDt   = DateSerial(utcST.wYear,   utcST.wMonth,   utcST.wDay)   + _
              TimeSerial(utcST.wHour,   utcST.wMinute,  utcST.wSecond)
    GetUTCOffsetHours = CDbl(DateDiff("n", utcDt, localDt)) / 60
End Function

' Convert Unix timestamp (seconds since 1970-01-01 UTC) to Excel date serial in local time
' Returns 0 if ts is empty/invalid
Function UnixToSerial(ts As Variant) As Double
    On Error GoTo ErrExit
    If IsEmpty(ts) Or IsNull(ts) Then GoTo ErrExit
    If Trim(CStr(ts)) = "" Or CDbl(ts) = 0 Then GoTo ErrExit
    ' Excel date serial: days from 1900-01-00; Unix epoch = serial 25569
    UnixToSerial = CDbl(ts) / 86400# + 25569# + GetUTCOffsetHours() / 24#
    Exit Function
ErrExit:
    UnixToSerial = 0
End Function

' ===== Trim leading/trailing spaces AND tabs (VBA Trim() only removes spaces) =====
Function TrimWS(s As String) As String
    Dim result As String
    result = s
    Do While Len(result) > 0
        If Left(result, 1) = " " Or Left(result, 1) = Chr(9) Then
            result = Mid(result, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(result) > 0
        If Right(result, 1) = " " Or Right(result, 1) = Chr(9) Then
            result = Left(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimWS = result
End Function

' ===== Create push button on Sheet1 (called from PowerShell after module load) =====
Sub SetupButton()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ' Remove old button if exists
    On Error Resume Next
    ws.Shapes("btnRun").Delete
    On Error GoTo 0
    ' Place button below description text (row 24)
    Dim btnTop  As Double
    Dim btnLeft As Double
    btnTop  = ws.Rows(24).Top + 5
    btnLeft = ws.Columns(2).Left
    ' msoShapeRoundedRectangle = 5
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, btnLeft, btnTop, 300, 80)
    shp.Name = "btnRun"
    shp.OnAction = "NagiosModule.RunMacro"
    With shp.Fill
        .ForeColor.RGB = RGB(70, 130, 180)
    End With
    With shp.Line
        .ForeColor.RGB = RGB(50, 100, 150)
    End With
    With shp.TextFrame.Characters.Font
        .Color = RGB(255, 255, 255)
        .Size  = 14
        .Bold  = True
    End With
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment   = xlVAlignCenter
End Sub

' ===== Main entry point =====
Sub RunMacro()
    Dim objCachePath As String
    Dim statusDatPath As String
    Dim useStatusDat As Boolean

    On Error GoTo ErrHandler

    ' Select objects.cache
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select objects.cache file"
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path & "\"
        If .Show = True Then
            objCachePath = .SelectedItems(1)
        Else
            MsgBox "objects.cache was not selected. Aborted.", vbExclamation, "Aborted"
            Exit Sub
        End If
    End With

    ' Select status.dat (optional - cancel to skip)
    useStatusDat = False
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select status.dat file (Cancel to skip)"
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        If .Show = True Then
            statusDatPath = .SelectedItems(1)
            useStatusDat = True
        End If
    End With

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call ParseObjectsCache(objCachePath)

    If useStatusDat Then
        Call ParseStatusDat(statusDatPath)
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Done.", vbInformation, "Complete"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

' ===== Read UTF-8 file via ADODB.Stream =====
Function ReadUTF8File(filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadUTF8File = .ReadText
        .Close
    End With
End Function

' ===== Parse objects.cache -> Sheet 2 =====
Sub ParseObjectsCache(filePath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(2)
    ws.Cells.Clear

    Dim headers As Variant
    headers = Array("host_name", "service_description", "check_period", "check_command", _
                    "event_handler", "notification_period", "check_interval", "retry_interval", _
                    "max_check_attempts", "active_checks_enabled", "passive_checks_enabled", "event_handler_enabled")

    Dim i As Integer
    For i = 0 To 11
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    With ws.Range("A1:L1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' AutoFilter on header row
    ws.Range("A1").AutoFilter

    Dim content As String
    content = ReadUTF8File(filePath)

    Dim lines() As String
    lines = Split(content, vbLf)

    Dim row As Long
    row = 2
    Dim inService As Boolean
    inService = False
    Dim serviceData(11) As String
    Dim j As Integer
    Dim li As Long
    Dim line As String
    Dim key As String
    Dim value As String
    Dim tabPos As Integer

    For li = 0 To UBound(lines)
        line = lines(li)
        If Len(line) > 0 Then
            If Right(line, 1) = Chr(13) Then line = Left(line, Len(line) - 1)
        End If
        line = TrimWS(line)

        If line = "define service {" Then
            inService = True
            For j = 0 To 11
                serviceData(j) = ""
            Next j
        ElseIf line = "}" And inService Then
            inService = False
            For j = 0 To 11
                ws.Cells(row, j + 1).Value = serviceData(j)
            Next j
            row = row + 1
        ElseIf inService And Len(line) > 0 Then
            tabPos = InStr(line, Chr(9))
            If tabPos > 0 Then
                key   = Trim(Left(line, tabPos - 1))
                value = Trim(Mid(line, tabPos + 1))
                Select Case key
                    Case "host_name":              serviceData(0)  = value
                    Case "service_description":    serviceData(1)  = value
                    Case "check_period":           serviceData(2)  = value
                    Case "check_command":          serviceData(3)  = value
                    Case "event_handler":          serviceData(4)  = value
                    Case "notification_period":    serviceData(5)  = value
                    Case "check_interval":         serviceData(6)  = value
                    Case "retry_interval":         serviceData(7)  = value
                    Case "max_check_attempts":     serviceData(8)  = value
                    Case "active_checks_enabled":  serviceData(9)  = value
                    Case "passive_checks_enabled": serviceData(10) = value
                    Case "event_handler_enabled":  serviceData(11) = value
                End Select
            End If
        End If
    Next li

    ws.Columns("A:L").AutoFit
    ws.Activate
End Sub

' ===== Parse status.dat -> append downtime columns =====
Sub ParseStatusDat(filePath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(2)

    Dim lastDataRow As Long
    lastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Build lookup: "host_name|service_description" -> row number
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim lookupKey As String
    For r = 2 To lastDataRow
        lookupKey = ws.Cells(r, 1).Value & "|" & ws.Cells(r, 2).Value
        If Not dict.Exists(lookupKey) Then
            dict.Add lookupKey, r
        End If
    Next r

    ' Downtime columns start at column 13 (M)
    Dim baseCol As Integer
    baseCol = 13
    ws.Cells(1, baseCol).Value     = "is_in_effect"
    ws.Cells(1, baseCol + 1).Value = "start_time"
    ws.Cells(1, baseCol + 2).Value = "end_time"
    With ws.Range(ws.Cells(1, baseCol), ws.Cells(1, baseCol + 2))
        .Font.Bold = True
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
    End With

    Dim content As String
    content = ReadUTF8File(filePath)

    Dim lines() As String
    lines = Split(content, vbLf)

    Dim inDowntime As Boolean
    inDowntime = False
    Dim dtHostName    As String
    Dim dtServiceDesc As String
    Dim dtIsInEffect  As String
    Dim dtStartTime   As String
    Dim dtEndTime     As String

    Dim li    As Long
    Dim line  As String
    Dim key   As String
    Dim value As String
    Dim eqPos As Integer

    For li = 0 To UBound(lines)
        line = lines(li)
        If Len(line) > 0 Then
            If Right(line, 1) = Chr(13) Then line = Left(line, Len(line) - 1)
        End If
        line = TrimWS(line)

        If line = "servicedowntime {" Then
            inDowntime    = True
            dtHostName    = ""
            dtServiceDesc = ""
            dtIsInEffect  = ""
            dtStartTime   = ""
            dtEndTime     = ""
        ElseIf line = "}" And inDowntime Then
            inDowntime = False
            Dim matchKey As String
            matchKey = dtHostName & "|" & dtServiceDesc
            If dict.Exists(matchKey) Then
                Dim matchRow As Long
                matchRow = dict(matchKey)
                ' Find next empty downtime slot
                Dim col As Integer
                col = baseCol
                Do While ws.Cells(matchRow, col).Value <> ""
                    col = col + 3
                Loop
                ' Add header if this is a new column set
                If ws.Cells(1, col).Value = "" Then
                    ws.Cells(1, col).Value     = "is_in_effect"
                    ws.Cells(1, col + 1).Value = "start_time"
                    ws.Cells(1, col + 2).Value = "end_time"
                    With ws.Range(ws.Cells(1, col), ws.Cells(1, col + 2))
                        .Font.Bold = True
                        .Interior.Color = RGB(197, 90, 17)
                        .Font.Color = RGB(255, 255, 255)
                    End With
                End If
                ws.Cells(matchRow, col).Value     = dtIsInEffect
                Dim startSerial As Double
                Dim endSerial   As Double
                startSerial = UnixToSerial(dtStartTime)
                endSerial   = UnixToSerial(dtEndTime)
                With ws.Cells(matchRow, col + 1)
                    .Value        = startSerial
                    .NumberFormat = "yyyy/mm/dd hh:mm:ss"
                End With
                With ws.Cells(matchRow, col + 2)
                    .Value        = endSerial
                    .NumberFormat = "yyyy/mm/dd hh:mm:ss"
                End With
            End If
        ElseIf inDowntime And Len(line) > 0 Then
            eqPos = InStr(line, "=")
            If eqPos > 0 Then
                key   = Trim(Left(line, eqPos - 1))
                value = Trim(Mid(line, eqPos + 1))
                Select Case key
                    Case "host_name":           dtHostName    = value
                    Case "service_description": dtServiceDesc = value
                    Case "is_in_effect":        dtIsInEffect  = value
                    Case "start_time":          dtStartTime   = value
                    Case "end_time":            dtEndTime     = value
                End Select
            End If
        End If
    Next li

    ws.Columns.AutoFit

    ' Re-apply AutoFilter to include downtime columns (M/N/O)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1").AutoFilter
End Sub
'@

Write-Host "Creating Excel file..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Add()

    # Rename Sheet1 to Japanese name (PS handles Unicode correctly with BOM)
    $sheet1 = $workbook.Sheets(1)
    $sheet1.Name = "操作"

    # Remove extra sheets if any
    $initialCount = $workbook.Sheets.Count
    for ($idx = $initialCount; $idx -ge 2; $idx--) {
        $workbook.Sheets($idx).Delete()
    }

    # Add Sheet2 with Japanese name
    $sheet2 = $workbook.Sheets.Add([System.Reflection.Missing]::Value, $workbook.Sheets(1))
    $sheet2.Name = "サービス定義"

    # Add VBA module first (SetupButton is needed below)
    $vbaModule = $workbook.VBProject.VBComponents.Add(1)
    $vbaModule.Name = "NagiosModule"
    $vbaModule.CodeModule.AddFromString($vbaCode)

    # Create push button via VBA (avoids AddFormControl type-constant issues in PowerShell)
    $excel.Run("NagiosModule.SetupButton")

    # Set Japanese text on button (PS handles Unicode correctly)
    $sheet1.Shapes("btnRun").TextFrame.Characters().Text = "マクロ実行"

    # ===== Sheet1: description and usage =====
    $sheet1.Activate()
    $s = $sheet1

    # Column B width
    $s.Columns(2).ColumnWidth = 90

    # Title (row 1)
    $s.Rows(1).RowHeight = 36
    $s.Cells(1, 2).Value = "Nagios サービス定義 Excel 出力ツール"
    $s.Cells(1, 2).Font.Size = 18
    $s.Cells(1, 2).Font.Bold = $true
    $s.Cells(1, 2).Font.Color = 0xC47244  # RGB(68,114,196) steel blue

    # ■ 機能説明 (row 3-6)
    $s.Cells(3, 2).Value = "■ 機能説明"
    $s.Cells(3, 2).Font.Bold = $true
    $s.Cells(4, 2).Value = "  ・Nagios の objects.cache ファイルに定義されたサービス情報を読み込み、「サービス定義」シートに一覧出力します。"
    $s.Cells(5, 2).Value = "  ・status.dat を指定した場合は、ダウンタイム情報（is_in_effect / start_time / end_time）を対象サービス行に追記します。"
    $s.Cells(6, 2).Value = "  ・start_time / end_time は Unix タイムスタンプを現地時間（yyyy/mm/dd hh:mm:ss）に変換して出力します。"

    # ■ 利用方法 (row 8-13)
    $s.Cells(8, 2).Value = "■ 利用方法"
    $s.Cells(8, 2).Font.Bold = $true
    $s.Cells(9,  2).Value = "  ① 「マクロ実行」ボタンをクリックします。"
    $s.Cells(10, 2).Value = "  ② ファイル選択ダイアログが表示されます。objects.cache ファイルを選択してください。"
    $s.Cells(11, 2).Value = "       （初期ディレクトリはこの Excel ファイルと同じフォルダです）"
    $s.Cells(12, 2).Value = "  ③ 続けて status.dat ファイルの選択ダイアログが表示されます。不要な場合はキャンセルしてください。"
    $s.Cells(13, 2).Value = "  ④ 処理完了後、「サービス定義」シートにデータが出力されます。"

    # ■ 出力内容 (row 15-18)
    $s.Cells(15, 2).Value = "■ 出力内容（「サービス定義」シート）"
    $s.Cells(15, 2).Font.Bold = $true
    $s.Cells(16, 2).Value = "  ・A ～ L 列：サービス定義情報（host_name / service_description / check_command など 12 項目）"
    $s.Cells(17, 2).Value = "  ・M ～ O 列：ダウンタイム情報（is_in_effect / start_time / end_time）※ status.dat 指定時のみ"
    $s.Cells(18, 2).Value = "  ・1 行目にオートフィルターを設定済み。マクロ実行のたびにシート内容はクリアされます。"

    # ■ 注意事項 (row 20-22)
    $s.Cells(20, 2).Value = "■ 注意事項"
    $s.Cells(20, 2).Font.Bold = $true
    $s.Cells(21, 2).Value = "  ・マクロを実行するたびに「サービス定義」シートの内容はクリアされます。"
    $s.Cells(22, 2).Value = "  ・本ファイルはマクロ有効ブック（.xlsm）です。Excel のセキュリティ警告が表示された場合はコンテンツを有効にしてください。"

    # Row 24: button (placed by SetupButton)

    # Save as xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
    $workbook.SaveAs($OutputPath, 52)
    Write-Host "Done: $OutputPath"

} catch {
    Write-Error "Error: $_"
    Write-Host ""
    Write-Host "[If VBProject access is denied]"
    Write-Host "Go to: Excel -> File -> Options -> Trust Center -> Trust Center Settings"
    Write-Host "       -> Macro Settings -> Check 'Trust access to the VBA project object model'"
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
}
