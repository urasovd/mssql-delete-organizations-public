[CmdletBinding()]
param(
    [ValidateSet('Summary','Access','Check','Load','Unbind')]
    [string]$Action = 'Summary',
    [string]$ToolkitUrl = 'http://127.0.0.1:6013',
    [string]$Processor = (Join-Path $PSScriptRoot '..\build\1c\ПодготовкаКопииПослеВыделения.epf'),
    [string[]]$PathSubstitution = @(),
    [int]$BatchFiles = 500,
    [int]$BatchMB = 150,
    [switch]$IncludeVolumesWithFiles,
    [string]$ConfirmationToken,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-volume-files'))
)
<#
    Attached files stored in file volumes, through the processor
    1c/ПодготовкаКопииПослеВыделения (build it with Build-1CTools.ps1) and
    MCP_Toolkit opened in a thin client of the infobase.

      Summary  files on volumes by catalog and volume.
      Access   whether the 1C server sees every volume folder.  The server reads
               files under its own account; when it cannot, save a credential
               for it in the processor form ("Сохранить для сервера 1С") or
               copy the files and pass -PathSubstitution 'from => to'.
      Check    every file exists and has the size stored in the base.  Reads
               no file content.
      Load     files into the infobase, one transaction per file.  Files on the
               volumes are never deleted or changed.  Repeatable: loaded files
               leave the list.  Token LOAD-FILES.
      Unbind   new files are stored in the infobase, volume paths are cleared.
               Volumes that still have files are kept unless
               -IncludeVolumesWithFiles (for files missing from the storage).
               Token UNBIND-VOLUMES.

    Check and Load write files.csv-style errors with the owner and file name:
    files missing from the storage are what users have to attach again.
#>
$ErrorActionPreference = 'Stop'

if ($Action -eq 'Load' -and $ConfirmationToken -ne 'LOAD-FILES') { throw 'Load changes the infobase and requires -ConfirmationToken LOAD-FILES.' }
if ($Action -eq 'Unbind' -and $ConfirmationToken -ne 'UNBIND-VOLUMES') { throw 'Unbind changes the infobase and requires -ConfirmationToken UNBIND-VOLUMES.' }
$Processor = (Resolve-Path -LiteralPath $Processor).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

function ConvertTo-1CString([string]$Value) { '"' + $Value.Replace('"', '""') + '"' }

function Invoke-Processor([string]$Code, [int]$TimeoutSec = 175) {
    $full = "О = ВнешниеОбработки.Создать($(ConvertTo-1CString $Processor), Ложь);`n$Code`n" +
        "Запись = Новый ЗаписьJSON; Запись.УстановитьСтроку(); ЗаписатьJSON(Запись, Результат); Результат = Запись.Закрыть();"
    $body = @{ code = $full } | ConvertTo-Json -Compress
    $answer = Invoke-RestMethod -Uri "$ToolkitUrl/api/execute_code" -Method Post -ContentType 'application/json; charset=utf-8' `
        -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec $TimeoutSec
    if (-not $answer.success) { throw "1C error: $($answer.error)" }
    # 1C serialized Результат itself; in JSON response mode the toolkit wraps that text in quotes once more.
    $value = ([string]$answer.data) | ConvertFrom-Json
    if ($value -is [string]) { try { $value = $value | ConvertFrom-Json } catch { } }
    return $value
}

$substitution = ($PathSubstitution | ForEach-Object { "П.ПодменаПутей.Добавить($(ConvertTo-1CString $_));" }) -join ' '
$summary = [ordered]@{ generated_at_local = (Get-Date).ToString('s'); action = $Action; toolkit = $ToolkitUrl
                       infobase = (Invoke-Processor 'Результат = СтрокаСоединенияИнформационнойБазы();') }

switch ($Action) {
    'Summary' {
        $result = Invoke-Processor 'Результат = О.СводкаФайловНаТомах();'
        $result.Справочники | Select-Object Справочник, Файлов, Размер | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'catalogs.csv') -NoTypeInformation -Encoding UTF8
        $result.Тома | Select-Object Том, Путь, Файлов, Размер | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'volumes.csv') -NoTypeInformation -Encoding UTF8
        $summary.files = $result.Файлов; $summary.bytes = $result.Размер; $summary.storage_mode = $result.СпособХранения
        $summary.volumes = @($result.Тома | ForEach-Object { "$($_.Том): $($_.Файлов)" })
    }
    'Access' {
        $paths = ($PathSubstitution | ForEach-Object { "Подмена.Добавить($(ConvertTo-1CString $_));" }) -join ' '
        $result = Invoke-Processor "Подмена = Новый Массив; $paths Результат = О.ДоступСервераКТомам(Подмена);"
        $summary.server_temp_directory = $result.КаталогВременныхФайлов
        $summary.volumes = @($result.Тома | ForEach-Object { [ordered]@{ volume = $_.Том; path = $_.Путь; available = $_.Доступен } })
        $summary.result = if (@($result.Тома | Where-Object { -not $_.Доступен }).Count) { 'NO_ACCESS' } else { 'OK' }
    }
    { $_ -in 'Check', 'Load' } {
        $method = if ($Action -eq 'Load') { 'ЗагрузитьФайлыВБазу' } else { 'ПроверитьФайлыНаТомах' }
        $position = 0; $done = 0; $errors = @(); $timeouts = 0; $watch = [Diagnostics.Stopwatch]::StartNew()
        while ($true) {
            try {
                $batch = Invoke-Processor ("П = О.ПараметрыФайлов(); П.Порция = $BatchFiles; П.ПорцияМБ = $BatchMB; " +
                    "П.Позиция = $position; $substitution Результат = О.$method(П);")
            } catch [System.Threading.Tasks.TaskCanceledException], [System.Net.Http.HttpRequestException] {
                # The toolkit stops waiting while 1C finishes the call: wait until it answers, then go on.
                $timeouts++
                Write-Host ("{0:s} toolkit timeout #{1}" -f (Get-Date), $timeouts)
                if ($timeouts -gt 20) { throw }
                do { $alive = $true; try { [void](Invoke-Processor 'Результат = 1;') } catch { $alive = $false } } until ($alive)
                continue
            }
            $position = $batch.Позиция; $done += $batch.Успешно; $errors += @($batch.Ошибки)
            Write-Host ("{0:s} {1}: total={2} done={3} errors={4} t={5}s" -f (Get-Date), $Action, $batch.Всего, $done, $errors.Count, [int]$watch.Elapsed.TotalSeconds)
            if ($batch.Завершено -or $batch.Обработано -eq 0) { break }
        }
        $errorsPath = Join-Path $OutputDirectory 'file-errors.csv'
        $errors | Select-Object Владелец, Файл, Путь, Ошибка, Объект | Export-Csv -LiteralPath $errorsPath -NoTypeInformation -Encoding UTF8
        $summary.done = $done; $summary.errors = $errors.Count; $summary.timeouts = $timeouts
        $summary.seconds = [int]$watch.Elapsed.TotalSeconds; $summary.errors_csv = $errorsPath
        $summary.remaining_on_volumes = (Invoke-Processor 'Результат = О.СводкаФайловНаТомах();').Файлов
    }
    'Unbind' {
        $include = if ($IncludeVolumesWithFiles) { 'Истина' } else { 'Ложь' }
        $result = Invoke-Processor "Результат = О.ОтвязатьТома(Ложь, $include);"
        $summary.storage_mode = $result.СпособХранения; $summary.unbound_volumes = $result.ОтвязаноТомов
        $summary.volumes_with_files = @($result.ТомаСФайлами)
    }
}

$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 5
