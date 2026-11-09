[CmdletBinding()]
param(
    [string]$ToolkitUrl = 'http://127.0.0.1:6003',
    [switch]$RecalcTotals,
    [string[]]$TotalsRegister = @('*'),
    [switch]$SequenceReport,
    [switch]$StartAccessUpdate,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-1c-maintenance'))
)
<#
    Post-delete maintenance inside 1C, through the embedded server of
    1c-mcp-toolkit (MCP_Toolkit.epf opened with
    /C "startup;mode=embedded;port=6003"; anonymization off, JSON answers).

      -RecalcTotals       ПересчитатьИтоги() for every accumulation and
                          accounting register that uses totals, one register per
                          call.  TiI -RecalcTotals skipped three registers on
                          bit_copy (Хозрасчетный among them).
      -TotalsRegister     register names (wildcards) to recalculate, '*' by
                          default; after TiI -RecalcTotals only the registers
                          TiI left stale are worth a call.
      -SequenceReport     boundaries of every document sequence (query
                          Последовательность.<Имя>.Границы).  Restoring a
                          sequence re-posts documents, so it is left to the user.
      -StartAccessUpdate  starts the BSP record-level access update, as the
                          "Обновление доступа" command does (background job).

    Every step writes its own CSV; summary.json collects them.
#>
$ErrorActionPreference = 'Stop'
if (-not ($RecalcTotals -or $SequenceReport -or $StartAccessUpdate)) { throw 'Choose at least one of -RecalcTotals, -SequenceReport, -StartAccessUpdate.' }

function Invoke-1C([string]$Code, [int]$TimeoutSec = 7200) {
    # 1C serializes Результат itself, so the answer does not depend on the
    # toolkit's response format setting (a new 1C user gets text, not JSON).
    $Code += "`nЗаписьJSON = Новый ЗаписьJSON; ЗаписьJSON.УстановитьСтроку();`nЗаписатьJSON(ЗаписьJSON, Результат);`nРезультат = ЗаписьJSON.Закрыть();"
    $body = @{ code = $Code } | ConvertTo-Json -Compress
    $answer = Invoke-RestMethod -Uri "$ToolkitUrl/api/execute_code" -Method Post -ContentType 'application/json; charset=utf-8' `
        -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec $TimeoutSec
    if (-not $answer.success) { throw "1C error: $($answer.error)" }
    $value = ([string]$answer.data) | ConvertFrom-Json
    # In JSON response mode the toolkit wraps the string in quotes once more;
    # a string result (the user name) stays a string.
    if ($value -is [string]) { try { $value = $value | ConvertFrom-Json } catch { } }
    return $value
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$summary = [ordered]@{ generated_at_local=(Get-Date).ToString('s'); toolkit=$ToolkitUrl
                       user=(Invoke-1C 'Результат = ПользователиИнформационнойБазы.ТекущийПользователь().Имя;') }

if ($RecalcTotals) {
    $registers = Invoke-1C @'
Массив = Новый Массив;
Для Каждого М Из Метаданные.РегистрыНакопления Цикл
	Если РегистрыНакопления[М.Имя].ПолучитьИспользованиеИтогов() Тогда
		Массив.Добавить(Новый Структура("Вид, Имя", "РегистрыНакопления", М.Имя));
	КонецЕсли;
КонецЦикла;
Для Каждого М Из Метаданные.РегистрыБухгалтерии Цикл
	Если РегистрыБухгалтерии[М.Имя].ПолучитьИспользованиеИтогов() Тогда
		Массив.Добавить(Новый Структура("Вид, Имя", "РегистрыБухгалтерии", М.Имя));
	КонецЕсли;
КонецЦикла;
Результат = Массив;
'@
    $registers = @($registers | Where-Object { $name = $_.Имя; @($TotalsRegister | Where-Object { $name -like $_ }).Count -gt 0 })
    if ($registers.Count -eq 0) { throw "No register with totals matches -TotalsRegister $($TotalsRegister -join ', ')." }
    $rows = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($register in @($registers)) {
        $index++
        $started = Get-Date
        $status = 'OK'
        try { [void](Invoke-1C "$($register.Вид).$($register.Имя).ПересчитатьИтоги(); Результат = Истина;") }
        catch { $status = $_.Exception.Message }
        $seconds = [Math]::Round(((Get-Date) - $started).TotalSeconds, 1)
        $rows.Add([pscustomobject]@{ kind=$register.Вид; register=$register.Имя; seconds=$seconds; status=$status })
        Write-Host ("Totals {0}/{1} {2}.{3}: {4} s {5}" -f $index, @($registers).Count, $register.Вид, $register.Имя, $seconds, $status)
    }
    $path = Join-Path $OutputDirectory 'recalc-totals.csv'
    $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    $summary.recalc_totals = [ordered]@{ registers=$rows.Count; failed=@($rows | Where-Object status -ne 'OK').Count
                                         seconds=[Math]::Round(($rows | Measure-Object seconds -Sum).Sum, 1); csv=$path }
}

if ($SequenceReport) {
    $sequences = Invoke-1C @'
Массив = Новый Массив;
Для Каждого М Из Метаданные.Последовательности Цикл
	Запрос = Новый Запрос("ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Границ, МИНИМУМ(Г.Период) КАК Минимум, МАКСИМУМ(Г.Период) КАК Максимум
	|ИЗ Последовательность." + М.Имя + ".Границы КАК Г");
	Выборка = Запрос.Выполнить().Выбрать();
	Выборка.Следующий();
	Массив.Добавить(Новый Структура("Последовательность, Границ, Минимум, Максимум",
		М.Имя, Выборка.Границ, ?(Выборка.Минимум = NULL, Неопределено, Выборка.Минимум),
		?(Выборка.Максимум = NULL, Неопределено, Выборка.Максимум)));
КонецЦикла;
Результат = Массив;
'@
    $path = Join-Path $OutputDirectory 'sequences.csv'
    @($sequences) | ForEach-Object { [pscustomobject]@{ sequence=$_.Последовательность; boundaries=$_.Границ; min_period=$_.Минимум; max_period=$_.Максимум } } |
        Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    $summary.sequences = [ordered]@{ count=@($sequences).Count; csv=$path }
}

if ($StartAccessUpdate) {
    # Universal record-level restriction: start the background update.
    # Standard restriction (bit_copy): rebuild the BSP access registers
    # and user roles, as "Обновление доступа" does in that mode.
    $universal = Invoke-1C 'Результат = Константы.ОграничиватьДоступНаУровнеЗаписейУниверсально.Получить();'
    $steps = if ($universal) { @('УправлениеДоступомСлужебный.ЗапуститьОбновлениеДоступаНаУровнеЗаписей(Истина)') } else { @(
        'РегистрыСведений.ТаблицыГруппДоступа.ОбновитьДанныеРегистра()',
        'РегистрыСведений.ЗначенияГруппДоступа.ОбновитьДанныеРегистра()',
        'РегистрыСведений.ГруппыЗначенийДоступа.ОбновитьДанныеРегистра()',
        'РегистрыСведений.НаборыЗначенийДоступа.ОбновитьДанныеРегистра()',
        'УправлениеДоступом.ОбновитьРолиПользователей()') }
    $rows = foreach ($step in $steps) {
        $started = Get-Date
        $status = 'OK'
        try { [void](Invoke-1C "$step; Результат = Истина;") } catch { $status = $_.Exception.Message }
        $seconds = [Math]::Round(((Get-Date) - $started).TotalSeconds, 1)
        Write-Host ("Access {0}: {1} s {2}" -f $step, $seconds, $status)
        [pscustomobject]@{ step=$step; seconds=$seconds; status=$status }
    }
    $path = Join-Path $OutputDirectory 'access-update.csv'
    $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    $summary.access_update = [ordered]@{ mode=if ($universal) { 'universal' } else { 'standard' }
                                         failed=@($rows | Where-Object status -ne 'OK').Count; csv=$path }
}

$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 6
