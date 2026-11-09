[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidateSet('Osv','Plan','Prepare','Registers','Postings','Convert','Update','Reconcile')] [string]$Step,
    # Osv, Plan, Postings: COM connection, no 1C client and no startup dialogs.
    [string]$ConnectionString,
    [string]$Processor = (Join-Path $PSScriptRoot '..\build\1c\ПодготовкаКопииПослеВыделения.epf'),
    # Plan, Postings: Configuration.xml of the typical configuration (/DumpConfigToFiles -listFile with "Configuration").
    [string]$TargetConfigurationXml,
    # Reconcile: the infobase before conversion.
    [string]$SourceConnectionString,
    # Registers: file infobase created from the CF.
    [string]$TargetFileBase,
    # Convert
    [switch]$OnlyUpdateDb,
    [string]$InfoBase,
    [string]$TargetCf,
    [string]$User = 'admin',
    [string]$Password = '',
    [string]$PlatformDirectory = 'C:\Program Files\1cv8\8.3.27.2214\bin',
    # Osv
    [datetime]$PeriodTo = '2999-12-31',
    [string]$CompareWith,
    [int]$BatchDocuments = 50,
    [string]$ConfirmationToken,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot ('..\artifacts\' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-convert-' + $Step.ToLower()))
)
<#
    Turning an extracted copy of a modified configuration (BIT) into a typical
    one (БП КОРП).  Run on a separate copy.  Order and results:
    docs/CONVERT-TO-TYPICAL.md.

      Osv       OSV of Хозрасчетный by account through 1C; -CompareWith compares
                it with an earlier osv.csv.  Read-only.
      Plan      documents and registers of the infobase that are not in the
                typical configuration; document kinds with movements.  Read-only.
      Prepare   date ban, record-level access and object versioning off, as they
                stop or slow the steps below; service register НеудаленныеОбъекты
                cleared, its records become non-unique on restructuring.
                Token CONVERT-TO-TYPICAL.
      Registers records of information registers whose dimension holds a value of
                a type absent from the CF.  Such values turn empty on restructuring
                and the records become non-unique, and the operation is rolled back.
                Same for values of enumerations absent from the CF.  Records with
                kept values stay.  Run after Postings.  Token CONVERT-TO-TYPICAL.
      Postings  movements of those documents go to "Операция (бух)", by the
                processor 1c/ПодготовкаКопииПослеВыделения: only registers that
                stay in the typical configuration.  Token CONVERT-TO-TYPICAL.
      Convert   sessions closed, scheduled jobs denied, all extensions deleted,
                the configuration replaced by -TargetCf, database restructured.
                Objects absent from the CF are removed with their data.
                Token CONVERT-TO-TYPICAL.

      Update    update handlers of the new configuration, on the server.  A client
                would stop at the startup dialogs of a copy.  Token CONVERT-TO-TYPICAL.

      Reconcile the converted copy against -SourceConnectionString through 1C, as SQL
                tables differ in columns now: rows and resource sums of every common
                accumulation and information register, Хозрасчетный by account and
                subconto.  Read-only.

    Osv, Plan, Prepare, Registers, Postings, Update and Reconcile run through the COM connector (-ConnectionString
    'Srvr="server";Ref="base";Usr="admin";Pwd=""').  Convert closes all sessions.
#>
$ErrorActionPreference = 'Stop'
if ($Step -in 'Prepare', 'Registers', 'Postings', 'Convert', 'Update' -and $ConfirmationToken -ne 'CONVERT-TO-TYPICAL') { throw "$Step changes the infobase: pass -ConfirmationToken CONVERT-TO-TYPICAL." }
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path

function ConvertTo-1CString([string]$Value) { '"' + $Value.Replace('"', '""') + '"' }

$script:ProcessorObject = $null

function Invoke-1C([string]$Code) {
    $full = "О = ЭтотОбъект;`n$Code`nЗапись = Новый ЗаписьJSON; Запись.УстановитьСтроку(); ЗаписатьJSON(Запись, Результат); Результат = Запись.Закрыть();"
    if (-not $script:ProcessorObject) {
        $connection = (New-Object -ComObject V83.COMConnector).Connect($ConnectionString)
        # The COM adapter of PowerShell does not see properties of the connection: late binding.
        $manager = [System.__ComObject].InvokeMember('ВнешниеОбработки', [Reflection.BindingFlags]::GetProperty, $null, $connection, $null)
        $script:ProcessorObject = [System.__ComObject].InvokeMember('Создать', [Reflection.BindingFlags]::InvokeMethod, $null, $manager, @((Resolve-Path -LiteralPath $Processor).Path, $false))
    }
    $text = [System.__ComObject].InvokeMember('ВыполнитьКод', [Reflection.BindingFlags]::InvokeMethod, $null, $script:ProcessorObject, @($full))
    return ($text | ConvertFrom-Json)
}

function Invoke-Processor([string]$Code) { Invoke-1C $Code }

function Get-TargetObjects([string]$Path) {
    if (-not $Path) { throw "$Step needs -TargetConfigurationXml." }
    $xml = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $names = @{}
    foreach ($kind in 'Document', 'AccountingRegister', 'AccumulationRegister', 'InformationRegister', 'CalculationRegister') {
        $names[$kind] = [Collections.Generic.HashSet[string]]::new([string[]]@([regex]::Matches($xml, "<$kind>([^<]+)</$kind>") | ForEach-Object { $_.Groups[1].Value }))
    }
    return $names
}

$summary = [ordered]@{ generated_at_local = (Get-Date).ToString('s'); step = $Step }

switch ($Step) {
    'Osv' {
        $rows = Invoke-1C @"
Запрос = Новый Запрос("ВЫБРАТЬ
|	Т.Счет КАК СчетСсылка,
|	СУММА(Т.СуммаОборотДт) КАК СуммаОборотДт, СУММА(Т.СуммаОборотКт) КАК СуммаОборотКт,
|	СУММА(Т.СуммаКонечныйОстатокДт) КАК СуммаОстатокДт, СУММА(Т.СуммаКонечныйОстатокКт) КАК СуммаОстатокКт,
|	СУММА(Т.ВалютнаяСуммаКонечныйОстатокДт) КАК ВалютаОстатокДт, СУММА(Т.ВалютнаяСуммаКонечныйОстатокКт) КАК ВалютаОстатокКт,
|	СУММА(Т.КоличествоКонечныйОстатокДт) КАК КоличествоОстатокДт, СУММА(Т.КоличествоКонечныйОстатокКт) КАК КоличествоОстатокКт,
|	СУММА(Т.СуммаНУКонечныйОстатокДт) КАК НУОстатокДт, СУММА(Т.СуммаНУКонечныйОстатокКт) КАК НУОстатокКт
|ИЗ
|	РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(, &Конец) КАК Т
|СГРУППИРОВАТЬ ПО Т.Счет");
Запрос.УстановитьПараметр("Конец", Дата($($PeriodTo.ToString('yyyy,M,d')),23,59,59));
Результат = Новый Массив;
Выборка = Запрос.Выполнить().Выбрать();
Пока Выборка.Следующий() Цикл
	Строка = Новый Структура("Счет", Выборка.СчетСсылка.Код);
	Для Каждого Поле Из СтрРазделить("СуммаОборотДт,СуммаОборотКт,СуммаОстатокДт,СуммаОстатокКт,ВалютаОстатокДт,ВалютаОстатокКт,КоличествоОстатокДт,КоличествоОстатокКт,НУОстатокДт,НУОстатокКт", ",") Цикл
		Строка.Вставить(Поле, ?(Выборка[Поле] = NULL, 0, Выборка[Поле]));
	КонецЦикла;
	Результат.Добавить(Строка);
КонецЦикла;
"@
        $path = Join-Path $OutputDirectory 'osv.csv'
        $rows | Sort-Object Счет | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
        $summary.accounts = @($rows).Count
        $summary.csv = $path
        if ($CompareWith) {
            $before = @{}; Import-Csv -LiteralPath $CompareWith | ForEach-Object { $before[$_.Счет] = $_ }
            $after = @{}; Import-Csv -LiteralPath $path | ForEach-Object { $after[$_.Счет] = $_ }
            $fields = @((Import-Csv -LiteralPath $path | Select-Object -First 1).PSObject.Properties.Name | Where-Object { $_ -ne 'Счет' })
            $differences = foreach ($account in @($before.Keys) + @($after.Keys) | Sort-Object -Unique) {
                foreach ($field in $fields) {
                    $a = if ($before[$account]) { [decimal]$before[$account].$field } else { [decimal]0 }
                    $b = if ($after[$account]) { [decimal]$after[$account].$field } else { [decimal]0 }
                    if ($a -ne $b) { [pscustomobject]@{ account = $account; field = $field; before = $a; after = $b; difference = $b - $a } }
                }
            }
            $differencesPath = Join-Path $OutputDirectory 'osv-differences.csv'
            @($differences) | Export-Csv -LiteralPath $differencesPath -NoTypeInformation -Encoding UTF8
            $summary.compared_with = $CompareWith
            $summary.differences = @($differences).Count
            $summary.result = if (@($differences).Count) { 'FAIL' } else { 'PASS' }
        }
    }
    'Plan' {
        $target = Get-TargetObjects $TargetConfigurationXml
        $metadata = Invoke-1C @'
Результат = Новый Массив;
Для Каждого М Из Метаданные.Документы Цикл
	Запрос = Новый Запрос("ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Всего, ЕСТЬNULL(СУММА(ВЫБОР КОГДА Т.Проведен ТОГДА 1 ИНАЧЕ 0 КОНЕЦ), 0) КАК Проведено ИЗ Документ." + М.Имя + " КАК Т");
	Выборка = Запрос.Выполнить().Выбрать(); Выборка.Следующий();
	Результат.Добавить(Новый Структура("Имя, Представление, Всего, Проведено, ИзРасширения", М.Имя, М.Представление(), Выборка.Всего, Выборка.Проведено, М.РасширениеКонфигурации() <> Неопределено));
КонецЦикла;
'@
        $gone = @($metadata | Where-Object { -not $target.Document.Contains($_.Имя) })
        $gone | Select-Object Имя, Представление, Всего, Проведено, ИзРасширения | Sort-Object Всего -Descending |
            Export-Csv -LiteralPath (Join-Path $OutputDirectory 'documents-absent-in-typical.csv') -NoTypeInformation -Encoding UTF8
        $kinds = Invoke-Processor 'Результат = О.ВидыДокументовСДвижениями();'
        $goneNames = [Collections.Generic.HashSet[string]]::new([string[]]@($gone.Имя))
        $withMovements = @($kinds | Where-Object { $goneNames.Contains($_.Документ) })
        $withMovements | Select-Object Документ, Представление, Документов, Движений, Регистры |
            Export-Csv -LiteralPath (Join-Path $OutputDirectory 'documents-with-movements.csv') -NoTypeInformation -Encoding UTF8
        $summary.documents_absent = $gone.Count
        $summary.documents_absent_with_data = @($gone | Where-Object { $_.Всего -gt 0 }).Count
        $summary.document_rows_absent = ($gone | Measure-Object Всего -Sum).Sum
        $summary.kinds_with_movements = @($withMovements | ForEach-Object { "$($_.Документ): $($_.Документов) документов, $($_.Движений) движений" })
    }
    'Update' {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $summary.update = Invoke-1C @'
ДоОбновления = Метаданные.Версия;
ОбновлениеИнформационнойБазы.ВыполнитьОбновлениеИнформационнойБазы();
Результат = Новый Структура("Конфигурация, Версия, НужноОбновление", Метаданные.Имя, ДоОбновления, ОбновлениеИнформационнойБазы.НеобходимоОбновлениеИнформационнойБазы());
'@
        $summary.minutes = [Math]::Round($watch.Elapsed.TotalMinutes, 1)
    }
    'Reconcile' {
        if (-not $SourceConnectionString) { throw 'Reconcile needs -SourceConnectionString.' }
        $code = @'
Регистры = Новый Соответствие;
Для Каждого Вид Из СтрРазделить("РегистрыНакопления,РегистрыСведений", ",") Цикл
	Для Каждого Р Из Метаданные[Вид] Цикл
		Поля = Новый Массив;
		Поля.Добавить("КОЛИЧЕСТВО(*) КАК Строк_");
		Имена = Новый Массив;
		Для Каждого Рес Из Р.Ресурсы Цикл
			Если Рес.Тип.Типы().Количество() = 1 И Рес.Тип.СодержитТип(Тип("Число")) Тогда
				Поля.Добавить("ЕСТЬNULL(СУММА(Т." + Рес.Имя + "), 0) КАК " + Рес.Имя);
				Имена.Добавить(Рес.Имя);
			КонецЕсли;
		КонецЦикла;
		Запрос = Новый Запрос("ВЫБРАТЬ " + СтрСоединить(Поля, ", ") + " ИЗ " + ?(Вид = "РегистрыНакопления", "РегистрНакопления.", "РегистрСведений.") + Р.Имя + " КАК Т");
		В = Запрос.Выполнить().Выбрать(); В.Следующий();
		Итог = Новый Соответствие;
		Итог.Вставить("Строк_", В.Строк_);
		Для Каждого Имя Из Имена Цикл Итог.Вставить(Имя, В[Имя]); КонецЦикла;
		Регистры.Вставить(Р.Имя, Итог);
	КонецЦикла;
КонецЦикла;
Запрос = Новый Запрос("ВЫБРАТЬ Т.Счет КАК СчетСсылка, Т.Субконто1 КАК С1, Т.Субконто2 КАК С2, Т.Субконто3 КАК С3,
|	СУММА(Т.СуммаОборотДт) КАК Дт, СУММА(Т.СуммаОборотКт) КАК Кт, СУММА(Т.СуммаКонечныйОстатокДт) КАК ОстДт, СУММА(Т.СуммаКонечныйОстатокКт) КАК ОстКт
|ИЗ РегистрБухгалтерии.Хозрасчетный.ОстаткиИОбороты(, &Конец, , , , , ) КАК Т
|СГРУППИРОВАТЬ ПО Т.Счет, Т.Субконто1, Т.Субконто2, Т.Субконто3");
Запрос.УстановитьПараметр("Конец", Дата(2999, 12, 31, 23, 59, 59));
Субконто = Новый Массив;
В = Запрос.Выполнить().Выбрать();
Пока В.Следующий() Цикл
	Ключ = В.СчетСсылка.Код;
	Для Каждого Поле Из СтрРазделить("С1,С2,С3", ",") Цикл
		Знач_ = В[Поле];
		Ключ = Ключ + "|" + ?(Знач_ = NULL Или Знач_ = Неопределено, "", ?(ОбщегоНазначения.ЭтоСсылка(ТипЗнч(Знач_)) И Не Перечисления.ТипВсеСсылки().СодержитТип(ТипЗнч(Знач_)), Строка(Знач_.УникальныйИдентификатор()), Строка(Знач_)));
	КонецЦикла;
	Субконто.Добавить(Новый Структура("Ключ, Дт, Кт, ОстДт, ОстКт", Ключ, В.Дт, В.Кт, В.ОстДт, В.ОстКт));
КонецЦикла;
Результат = Новый Структура("Регистры, Субконто", Регистры, Субконто);
'@
        $target = Invoke-1C $code
        $script:ProcessorObject = $null
        $ConnectionString = $SourceConnectionString
        $source = Invoke-1C $code
        $rows = foreach ($register in $source.Регистры.PSObject.Properties) {
            $after = $target.Регистры.PSObject.Properties[$register.Name]
            if (-not $after) { [pscustomobject]@{ register = $register.Name; field = ''; source = $register.Value.Строк_; target = ''; status = 'ABSENT_IN_TARGET' }; continue }
            foreach ($field in $register.Value.PSObject.Properties) {
                $value = $after.Value.PSObject.Properties[$field.Name]
                if (-not $value) { continue }
                if ([decimal]$field.Value -ne [decimal]$value.Value) { [pscustomobject]@{ register = $register.Name; field = $field.Name; source = $field.Value; target = $value.Value; status = 'DIFFERENT' } }
            }
        }
        $rows | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'registers.csv') -NoTypeInformation -Encoding UTF8
        $absent = @($rows | Where-Object { $_.status -eq 'ABSENT_IN_TARGET' })
        $summary.registers_compared = @($source.Регистры.PSObject.Properties).Count - $absent.Count
        $summary.registers_absent_in_typical = $absent.Count
        $summary.registers_absent_with_rows = @($absent | Where-Object { [decimal]$_.source -gt 0 }).Count
        $summary.register_differences = @($rows | Where-Object { $_.status -eq 'DIFFERENT' } | ForEach-Object { "$($_.register).$($_.field): $($_.source) -> $($_.target)" })
        $before = @{}; $source.Субконто | ForEach-Object { $before[$_.Ключ] = $_ }
        $after = @{}; $target.Субконто | ForEach-Object { $after[$_.Ключ] = $_ }
        $subconto = foreach ($key in @($before.Keys) + @($after.Keys) | Sort-Object -Unique) {
            $a = $before[$key]; $b = $after[$key]
            foreach ($field in 'Дт', 'Кт', 'ОстДт', 'ОстКт') {
                $x = if ($a) { [decimal]$a.$field } else { [decimal]0 }
                $y = if ($b) { [decimal]$b.$field } else { [decimal]0 }
                if ($x -ne $y) { [pscustomobject]@{ key = $key; field = $field; source = $x; target = $y } }
            }
        }
        @($subconto) | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'osv-subconto-differences.csv') -NoTypeInformation -Encoding UTF8
        $summary.subconto_rows = @($source.Субконто).Count
        $summary.subconto_differences = @($subconto).Count
    }
    'Prepare' {
        $summary.constants = Invoke-1C @'
Результат = Новый Массив;
Для Каждого Имя Из СтрРазделить("ИспользоватьДатыЗапретаИзменения,ОграничиватьДоступНаУровнеЗаписей,ИспользоватьВерсионированиеОбъектов", ",") Цикл
	Если Метаданные.Константы.Найти(Имя) = Неопределено Тогда
		Результат.Добавить(Новый Структура("Константа, Было", Имя, "нет в конфигурации"));
		Продолжить;
	КонецЕсли;
	Результат.Добавить(Новый Структура("Константа, Было", Имя, Константы[Имя].Получить()));
	Константы[Имя].Установить(Ложь);
КонецЦикла;
Для Каждого Имя Из СтрРазделить("НеудаленныеОбъекты", ",") Цикл
	Если Метаданные.РегистрыСведений.Найти(Имя) = Неопределено Тогда
		Продолжить;
	КонецЕсли;
	Запрос = Новый Запрос("ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК Н ИЗ РегистрСведений." + Имя);
	Выборка = Запрос.Выполнить().Выбрать(); Выборка.Следующий();
	Результат.Добавить(Новый Структура("Регистр, Было", Имя, Выборка.Н));
	Набор = РегистрыСведений[Имя].СоздатьНаборЗаписей();
	Набор.ОбменДанными.Загрузка = Истина;
	Набор.Записать();
КонецЦикла;
'@
    }
    'Registers' {
        if (-not $TargetFileBase) { throw 'Registers needs -TargetFileBase.' }
        $describe = @'
Результат = Новый Массив;
Для Каждого Р Из Метаданные.РегистрыСведений Цикл
	Изм = Новый Массив;
	Для Каждого М Из Р.Измерения Цикл
		Т = Новый Массив;
		Для Каждого Ти Из М.Тип.Типы() Цикл
			Об = Метаданные.НайтиПоТипу(Ти);
			Т.Добавить(?(Об = Неопределено, Строка(Ти), Об.ПолноеИмя()));
		КонецЦикла;
		Изм.Добавить(Новый Структура("Имя, Типы, ДлинаСтроки", М.Имя, Т, М.Тип.КвалификаторыСтроки.Длина));
	КонецЦикла;
	Результат.Добавить(Новый Структура("Имя, Измерения", Р.Имя, Изм));
КонецЦикла;
Перечисления_ = Новый Структура;
Для Каждого П Из Метаданные.Перечисления Цикл
	Зн = Новый Массив;
	Для Каждого З Из П.ЗначенияПеречисления Цикл Зн.Добавить(З.Имя); КонецЦикла;
	Перечисления_.Вставить(П.Имя, Зн);
КонецЦикла;
Результат = Новый Структура("Регистры, Перечисления", Результат, Перечисления_);
'@
        $current = Invoke-1C $describe
        $own = $script:ProcessorObject
        $script:ProcessorObject = $null
        $ownConnection = $ConnectionString
        $ConnectionString = "File=`"$TargetFileBase`";"
        $typicalDescription = Invoke-1C $describe
        $typical = @{}; $typicalDescription.Регистры | ForEach-Object { $typical[$_.Имя] = $_ }
        $ConnectionString = $ownConnection
        $script:ProcessorObject = $own
        # Values of an enumeration absent from the CF turn empty the same way.
        $goneValues = @{}
        foreach ($enum in $current.Перечисления.PSObject.Properties) {
            $typicalValues = $typicalDescription.Перечисления.PSObject.Properties[$enum.Name]
            if (-not $typicalValues) { continue }
            $gone = @($enum.Value | Where-Object { $typicalValues.Value -notcontains $_ })
            if ($gone.Count) { $goneValues["Перечисление.$($enum.Name)"] = $gone }
        }
        $summary.enum_values_absent = @($goneValues.Keys | Sort-Object | ForEach-Object { "$_ : $($goneValues[$_] -join ', ')" })
        $rows = foreach ($register in $current.Регистры) {
            $target = $typical[$register.Имя]
            if (-not $target) { continue }
            $kept = @{}; $target.Измерения | ForEach-Object { $kept[$_.Имя] = $_.Типы }
            $lengths = @{}; $target.Измерения | ForEach-Object { $lengths[$_.Имя] = $_.ДлинаСтроки }
            foreach ($dimension in $register.Измерения) {
                if (-not $kept.ContainsKey($dimension.Имя)) { [pscustomobject]@{ register = $register.Имя; dimension = $dimension.Имя; deleted = 0; status = 'DIMENSION_REMOVED' }; continue }
                $types = $kept[$dimension.Имя]
                if ($lengths[$dimension.Имя] -gt 0 -and $dimension.ДлинаСтроки -gt $lengths[$dimension.Имя] -or ($dimension.ДлинаСтроки -eq 0 -and $lengths[$dimension.Имя] -gt 0)) {
                    [pscustomobject]@{ register = $register.Имя; dimension = $dimension.Имя; found = 0; deleted = 0; status = "STRING_SHORTER $($dimension.ДлинаСтроки)->$($lengths[$dimension.Имя])" }
                }
                $values = @(foreach ($type in $dimension.Типы) {
                    if ($goneValues[$type] -and $types -contains $type) {
                        $manager = $type -replace '^Перечисление\.', ''
                        foreach ($value in $goneValues[$type]) { "Значения.Добавить(Перечисления.$manager.$value);" }
                    }
                })
                if (@($dimension.Типы | Where-Object { $types -notcontains $_ }).Count -eq 0 -and $values.Count -eq 0) { continue }
                $allowed = ConvertTo-1CString ($types -join ',')
                $result = Invoke-1C @"
ИмяРегистра = "$($register.Имя)"; ИмяИзмерения = "$($dimension.Имя)";
Оставить = Новый Соответствие;
Для Каждого Имя Из СтрРазделить($allowed, ",") Цикл Оставить.Вставить(Имя, Истина); КонецЦикла;
МетаданныеРегистра = Метаданные.РегистрыСведений[ИмяРегистра];
ПоРегистратору = МетаданныеРегистра.РежимЗаписи = Метаданные.СвойстваОбъектов.РежимЗаписиРегистра.ПодчинениеРегистратору;
Периодический = МетаданныеРегистра.ПериодичностьРегистраСведений <> Метаданные.СвойстваОбъектов.ПериодичностьРегистраСведений.Непериодический;
Запрос = Новый Запрос("ВЫБРАТЬ РАЗЛИЧНЫЕ ТИПЗНАЧЕНИЯ(Т." + ИмяИзмерения + ") КАК Тип ИЗ РегистрСведений." + ИмяРегистра + " КАК Т");
Удалено = 0; Найдено = 0; Статус = "OK";
Выборка = Запрос.Выполнить().Выбрать();
Пока Выборка.Следующий() Цикл
	Об = ?(Выборка.Тип = NULL, Неопределено, Метаданные.НайтиПоТипу(Выборка.Тип));
	Если Об = Неопределено Или Оставить[Об.ПолноеИмя()] <> Неопределено Тогда Продолжить; КонецЕсли;
	Строки = Новый Запрос("ВЫБРАТЬ * ИЗ РегистрСведений." + ИмяРегистра + " КАК Т ГДЕ ТИПЗНАЧЕНИЯ(Т." + ИмяИзмерения + ") = &Тип");
	Строки.УстановитьПараметр("Тип", Выборка.Тип);
	Записи = Строки.Выполнить().Выгрузить();
	Найдено = Найдено + Записи.Количество();
	Если ПоРегистратору Тогда Статус = "RECORDER_SUBORDINATE"; Продолжить; КонецЕсли;
	Для Каждого Запись Из Записи Цикл
		Набор = РегистрыСведений[ИмяРегистра].СоздатьНаборЗаписей();
		Набор.ОбменДанными.Загрузка = Истина;
		Для Каждого Изм Из МетаданныеРегистра.Измерения Цикл
			Набор.Отбор[Изм.Имя].Установить(Запись[Изм.Имя]);
		КонецЦикла;
		Если Периодический Тогда Набор.Отбор.Период.Установить(Запись.Период); КонецЕсли;
		Набор.Записать();
		Удалено = Удалено + 1;
	КонецЦикла;
КонецЦикла;
Значения = Новый Массив;
$($values -join ' ')
Если Значения.Количество() > 0 Тогда
	Строки = Новый Запрос("ВЫБРАТЬ * ИЗ РегистрСведений." + ИмяРегистра + " КАК Т ГДЕ Т." + ИмяИзмерения + " В (&Значения)");
	Строки.УстановитьПараметр("Значения", Значения);
	Записи = Строки.Выполнить().Выгрузить();
	Найдено = Найдено + Записи.Количество();
	Если ПоРегистратору И Записи.Количество() > 0 Тогда Статус = "RECORDER_SUBORDINATE"; Иначе
	Для Каждого Запись Из Записи Цикл
		Набор = РегистрыСведений[ИмяРегистра].СоздатьНаборЗаписей();
		Набор.ОбменДанными.Загрузка = Истина;
		Для Каждого Изм Из МетаданныеРегистра.Измерения Цикл
			Набор.Отбор[Изм.Имя].Установить(Запись[Изм.Имя]);
		КонецЦикла;
		Если Периодический Тогда Набор.Отбор.Период.Установить(Запись.Период); КонецЕсли;
		Набор.Записать();
		Удалено = Удалено + 1;
	КонецЦикла;
	КонецЕсли;
КонецЕсли;
Результат = Новый Структура("Найдено, Удалено, Статус", Найдено, Удалено, Статус);
"@
                if ($result.Найдено) { Write-Host ("{0}.{1}: found {2}, deleted {3}, {4}" -f $register.Имя, $dimension.Имя, $result.Найдено, $result.Удалено, $result.Статус) }
                [pscustomobject]@{ register = $register.Имя; dimension = $dimension.Имя; found = $result.Найдено; deleted = $result.Удалено; status = $result.Статус }
            }
        }
        $rows | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'registers.csv') -NoTypeInformation -Encoding UTF8
        $summary.dimensions_narrowed = @($rows).Count
        $summary.rows_deleted = ($rows | Measure-Object deleted -Sum).Sum
        $summary.not_deleted = @($rows | Where-Object { $_.status -ne 'OK' } | ForEach-Object { "$($_.register).$($_.dimension): $($_.status)" })
    }
    'Postings' {
        $target = Get-TargetObjects $TargetConfigurationXml
        $plan = Invoke-Processor 'Результат = О.ВидыДокументовСДвижениями();'
        $metadata = Invoke-1C 'Результат = Новый Массив; Для Каждого М Из Метаданные.Документы.ОперацияБух.Движения Цикл Результат.Добавить(М.ПолноеИмя()); КонецЦикла;'
        $names = @($plan | Where-Object { -not $target.Document.Contains($_.Документ) } | ForEach-Object Документ)
        $registers = @($metadata | Where-Object {
            $kind, $name = $_ -split '\.'
            ($kind -eq 'РегистрБухгалтерии' -and $target.AccountingRegister.Contains($name)) -or
            ($kind -eq 'РегистрНакопления' -and $target.AccumulationRegister.Contains($name)) -or
            ($kind -eq 'РегистрСведений' -and $target.InformationRegister.Contains($name))
        })
        $summary.document_kinds = $names
        $summary.registers = $registers.Count
        if ($names.Count) {
            $setup = "П = О.ПараметрыПроводок(); П.Порция = $BatchDocuments; " +
                (($names | ForEach-Object { "П.ВидыДокументов.Добавить($(ConvertTo-1CString $_));" }) -join ' ') + ' ' +
                (($registers | ForEach-Object { "П.Регистры.Добавить($(ConvertTo-1CString $_));" }) -join ' ')
            $position = 0; $done = 0; $movements = 0; $errors = @(); $notMoved = @{}
            while ($true) {
                $batch = Invoke-Processor "$setup П.Позиция = $position; Результат = О.ПеренестиПроводкиВОперации(П);"
                $position = $batch.Позиция; $done += $batch.Успешно; $movements += $batch.Движений; $errors += @($batch.Ошибки)
                foreach ($name in ($batch.НеПереносится -split ',\s*' | Where-Object { $_ })) { $notMoved[$name] = $true }
                Write-Host ("{0:s} Postings: total={1} done={2} movements={3} errors={4}" -f (Get-Date), $batch.Всего, $done, $movements, $errors.Count)
                if ($batch.Завершено -or $batch.Обработано -eq 0) { break }
            }
            @($errors) | Select-Object Объект, Ошибка | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'errors.csv') -NoTypeInformation -Encoding UTF8
            $summary.documents = $done; $summary.movements = $movements; $summary.errors = $errors.Count
            $summary.registers_not_moved = @($notMoved.Keys | Sort-Object)
        }
    }
    'Convert' {
        if (-not $InfoBase -or -not $TargetCf) { throw 'Convert needs -InfoBase and -TargetCf.' }
        if ($InfoBase -notmatch '^(?<host>[^\\:]+)(:\d+)?\\(?<name>.+)$') { throw "InfoBase must be <server>\<name>: $InfoBase" }
        $name = $Matches.name
        $designer = Join-Path $PlatformDirectory '1cv8.exe'
        $rac = Join-Path $PlatformDirectory 'rac.exe'
        $cluster = ((& $rac cluster list) | Select-String '^cluster\s*:\s*(\S+)' | Select-Object -First 1).Matches[0].Groups[1].Value
        $lines = & $rac infobase summary list --cluster=$cluster
        $id = $null
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^infobase\s*:\s*(\S+)') { $candidate = $Matches[1] }
            if ($lines[$i] -match '^name\s*:\s*"?(?<n>[^"]+)"?\s*$' -and $Matches.n -eq $name) { $id = $candidate; break }
        }
        if (-not $id) { throw "Infobase $name is not registered in cluster $cluster." }
        & $rac infobase update --cluster=$cluster --infobase=$id --infobase-user=$User --infobase-pwd=$Password --scheduled-jobs-deny=on | Out-Null
        foreach ($session in @((& $rac session list --cluster=$cluster --infobase=$id) | Select-String '^session\s*:\s*(\S+)' | ForEach-Object { $_.Matches[0].Groups[1].Value })) {
            & $rac session terminate --cluster=$cluster --session=$session | Out-Null
        }
        $steps = [ordered]@{}
        if (-not $OnlyUpdateDb) {
            $steps['delete-extensions'] = @('/DeleteCfg', '-AllExtensions')
            $steps['load-cf'] = @('/LoadCfg', "`"$TargetCf`"")
        }
        $steps['update-db'] = @('/UpdateDBCfg')
        $results = foreach ($stepName in $steps.Keys) {
            $log = Join-Path $OutputDirectory "$stepName.log"
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $process = Start-Process -FilePath $designer -Wait -PassThru -ArgumentList (@('DESIGNER', '/S', $InfoBase, '/N', $User, '/P', "`"$Password`"",
                '/DisableStartupDialogs', '/DisableStartupMessages') + $steps[$stepName] + @('/Out', "`"$log`""))
            $text = if (Test-Path -LiteralPath $log) { (Get-Content -LiteralPath $log -Raw).Trim() } else { '' }
            Write-Host ("{0}: exit {1}, {2} min" -f $stepName, $process.ExitCode, [Math]::Round($watch.Elapsed.TotalMinutes, 1))
            if ($process.ExitCode -ne 0) { throw "$stepName failed (exit $($process.ExitCode)): $text" }
            [pscustomobject]@{ step = $stepName; minutes = [Math]::Round($watch.Elapsed.TotalMinutes, 1); log = $log }
        }
        $summary.steps = @($results)
    }
}

$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'summary.json') -Encoding UTF8
$summary | ConvertTo-Json -Depth 5
