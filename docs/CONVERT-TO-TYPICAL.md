# Превращение выделенной копии БИТ в типовую БП КОРП

Шаг после выделения: копия доработанной конфигурации становится типовой, остатки
сохраняются. Запускается на **отдельной копии**: результат выделения (база БИТ)
остаётся как базовый продукт и как источник сверки. Скрипт —
[`scripts/52-convert-to-typical.ps1`](../scripts/52-convert-to-typical.ps1), первый
прогон — [разбор прогонов](CASE-STUDY.md#превращение-в-бп-корп).

Распределённая база не создаётся: в копии одна организация.

## Подготовка

1. Копия: бэкап результата `Compact` → `RESTORE` в новую базу → `rac infobase create`
   с `--date-offset=2000` и `--scheduled-jobs-deny=on`.
2. Файловая база из cf и список её объектов:

   ```powershell
   1cv8 CREATEINFOBASE File="<cfbase>" /UseTemplate "<typical.cf>"
   1cv8 DESIGNER /F "<cfbase>" /DumpConfigToFiles "<cfdump>" -listFile <файл со строкой Configuration>
   ```

Все шаги, кроме `Convert`, идут через COM-соединение. Клиент 1С на копии не нужен: он
останавливается на стартовых окнах, а при заблокированном рабочем столе их не закрыть.
`Build-1CTools.ps1` после изменения обработки запускать вручную.

## Шаги

| Шаг `-Step` | Что делает | Пример прогона |
|---|---|---|
| `Osv` | эталон ОСВ по счетам | 103 счёта, 20 с |
| `Plan` | документы и регистры, которых нет в cf; только чтение | 237 видов документов, данные у 17 |
| `Prepare` | снять дату запрета, ограничение доступа, версионирование; очистить `НеудаленныеОбъекты` | секунды |
| `Postings` | движения документов, которых нет в cf, в «Операция (бух)» | 34 документа, 152 проводки, 1 мин |
| `Osv -CompareWith` | контроль | 0 расхождений |
| бэкап копии | откат к этой точке | 9,4 ГБ |
| `Registers` | удалить записи регистров сведений со значениями, которых нет в cf | 32 516 записей, 2 мин |
| `Convert` | удалить расширения, загрузить cf, реструктуризация | 2 + 14 + 35 мин |
| `Osv -CompareWith` | контроль | 0 расхождений |
| `Update` | обработчики обновления на сервере | 2 мин |
| `Osv -CompareWith`, `Reconcile` | контроль: счета, субконто, регистры | см. отчёт |

Без неудачных попыток — около часа. Токен изменяющих шагов — `CONVERT-TO-TYPICAL`.

```powershell
$cs = 'Srvr="server";Ref="<копия>";Usr="admin";Pwd=""'
.\scripts\52-convert-to-typical.ps1 -Step Osv -ConnectionString $cs -OutputDirectory <osv-0>
.\scripts\52-convert-to-typical.ps1 -Step Plan -ConnectionString $cs -TargetConfigurationXml <cfdump>\Configuration.xml
.\scripts\52-convert-to-typical.ps1 -Step Prepare -ConnectionString $cs -ConfirmationToken CONVERT-TO-TYPICAL
.\scripts\52-convert-to-typical.ps1 -Step Postings -ConnectionString $cs -TargetConfigurationXml <cfdump>\Configuration.xml -ConfirmationToken CONVERT-TO-TYPICAL
.\scripts\52-convert-to-typical.ps1 -Step Registers -ConnectionString $cs -TargetFileBase <cfbase> -ConfirmationToken CONVERT-TO-TYPICAL
.\scripts\52-convert-to-typical.ps1 -Step Convert -InfoBase 'server\<копия>' -TargetCf <typical.cf> -ConfirmationToken CONVERT-TO-TYPICAL
.\scripts\52-convert-to-typical.ps1 -Step Update -ConnectionString $cs -ConfirmationToken CONVERT-TO-TYPICAL
.\scripts\52-convert-to-typical.ps1 -Step Osv -ConnectionString $cs -CompareWith <osv-0>\osv.csv
.\scripts\52-convert-to-typical.ps1 -Step Reconcile -ConnectionString $cs -SourceConnectionString 'Srvr="server";Ref="<база БИТ>";Usr="admin";Pwd=""'
```

После неудачной реструктуризации cf загружать заново не нужно: `Convert -OnlyUpdateDb`.
Сверка `45` после превращения не работает: у таблиц другие колонки, таблиц расширений нет.

## Почему падает реструктуризация

«Записи регистра сведений стали неуникальными» и откат всей операции. Значение
измерения становится пустым, и несколько записей совпадают. Причины, которые встретились:

- у измерения в cf нет типа значения (объект БИТ): `БезопасноеХранилищеДанных`;
- в типовом перечислении нет значения БИТ (`бит_ар_ПеревыставлениеУслуг`): `ГруппыЗначенийДоступа`;
- служебный `НеудаленныеОбъекты`.

Первые две находит и удаляет `Registers`. Третий чистит `Prepare`. Укорочение строковых
измерений `Registers` только показывает (`STRING_SHORTER`); в прогоне их не было.

## Что теряется

Движения в регистрах БИТ и данные объектов БИТ: это задача превращения. Шаг не сохраняет:

- документы БИТ без движений в типовых регистрах и все их данные;
- присоединённые файлы объектов БИТ;
- движения типового документа в регистре, которого нет в его движениях в cf. Платформа
  удаляет их при реструктуризации, и ни `Plan`, ни `Postings` их пока не видят;
- данные объекта с другим `uuid` в cf при том же имени: платформа считает его новым и
  создаёт пустую таблицу. `Plan` этого пока не проверяет, сравнение делалось по выгрузкам.

## Известные ограничения

- `Convert` не снимает конфигурацию с поддержки вручную: `/LoadCfg` заменяет её целиком.
- Дата запрета, ограничение доступа и версионирование остаются выключенными: включить
  перед выдачей.
- Не входит в стадии `Invoke-Extraction`.
