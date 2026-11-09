# Скрипты

## Конвейер

| Скрипт | Что делает | Меняет базу |
|---|---|---|
| `Invoke-Extraction.ps1` | точка входа: стадии `Prepare`, `DryRun`, `Apply`, `Verify`, `Files`, `Compact` по файлу настроек | `Prepare` и `Apply` — только с токеном |
| `40-build-reference-catalog.ps1` | каталог ссылок из схемы хранения: классы колонок, имена типов и реквизитов (SQLCLR `sqlclr/`) | нет |
| `41-build-keep-closure.ps1` | KEEP-замыкание сохраняемых организаций, DELETE как дополнение, каскад, граница, план правил | нет |
| `42-apply-rule-actions.ps1` | `CLEAR_FIELD` / `CLEAR_ROW` из плана `41` | `-Apply`, токен `CLEAR-RULE-ACTIONS` |
| `11-batched-delete.ps1` | пакетное удаление по замороженному манифесту, проверка границы | `-Apply`, токен `DELETE-ALL-MARKED` |
| `43-row-count-snapshot.ps1` | снимок числа строк, сравнение снимков | нет |
| `44-classify-tii-log.ps1` | разбор журнала ТиИ: `deleted-by-run` / `pre-existing` | нет |
| `45-reconcile-registers.ps1` | сверка движений регистров (бухгалтерии, накопления, сведений, расчёта) с копией до удаления | нет |
| `46-run-tii.ps1` | ТиИ конфигуратором, подготовка кластера | `Repair` — токен `TII-REPAIR` |
| `47-post-delete-1c-maintenance.ps1` | через `MCP_Toolkit`: пересчёт итогов (все регистры или `-TotalsRegister`), границы последовательностей, обновление доступа | да (через 1С) |
| `48-sql-maintenance.ps1` | отчёт о файлах и индексах; сжатие, перестроение | `-Shrink` — токен `SHRINK-FILES` |
| `49-check-stale-totals.ps1` | строки итогов, остатков и срезов со ссылкой на удалённую организацию; список регистров для `47 -TotalsRegister` | нет |
| `50-volume-files.ps1` | файлы на томах через обработку `1c/`: сводка, доступ сервера, проверка, загрузка в базу, отвязка томов | `Load` — `LOAD-FILES`, `Unbind` — `UNBIND-VOLUMES` |
| `51-reload-dt.ps1` | выгрузка в `.dt` и загрузка обратно (или в другую базу), сравнение числа строк: меняется только `BinaryData` | токен `RELOAD-DT` |
| `52-convert-to-typical.ps1` | превращение копии в типовую конфигурацию: [docs/CONVERT-TO-TYPICAL.md](../docs/CONVERT-TO-TYPICAL.md) |
| `Build-1CTools.ps1` | сборка обработок `1c/` из XML во временной файловой базе | нет |
| `lib/Organizations.ps1` | разбор ссылок на организации (ID, навигационная ссылка, GUID) | — |

Организации во всех скриптах задаются ссылками, можно несколько:
`-Organization` (`41`, `45`), `-PreservedOrganization` (`11`). Где у
организации есть код, можно `-OrganizationCode` / `-PreservedOrganizationCode`.
