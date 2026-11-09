[CmdletBinding()]
param(
    # Run settings, e.g. configurations\zup\example.settings.json.
    [Parameter(Mandatory)] [string]$Settings,
    [Parameter(Mandatory)] [ValidateSet('Prepare','DryRun','Apply','Verify','Files','Compact')] [string]$Stage,
    # Prepare: MARK-OTHER-ORGANIZATIONS.  Apply: DELETE-ALL-MARKED.  Files: LOAD-FILES.  Compact: RELOAD-DT.
    [string]$ConfirmationToken,
    # Apply / Verify: the dry-run directory to use; the latest one by default.
    [string]$DryRunDirectory,
    # Apply without a fresh backup (the database is a disposable copy).
    [switch]$SkipBackup
)
<#
    One entry point for an extraction run, the same for every configuration.
    Everything that differs between runs is in the settings file; everything
    decided about objects is in policies\.

      Prepare  marks every organization except the kept ones for deletion
               (_Marked), as 41 requires; without the token only lists them.
      DryRun   40 catalog -> 41 KEEP/DELETE -> 42 rules (read-only) ->
               11 delete (read-only, rules simulated).  Writes run.json with
               the frozen manifest hashes and stops unless the gate is PASS.
      Apply    backup, 43 snapshot, 42 and 11 with -Apply against the frozen
               manifests of the dry run, repeat dry run of 11 (must be 0
               rows), 43 snapshot after.
      Verify   40 + 41 on the result (must be PASS with DELETE 0) and, when
               source_database is set, 45 register reconciliation against
               the copy before deletion.  1C steps (update, ТиИ 46, stale
               totals 49, totals and access 47) need a 1C user and are
               printed, not run.
      Files    attached files from file volumes into the infobase (50 with
               the processor 1c\ПодготовкаКопииПослеВыделения and MCP_Toolkit
               open in a thin client): server access, check, and with the
               token load + unbind of the emptied volumes.  Files missing from
               the storage are listed for users to attach again.
      Compact  before handover: backup, 51 dump to .dt and load back into the
               same infobase (drops binary data of deleted objects; only
               BinaryData may change), 48 shrink in SIMPLE recovery.  1C user
               from 51 defaults.

    Settings (JSON):
      server, database           SQL Server and the database being cut
      organizations              kept organizations by reference (ID, 1C
                                 navigation link or GUID), one or more
      catalog_profile            base | aggressive (policies\catalogs-*.csv)
      artifacts                  output root, relative to the repository
      source_database            optional: copy before deletion, for 45
      reconcile_period_to        optional: last date reconciled by 45
      backup_directory           optional: where Apply writes the backup
      policy_layers              optional: { "catalogs": [...], "boundary_rules": [...] },
                                 decisions of this run over the common policies,
                                 paths relative to the settings file
      infobase                   optional: Srvr\Ref for the printed 1C steps
      toolkit_url                optional: MCP_Toolkit address for Files
                                 (http://127.0.0.1:6013 by default)
      volume_path_substitution   optional: ["from => to", ...] when the 1C
                                 server reads copies of the volume files
#>
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$config = Get-Content -LiteralPath $Settings -Raw | ConvertFrom-Json
foreach ($name in 'server','database','organizations','catalog_profile','artifacts') {
    if (-not $config.$name) { throw "Settings: '$name' is required." }
}
$organizations = @($config.organizations)
$artifacts = Join-Path $root $config.artifacts
# Common decisions first, then the decisions of this run (policy_layers).
$settingsDirectory = Split-Path -Parent (Resolve-Path -LiteralPath $Settings).Path
function Resolve-Layer([string]$Path) { if ([System.IO.Path]::IsPathRooted($Path)) { $Path } else { Join-Path $settingsDirectory $Path } }
$catalogPolicies = @(Join-Path $root "policies\catalogs-$($config.catalog_profile).csv") + @($config.policy_layers.catalogs | Where-Object { $_ } | ForEach-Object { Resolve-Layer $_ })
$boundaryRules = @(Join-Path $root 'policies\boundary-rules.csv') + @($config.policy_layers.boundary_rules | Where-Object { $_ } | ForEach-Object { Resolve-Layer $_ })
New-Item -ItemType Directory -Force -Path $artifacts | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'

function Step([string]$Title) { Write-Host ''; Write-Host "== $Title" -ForegroundColor Cyan }
function Read-Summary([string]$Directory) { Get-Content -LiteralPath (Join-Path $Directory 'summary.json') -Raw | ConvertFrom-Json }
function Invoke-Sql([string]$Sql) {
    $connection = [System.Data.SqlClient.SqlConnection]::new("Server=$($config.server);Database=$($config.database);Integrated Security=True;Encrypt=False;Application Name=Codex-Extraction")
    try {
        $connection.Open()
        $command = $connection.CreateCommand(); $command.CommandTimeout = 0; $command.CommandText = $Sql
        $table = [System.Data.DataTable]::new(); $table.Load($command.ExecuteReader())
        Write-Output -NoEnumerate $table
    }
    finally { $connection.Dispose() }
}
function Get-LatestDryRun {
    if ($DryRunDirectory) { return (Resolve-Path -LiteralPath $DryRunDirectory).Path }
    # Only directories this script created; older layouts have e.g. rules-dryrun.
    $latest = Get-ChildItem -LiteralPath $artifacts -Directory | Where-Object Name -match '^\d{8}-\d{6}-dryrun$' | Sort-Object Name | Select-Object -Last 1
    if (-not $latest) { throw "No dry run in $artifacts; run -Stage DryRun first." }
    return $latest.FullName
}
function Invoke-DeleteDryRun([string]$Run, [string]$Catalog, [string]$Closure, [string]$Plan, [string]$Output) {
    $closureSummary = Read-Summary $Closure
    $arguments = @{
        Server=$config.server; Database=$config.database
        ObjectManifest=(Join-Path $Closure 'delete-manifest-with-cascade.csv'); ExpectedManifestSha256=$closureSummary.delete_manifest_with_cascade_sha256
        FixedEdgesManifest=(Join-Path $Catalog 'fixed-reference-edges.csv')
        PreservedOrganization=$organizations
        ProtectedRegisterRowsManifest=(Join-Path $Closure 'organization-register-rows.csv'); ExpectedProtectedRegisterRowsSha256=$closureSummary.organization_register_rows_sha256
        OutputDirectory=$Output
    }
    if ($Plan) { $arguments.SimulatedRuleActionsPlan = $Plan }
    & (Join-Path $PSScriptRoot '11-batched-delete.ps1') @arguments *> (Join-Path $Run "$(Split-Path -Leaf $Output).log")
    return (Read-Summary $Output)
}

switch ($Stage) {
    'Prepare' {
        Step 'Organizations'
        $catalog = Join-Path $artifacts "$stamp-prepare-catalog"
        & (Join-Path $PSScriptRoot '40-build-reference-catalog.ps1') -Server $config.server -Database $config.database -OutputDirectory $catalog *> "$catalog.log"
        $orgType = @(Import-Csv (Join-Path $catalog 'type-names.csv') | Where-Object { $_.logical_name -eq 'Организации' -and $_.type_table -match '^Reference[0-9]+$' })[0].type_table
        . (Join-Path $PSScriptRoot 'lib\Organizations.ps1')
        $filter = Get-OrganizationFilter $organizations @()
        $rows = Invoke-Sql "SELECT sys.fn_varbintohexstr(_IDRRef) AS id_hex, RTRIM(_Description) AS name, CASE WHEN $($filter.sql) THEN 'KEEP' ELSE 'DELETE' END AS fate, CONVERT(int,CONVERT(varbinary(1),_Marked)) AS marked FROM dbo._$orgType ORDER BY fate DESC, name;"
        $rows | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
        if (@($rows.Rows | Where-Object fate -eq 'KEEP').Count -ne $filter.expected) { throw "Not every kept organization was found ($($filter.label))." }
        $toMark = @($rows.Rows | Where-Object { $_.fate -eq 'DELETE' -and [int]$_.marked -eq 0 }).Count
        if ($ConfirmationToken -ne 'MARK-OTHER-ORGANIZATIONS') { Write-Host "$toMark organization(s) to mark. Repeat with -ConfirmationToken MARK-OTHER-ORGANIZATIONS."; return }
        $marked = Invoke-Sql "UPDATE dbo._$orgType SET _Marked=0x01 WHERE NOT ($($filter.sql)) AND _Marked=0x00; SELECT @@ROWCOUNT AS rows_;"
        Write-Host "Marked: $($marked.Rows[0].rows_)"
    }
    'DryRun' {
        $run = Join-Path $artifacts "$stamp-dryrun"
        New-Item -ItemType Directory -Force -Path $run | Out-Null
        Step '40 reference catalog'
        $catalog = Join-Path $run 'catalog'
        & (Join-Path $PSScriptRoot '40-build-reference-catalog.ps1') -Server $config.server -Database $config.database -OutputDirectory $catalog *> (Join-Path $run 'catalog.log')
        $catalogSummary = Read-Summary $catalog
        Write-Host "coverage_gate=$($catalogSummary.coverage_gate)"
        if ($catalogSummary.coverage_gate -ne 'PASS') { throw 'Catalog coverage gate is not PASS; see catalog\unresolved-columns.csv.' }

        Step '41 KEEP / DELETE'
        $closure = Join-Path $run 'closure'
        & (Join-Path $PSScriptRoot '41-build-keep-closure.ps1') -Server $config.server -Database $config.database -Organization $organizations `
            -CatalogDirectory $catalog -CatalogPolicyPath $catalogPolicies -BoundaryRulesPath $boundaryRules -OutputDirectory $closure *> (Join-Path $run 'closure.log')
        $closureSummary = Read-Summary $closure
        Write-Host ("gate={0} keep={1} delete={2} cascade={3} boundary={4} clear_row_candidates={5} cascade_blocked={6} unmarked_organizations={7}" -f
            $closureSummary.gate, $closureSummary.keep_objects, $closureSummary.delete_objects, $closureSummary.cascade_objects,
            $closureSummary.boundary_targets, $closureSummary.clear_row_candidates, $closureSummary.cascade_blocked_objects, $closureSummary.delete_unmarked_organizations)
        if ($closureSummary.skipped_policy_objects) { Write-Host "Policy objects not in this configuration: $(@($closureSummary.skipped_policy_objects).Count)" }
        if ($closureSummary.gate -ne 'PASS') {
            throw "41 gate is BLOCK. Decide each reference in closure\boundary-references.csv, cascade-blocked.csv, clear-row-candidates.csv (policies\), or run -Stage Prepare for unmarked organizations."
        }

        Step '42 rule actions (read-only)'
        $rules = Join-Path $run 'rules'
        & (Join-Path $PSScriptRoot '42-apply-rule-actions.ps1') -Server $config.server -Database $config.database `
            -RuleActions (Join-Path $closure 'rule-actions.csv') -ExpectedRuleActionsSha256 $closureSummary.rule_actions_sha256 `
            -ObjectManifest (Join-Path $closure 'delete-manifest-with-cascade.csv') -ExpectedManifestSha256 $closureSummary.delete_manifest_with_cascade_sha256 `
            -OutputDirectory $rules *> (Join-Path $run 'rules.log')
        $plan = Join-Path $rules 'rule-actions-plan.csv'
        Import-Csv $plan | Where-Object { [int64]$_.rows -gt 0 } | Format-Table rule,action,table,rows -AutoSize | Out-String | Write-Host

        Step '11 delete (read-only)'
        $delete = Invoke-DeleteDryRun $run $catalog $closure $plan (Join-Path $run 'delete')
        Write-Host "rows=$($delete.total_rows) boundary=$($delete.keep_boundary_gate.result) conflicts=$($delete.keep_boundary_gate.conflict_rows)"
        if ($delete.keep_boundary_gate.result -notlike 'PASS*') { throw '11 boundary gate is not PASS; see delete\.' }

        [ordered]@{
            settings=(Resolve-Path -LiteralPath $Settings).Path; generated_at_local=(Get-Date).ToString('s')
            database=$config.database; organizations=@($closureSummary.organization_ids); catalog_profile=$config.catalog_profile
            keep_objects=$closureSummary.keep_objects; delete_objects=$closureSummary.delete_objects; delete_rows=$delete.total_rows
            delete_manifest_sha256=$closureSummary.delete_manifest_with_cascade_sha256; rule_actions_sha256=$closureSummary.rule_actions_sha256
            protected_register_rows_sha256=$closureSummary.organization_register_rows_sha256
            gate='PASS'
        } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $run 'run.json') -Encoding UTF8
        Step 'Dry run PASS'
        Write-Host "$run"
        Write-Host "Next: -Stage Apply -ConfirmationToken DELETE-ALL-MARKED"
    }
    'Apply' {
        if ($ConfirmationToken -ne 'DELETE-ALL-MARKED') { throw 'Apply deletes data: pass -ConfirmationToken DELETE-ALL-MARKED.' }
        $run = Get-LatestDryRun
        $catalog = Join-Path $run 'catalog'; $closure = Join-Path $run 'closure'
        $closureSummary = Read-Summary $closure
        if ($closureSummary.gate -ne 'PASS' -or $closureSummary.database -ne $config.database) { throw "Dry run $run is not a PASS for $($config.database)." }
        $apply = Join-Path $run "apply-$stamp"
        New-Item -ItemType Directory -Force -Path $apply | Out-Null

        Step 'Backup'
        if ($SkipBackup) { Write-Host 'Skipped (-SkipBackup).' }
        else {
            if (-not $config.backup_directory) { throw "Settings: set backup_directory, or pass -SkipBackup for a disposable copy." }
            $backup = Join-Path $config.backup_directory "$($config.database)_before_delete_$stamp.bak"
            [void](Invoke-Sql "BACKUP DATABASE [$($config.database)] TO DISK=N'$backup' WITH CHECKSUM, COMPRESSION, INIT; RESTORE VERIFYONLY FROM DISK=N'$backup' WITH CHECKSUM; SELECT 1 AS ok;")
            Write-Host $backup
        }

        Step '43 snapshot before'
        & (Join-Path $PSScriptRoot '43-row-count-snapshot.ps1') -Server $config.server -Database $config.database -Label before -CatalogDirectory $catalog -OutputDirectory (Join-Path $apply 'snapshot-before') *> (Join-Path $apply 'snapshot-before.log')

        Step '42 rule actions'
        & (Join-Path $PSScriptRoot '42-apply-rule-actions.ps1') -Server $config.server -Database $config.database `
            -RuleActions (Join-Path $closure 'rule-actions.csv') -ExpectedRuleActionsSha256 $closureSummary.rule_actions_sha256 `
            -ObjectManifest (Join-Path $closure 'delete-manifest-with-cascade.csv') -ExpectedManifestSha256 $closureSummary.delete_manifest_with_cascade_sha256 `
            -OutputDirectory (Join-Path $apply 'rules') -Apply -ConfirmationToken CLEAR-RULE-ACTIONS *> (Join-Path $apply 'rules.log')
        $rulesSummary = Read-Summary (Join-Path $apply 'rules'); Write-Host "rule rows changed=$($rulesSummary.changed_rows)"

        Step '11 delete'
        & (Join-Path $PSScriptRoot '11-batched-delete.ps1') -Server $config.server -Database $config.database `
            -ObjectManifest (Join-Path $closure 'delete-manifest-with-cascade.csv') -ExpectedManifestSha256 $closureSummary.delete_manifest_with_cascade_sha256 `
            -FixedEdgesManifest (Join-Path $catalog 'fixed-reference-edges.csv') -PreservedOrganization $organizations `
            -ProtectedRegisterRowsManifest (Join-Path $closure 'organization-register-rows.csv') -ExpectedProtectedRegisterRowsSha256 $closureSummary.organization_register_rows_sha256 `
            -OutputDirectory (Join-Path $apply 'delete') -Apply -ConfirmationToken DELETE-ALL-MARKED *> (Join-Path $apply 'delete.log')
        $deleted = Read-Summary (Join-Path $apply 'delete')
        Write-Host "deleted rows=$($deleted.total_rows) boundary=$($deleted.keep_boundary_gate.result)"

        Step '11 repeat dry run'
        $repeat = Invoke-DeleteDryRun $apply $catalog $closure $null (Join-Path $apply 'repeat-dryrun')
        Write-Host "rows left=$($repeat.total_rows)"

        Step '43 snapshot after'
        & (Join-Path $PSScriptRoot '43-row-count-snapshot.ps1') -Server $config.server -Database $config.database -Label after -CatalogDirectory $catalog `
            -CompareWith (Join-Path $apply 'snapshot-before\row-counts.csv') -OutputDirectory (Join-Path $apply 'snapshot-after') *> (Join-Path $apply 'snapshot-after.log')
        $after = Read-Summary (Join-Path $apply 'snapshot-after')
        Write-Host "row_delta=$($after.row_delta) grown_tables=$($after.grown_tables)"
        if ([int64]$repeat.total_rows -ne 0 -or [int64]$after.grown_tables -ne 0) { throw 'Apply finished with rows left or tables that grew; see the snapshot diff.' }
        Write-Host "Next: -Stage Verify"
    }
    'Verify' {
        $run = Get-LatestDryRun
        $verify = Join-Path $run "verify-$stamp"
        New-Item -ItemType Directory -Force -Path $verify | Out-Null
        Step '40 + 41 on the result'
        $catalog = Join-Path $verify 'catalog'
        & (Join-Path $PSScriptRoot '40-build-reference-catalog.ps1') -Server $config.server -Database $config.database -OutputDirectory $catalog *> (Join-Path $verify 'catalog.log')
        & (Join-Path $PSScriptRoot '41-build-keep-closure.ps1') -Server $config.server -Database $config.database -Organization $organizations `
            -CatalogDirectory $catalog -CatalogPolicyPath $catalogPolicies -BoundaryRulesPath $boundaryRules -OutputDirectory (Join-Path $verify 'closure') *> (Join-Path $verify 'closure.log')
        $closureSummary = Read-Summary (Join-Path $verify 'closure')
        Write-Host "gate=$($closureSummary.gate) keep=$($closureSummary.keep_objects) delete=$($closureSummary.delete_objects) missing_references=$($closureSummary.missing_reference_targets)"

        if ($config.source_database) {
            Step '45 register reconciliation'
            $periodTo = if ($config.reconcile_period_to) { [datetime]$config.reconcile_period_to } else { [datetime]'2999-12-31' }
            & (Join-Path $PSScriptRoot '45-reconcile-registers.ps1') -Server $config.server -SourceDatabase $config.source_database -TargetDatabase $config.database `
                -Organization $organizations -CatalogDirectory $catalog -PeriodTo $periodTo -OutputDirectory (Join-Path $verify 'reconcile') *> (Join-Path $verify 'reconcile.log')
            $reconcile = Read-Summary (Join-Path $verify 'reconcile')
            Write-Host "result=$($reconcile.result) missing=$($reconcile.missing_in_target) extra=$($reconcile.extra_in_target) resource_mismatches=$($reconcile.resource_mismatches)"
        }

        Step '1C steps (need a 1C user)'
        $infobase = if ($config.infobase) { $config.infobase } else { "$($config.server)\$($config.database)" }
        Write-Host "1C client: /C ЗапуститьОбновлениеИнформационнойБазы   (update handlers, also when the version did not change)"
        Write-Host "scripts\46-run-tii.ps1 -InfoBase '$infobase' -User <user> -Profile CheckAndTotals -PrepareCluster -DeleteManifest '$(Join-Path $run 'closure\delete-manifest-with-cascade.csv')'"
        Write-Host "scripts\49-check-stale-totals.ps1 -Server $($config.server) -Database $($config.database) -CatalogDirectory '$catalog'"
        Write-Host "scripts\47-post-delete-1c-maintenance.ps1 -RecalcTotals -TotalsRegister <stale_registers of 49> -SequenceReport -StartAccessUpdate   (MCP_Toolkit open in 1C)"
        Write-Host "Invoke-Extraction -Stage Files   (attached files from volumes into the infobase)"
        Write-Host "Invoke-Extraction -Stage Compact (before handover: .dt reload drops binary data of deleted objects, then shrink)"
    }
    'Compact' {
        if ($ConfirmationToken -ne 'RELOAD-DT') { throw 'Compact reloads the infobase from .dt and shrinks its files: pass -ConfirmationToken RELOAD-DT.' }
        $run = Get-LatestDryRun
        $compact = Join-Path $run "compact-$stamp"
        New-Item -ItemType Directory -Force -Path $compact | Out-Null
        $infobase = if ($config.infobase) { $config.infobase } else { "$($config.server)\$($config.database)" }
        if (-not $config.backup_directory) { throw 'Settings: set backup_directory (the .dt and the backup before reload go there).' }

        Step 'Backup'
        if ($SkipBackup) { Write-Host 'Skipped (-SkipBackup).' }
        else {
            $backup = Join-Path $config.backup_directory "$($config.database)_before_dt_$stamp.bak"
            [void](Invoke-Sql "BACKUP DATABASE [$($config.database)] TO DISK=N'$backup' WITH COPY_ONLY, CHECKSUM, COMPRESSION, INIT; SELECT 1 AS ok;")
            Write-Host $backup
        }

        Step '51 dump to .dt and load back'
        $dt = Join-Path $config.backup_directory "$($config.database)_$stamp.dt"
        & (Join-Path $PSScriptRoot '51-reload-dt.ps1') -InfoBase $infobase -Database $config.database -SqlServer $config.server `
            -CatalogDirectory (Join-Path $run 'catalog') -DtPath $dt -ConfirmationToken RELOAD-DT -OutputDirectory (Join-Path $compact 'reload') *> (Join-Path $compact 'reload.log')
        $reload = Read-Summary (Join-Path $compact 'reload')
        Write-Host "result=$($reload.result) BinaryData rows $($reload.binary_data_rows_before) -> $($reload.binary_data_rows_after), $($reload.total_minutes) min"
        if ($reload.result -ne 'PASS') { throw "Reload changed row counts outside BinaryData; see $compact\reload." }

        Step '48 shrink'
        & (Join-Path $PSScriptRoot '48-sql-maintenance.ps1') -Server $config.server -Database $config.database -Shrink -SetSimpleRecovery -ConfirmationToken SHRINK-FILES `
            -OutputDirectory (Join-Path $compact 'shrink') *> (Join-Path $compact 'shrink.log')
        $shrink = Read-Summary (Join-Path $compact 'shrink')
        Write-Host ("files after: " + (($shrink.files_after | ForEach-Object { "$($_.name) $($_.size_mb) MB" }) -join ', '))
        Write-Host "The .dt and the backup stay in $($config.backup_directory); delete them when the base is accepted."
    }
    'Files' {
        $run = Get-LatestDryRun
        $files = Join-Path $run "files-$stamp"
        $script = Join-Path $PSScriptRoot '50-volume-files.ps1'
        $toolkit = if ($config.toolkit_url) { $config.toolkit_url } else { 'http://127.0.0.1:6013' }
        $substitution = @($config.volume_path_substitution | Where-Object { $_ })
        if (-not (Test-Path -LiteralPath (Join-Path $root 'build\1c\ПодготовкаКопииПослеВыделения.epf'))) {
            & (Join-Path $PSScriptRoot 'Build-1CTools.ps1') | Out-Null
        }

        Step '50 summary and server access'
        $summary = & $script -Action Summary -ToolkitUrl $toolkit -OutputDirectory (Join-Path $files 'summary') | ConvertFrom-Json
        Write-Host "infobase=$($summary.infobase) files_on_volumes=$($summary.files) storage_mode=$($summary.storage_mode)"
        if ([int64]$summary.files -eq 0) { Write-Host 'No files on volumes.'; return }
        $access = & $script -Action Access -ToolkitUrl $toolkit -PathSubstitution $substitution -OutputDirectory (Join-Path $files 'access') | ConvertFrom-Json
        if (-not @($access.volumes | Where-Object { $_.path })) {
            Write-Host "Volumes are already unbound; $($summary.files) file(s) left on them were missing from the storage (see the errors of the last load)."
            return
        }
        if ($access.result -ne 'OK') {
            throw "The 1C server ($($access.server_temp_directory)) does not see every volume. Save a credential for it in the processor form or set volume_path_substitution; see $files\access\summary.json."
        }

        Step '50 check (no file content is read)'
        $check = & $script -Action Check -ToolkitUrl $toolkit -PathSubstitution $substitution -OutputDirectory (Join-Path $files 'check') | ConvertFrom-Json
        Write-Host "ok=$($check.done) missing_or_different=$($check.errors) -> $($check.errors_csv)"
        if ($ConfirmationToken -ne 'LOAD-FILES') { Write-Host 'Repeat with -ConfirmationToken LOAD-FILES to load the files and unbind the emptied volumes.'; return }

        Step '50 load'
        $load = & $script -Action Load -ToolkitUrl $toolkit -PathSubstitution $substitution -ConfirmationToken LOAD-FILES -OutputDirectory (Join-Path $files 'load') | ConvertFrom-Json
        Write-Host "loaded=$($load.done) errors=$($load.errors) remaining_on_volumes=$($load.remaining_on_volumes) -> $($load.errors_csv)"

        Step '50 unbind volumes'
        $unbind = & $script -Action Unbind -ToolkitUrl $toolkit -ConfirmationToken UNBIND-VOLUMES -OutputDirectory (Join-Path $files 'unbind') | ConvertFrom-Json
        Write-Host "storage_mode=$($unbind.storage_mode) unbound=$($unbind.unbound_volumes) kept_with_files=$($unbind.volumes_with_files -join ', ')"
        if ($unbind.volumes_with_files) {
            Write-Host "Files left on volumes are listed in $($load.errors_csv): users attach them again. Then: scripts\50-volume-files.ps1 -Action Unbind -IncludeVolumesWithFiles -ConfirmationToken UNBIND-VOLUMES"
        }
    }
}
