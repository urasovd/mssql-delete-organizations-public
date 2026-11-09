<#
    Kept organizations by reference, shared by scripts 11, 41 and 45.

    A reference is the organization's _IDRRef in any of the forms a user can
    copy out of 1C or SQL:
      0xaabbccddeeff00112233445566778899            SQL binary(16)
      aabbccddeeff00112233445566778899              the same without 0x
      e1cib/data/Справочник.Организации?ref=aabb...  1C navigation link
      66778899-4455-2233-aabb-ccddeeff0011          УникальныйИдентификатор()
    The navigation link carries _IDRRef as is; the GUID of
    УникальныйИдентификатор() stores the same 16 bytes in another order.
#>

function ConvertTo-IDRRefHex([string]$Reference) {
    $value = $Reference.Trim()
    if ($value -match 'ref=(?<hex>[0-9A-Fa-f]{32})') { return '0x' + $Matches.hex.ToLowerInvariant() }
    if ($value -match '^(0x)?(?<hex>[0-9A-Fa-f]{32})$') { return '0x' + $Matches.hex.ToLowerInvariant() }
    if ($value -match '^\{?(?<a>[0-9A-Fa-f]{8})-(?<b>[0-9A-Fa-f]{4})-(?<c>[0-9A-Fa-f]{4})-(?<d>[0-9A-Fa-f]{4})-(?<e>[0-9A-Fa-f]{12})\}?$') {
        return '0x' + ($Matches.d + $Matches.e + $Matches.c + $Matches.b + $Matches.a).ToLowerInvariant()
    }
    throw "Not an organization reference: '$Reference'. Use 0x + 32 hex digits, a 1C navigation link or a GUID."
}

function Get-OrganizationFilter([string[]]$Reference, [string[]]$Code, [string]$Alias = '') {
    # SQL predicate on the organization catalog; exactly one of the two lists.
    # Codes are kept for configurations whose organizations have one (BIT).
    $references = @($Reference | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $codes = @($Code | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (($references.Count -gt 0) -eq ($codes.Count -gt 0)) { throw 'Specify the kept organizations either by reference (-Organization) or by code (-OrganizationCode).' }
    $prefix = if ($Alias) { "$Alias." } else { '' }
    if ($references.Count) {
        $ids = @($references | ForEach-Object { ConvertTo-IDRRefHex $_ } | Sort-Object -Unique)
        return [pscustomobject]@{ sql="$($prefix)_IDRRef IN ($($ids -join ','))"; expected=$ids.Count; label=($ids -join ',') }
    }
    $quoted = @($codes | ForEach-Object { "N'" + $_.Trim().Replace("'","''") + "'" } | Sort-Object -Unique)
    return [pscustomobject]@{ sql="RTRIM($($prefix)_Code) IN ($($quoted -join ','))"; expected=$quoted.Count; label=($codes -join ',') }
}
